# app.py
import os
import time
import uuid
import base64
import re
import unicodedata
import difflib
import json
import random
from flask import Flask, render_template_string, request, redirect, url_for, send_file, make_response
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementClickInterceptedException
from bs4 import BeautifulSoup
import pandas as pd
import io

# ---------- CONFIGURATION ----------
URL_BASE = "https://www.faciltecnologia.com.br/consigfacil/maranhao/geracao_senha.php?tipo=1"
EXCEL_PATH = "clientes.xlsx"
CAPTURE_DIR = "captures"
os.makedirs(CAPTURE_DIR, exist_ok=True)

# Performance / behavior tuning
CAPTURE_ENABLED = os.environ.get("CAPTURE_ENABLED", "false").lower() == "true"
NAV_TIMEOUT = int(os.environ.get("NAV_TIMEOUT", 45))
FIND_TIMEOUT = int(os.environ.get("FIND_TIMEOUT", 20))
POLL_INTERVAL = float(os.environ.get("POLL_INTERVAL", 0.4))
POLL_AFTER_CLICK = float(os.environ.get("POLL_AFTER_CLICK", 8.0))
MAX_QUESTION_STEPS = int(os.environ.get("MAX_QUESTION_STEPS", 8))

# If chromedriver is not in PATH, set this to its full path or set CHROMEDRIVER_PATH env var
CHROMEDRIVER_PATH = os.environ.get("CHROMEDRIVER_PATH") or "/usr/local/bin/chromedriver"

# Google Sheets env names
GSHEET_SA_JSON_ENV = "GSHEET_SA_JSON"
GSHEET_SPREADSHEET_ID_ENV = "GSHEET_SPREADSHEET_ID"
GSHEET_SHEET_NAME_ENV = "GSHEET_SHEET_NAME"

# Proxy list (optional) - format: "host:port,host2:port" or "user:pass@host:port,..."
PROXY_LIST = []
_raw_proxies = os.environ.get("PROXY_LIST", "")
if _raw_proxies:
    PROXY_LIST = [p.strip() for p in _raw_proxies.split(",") if p.strip()]

app = Flask(__name__)

# In-memory headless browser sessions and temporary download tokens
SESSIONS = {}  # session_id -> {"driver": webdriver, "meta": {...}, "created": timestamp, "password": None}

# ---------- HTML template ----------
TEMPLATE = """
<!doctype html>
<html>
<head><meta charset="utf-8"><title>Gerar senha</title>
<style>
body{font-family:Arial,Helvetica,sans-serif;margin:20px}
input[type=text]{width:320px;padding:6px}
button{padding:8px 12px;margin-top:8px}
.box{border:1px solid #ddd;padding:12px;border-radius:6px;margin-top:12px}
.small{font-size:0.9em;color:#555}
pre{background:#f7f7f7;padding:10px;border-radius:6px;white-space:pre-wrap}
</style>
</head>
<body>
  <h2>Gerar senha</h2>

  <div class="box">
    <form method="post" action="/start_headless">
      <label><b>CPF</b></label><br>
      <input name="cpf" required value="{{ cpf or '' }}"><br><br>
      <label><b>Matrícula</b></label><br>
      <input name="matricula" required value="{{ mat or '' }}"><br><br>
      <button type="submit">Gerar CAPTCHA</button>
    </form>
    <p class="small">O CAPTCHA aparecerá abaixo. Digite o texto e envie para continuar.</p>
  </div>

  {% if captcha_data %}
  <div class="box">
    <h4>CAPTCHA</h4>
    <img src="{{ captcha_data }}" alt="captcha"><br><br>
    <form method="post" action="/submit_captcha">
      <input type="hidden" name="session_id" value="{{ session_id }}">
      <input type="hidden" name="cpf" value="{{ cpf }}">
      <input type="hidden" name="matricula" value="{{ mat }}">
      <label>Texto do CAPTCHA</label><br>
      <input name="captcha_text" required><br><br>
      <button type="submit">Enviar CAPTCHA e aguardar senha</button>
    </form>
  </div>
  {% endif %}

  {% if result %}
  <div class="box">
    <h4>Resultado</h4>
    <pre>{{ result }}</pre>
    {% if download_token %}
      <p><a href="{{ url_for('download_password', token=download_token) }}">Baixar senha (arquivo .txt)</a></p>
    {% endif %}
  </div>
  {% endif %}

  <hr>
  <p class="small">Capturas e logs são salvos em <code>captures/</code> apenas quando CAPTURE_ENABLED é True.</p>
</body>
</html>
"""

# ---------- utilities ----------
def normalize_colname(s):
    if not isinstance(s, str): return s
    s = s.strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s = s.replace(' ', '_').upper()
    s = re.sub(r'[^A-Z0-9_]', '', s)
    return s

def load_clients_local(path=EXCEL_PATH):
    if not os.path.exists(path):
        raise FileNotFoundError(f"{path} não encontrado")
    df = pd.read_excel(path, dtype=str)
    df.columns = [normalize_colname(c) for c in df.columns]
    if 'CPF' in df.columns:
        df['CPF'] = df['CPF'].astype(str).str.replace(r'\D','',regex=True)
    if 'MATRICULA' in df.columns:
        df['MATRICULA'] = df['MATRICULA'].astype(str).str.replace(r'\D','',regex=True)
    return df

# ---------- Google Sheets integration ----------
def load_gspread_client_from_env():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except Exception as e:
        raise RuntimeError("Dependências gspread/google-auth não instaladas") from e

    b64 = os.environ.get(GSHEET_SA_JSON_ENV)
    if not b64:
        raise RuntimeError("GSHEET_SA_JSON não definido")
    try:
        raw = base64.b64decode(b64)
        info = json.loads(raw)
    except Exception as e:
        raise RuntimeError(f"Erro ao decodificar GSHEET_SA_JSON: {e}") from e
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(creds)
    return client

def load_clients_from_gsheets():
    spreadsheet_id = os.environ.get(GSHEET_SPREADSHEET_ID_ENV)
    if not spreadsheet_id:
        raise RuntimeError("GSHEET_SPREADSHEET_ID não definido")
    sheet_name = os.environ.get(GSHEET_SHEET_NAME_ENV) or None
    client = load_gspread_client_from_env()
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.sheet1 if not sheet_name else sh.worksheet(sheet_name)
    rows = ws.get_all_records()
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df.columns = [normalize_colname(c) for c in df.columns]
    if 'CPF' in df.columns:
        df['CPF'] = df['CPF'].astype(str).str.replace(r'\D','',regex=True)
    if 'MATRICULA' in df.columns:
        df['MATRICULA'] = df['MATRICULA'].astype(str).str.replace(r'\D','',regex=True)
    return df

def load_clients_auto():
    """Tenta carregar do Google Sheets; se falhar, usa o Excel local."""
    try:
        df = load_clients_from_gsheets()
        print("Dados carregados do Google Sheets")
        return df
    except Exception as e:
        print("Aviso: não foi possível carregar do Google Sheets:", e)
        try:
            df = load_clients_local()
            print("Dados carregados do Excel local")
            return df
        except Exception as e2:
            raise RuntimeError(f"Falha ao carregar dados: {e} ; {e2}")

def save_capture(session_id, tag, html):
    if not CAPTURE_ENABLED:
        return None
    try:
        fname = os.path.join(CAPTURE_DIR, f"{session_id}_{tag}_{int(time.time())}.html")
        with open(fname, "w", encoding="utf-8") as f:
            f.write(html or "")
        print(f"Capture saved: {fname}")
        return fname
    except Exception:
        return None

def extract_senha_from_html(html):
    try:
        soup = BeautifulSoup(html, "html.parser")
        for p in soup.find_all(["p", "div", "span"]):
            txt = p.get_text(" ", strip=True)
            if "nova senha" in txt.lower():
                b = p.find("b")
                if b and b.get_text(strip=True):
                    return b.get_text(strip=True)
                m = re.search(r'Nova senha[:\s]*([A-Z0-9\-]{4,})', txt, re.IGNORECASE)
                if m:
                    return m.group(1).strip()
                m2 = re.search(r'([A-Z0-9]{6,})', txt, re.IGNORECASE)
                if m2:
                    return m2.group(1).strip()
        body_text = soup.get_text(" ", strip=True)
        m = re.search(r'Nova senha[:\s]*([A-Z0-9\-]{4,})', body_text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        m2 = re.search(r'Nova senha[:\s]*<b[^>]*>([A-Z0-9\-]{4,})</b>', html, re.IGNORECASE)
        if m2:
            return m2.group(1).strip()
    except Exception:
        pass
    return None

def extract_question_and_options_from_html(html):
    soup = BeautifulSoup(html, "html.parser")
    form = soup.find("form", action=lambda a: a and "geracao_senha_validar.php" in a)
    if not form:
        radios = soup.find_all("input", {"name":"resposta","type":"radio"})
        if not radios:
            return "", []
        form = soup
    p = form.find("p")
    pergunta = p.get_text(" ", strip=True) if p else ""
    options = []
    for r in form.find_all("input", {"name":"resposta","type":"radio"}):
        val = r.get("value","")
        text = ""
        label = r.find_parent("label")
        if label:
            text = label.get_text(" ", strip=True)
        else:
            sib = r.find_next_sibling()
            if sib:
                text = sib.get_text(" ", strip=True)
        options.append((val, text))
    return pergunta, options

# ---------- selenium helpers ----------
def start_selenium_headless(proxy=None):
    opts = webdriver.ChromeOptions()
    if os.environ.get("HEADLESS","true").lower() != "false":
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--window-size=1200,900")
    opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36")

    if proxy:
        opts.add_argument(f'--proxy-server={proxy}')

    # Try to start Chrome with provided chromedriver path or default
    try:
        if CHROMEDRIVER_PATH and os.path.exists(CHROMEDRIVER_PATH):
            from selenium.webdriver.chrome.service import Service as ChromeService
            service = ChromeService(executable_path=CHROMEDRIVER_PATH)
            driver = webdriver.Chrome(service=service, options=opts)
        else:
            driver = webdriver.Chrome(options=opts)
    except Exception as e:
        # fallback: try without service
        try:
            driver = webdriver.Chrome(options=opts)
        except Exception as e2:
            raise RuntimeError(f"Falha ao iniciar Chrome/Chromedriver: {e} / {e2}") from e2

    driver.set_window_size(1200, 900)
    return driver

def navigate_via_proxysite(driver, target_url, timeout=NAV_TIMEOUT):
    """
    Força uso do proxysite https://www.proxysite.com/pt/
    - preenche input[name="d"] com a URL alvo
    - clica no botão 'Ir' (type=submit)
    - aguarda até encontrar o campo CPF no conteúdo proxied
    """
    PROXY_URL = "https://www.proxysite.com/pt/"

    try:
        # abrir proxysite explicitamente
        driver.get(PROXY_URL)
        # dar um tempo inicial para o JS carregar
        time.sleep(1.2)

        # tentar localizar o input[name="d"] de forma robusta
        input_el = None
        try:
            input_el = driver.find_element(By.NAME, "d")
        except Exception:
            # tentar por seletor CSS como fallback
            try:
                input_el = driver.find_element(By.CSS_SELECTOR, "input[name='d']")
            except Exception:
                input_el = None

        if not input_el:
            # salvar captura para debug e falhar com mensagem clara
            save_capture("proxysite", "no_input_name_d", driver.page_source)
            raise RuntimeError("Campo name='d' não encontrado no proxysite.")

        # limpar e colar a URL alvo
        try:
            driver.execute_script("arguments[0].value = '';", input_el)
            input_el.clear()
        except Exception:
            pass
        try:
            input_el.send_keys(target_url)
        except Exception:
            # fallback: setar via JS
            try:
                driver.execute_script("arguments[0].value = arguments[1];", input_el, target_url)
            except Exception:
                pass

        # localizar e clicar no botão 'Ir' (type=submit com texto 'Ir')
        btn = None
        try:
            # primeiro por texto do botão
            btn = driver.find_element(By.XPATH, "//button[@type='submit' and normalize-space(text())='Ir']")
        except Exception:
            try:
                # fallback: qualquer botão submit no form do input
                form = input_el.find_element(By.XPATH, "ancestor::form")
                if form:
                    try:
                        btn = form.find_element(By.XPATH, ".//button[@type='submit']")
                    except Exception:
                        btn = None
            except Exception:
                btn = None

        if not btn:
            # salvar captura para debug e falhar
            save_capture("proxysite", "no_submit_button", driver.page_source)
            raise RuntimeError("Botão 'Ir' (type='submit') não encontrado no proxysite.")

        # clicar no botão de forma robusta
        try:
            btn.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", btn)
            except Exception:
                pass

        # aguardar o conteúdo proxied carregar e expor o campo CPF
        end = time.time() + timeout
        last_html = None
        while time.time() < end:
            time.sleep(0.6)
            try:
                # se o conteúdo proxied já expôs o campo CPF, retornamos
                if driver.find_elements(By.ID, "cpf") or driver.find_elements(By.NAME, "cpf"):
                    return driver
            except Exception:
                pass
            try:
                last_html = driver.page_source
            except Exception:
                last_html = None

        # timeout: salvar captura para análise
        save_capture("proxysite", "timeout_after_submit", last_html or "")
        raise RuntimeError("Timeout aguardando conteúdo proxied no proxysite. Verifique captures.")
    except Exception:
        # garantir que qualquer captura seja salva antes de propagar o erro
        try:
            save_capture("proxysite", "exception", getattr(driver, "page_source", "") or "")
        except Exception:
            pass
        raise

def find_element_across_frames(driver, by, value, timeout=FIND_TIMEOUT):
    end = time.time() + timeout

    def try_find_in_current():
        try:
            return driver.find_element(by, value)
        except Exception:
            return None

    def search_frames(visited_frames):
        el = try_find_in_current()
        if el:
            return el
        try:
            frames = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            frames = []
        for idx, fr in enumerate(frames):
            try:
                fid = f"{idx}:{fr.get_attribute('id') or ''}:{fr.get_attribute('name') or ''}:{(fr.get_attribute('src') or '')[:120]}"
                if fid in visited_frames:
                    continue
                visited_frames.add(fid)
                try:
                    driver.switch_to.frame(fr)
                except Exception:
                    try:
                        driver.switch_to.frame(idx)
                    except Exception:
                        continue
                el = try_find_in_current()
                if el:
                    return el
                found = search_frames(visited_frames)
                if found:
                    return found
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
            except Exception:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
        return None

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    visited = set()
    while time.time() < end:
        found = search_frames(visited)
        if found:
            return found
        time.sleep(0.25)
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    return None

def ensure_on_target_context_and_find(driver, target_selector=("id","cpf"), timeout=FIND_TIMEOUT):
    by_map = {"id": By.ID, "name": By.NAME, "xpath": By.XPATH, "css": By.CSS_SELECTOR}
    by = by_map.get(target_selector[0], By.ID)
    value = target_selector[1]
    el = find_element_across_frames(driver, by, value, timeout=timeout)
    if el:
        return el
    try:
        handles = driver.window_handles
        for h in handles:
            try:
                driver.switch_to.window(h)
                el = find_element_across_frames(driver, by, value, timeout=3)
                if el:
                    return el
            except Exception:
                pass
    except Exception:
        pass
    save_capture("proxysite", "no_cpf_found", driver.page_source)
    raise RuntimeError("Campo 'cpf' não encontrado no conteúdo proxied.")

def click_next_button_selenium(driver):
    btn = None
    try:
        btn = driver.find_element(By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'próximo') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'proximo')]")
    except Exception:
        btn = None
    if not btn:
        try:
            btn = driver.find_element(By.NAME, "btnentrar")
        except Exception:
            btn = None
    if not btn:
        try:
            form = driver.find_element(By.TAG_NAME, "form")
            btn = form.find_element(By.XPATH, ".//button")
        except Exception:
            btn = None
    if not btn:
        return False
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.12)
        btn.click()
        time.sleep(0.18)
        return True
    except ElementClickInterceptedException:
        try:
            ActionChains(driver).move_to_element(btn).click(btn).perform()
            time.sleep(0.18)
            return True
        except Exception:
            pass
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(0.18)
        return True
    except Exception:
        pass
    try:
        parent = btn.find_element(By.XPATH, "ancestor::form")
        driver.execute_script("arguments[0].submit();", parent)
        time.sleep(0.18)
        return True
    except Exception:
        pass
    return False

def check_voltar_button_present(html):
    try:
        soup = BeautifulSoup(html, "html.parser")
        btn = soup.find("button", onclick=lambda v: v and "index_servidor.php" in v)
        return bool(btn)
    except Exception:
        return False

# ---------- routes ----------
@app.route('/', methods=['GET'])
def index():
    return render_template_string(TEMPLATE, captcha_data=None, result=None, session_id=None, cpf=None, mat=None, download_token=None)

@app.route('/start_headless', methods=['POST'])
def start_headless():
    cpf = request.form.get('cpf','').strip()
    matricula = request.form.get('matricula','').strip()

    # escolher proxy aleatório se houver PROXY_LIST
    proxy = None
    if PROXY_LIST:
        proxy = random.choice(PROXY_LIST)

    try:
        driver = start_selenium_headless(proxy=proxy)
    except Exception as e:
        return f"Erro ao iniciar Selenium: {e}", 500

    try:
        try:
            navigate_via_proxysite(driver, URL_BASE, timeout=NAV_TIMEOUT)
        except Exception as e:
            try:
                driver.quit()
            except:
                pass
            return f"Erro ao abrir via proxysite: {e}", 500

        try:
            el_cpf = ensure_on_target_context_and_find(driver, target_selector=("id","cpf"), timeout=FIND_TIMEOUT)
        except Exception as e:
            try:
                driver.quit()
            except:
                pass
            return f"Erro ao localizar campo CPF no conteúdo proxied: {e}", 500

        try:
            driver.execute_script("arguments[0].value = arguments[1];", el_cpf, cpf)
        except Exception:
            try:
                driver.execute_script("document.getElementById('cpf').value = arguments[0];", cpf)
            except Exception:
                pass

        try:
            el_mat = find_element_across_frames(driver, By.ID, "login", timeout=2)
            if el_mat:
                driver.execute_script("arguments[0].value = arguments[1];", el_mat, matricula)
            else:
                el_mat2 = find_element_across_frames(driver, By.ID, "matricula", timeout=2)
                if el_mat2:
                    driver.execute_script("arguments[0].value = arguments[1];", el_mat2, matricula)
                else:
                    driver.execute_script("var e=document.getElementById('login')||document.getElementById('matricula'); if(e) e.value = arguments[0];", matricula)
        except Exception:
            pass

        # try to find captcha image quickly without heavy parsing
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        img = soup.find("img", class_="imagem-captcha") or soup.find("img", src=lambda s: s and "captcha.php" in s)
        if not img or not img.get("src"):
            driver.quit()
            return "Não foi possível localizar o CAPTCHA na página.", 500
        src = img.get("src")
        from urllib.parse import urljoin
        captcha_url = urljoin(driver.current_url, src)
        import requests
        sess = requests.Session()
        for c in driver.get_cookies():
            sess.cookies.set(c['name'], c['value'], domain=c.get('domain'))
        resp = sess.get(captcha_url, timeout=10)
        if resp.status_code != 200:
            driver.quit()
            return f"Erro ao baixar captcha (HTTP {resp.status_code})", 500
        b64 = base64.b64encode(resp.content).decode("ascii")
        mime = resp.headers.get("Content-Type", "image/png") or "image/png"
        data_uri = f"data:{mime};base64,{b64}"
        session_id = str(uuid.uuid4())[:8]
        SESSIONS[session_id] = {"driver": driver, "meta": {"cpf": cpf, "matricula": matricula}, "created": time.time(), "password": None}
        return render_template_string(TEMPLATE, captcha_data=data_uri, session_id=session_id, cpf=cpf, mat=matricula, result="Resolva o CAPTCHA e envie.")
    except Exception as e:
        try:
            driver.quit()
        except:
            pass
        return f"Erro durante preparação: {e}", 500

@app.route('/submit_captcha', methods=['POST'])
def submit_captcha():
    session_id = request.form.get('session_id')
    captcha_text = request.form.get('captcha_text','').strip()
    cpf = request.form.get('cpf','').strip()
    matricula = request.form.get('matricula','').strip()
    if not session_id or session_id not in SESSIONS:
        return render_template_string(TEMPLATE, captcha_data=None, result="Sessão inválida ou expirada. Gere o CAPTCHA novamente.", session_id=None, cpf=None, mat=None)
    info = SESSIONS[session_id]
    driver = info['driver']
    try:
        try:
            el = driver.find_element(By.NAME, "captcha")
        except Exception:
            try:
                el = driver.find_element(By.ID, "captcha")
            except Exception:
                el = None
        if el:
            driver.execute_script("arguments[0].value = arguments[1];", el, captcha_text)

        clicked = click_next_button_selenium(driver)
        if not clicked:
            try:
                form = driver.find_element(By.TAG_NAME, "form")
                driver.execute_script("arguments[0].submit();", form)
            except Exception:
                pass

        # short polling for immediate password
        elapsed = 0.0
        found = None
        while elapsed < 6.0:
            html = driver.page_source
            senha = extract_senha_from_html(html)
            if senha:
                found = senha
                break
            if check_voltar_button_present(html):
                senha2 = extract_senha_from_html(html)
                if senha2:
                    found = senha2
                    break
            time.sleep(POLL_INTERVAL)
            elapsed += POLL_INTERVAL

        if found:
            try:
                driver.quit()
            except:
                pass
            SESSIONS.pop(session_id, None)
            token = str(uuid.uuid4())
            SESSIONS[token] = {"password_text": f"CPF: {cpf}\nNova senha: {found}"}
            return render_template_string(TEMPLATE, captcha_data=None, result=f"CPF: {cpf}\nNova senha: {found}", session_id=None, cpf=None, mat=None, download_token=token)

        # load clients once (uses Google Sheets if configured, otherwise Excel local)
        try:
            df = load_clients_auto()
        except Exception as e:
            try:
                driver.quit()
            except:
                pass
            return render_template_string(TEMPLATE, captcha_data=None, result=f"Erro ao carregar planilha: {e}", session_id=session_id, cpf=cpf, mat=matricula)

        cur_html = driver.page_source
        for step in range(MAX_QUESTION_STEPS):
            senha = extract_senha_from_html(cur_html)
            if senha:
                try:
                    driver.quit()
                except:
                    pass
                SESSIONS.pop(session_id, None)
                token = str(uuid.uuid4())
                SESSIONS[token] = {"password_text": f"CPF: {cpf}\nNova senha: {senha}"}
                return render_template_string(TEMPLATE, captcha_data=None, result=f"CPF: {cpf}\nNova senha: {senha}", session_id=None, cpf=None, mat=None, download_token=token)

            pergunta, options = extract_question_and_options_from_html(cur_html)
            if not pergunta:
                break

            cpf_norm = cpf.replace(".","").replace("-","")
            df_match = df[(df['CPF'] == cpf_norm) & (df['MATRICULA'] == matricula)]
            if df_match.empty:
                try:
                    driver.quit()
                except:
                    pass
                return render_template_string(TEMPLATE, captcha_data=None, result="Cliente não encontrado na planilha.", session_id=session_id, cpf=cpf, mat=matricula)
            row = df_match.iloc[0]

            pergunta_clean = re.sub(r'[^0-9A-Za-zÀ-ÿ\s]', ' ', pergunta).strip()
            col_candidate = normalize_colname(pergunta_clean)
            resposta_valor = None

            if col_candidate in row.index:
                resposta_valor = str(row[col_candidate]).strip()

            if not resposta_valor:
                for c in row.index:
                    if not isinstance(c, str):
                        continue
                    c_clean = c.strip().lower()
                    if c_clean and (c_clean in pergunta_clean.lower() or pergunta_clean.lower() in c_clean):
                        resposta_valor = str(row[c]).strip()
                        col_candidate = c
                        break

            if not resposta_valor:
                names = [str(c) for c in row.index if isinstance(c, str)]
                matches = difflib.get_close_matches(pergunta_clean, names, n=1, cutoff=0.6)
                if matches:
                    m = matches[0]
                    resposta_valor = str(row[m]).strip()
                    col_candidate = m

            if not resposta_valor:
                row_vals = {k: (str(v).strip() if pd.notna(v) else "") for k, v in row.items()}
                for val, text in options:
                    txt_norm = re.sub(r'[^0-9A-Za-zÀ-ÿ\s]', ' ', (text or "")).strip().lower()
                    for k, v in row_vals.items():
                        if not v:
                            continue
                        v_norm = v.strip().lower()
                        if v_norm == txt_norm or v_norm in txt_norm or txt_norm in v_norm:
                            resposta_valor = str(val).strip() if val else str(v).strip()
                            col_candidate = k
                            break
                    if resposta_valor:
                        break

            if not resposta_valor:
                # extract a helpful message from the page to show to the user
                try:
                    soup_dbg = BeautifulSoup(cur_html, "html.parser")
                    body_text = soup_dbg.get_text(" ", strip=True)
                    snippet = body_text[:800]
                except Exception:
                    snippet = cur_html[:800] if cur_html else "Sem conteúdo disponível."
                try:
                    driver.quit()
                except:
                    pass
                SESSIONS.pop(session_id, None)
                return render_template_string(TEMPLATE, captcha_data=None, result=f"Não encontrei resposta automática. Página retornou:\n\n{snippet}", session_id=session_id, cpf=cpf, mat=matricula)

            selected = False
            try:
                radio = driver.find_element(By.XPATH, f"//input[@name='resposta' and (@value='{resposta_valor}' or normalize-space(@value)='{resposta_valor}')]")
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", radio)
                radio.click()
                time.sleep(0.12)
                selected = True
            except Exception:
                selected = False

            if not selected:
                for val, text in options:
                    try:
                        if (str(val).strip() == str(resposta_valor).strip()) or (text and str(text).strip() and (str(text).strip() == str(resposta_valor).strip() or str(resposta_valor).strip() in str(text).strip())):
                            try:
                                label = driver.find_element(By.XPATH, f"//label[.//text()[contains(., \"{text}\")]]")
                                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", label)
                                label.click()
                                time.sleep(0.12)
                                selected = True
                                break
                            except Exception:
                                pass
                    except Exception:
                        pass

            if not selected:
                for val, text in options:
                    try:
                        if str(resposta_valor).strip() and str(text).strip() and str(resposta_valor).strip() in str(text).strip():
                            label = driver.find_element(By.XPATH, f"//label[.//text()[contains(., \"{text}\")]]")
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", label)
                            label.click()
                            time.sleep(0.12)
                            selected = True
                            break
                    except Exception:
                        pass

            if not selected:
                try:
                    driver.quit()
                except:
                    pass
                SESSIONS.pop(session_id, None)
                return render_template_string(TEMPLATE, captcha_data=None, result="Não foi possível marcar a opção no navegador; verifique a página.", session_id=session_id, cpf=cpf, mat=matricula)

            clicked = click_next_button_selenium(driver)
            if not clicked:
                try:
                    driver.quit()
                except:
                    pass
                SESSIONS.pop(session_id, None)
                return render_template_string(TEMPLATE, captcha_data=None, result="Falha ao clicar Próximo; verifique a página.", session_id=session_id, cpf=cpf, mat=matricula)

            # polling after click (shorter)
            found_senha = None
            end_poll = time.time() + POLL_AFTER_CLICK
            while time.time() < end_poll:
                cur_html = driver.page_source
                senha = extract_senha_from_html(cur_html)
                if senha:
                    found_senha = senha
                    break
                if check_voltar_button_present(cur_html):
                    senha2 = extract_senha_from_html(cur_html)
                    if senha2:
                        found_senha = senha2
                        break
                time.sleep(POLL_INTERVAL)
            if found_senha:
                try:
                    driver.quit()
                except:
                    pass
                SESSIONS.pop(session_id, None)
                token = str(uuid.uuid4())
                SESSIONS[token] = {"password_text": f"CPF: {cpf}\nNova senha: {found_senha}"}
                return render_template_string(TEMPLATE, captcha_data=None, result=f"CPF: {cpf}\nNova senha: {found_senha}", session_id=None, cpf=None, mat=None, download_token=token)

            cur_html = driver.page_source

        # end loop
        try:
            snippet = BeautifulSoup(driver.page_source, "html.parser").get_text(" ", strip=True)[:800]
        except Exception:
            snippet = "Sem conteúdo disponível."
        try:
            driver.quit()
        except:
            pass
        SESSIONS.pop(session_id, None)
        return render_template_string(TEMPLATE, captcha_data=None, result=f"Não foi possível obter a senha automaticamente. Página final:\n\n{snippet}", session_id=None, cpf=cpf, mat=matricula)
    except Exception as e:
        try:
            last = driver.page_source
        except Exception:
            last = ""
        try:
            driver.quit()
        except:
            pass
        SESSIONS.pop(session_id, None)
        snippet = ""
        try:
            snippet = BeautifulSoup(last, "html.parser").get_text(" ", strip=True)[:800]
        except Exception:
            snippet = last[:800] if last else ""
        return render_template_string(TEMPLATE, captcha_data=None, result=f"Erro durante processamento: {e}\n\nPágina/trecho:\n{snippet}", session_id=None, cpf=cpf, mat=matricula)

@app.route('/download/<token>', methods=['GET'])
def download_password(token):
    info = SESSIONS.get(token)
    if not info or 'password_text' not in info:
        return "Token inválido ou expirado", 404
    text = info['password_text']
    buf = io.BytesIO()
    buf.write(text.encode("utf-8"))
    buf.seek(0)
    resp = make_response(send_file(buf, as_attachment=True, download_name="senha.txt", mimetype="text/plain"))
    SESSIONS.pop(token, None)
    return resp

@app.route('/close_all', methods=['POST'])
def close_all():
    for sid, info in list(SESSIONS.items()):
        try:
            if 'driver' in info and info['driver']:
                info['driver'].quit()
        except:
            pass
        SESSIONS.pop(sid, None)
    return redirect(url_for('index'))

# ---------- run ----------
if __name__ == '__main__':
    if not os.path.exists(EXCEL_PATH):
        print(f"Atenção: {EXCEL_PATH} não encontrado. O app tentará usar Google Sheets se as variáveis estiverem configuradas.")
    # Allow PORT env override for local testing
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
