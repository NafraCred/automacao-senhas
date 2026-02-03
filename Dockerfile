# Dockerfile minimal for Flask + Selenium + Chrome
FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive PYTHONUNBUFFERED=1 LANG=C.UTF-8
WORKDIR /app

# Instala dependências do sistema para Chrome headless
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget unzip ca-certificates fonts-liberation libnss3 libxss1 libasound2 \
    libatk1.0-0 libatk-bridge2.0-0 libgtk-3-0 libgbm1 curl && rm -rf /var/lib/apt/lists/*

# Instala Google Chrome
RUN wget -q -O /tmp/chrome.deb https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb \
    && apt-get update && apt-get install -y /tmp/chrome.deb && rm -f /tmp/chrome.deb

# Instala Chromedriver compatível com versão do Chrome
RUN CHROME_VER=$(google-chrome --product-version | cut -d. -f1) && \
    LATEST=$(curl -sS https://chromedriver.storage.googleapis.com/LATEST_RELEASE_$CHROME_VER) && \
    wget -q -O /tmp/chromedriver.zip "https://chromedriver.storage.googleapis.com/${LATEST}/chromedriver_linux64.zip" && \
    unzip /tmp/chromedriver.zip -d /usr/local/bin/ && chmod +x /usr/local/bin/chromedriver && rm /tmp/chromedriver.zip

# Instala dependências Python
COPY requirements.txt /app/
RUN pip install --upgrade pip && pip install -r /app/requirements.txt

# Copia o código da aplicação
COPY . /app
RUN mkdir -p /app/captures

# Variáveis de ambiente padrão
ENV HEADLESS=true CAPTURE_ENABLED=false PORT=8080
EXPOSE 8080

# Comando para iniciar o servidor Flask via Gunicorn
CMD ["gunicorn", "app:app", "-b", "0.0.0.0:8080", "--workers", "1", "--threads", "4", "--timeout", "120"]
