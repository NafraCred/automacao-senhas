FROM browserless/chrome:latest

USER root
WORKDIR /app

# Instala Python e pip
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 python3-pip python3-venv curl unzip && rm -rf /var/lib/apt/lists/*

# Copia requirements e código
COPY requirements.txt /app/
COPY . /app

# Atualiza pip/setuptools/wheel e força reinstalação do blinker antes de instalar requirements
# --no-cache-dir evita cache grande no layer
RUN pip3 install --upgrade pip setuptools wheel && \
    pip3 install --no-cache-dir --upgrade --ignore-installed blinker && \
    pip3 install --no-cache-dir -r /app/requirements.txt

# Cria pasta de capturas
RUN mkdir -p /app/captures && chmod -R 755 /app

ENV HEADLESS=true CAPTURE_ENABLED=false PORT=8080
EXPOSE 8080

CMD ["gunicorn", "app:app", "-b", "0.0.0.0:8080", "--workers", "1", "--threads", "4", "--timeout", "120"]
