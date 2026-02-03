FROM seleniarm/selenium-base:latest

USER root
WORKDIR /app

# Instala Python e dependências
RUN apt-get update && apt-get install -y python3 python3-pip

COPY requirements.txt /app/
RUN pip3 install --upgrade pip && pip3 install -r /app/requirements.txt

COPY . /app
RUN mkdir -p /app/captures

ENV HEADLESS=true CAPTURE_ENABLED=false PORT=8080
EXPOSE 8080

CMD ["gunicorn", "app:app", "-b", "0.0.0.0:8080", "--workers", "1", "--threads", "4", "--timeout", "120"]
