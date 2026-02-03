\# Criacao Senha - Deploy no Render (Docker)



\## Arquivos no repositório

\- `app.py` - aplicação Flask + Selenium (já pronta)

\- `requirements.txt` - dependências Python

\- `Dockerfile` - imagem Docker com Chrome + Chromedriver

\- `.dockerignore` - arquivos a ignorar no build

\- `Procfile` - (opcional) comando para plataformas que usam Procfile

\- `sa.b64` - (opcional) base64 da chave de serviço Google (recomenda-se usar Secrets do Render)



\## Como preparar o repositório

1\. Coloque todos os arquivos na raiz do repositório.

2\. Faça commit e push para o GitHub:

```bash

git init

git add .

git commit -m "Initial commit - app + Dockerfile"

git branch -M main

git remote add origin https://github.com/SEU\_USUARIO/SEU\_REPO.git

git push -u origin main



