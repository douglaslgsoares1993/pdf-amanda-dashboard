# render_setup.py — Prepara deploy no Render
# Executar UMA VEZ pelo Douglas: python render_setup.py

import os
import sys
import subprocess
from pathlib import Path

PASTA = Path(__file__).parent

def titulo(texto):
    print(f"\n{'═'*56}")
    print(f"  {texto}")
    print(f"{'═'*56}")

def ok(texto):
    print(f"  ✔  {texto}")

def info(texto):
    print(f"  →  {texto}")

def aviso(texto):
    print(f"  ⚠  {texto}")

# ── 1. Verifica git ─────────────────────────────────────────────
titulo("VERIFICANDO DEPENDÊNCIAS")
try:
    subprocess.run(["git","--version"], capture_output=True, check=True)
    ok("git instalado")
except (subprocess.CalledProcessError, FileNotFoundError):
    aviso("git não encontrado.")
    info("Instale em: https://git-scm.com/downloads")
    info("Após instalar, rode este script novamente.")
    sys.exit(1)

# ── 2. Inicia repositório git se necessário ─────────────────────
git_dir = PASTA / ".git"
if not git_dir.exists():
    subprocess.run(["git","init"], cwd=PASTA, capture_output=True)
    ok("Repositório git criado")
else:
    ok("Repositório git já existe")

# ── 3. Gera requirements_render.txt ────────────────────────────
titulo("GERANDO ARQUIVOS DE DEPLOY")

req_render = PASTA / "requirements_render.txt"
req_render.write_text(
    "streamlit>=1.32.0\n"
    "pandas\n"
    "plotly\n"
    "openpyxl\n"
    "python-dotenv\n"
    "pdfplumber\n"
    "PyMuPDF\n"
    "pypdf\n"
    "python-docx\n"
    "rapidfuzz\n"
    "psycopg2-binary\n",
    encoding="utf-8"
)
ok("requirements_render.txt gerado")

# ── 4. Gera render.yaml ─────────────────────────────────────────
render_yaml = PASTA / "render.yaml"
render_yaml.write_text(
    "services:\n"
    "  - type: web\n"
    "    name: pdf-amanda-dashboard\n"
    "    env: python\n"
    "    buildCommand: pip install -r requirements_render.txt\n"
    "    startCommand: >\n"
    "      streamlit run dashboard_app.py\n"
    "      --server.port $PORT\n"
    "      --server.address 0.0.0.0\n"
    "      --server.headless true\n"
    "    plan: free\n",
    encoding="utf-8"
)
ok("render.yaml gerado")

# ── 5. Gera .streamlit/config.toml ──────────────────────────────
pasta_st = PASTA / ".streamlit"
pasta_st.mkdir(exist_ok=True)
(pasta_st / "config.toml").write_text(
    "[server]\n"
    "headless = true\n"
    "enableCORS = false\n"
    "enableXsrfProtection = false\n\n"
    "[theme]\n"
    'primaryColor = "#185FA5"\n'
    'backgroundColor = "#FFFFFF"\n'
    'secondaryBackgroundColor = "#F7FBFF"\n'
    'textColor = "#1a1a1a"\n'
    'font = "sans serif"\n',
    encoding="utf-8"
)
ok(".streamlit/config.toml gerado")

# ── 6. Gera Procfile ────────────────────────────────────────────
(PASTA / "Procfile").write_text(
    "web: streamlit run dashboard_app.py "
    "--server.port $PORT "
    "--server.address 0.0.0.0 "
    "--server.headless true\n",
    encoding="utf-8"
)
ok("Procfile gerado")

# ── 7. Gera .gitignore ──────────────────────────────────────────
gitignore = PASTA / ".gitignore"
if not gitignore.exists():
    gitignore.write_text(
        "__pycache__/\n"
        "*.pyc\n"
        ".env\n"
        "SAIDA/erros.log\n"
        "*.log\n",
        encoding="utf-8"
    )
    ok(".gitignore gerado")

# ── 8. Instruções finais ─────────────────────────────────────────
titulo("DEPLOY NO RENDER — PASSO A PASSO")

print("""
  Todos os arquivos foram gerados. Agora siga estes passos:

  ── PASSO 1 ─────────────────────────────────────────────────
  Crie uma conta gratuita em:
    https://render.com

  ── PASSO 2 ─────────────────────────────────────────────────
  Crie um repositório no GitHub:
    https://github.com/new
    Nome sugerido: pdf-amanda-dashboard
    Deixe como PRIVADO se quiser proteger os dados

  ── PASSO 3 ─────────────────────────────────────────────────
  No terminal desta pasta, rode os comandos:

    git add .
    git commit -m "deploy inicial pdf amanda"
    git remote add origin https://github.com/SEU_USUARIO/pdf-amanda-dashboard.git
    git push -u origin main

  (substitua SEU_USUARIO pelo seu usuário do GitHub)

  ── PASSO 4 ─────────────────────────────────────────────────
  No Render:
    1. Clique em "New +" → "Web Service"
    2. Conecte sua conta do GitHub
    3. Selecione o repositório pdf-amanda-dashboard
    4. Configurações detectadas automaticamente pelo render.yaml
    5. Clique em "Create Web Service"
    6. Aguarde 2-3 minutos

  ── PASSO 5 ─────────────────────────────────────────────────
  Seu link será:
    https://pdf-amanda-dashboard.onrender.com

  ── BANCO DE DADOS ───────────────────────────────────────────
  OPÇÃO A — Supabase (dados sempre atualizados na nuvem):
    Adicione no .env do Render (Environment Variables):
      SUPABASE_URL = sua_url_do_supabase
      SUPABASE_KEY = sua_chave_do_supabase
    O sistema detecta e usa automaticamente.

  OPÇÃO B — Commit do banco (mais simples):
    Após processar novos PDFs localmente, rode:
      git add SAIDA/procedimentos.db
      git commit -m "atualiza dados"
      git push
    O Render faz redeploy em ~1 minuto.

  ── ATENÇÃO ──────────────────────────────────────────────────
  O plano gratuito do Render hiberna após 15 min sem acesso.
  A primeira abertura pode demorar ~30 segundos para acordar.
  Para acesso instantâneo sempre, use o plano pago ($7/mês).
""")

titulo("CONFIGURAÇÃO CONCLUÍDA")
ok("Todos os arquivos de deploy gerados com sucesso")
info("Siga os passos acima para publicar o dashboard")
print()
