"""Instalador automático do Sistema PDF Amanda."""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
import warnings
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
REQ_PATH = BASE_DIR / "requirements.txt"

PACOTES = [
    "pdfplumber",
    "pymupdf",
    "pypdf",
    "python-docx",
    "openpyxl",
    "python-dotenv",
    "requests",
    "pillow",
    "pdf2image",
    "pytesseract",
    "rapidfuzz",
    "colorama",
    "tqdm",
    "groq",
    "google-generativeai",
    "openai",
    "anthropic",
]

IMPORTS = {
    "pdfplumber": "pdfplumber",
    "pymupdf": "fitz",
    "pypdf": "pypdf",
    "python-docx": "docx",
    "openpyxl": "openpyxl",
    "python-dotenv": "dotenv",
    "requests": "requests",
    "pillow": "PIL",
    "pdf2image": "pdf2image",
    "pytesseract": "pytesseract",
    "rapidfuzz": "rapidfuzz",
    "colorama": "colorama",
    "tqdm": "tqdm",
    "groq": "groq",
    "google-generativeai": "google.generativeai",
    "openai": "openai",
    "anthropic": "anthropic",
}


def configurar_utf8() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def executar(cmd: list[str], timeout: int = 180) -> tuple[bool, str]:
    try:
        proc = subprocess.run(
            cmd,
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
        texto = (proc.stdout or "") + (proc.stderr or "")
        return proc.returncode == 0, texto.strip()
    except Exception as exc:
        return False, str(exc)


def detectar_sistema() -> tuple[str, str]:
    sistema = platform.system() or "Sistema desconhecido"
    maquina = platform.machine() or "arquitetura desconhecida"
    versao = platform.release()
    if sistema == "Linux":
        try:
            texto = Path("/proc/version").read_text(encoding="utf-8", errors="ignore").lower()
            if "microsoft" in texto or "wsl" in texto:
                sistema = "WSL/Linux"
        except Exception:
            pass
    bits = "64-bit" if sys.maxsize > 2**32 else "32-bit"
    return f"{sistema} {versao} {bits}", maquina


def verificar_python() -> bool:
    versao = platform.python_version()
    print(f"Python detectado: {versao}")
    if sys.version_info >= (3, 10):
        return True

    sistema = platform.system()
    print("\nPython 3.10 ou mais novo é necessário.")
    if sistema == "Windows":
        print("Instale pelo link: https://www.python.org/downloads/")
        print("Durante a instalação, marque a opção 'Add Python to PATH'.")
    elif sistema == "Darwin":
        print("No macOS, instale com: brew install python3")
        print("Ou baixe pelo link: https://www.python.org/downloads/")
    else:
        print("No Linux, tente: sudo apt-get install -y python3 python3-pip")
    return False


def garantir_pip() -> bool:
    print("\nAtualizando pip...")
    ok, saida = executar([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], timeout=240)
    if not ok:
        print("Aviso: não consegui atualizar o pip. Vou continuar mesmo assim.")
        if saida:
            print(saida[-500:])
    return True


def escrever_requirements() -> None:
    REQ_PATH.write_text("\n".join(PACOTES) + "\n", encoding="utf-8")


def instalar_pacotes() -> tuple[list[str], list[str]]:
    instalados: list[str] = []
    falharam: list[str] = []
    print("\nInstalando bibliotecas Python...")
    for pacote in PACOTES:
        print(f"  - {pacote}...", end=" ", flush=True)
        ok, _ = executar([sys.executable, "-m", "pip", "install", "-q", pacote], timeout=300)
        if not ok:
            ok, _ = executar([sys.executable, "-m", "pip", "install", "-q", "--user", pacote], timeout=300)
        if ok:
            print("OK")
            instalados.append(pacote)
        else:
            print("AVISO")
            falharam.append(pacote)
    return instalados, falharam


def detectar_gerenciador_linux() -> str:
    for nome in ("apt-get", "dnf", "pacman"):
        if shutil.which(nome):
            return nome
    return ""


def instalar_ocr_opcional() -> str:
    print("\nVerificando OCR/Tesseract (opcional)...")
    if shutil.which("tesseract"):
        return "disponível"

    sistema = platform.system()
    comandos: list[list[str]] = []
    if sistema == "Windows":
        if shutil.which("winget"):
            comandos.append(["winget", "install", "--id", "UB-Mannheim.TesseractOCR", "-e", "--silent"])
        if shutil.which("choco"):
            comandos.append(["choco", "install", "tesseract", "-y"])
    elif sistema == "Darwin":
        if shutil.which("brew"):
            comandos.extend([["brew", "install", "tesseract"], ["brew", "install", "poppler"]])
        else:
            print("Homebrew não encontrado. Para OCR, instale: https://brew.sh/")
    else:
        gerente = detectar_gerenciador_linux()
        if gerente == "apt-get":
            comandos.append(["sudo", "apt-get", "install", "-y", "tesseract-ocr", "tesseract-ocr-por", "poppler-utils"])
        elif gerente == "dnf":
            comandos.append(["sudo", "dnf", "install", "-y", "tesseract", "tesseract-langpack-por", "poppler-utils"])
        elif gerente == "pacman":
            comandos.append(["sudo", "pacman", "-S", "--noconfirm", "tesseract", "tesseract-data-por", "poppler"])

    for cmd in comandos:
        print("Tentando:", " ".join(cmd))
        ok, _ = executar(cmd, timeout=240)
        if ok and shutil.which("tesseract"):
            return "disponível"

    if sistema == "Windows":
        print(
            "Para habilitar leitura de PDFs escaneados, instale o Tesseract:\n"
            "https://github.com/UB-Mannheim/tesseract/wiki\n"
            "Não é obrigatório — o sistema funciona sem ele para PDFs normais."
        )
        print(
            "Para Poppler no Windows, veja:\n"
            "https://github.com/oschwartz10612/poppler-windows"
        )
    return "não disponível (opcional)"


def validar_env() -> tuple[str, dict[str, bool]]:
    chaves = ["GROQ_API_KEY", "GEMINI_API_KEY", "OPENAI_API_KEY", "ANTHROPIC_API_KEY"]
    if not ENV_PATH.exists():
        ENV_PATH.write_text("\n".join(f"{chave}=" for chave in chaves) + "\n", encoding="utf-8")
        print("\nArquivo .env criado. Adicione uma chave se desejar análise inteligente. Não é obrigatório.")

    conteudo = ENV_PATH.read_text(encoding="utf-8", errors="ignore").splitlines()
    valores: dict[str, str] = {}
    for linha in conteudo:
        if "=" in linha and not linha.strip().startswith("#"):
            chave, valor = linha.split("=", 1)
            valores[chave.strip()] = valor.strip().strip('"').strip("'")
    status = {chave: bool(valores.get(chave)) for chave in chaves}
    print("\nVerificando chaves de IA:")
    for chave in chaves:
        simbolo = "✔" if status[chave] else "✗"
        texto = "encontrada" if status[chave] else "não configurada"
        print(f"  {simbolo} {chave} {texto}")
    provedor = "nenhum"
    for nome, chave in [("Groq", "GROQ_API_KEY"), ("Gemini", "GEMINI_API_KEY"), ("OpenAI", "OPENAI_API_KEY"), ("Anthropic", "ANTHROPIC_API_KEY")]:
        if status[chave]:
            provedor = nome
            break
    return provedor, status


def testar_imports() -> dict[str, bool]:
    print("\nFazendo teste de sanidade das bibliotecas...")
    resultado: dict[str, bool] = {}
    for pacote, modulo in IMPORTS.items():
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                __import__(modulo)
            resultado[pacote] = True
        except Exception:
            resultado[pacote] = False
    return resultado


def criar_atalhos() -> None:
    if platform.system() == "Windows":
        (BASE_DIR / "INICIAR.bat").write_text(
            "@echo off\n"
            "where python >nul 2>nul\n"
            "if %errorlevel%==0 (\n"
            "  python INICIAR.py\n"
            ") else (\n"
            "  where py >nul 2>nul\n"
            "  if %errorlevel%==0 (\n"
            "    py -3 INICIAR.py\n"
            "  ) else (\n"
            "    echo Python nao encontrado. Execute INSTALAR.py primeiro.\n"
            "  )\n"
            ")\n"
            "pause\n",
            encoding="utf-8",
        )
    else:
        sh = BASE_DIR / "iniciar.sh"
        sh.write_text("#!/bin/bash\npython3 INICIAR.py\n", encoding="utf-8")
        try:
            sh.chmod(0o755)
        except Exception:
            pass


def criar_ajuda() -> None:
    try:
        from INICIAR import gerar_ajuda

        gerar_ajuda()
    except Exception:
        texto = "Execute python INICIAR.py e escolha uma opção no menu.\n"
        (BASE_DIR / "AJUDA.txt").write_text(texto, encoding="utf-8")


def relatorio_final(sistema: str, python_ok: bool, imports: dict[str, bool], ocr: str, provedor: str) -> None:
    print("\n╔══════════════════════════════════════╗")
    print("║    INSTALAÇÃO CONCLUÍDA              ║")
    print("╠══════════════════════════════════════╣")
    print(f"║ Sistema: {sistema[:27]:<27} ║")
    print(f"║ Python:  {platform.python_version():<27} ║")
    print("╠══════════════════════════════════════╣")
    principais = ["pdfplumber", "pymupdf", "pypdf", "python-docx", "openpyxl"]
    for pacote in principais:
        status = "✔ instalado" if imports.get(pacote) else "✗ verificar"
        print(f"║ {pacote[:14]:<14} {status:<20} ║")
    print(f"║ OCR/Tesseract   {ocr[:20]:<20} ║")
    ia = f"✔ chave configurada ({provedor})" if provedor != "nenhum" else "opcional não configurada"
    print(f"║ IA              {ia[:20]:<20} ║")
    print("╠══════════════════════════════════════╣")
    comando = "python INICIAR.py" if python_ok else "instale Python 3.10+"
    print(f"║ Execute: {comando:<27} ║")
    print("╚══════════════════════════════════════╝")


def main() -> int:
    configurar_utf8()
    sistema, maquina = detectar_sistema()
    print(f"Sistema detectado: {sistema} ({maquina})")

    python_ok = verificar_python()
    if not python_ok:
        relatorio_final(sistema, False, {}, "não verificado", "nenhum")
        input("\nPressione ENTER para sair...")
        return 1

    garantir_pip()
    escrever_requirements()
    instalar_pacotes()
    ocr = instalar_ocr_opcional()
    provedor, _ = validar_env()
    imports = testar_imports()
    criar_atalhos()
    criar_ajuda()
    relatorio_final(sistema, True, imports, ocr, provedor)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nInstalação interrompida.")
        raise SystemExit(130)
