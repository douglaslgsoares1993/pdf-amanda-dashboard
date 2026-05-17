"""Gerencia pastas e arquivos PDF escolhidos pela usuária."""

from __future__ import annotations

import json
import os
import platform
from dataclasses import dataclass
from difflib import get_close_matches
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
CONFIG_USUARIO = BASE_DIR / "config_usuario.json"


@dataclass
class FontePDF:
    caminho: Path
    origem: str
    tipo: str


def _ler_json() -> dict[str, Any]:
    if not CONFIG_USUARIO.exists():
        return {}
    try:
        return json.loads(CONFIG_USUARIO.read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_config_usuario(config: dict[str, Any]) -> None:
    """Salva apenas preferências de caminhos, nunca chaves ou senhas."""
    seguro = {
        "ultima_pasta_usada": config.get("ultima_pasta_usada", ""),
        "pastas_adicionais": config.get("pastas_adicionais", [])[:20],
        "arquivos_avulsos_recentes": config.get("arquivos_avulsos_recentes", [])[:10],
    }
    CONFIG_USUARIO.write_text(json.dumps(seguro, ensure_ascii=False, indent=2), encoding="utf-8")


def normalizar_caminho(entrada: str) -> Path:
    """Aceita caminhos Windows, WSL, macOS/Linux, relativos e com aspas."""
    texto = (entrada or "").strip().strip('"').strip("'").strip()
    if not texto:
        return BASE_DIR
    if platform.system().lower() == "windows":
        if texto.startswith("/mnt/") and len(texto) > 6:
            drive = texto[5]
            resto = texto[7:].replace("/", "\\")
            texto = f"{drive.upper()}:\\{resto}"
        texto = texto.replace("/", "\\") if ":" in texto[:4] else texto
    caminho = Path(os.path.expanduser(texto))
    if not caminho.is_absolute():
        caminho = (BASE_DIR / caminho).resolve()
    if not caminho.exists() and caminho.suffix == "":
        possivel_pdf = caminho.with_suffix(".pdf")
        if possivel_pdf.exists():
            caminho = possivel_pdf
    return caminho


def sugerir_caminho(caminho: Path) -> str:
    """Sugere nome parecido dentro da pasta pai, quando possível."""
    pai = caminho.parent if caminho.parent.exists() else BASE_DIR
    try:
        nomes = [p.name for p in pai.iterdir()]
    except Exception:
        return ""
    sugestoes = get_close_matches(caminho.name, nomes, n=1, cutoff=0.55)
    return sugestoes[0] if sugestoes else ""


def pdfs_da_pasta(pasta: Path, origem: str, tipo: str) -> list[FontePDF]:
    if not pasta.exists() or not pasta.is_dir():
        return []
    return [FontePDF(p, origem, tipo) for p in sorted(pasta.glob("*.pdf")) if p.is_file()]


def carregar_fontes_iniciais() -> tuple[list[FontePDF], dict[str, Any], str]:
    """Carrega a pasta da sessão anterior; se não existir, usa a pasta do script."""
    config = _ler_json()
    mensagem = ""
    ultima = config.get("ultima_pasta_usada")
    if ultima and Path(ultima).exists():
        pasta_principal = Path(ultima)
        mensagem = f"Usando pasta da última sessão: {pasta_principal}"
    else:
        pasta_principal = BASE_DIR
        if ultima:
            mensagem = "A pasta anterior não foi encontrada. Usando a pasta atual."
        config["ultima_pasta_usada"] = str(pasta_principal)

    fontes = pdfs_da_pasta(pasta_principal, str(pasta_principal), "pasta principal")
    for pasta_txt in config.get("pastas_adicionais", []):
        pasta = Path(pasta_txt)
        fontes.extend(pdfs_da_pasta(pasta, str(pasta), "pasta adicional"))
    for arquivo_txt in config.get("arquivos_avulsos_recentes", []):
        arquivo = Path(arquivo_txt)
        if arquivo.exists() and arquivo.is_file() and arquivo.suffix.lower() == ".pdf":
            fontes.append(FontePDF(arquivo, str(arquivo.parent), "arquivo avulso"))
    return remover_duplicados(fontes), config, mensagem


def remover_duplicados(fontes: list[FontePDF]) -> list[FontePDF]:
    vistos: set[str] = set()
    saida: list[FontePDF] = []
    for fonte in fontes:
        chave = str(fonte.caminho.resolve()).lower()
        if chave not in vistos:
            vistos.add(chave)
            saida.append(fonte)
    return saida


def resumo_fontes(fontes: list[FontePDF]) -> list[tuple[str, int]]:
    contagem: dict[str, int] = {}
    for fonte in fontes:
        rotulo = Path(fonte.origem).name or fonte.origem
        if fonte.tipo == "pasta principal":
            rotulo += " (pasta atual)"
        contagem[rotulo] = contagem.get(rotulo, 0) + 1
    return sorted(contagem.items())


def fontes_para_processamento(fontes: list[FontePDF]) -> list[dict[str, str]]:
    return [
        {"caminho": str(fonte.caminho), "origem": fonte.origem, "tipo_origem": fonte.tipo}
        for fonte in remover_duplicados(fontes)
    ]


def validar_pasta(entrada: str) -> tuple[Path | None, str]:
    caminho = normalizar_caminho(entrada)
    if not caminho.exists():
        sugestao = sugerir_caminho(caminho)
        msg = "Pasta não encontrada."
        if sugestao:
            msg += f" Você quis dizer: {sugestao}?"
        return None, msg
    if not caminho.is_dir():
        return None, "Esse caminho existe, mas não é uma pasta."
    qtd = len(list(caminho.glob("*.pdf")))
    if qtd == 0:
        return caminho, "Pasta válida, mas nenhum PDF foi encontrado nela."
    return caminho, f"Pasta válida. {qtd} PDF(s) encontrado(s)."


def validar_arquivo_pdf(entrada: str) -> tuple[Path | None, str]:
    caminho = normalizar_caminho(entrada)
    if not caminho.exists():
        sugestao = sugerir_caminho(caminho)
        msg = "Arquivo não encontrado."
        if sugestao:
            msg += f" Você quis dizer: {sugestao}?"
        return None, msg
    if not caminho.is_file():
        return None, "Esse caminho existe, mas não é um arquivo."
    if caminho.suffix.lower() != ".pdf":
        return None, "Esse arquivo não parece ser PDF."
    return caminho, "Arquivo PDF adicionado."


def atualizar_pasta_principal(config: dict[str, Any], pasta: Path) -> None:
    config["ultima_pasta_usada"] = str(pasta)
    salvar_config_usuario(config)


def adicionar_pasta_config(config: dict[str, Any], pasta: Path) -> None:
    pastas = [p for p in config.get("pastas_adicionais", []) if Path(p) != pasta]
    pastas.append(str(pasta))
    config["pastas_adicionais"] = pastas[-20:]
    salvar_config_usuario(config)


def adicionar_arquivo_config(config: dict[str, Any], arquivo: Path) -> None:
    arquivos = [p for p in config.get("arquivos_avulsos_recentes", []) if Path(p) != arquivo]
    arquivos.append(str(arquivo))
    config["arquivos_avulsos_recentes"] = arquivos[-10:]
    salvar_config_usuario(config)
