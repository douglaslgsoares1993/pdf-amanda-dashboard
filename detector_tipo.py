"""Detector de tipo de documento PDF."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
CONFIG_USUARIO = BASE_DIR / "config_usuario.json"

TIPO_LIVRE = "Documento Livre"
TIPO_ESTRUTURADO = "Relatório Estruturado"
TIPO_INDEFINIDO = "Tipo Indefinido"


def _ler_config() -> dict[str, Any]:
    if not CONFIG_USUARIO.exists():
        return {}
    try:
        return json.loads(CONFIG_USUARIO.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _salvar_config(config: dict[str, Any]) -> None:
    CONFIG_USUARIO.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def tipo_salvo(pdf_path: str | Path) -> str | None:
    config = _ler_config()
    info = config.get("tipos_documento", {}).get(str(Path(pdf_path).resolve()))
    if isinstance(info, dict):
        return info.get("tipo")
    if isinstance(info, str):
        return info
    return None


def salvar_tipo(pdf_path: str | Path, tipo: str, confianca: float = 1.0) -> None:
    config = _ler_config()
    config.setdefault("tipos_documento", {})[str(Path(pdf_path).resolve())] = {
        "tipo": tipo,
        "confianca": round(confianca, 3),
    }
    _salvar_config(config)


def _texto_primeiras_paginas(pdf_path: Path, max_paginas: int = 3) -> str:
    try:
        import pdfplumber

        partes: list[str] = []
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages[:max_paginas]:
                partes.append(page.extract_text() or "")
        return "\n".join(partes)
    except Exception:
        try:
            from pypdf import PdfReader

            reader = PdfReader(str(pdf_path))
            return "\n".join((p.extract_text() or "") for p in reader.pages[:max_paginas])
        except Exception:
            return ""


def detectar_tipo_pdf(pdf_path: str | Path, usar_cache: bool = True) -> dict[str, Any]:
    """Classifica PDF como livre, estruturado ou indefinido."""
    caminho = Path(pdf_path)
    if usar_cache:
        salvo = tipo_salvo(caminho)
        if salvo:
            return {"tipo": salvo, "confianca": 1.0, "motivos": ["tipo salvo em config_usuario.json"]}

    texto = _texto_primeiras_paginas(caminho)
    if not texto.strip():
        return {"tipo": TIPO_INDEFINIDO, "confianca": 0.0, "motivos": ["sem texto nas primeiras páginas"]}

    labels = ["Cirurgia:", "Cirurgião:", "Lateralidade:", "Origem do Paciente:", "Tipo Atendimento:", "Cod. Atendimento:"]
    contagens = {label: len(re.findall(re.escape(label), texto, re.I)) for label in labels}
    registros = len(re.findall(r"^\s*\d{5,}\s+.+?\s+\d{11}\s+\d{2}/\d{2}/\d{4}\s+\d{1,3}\s+", texto, re.M))
    cabecalho = "Relatório Detalhado de Procedimentos Cirúrgicos" in texto
    tabela = "Prontuário Nome CPF Dt Nascimento Idade" in texto
    motivos: list[str] = []
    pontos = 0

    if sum(1 for v in contagens.values() if v >= 5) >= 3:
        pontos += 3
        motivos.append("rótulos repetidos de procedimento")
    if registros >= 5:
        pontos += 3
        motivos.append("vários registros com prontuário/CPF/data")
    if cabecalho:
        pontos += 2
        motivos.append("cabeçalho institucional de relatório cirúrgico")
    if tabela:
        pontos += 2
        motivos.append("cabeçalho tabular de prontuário/nome/CPF")

    if pontos >= 6:
        tipo = TIPO_ESTRUTURADO
    elif pontos <= 2:
        tipo = TIPO_LIVRE
    else:
        tipo = TIPO_INDEFINIDO
    confianca = min(1.0, pontos / 10)
    salvar_tipo(caminho, tipo, confianca)
    return {"tipo": tipo, "confianca": confianca, "motivos": motivos, "contagens": contagens}
