"""Extração de texto, tabelas, metadados e campos por heurística."""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any


VALID_PUNCT = set(" \n\t\r.,;:!?()[]{}<>-_/\\|@#$%&*+=ºª°'\"")


def log(msg: str) -> None:
    from datetime import datetime

    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def limpar_texto(texto: str) -> str:
    """Normaliza acentuação, controles invisíveis e espaços repetidos."""
    if not texto:
        return ""
    texto = unicodedata.normalize("NFC", str(texto))
    texto = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", texto)
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n{4,}", "\n\n\n", texto)
    linhas = [linha.rstrip() for linha in texto.splitlines()]
    return "\n".join(linhas).strip()


def score_qualidade(texto: str) -> float:
    """Pontua a proporção de caracteres úteis no texto extraído."""
    if not texto or len(texto.strip()) < 20:
        return 0.0
    total = len(texto)
    validos = sum(1 for ch in texto if ch.isalnum() or ch in VALID_PUNCT)
    return round(validos / total, 4)


def _normalizar_meta(meta: dict[str, Any] | None) -> dict[str, str]:
    meta = meta or {}
    nomes = {
        "author": "autor",
        "creator": "criador",
        "creationDate": "data_criacao",
        "creationdate": "data_criacao",
        "producer": "produtor",
        "title": "titulo",
        "modDate": "data_modificacao",
    }
    saida = {"autor": "", "criador": "", "data_criacao": "", "produtor": "", "titulo": ""}
    for chave, valor in meta.items():
        limpa = str(chave).strip("/").strip()
        destino = nomes.get(limpa, nomes.get(limpa.lower()))
        if destino:
            saida[destino] = limpar_texto(str(valor or ""))
    return saida


def _extrair_pdfplumber(pdf_path: Path) -> tuple[list[dict[str, Any]], dict[str, str], list[dict[str, Any]]]:
    import pdfplumber

    paginas: list[dict[str, Any]] = []
    tabelas: list[dict[str, Any]] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        meta = _normalizar_meta(pdf.metadata)
        for i, page in enumerate(pdf.pages, start=1):
            texto = limpar_texto(page.extract_text(x_tolerance=1, y_tolerance=3) or "")
            paginas.append({"page": i, "text": texto})
            try:
                for ti, tabela in enumerate(page.extract_tables() or [], start=1):
                    tabela_limpa = [[limpar_texto(cel or "") for cel in linha] for linha in tabela if linha]
                    if tabela_limpa:
                        tabelas.append({"page": i, "table_index": ti, "rows": tabela_limpa})
            except Exception as exc:
                tabelas.append({"page": i, "table_index": 0, "error": str(exc), "rows": []})
    return paginas, meta, tabelas


def _extrair_pymupdf_blocks(pdf_path: Path) -> tuple[list[dict[str, Any]], dict[str, str], list[dict[str, Any]]]:
    import fitz

    paginas = []
    with fitz.open(str(pdf_path)) as doc:
        meta = _normalizar_meta(doc.metadata)
        for i, page in enumerate(doc, start=1):
            blocos = page.get_text("blocks") or []
            blocos = sorted(blocos, key=lambda b: (round(b[1], 1), round(b[0], 1)))
            texto = "\n".join(str(b[4]) for b in blocos if len(b) > 4 and str(b[4]).strip())
            paginas.append({"page": i, "text": limpar_texto(texto)})
    return paginas, meta, []


def _extrair_pymupdf_dict(pdf_path: Path) -> tuple[list[dict[str, Any]], dict[str, str], list[dict[str, Any]]]:
    import fitz

    paginas = []
    with fitz.open(str(pdf_path)) as doc:
        meta = _normalizar_meta(doc.metadata)
        for i, page in enumerate(doc, start=1):
            data = page.get_text("dict")
            linhas = []
            for bloco in data.get("blocks", []):
                for linha in bloco.get("lines", []):
                    partes = []
                    for span in linha.get("spans", []):
                        texto = span.get("text", "")
                        if texto.strip():
                            partes.append(texto)
                    if partes:
                        linhas.append(" ".join(partes))
            paginas.append({"page": i, "text": limpar_texto("\n".join(linhas))})
    return paginas, meta, []


def _extrair_pypdf(pdf_path: Path) -> tuple[list[dict[str, Any]], dict[str, str], list[dict[str, Any]]]:
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    meta = _normalizar_meta(dict(reader.metadata or {}))
    paginas = []
    for i, page in enumerate(reader.pages, start=1):
        paginas.append({"page": i, "text": limpar_texto(page.extract_text() or "")})
    return paginas, meta, []


def _extrair_ocr(pdf_path: Path) -> tuple[list[dict[str, Any]], dict[str, str], list[dict[str, Any]]]:
    import pytesseract
    from pdf2image import convert_from_path

    paginas = []
    imagens = convert_from_path(str(pdf_path), dpi=220)
    for i, image in enumerate(imagens, start=1):
        texto = pytesseract.image_to_string(image, lang="por+eng")
        paginas.append({"page": i, "text": limpar_texto(texto)})
    return paginas, {}, []


MOTORES = [
    ("pdfplumber", _extrair_pdfplumber),
    ("pymupdf_blocks", _extrair_pymupdf_blocks),
    ("pymupdf_dict", _extrair_pymupdf_dict),
    ("pypdf", _extrair_pypdf),
    ("ocr_tesseract", _extrair_ocr),
]


def extrair_pdf(pdf_path: str | Path) -> dict[str, Any]:
    """Executa a cascata de motores e retorna tudo em estrutura serializável."""
    caminho = Path(pdf_path)
    tamanho = caminho.stat().st_size if caminho.exists() else 0
    tentativas = []
    melhor: dict[str, Any] | None = None

    for nome, func in MOTORES:
        try:
            paginas, meta, tabelas = func(caminho)
            texto_total = "\n\n".join(p["text"] for p in paginas)
            score = score_qualidade(texto_total)
            candidato = {
                "arquivo": caminho.name,
                "caminho": str(caminho),
                "motor": nome,
                "score": score,
                "paginas": paginas,
                "numero_paginas": len(paginas),
                "tamanho_arquivo": tamanho,
                "metadados": meta,
                "tabelas": tabelas,
                "erro": "",
            }
            tentativas.append({"motor": nome, "score": score, "erro": ""})
            if melhor is None or score > melhor["score"]:
                melhor = candidato
            if score >= 0.45:
                break
        except Exception as exc:
            tentativas.append({"motor": nome, "score": 0.0, "erro": str(exc)})

    if melhor is None:
        melhor = {
            "arquivo": caminho.name,
            "caminho": str(caminho),
            "motor": "erro",
            "score": 0.0,
            "paginas": [],
            "numero_paginas": 0,
            "tamanho_arquivo": tamanho,
            "metadados": {},
            "tabelas": [],
            "erro": "Nenhum motor conseguiu extrair texto.",
        }
    melhor["tentativas"] = tentativas
    melhor["texto_completo"] = "\n\n".join(
        f"--- PÁGINA {p['page']} ---\n{p['text']}" for p in melhor["paginas"]
    )
    melhor["campos"] = extrair_campos(melhor["texto_completo"], melhor["arquivo"])
    return melhor


def _ocorrencias(patterns: list[str], texto: str, flags: int = re.I | re.M) -> list[str]:
    achados: list[str] = []
    for pattern in patterns:
        for match in re.finditer(pattern, texto, flags):
            valor = match.group(1) if match.groups() else match.group(0)
            valor = limpar_texto(re.sub(r"\s+", " ", valor))
            valor = valor.strip(" :-–—\t\r\n")
            if 1 < len(valor) <= 600 and valor not in achados:
                achados.append(valor)
            if len(achados) >= 3:
                return achados
    return achados


def extrair_campos(texto: str, nome_arquivo: str = "") -> dict[str, list[str]]:
    """Extrai campos médicos por regex, sem uso de IA."""
    texto = limpar_texto(texto)
    secao = r"(.{5,600}?)(?=\n\s*(?:Paciente|Data|CPF|CRM|Médico|Medico|Convênio|Convenio|Procedimento|CID|Diagnóstico|Diagnostico|Conclusão|Conclusao|Laudo|Hospital|Clínica|Clinica|Código|Codigo|Solicitante)\b|$)"
    campos = {
        "Paciente / Nome": _ocorrencias([
            r"(?:Paciente|Nome)\s*[:\-]\s*([A-ZÁ-Ú][A-ZÁ-Úa-zá-ú\s.'-]{3,120})",
        ], texto),
        "Data do Exame / Procedimento": _ocorrencias([
            r"(?:Data\s*(?:do)?\s*(?:Exame|Procedimento|Atendimento)|Realizado em|Emissão|Emissao)\s*[:\-]?\s*(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})",
            r"\b(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{4})\b",
        ], texto),
        "Data de Nascimento": _ocorrencias([
            r"(?:Nascimento|Data\s*Nasc\.?|DN)\s*[:\-]?\s*(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})",
        ], texto),
        "CPF": _ocorrencias([r"\b(\d{3}\.?\d{3}\.?\d{3}-?\d{2})\b"], texto),
        "CRM do médico": _ocorrencias([r"\b(CRM\s*[-:/]?\s*[A-Z]{0,2}\s*\d{3,8})\b"], texto),
        "Médico Responsável": _ocorrencias([
            r"(?:M[eé]dico\s*(?:Respons[aá]vel|Executante)?|Dr\.?|Dra\.?)\s*[:\-]?\s*([A-ZÁ-Ú][A-Za-zÁ-ú\s.'-]{4,120})",
        ], texto),
        "Convênio / Operadora / Plano": _ocorrencias([
            r"(?:Conv[eê]nio|Operadora|Plano)\s*[:\-]\s*([A-Za-zÁ-ú0-9\s./-]{2,100})",
        ], texto),
        "Número de Prontuário / Processo": _ocorrencias([
            r"(?:Prontu[aá]rio|Processo|Registro|Atendimento)\s*(?:n[ºo.]*)?\s*[:\-]?\s*([A-Z0-9./-]{3,50})",
        ], texto),
        "Procedimento / Nome do Exame": _ocorrencias([
            r"(?:Procedimento|Exame|Estudo)\s*[:\-]\s*([A-Za-zÁ-ú0-9\s,./()-]{4,180})",
            r"\b(Arteriografia\s+de\s+Membro|Angioplastia\s+Sem\s+Stent|Aortografia\s+Abdominal)\b",
        ], texto + "\n" + nome_arquivo),
        "CID": _ocorrencias([r"\bCID\s*[:\-]?\s*([A-Z]\d{2}(?:\.\d{1,2})?)\b", r"\b([A-Z]\d{2}\.\d{1,2})\b"], texto),
        "Diagnóstico / Conclusão / Laudo": _ocorrencias([
            r"(?:Diagn[oó]stico|Conclus[aã]o|Laudo)\s*[:\-]?\s*" + secao,
        ], texto, re.I | re.M | re.S),
        "Valores monetários (R$)": _ocorrencias([r"(R\$\s*\d{1,3}(?:\.\d{3})*,\d{2})"], texto),
        "Hospital / Clínica / Instituição": _ocorrencias([
            r"(?:Hospital|Cl[ií]nica|Institui[cç][aã]o|Centro M[eé]dico)\s*[:\-]?\s*([A-Za-zÁ-ú0-9\s.'-]{3,120})",
        ], texto),
        "Código TUSS / CBHPM": _ocorrencias([
            r"\b(?:TUSS|CBHPM|C[oó]digo)\s*[:\-]?\s*(\d{6,10})\b",
        ], texto),
        "Lateralidade": _ocorrencias([
            r"\b(direito|direita|esquerdo|esquerda|bilateral|membro inferior direito|membro inferior esquerdo)\b",
        ], texto),
        "Médico solicitante": _ocorrencias([
            r"(?:Solicitante|M[eé]dico\s+Solicitante)\s*[:\-]\s*([A-ZÁ-Ú][A-Za-zÁ-ú\s.'-]{4,120})",
        ], texto),
    }
    return {campo: valores[:3] for campo, valores in campos.items()}
