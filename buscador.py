"""Mecanismos de busca sem IA para os PDFs processados."""

from __future__ import annotations

import json
import math
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    from rapidfuzz import fuzz
except Exception:
    fuzz = None


STOPWORDS = {
    "a", "o", "as", "os", "um", "uma", "de", "da", "do", "das", "dos", "em", "no", "na", "nos", "nas",
    "por", "para", "com", "sem", "e", "ou", "que", "se", "ao", "aos", "à", "às", "como", "foi", "sao",
    "são", "ser", "ter", "ha", "há", "mais", "menos", "este", "esta", "esse", "essa", "pela", "pelo",
}


def limpar_nome_arquivo(texto: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", texto, flags=re.I).strip("_")[:60] or "busca"


def tokens(texto: str) -> list[str]:
    return re.findall(r"[A-Za-zÀ-ÿ0-9]{2,}", texto.lower())


def dobrar(texto: str) -> str:
    """Remove acentos para comparações tolerantes, preservando o original fora da busca."""
    normal = unicodedata.normalize("NFD", texto)
    return "".join(ch for ch in normal if unicodedata.category(ch) != "Mn").lower()


def trecho(texto: str, inicio: int, fim: int, termo: str | None = None, margem: int = 150) -> str:
    a = max(0, inicio - margem)
    b = min(len(texto), fim + margem)
    ctx = texto[a:b].replace("\n", " ")
    if termo:
        ctx = re.sub(re.escape(termo), lambda m: f"[>>> {m.group(0)} <<<]", ctx, flags=re.I)
    return re.sub(r"\s+", " ", ctx).strip()


def construir_indice(resultados: list[dict[str, Any]], caminho_saida: str | Path) -> dict[str, Any]:
    """Cria índice invertido e guarda textos, campos e tabelas para busca posterior."""
    indice: dict[str, Any] = {
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
        "documentos": {},
        "invertido": defaultdict(list),
    }
    for doc in resultados:
        arquivo = doc["arquivo"]
        doc_id = doc.get("id") or doc.get("caminho") or arquivo
        indice["documentos"][doc_id] = {
            "arquivo": arquivo,
            "caminho": doc.get("caminho", ""),
            "origem": doc.get("origem", ""),
            "motor": doc.get("motor", ""),
            "paginas": doc.get("paginas", []),
            "campos": doc.get("campos", {}),
            "tabelas": doc.get("tabelas", []),
        }
        for pagina in doc.get("paginas", []):
            texto = pagina.get("text", "")
            for match in re.finditer(r"[A-Za-zÀ-ÿ0-9]{2,}", texto):
                indice["invertido"][match.group(0).lower()].append({
                    "arquivo": arquivo,
                    "caminho": doc.get("caminho", ""),
                    "origem": doc.get("origem", ""),
                    "pagina": pagina.get("page", 0),
                    "posicao": match.start(),
                })
    indice["invertido"] = dict(indice["invertido"])
    Path(caminho_saida).write_text(json.dumps(indice, ensure_ascii=False, indent=2), encoding="utf-8")
    return indice


def carregar_indice(caminho: str | Path) -> dict[str, Any]:
    return json.loads(Path(caminho).read_text(encoding="utf-8"))


def busca_exata(indice: dict[str, Any], termo: str) -> list[dict[str, Any]]:
    resultados = []
    if not termo.strip():
        return resultados
    pattern = re.compile(re.escape(termo), re.I)
    termo_dobrado = dobrar(termo)
    termo_tokens = tokens(termo)
    if len(termo_tokens) == 1 and " " not in termo.strip():
        postings = indice.get("invertido", {}).get(termo_tokens[0].lower(), [])
        if postings:
            documentos = list(indice.get("documentos", {}).values())
            for post in postings[:1000]:
                doc = next(
                    (
                        d for d in documentos
                        if d.get("caminho") == post.get("caminho") or d.get("arquivo") == post.get("arquivo")
                    ),
                    None,
                )
                if not doc:
                    continue
                pagina = next((p for p in doc.get("paginas", []) if p.get("page") == post.get("pagina")), None)
                if not pagina:
                    continue
                texto = pagina.get("text", "")
                pos = int(post.get("posicao", 0))
                fim = pos + len(termo)
                resultados.append({
                    "tipo": "exata",
                    "arquivo": doc.get("arquivo", post.get("arquivo", "")),
                    "caminho": doc.get("caminho", post.get("caminho", "")),
                    "origem": doc.get("origem", post.get("origem", "")),
                    "pagina": post.get("pagina"),
                    "posicao": pos,
                    "score": 1.0,
                    "contexto": trecho(texto, pos, fim, texto[pos:fim]),
                })
            return resultados
    for _, doc in indice.get("documentos", {}).items():
        arquivo = doc.get("arquivo", "")
        for pagina in doc.get("paginas", []):
            texto = pagina.get("text", "")
            achou_literal = False
            for match in pattern.finditer(texto):
                achou_literal = True
                resultados.append({
                    "tipo": "exata",
                    "arquivo": arquivo,
                    "caminho": doc.get("caminho", ""),
                    "origem": doc.get("origem", ""),
                    "pagina": pagina.get("page"),
                    "posicao": match.start(),
                    "score": 1.0,
                    "contexto": trecho(texto, match.start(), match.end(), termo),
                })
            if achou_literal:
                continue
            texto_dobrado = dobrar(texto)
            inicio = 0
            while termo_dobrado and (pos := texto_dobrado.find(termo_dobrado, inicio)) >= 0:
                original = texto[pos:pos + len(termo)]
                resultados.append({
                    "tipo": "exata",
                    "arquivo": arquivo,
                    "caminho": doc.get("caminho", ""),
                    "origem": doc.get("origem", ""),
                    "pagina": pagina.get("page"),
                    "posicao": pos,
                    "score": 1.0,
                    "contexto": trecho(texto, pos, pos + len(termo), original),
                })
                inicio = pos + max(1, len(termo_dobrado))
    return resultados


def busca_regex(indice: dict[str, Any], expressao: str) -> list[dict[str, Any]]:
    resultados = []
    try:
        pattern = re.compile(expressao, re.I | re.M)
    except re.error as exc:
        return [{"erro": f"Regex inválida: {exc}"}]
    for _, doc in indice.get("documentos", {}).items():
        arquivo = doc.get("arquivo", "")
        for pagina in doc.get("paginas", []):
            texto = pagina.get("text", "")
            for match in pattern.finditer(texto):
                termo = match.group(0)
                resultados.append({
                    "tipo": "regex",
                    "arquivo": arquivo,
                    "caminho": doc.get("caminho", ""),
                    "origem": doc.get("origem", ""),
                    "pagina": pagina.get("page"),
                    "posicao": match.start(),
                    "score": 1.0,
                    "contexto": trecho(texto, match.start(), match.end(), termo),
                })
    return resultados


def _similaridade(a: str, b: str) -> float:
    a_norm = dobrar(a)
    b_norm = dobrar(b)
    if fuzz:
        return max(fuzz.ratio(a_norm, b_norm), fuzz.partial_ratio(a_norm, b_norm)) / 100.0
    from difflib import SequenceMatcher

    return SequenceMatcher(None, a_norm, b_norm).ratio()


def busca_fuzzy(indice: dict[str, Any], termo: str, minimo: float = 0.75) -> list[dict[str, Any]]:
    resultados = []
    termo_norm = termo.lower().strip()
    termo_base = dobrar(termo_norm)
    n = max(1, len(tokens(termo_norm)))
    for _, doc in indice.get("documentos", {}).items():
        arquivo = doc.get("arquivo", "")
        for pagina in doc.get("paginas", []):
            texto = pagina.get("text", "")
            palavras = list(re.finditer(r"[A-Za-zÀ-ÿ0-9]{2,}", texto))
            for i in range(0, len(palavras)):
                janela = palavras[i:i + n]
                if not janela:
                    continue
                candidato = " ".join(w.group(0) for w in janela)
                candidato_base = dobrar(candidato)
                if (
                    n == 1
                    and len(termo_base) >= 4
                    and termo_base[:3] != candidato_base[:3]
                    and termo_base not in candidato_base
                    and candidato_base not in termo_base
                ):
                    continue
                score = _similaridade(termo_norm, candidato.lower())
                if score >= minimo:
                    resultados.append({
                        "tipo": "fuzzy",
                        "arquivo": arquivo,
                        "caminho": doc.get("caminho", ""),
                        "origem": doc.get("origem", ""),
                        "pagina": pagina.get("page"),
                        "posicao": janela[0].start(),
                        "score": round(score, 3),
                        "contexto": trecho(texto, janela[0].start(), janela[-1].end(), candidato),
                    })
    return sorted(resultados, key=lambda r: r.get("score", 0), reverse=True)[:100]


def _chunks(texto: str, tamanho: int = 900, sobreposicao: int = 160) -> list[tuple[int, str]]:
    partes = []
    inicio = 0
    while inicio < len(texto):
        partes.append((inicio, texto[inicio:inicio + tamanho]))
        inicio += max(1, tamanho - sobreposicao)
    return partes


def busca_semantica(indice: dict[str, Any], consulta: str) -> list[dict[str, Any]]:
    q_tokens = [t for t in tokens(consulta) if t not in STOPWORDS]
    if not q_tokens:
        return []
    docs = []
    for _, doc in indice.get("documentos", {}).items():
        arquivo = doc.get("arquivo", "")
        for pagina in doc.get("paginas", []):
            for pos, chunk in _chunks(pagina.get("text", "")):
                tks = [t for t in tokens(chunk) if t not in STOPWORDS]
                if tks:
                    docs.append((arquivo, doc.get("caminho", ""), doc.get("origem", ""), pagina.get("page"), pos, chunk, Counter(tks)))
    df = Counter()
    for _, _, _, _, _, _, contagem in docs:
        for tk in contagem:
            df[tk] += 1
    total_docs = max(1, len(docs))
    resultados = []
    q_count = Counter(q_tokens)
    for arquivo, caminho, origem, pagina, pos, chunk, contagem in docs:
        score = 0.0
        for tk, qtf in q_count.items():
            if tk in contagem:
                idf = math.log((1 + total_docs) / (1 + df[tk])) + 1
                score += qtf * contagem[tk] * idf
        if score > 0:
            normalizado = min(1.0, score / (len(q_tokens) * 4))
            resultados.append({
                "tipo": "semantica",
                "arquivo": arquivo,
                "caminho": caminho,
                "origem": origem,
                "pagina": pagina,
                "posicao": pos,
                "score": round(normalizado, 3),
                "contexto": trecho(chunk, 0, min(len(chunk), 80), None),
            })
    return sorted(resultados, key=lambda r: r["score"], reverse=True)[:50]


def busca_tabelas(indice: dict[str, Any], termo: str) -> list[dict[str, Any]]:
    resultados = []
    termo_lower = termo.lower()
    for _, doc in indice.get("documentos", {}).items():
        arquivo = doc.get("arquivo", "")
        for tabela in doc.get("tabelas", []):
            for li, linha in enumerate(tabela.get("rows", []), start=1):
                for ci, celula in enumerate(linha, start=1):
                    valor = str(celula or "")
                    if termo_lower in valor.lower():
                        resultados.append({
                            "tipo": "tabela",
                            "arquivo": arquivo,
                            "caminho": doc.get("caminho", ""),
                            "origem": doc.get("origem", ""),
                            "pagina": tabela.get("page"),
                            "linha": li,
                            "coluna": ci,
                            "score": 1.0,
                            "contexto": trecho(valor, valor.lower().find(termo_lower), valor.lower().find(termo_lower) + len(termo), termo),
                        })
    return resultados


def busca_campo(indice: dict[str, Any], campo: str, termo: str = "") -> list[dict[str, Any]]:
    resultados = []
    campo_norm = campo.lower()
    termo_norm = termo.lower()
    for _, doc in indice.get("documentos", {}).items():
        arquivo = doc.get("arquivo", "")
        for nome_campo, valores in doc.get("campos", {}).items():
            if campo_norm in nome_campo.lower():
                for valor in valores:
                    if not termo_norm or termo_norm in valor.lower():
                        resultados.append({
                            "tipo": "campo",
                            "arquivo": arquivo,
                            "caminho": doc.get("caminho", ""),
                            "origem": doc.get("origem", ""),
                            "pagina": "",
                            "campo": nome_campo,
                            "score": 1.0,
                            "contexto": valor,
                        })
    return resultados


def formatar_resultados(resultados: list[dict[str, Any]], limite: int = 50) -> str:
    if not resultados:
        return "Nenhum resultado encontrado."
    linhas = []
    for i, item in enumerate(resultados[:limite], start=1):
        if "erro" in item:
            linhas.append(item["erro"])
            continue
        caminho = item.get("caminho") or item.get("origem") or item.get("arquivo", "")
        loc = f"Encontrado em: {caminho or item.get('arquivo', '')}"
        if item.get("arquivo") and item.get("arquivo") not in loc:
            loc += f" — {item.get('arquivo')}"
        if item.get("pagina"):
            loc += f" — Página {item.get('pagina')}"
        if item.get("linha"):
            loc += f" | linha {item.get('linha')} coluna {item.get('coluna')}"
        score = item.get("score")
        score_txt = f" | score {score:.3f}" if isinstance(score, (int, float)) else ""
        linhas.append(f"{i}. {loc}{score_txt}\n   {item.get('contexto', '')}")
    if len(resultados) > limite:
        linhas.append(f"... {len(resultados) - limite} resultados adicionais omitidos no terminal.")
    return "\n".join(linhas)


def busca_sql_estruturada(**filtros: Any) -> list[dict[str, Any]]:
    """Busca na base SQLite da camada estruturada."""
    from banco import buscar_registros

    return buscar_registros(**filtros)
