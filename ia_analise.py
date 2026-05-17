"""Análise narrativa opcional com IA, sempre com fallback gracioso."""

from __future__ import annotations

import json
from typing import Any

from config import CONFIG, AIConfig


FALLBACK = {
    "tipo_documento": "Documento médico",
    "resumo": "Análise por IA indisponível; resumo baseado apenas nos campos extraídos por heurística.",
    "procedimentos": [],
    "dados_clinicos": [],
    "valores": [],
    "pendencias": [],
    "conclusao": "Consulte os campos extraídos e o texto completo do documento.",
}


def _prompt(texto: str, campos: dict[str, list[str]]) -> str:
    texto_limitado = texto[:18000]
    return f"""
Você analisará um documento médico em português. Retorne apenas JSON válido, sem markdown.
Campos obrigatórios:
tipo_documento, resumo, procedimentos, dados_clinicos, valores, pendencias, conclusao.

Tarefas:
- Identificar o tipo de documento.
- Resumir em linguagem clara para leigo.
- Listar procedimentos realizados com explicação simples.
- Identificar dados clínicos relevantes, diagnóstico e achados.
- Apontar valores e coberturas de plano mencionados.
- Identificar pendências, solicitações ou encaminhamentos.
- Gerar uma conclusão em português claro.

Campos extraídos automaticamente:
{json.dumps(campos, ensure_ascii=False, indent=2)}

Texto extraído:
{texto_limitado}
""".strip()


def _parse_json(conteudo: str) -> dict[str, Any]:
    conteudo = conteudo.strip()
    if conteudo.startswith("```"):
        conteudo = conteudo.strip("`")
        conteudo = conteudo.replace("json\n", "", 1).replace("JSON\n", "", 1)
    inicio = conteudo.find("{")
    fim = conteudo.rfind("}")
    if inicio >= 0 and fim >= inicio:
        conteudo = conteudo[inicio:fim + 1]
    data = json.loads(conteudo)
    saida = FALLBACK.copy()
    for chave in saida:
        if chave in data:
            saida[chave] = data[chave]
    return saida


def fallback_por_campos(campos: dict[str, list[str]], motivo: str = "") -> dict[str, Any]:
    saida = FALLBACK.copy()
    proc = campos.get("Procedimento / Nome do Exame", [])
    diag = campos.get("Diagnóstico / Conclusão / Laudo", [])
    valores = campos.get("Valores monetários (R$)", [])
    saida["procedimentos"] = proc
    saida["dados_clinicos"] = diag
    saida["valores"] = valores
    saida["resumo"] = "IA indisponível. Campos principais extraídos: " + "; ".join(
        f"{k}: {', '.join(v)}" for k, v in campos.items() if v
    )[:900]
    if motivo:
        saida["observacao_ia"] = motivo
    return saida


def analisar_documento(texto: str, campos: dict[str, list[str]], config: AIConfig = CONFIG) -> dict[str, Any]:
    """Chama o provedor configurado; qualquer erro retorna fallback."""
    if not config.enabled:
        return fallback_por_campos(campos, "Nenhuma chave de IA configurada.")
    prompt = _prompt(texto, campos)
    try:
        if config.provider == "groq":
            from groq import Groq

            client = Groq(api_key=config.api_key, timeout=20.0)
            resp = client.chat.completions.create(
                model=config.model or "llama-3.1-8b-instant",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                response_format={"type": "json_object"},
            )
            return _parse_json(resp.choices[0].message.content or "{}")

        if config.provider == "gemini":
            import google.generativeai as genai

            genai.configure(api_key=config.api_key)
            model = genai.GenerativeModel(config.model or "gemini-1.5-flash")
            resp = model.generate_content(prompt, generation_config={"temperature": 0.1})
            return _parse_json(resp.text or "{}")

        if config.provider == "openai":
            from openai import OpenAI

            client = OpenAI(api_key=config.api_key, timeout=20.0)
            resp = client.chat.completions.create(
                model=config.model or "gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                response_format={"type": "json_object"},
            )
            return _parse_json(resp.choices[0].message.content or "{}")

        if config.provider == "anthropic":
            from anthropic import Anthropic

            client = Anthropic(api_key=config.api_key, timeout=20.0)
            resp = client.messages.create(
                model=config.model or "claude-3-5-haiku-latest",
                max_tokens=1800,
                temperature=0.1,
                messages=[{"role": "user", "content": prompt}],
            )
            conteudo = "".join(getattr(part, "text", "") for part in resp.content)
            return _parse_json(conteudo or "{}")

        return fallback_por_campos(campos, f"Provedor não suportado: {config.provider}")
    except Exception as exc:
        return fallback_por_campos(campos, f"Falha na IA ({config.provider}): {exc}")
