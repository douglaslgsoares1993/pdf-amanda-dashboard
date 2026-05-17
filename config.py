"""Configuração central do sistema de PDFs."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover - fallback para ambientes mínimos
    load_dotenv = None


BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"


@dataclass
class AIConfig:
    provider: str | None
    api_key: str | None
    model: str | None
    available_keys: dict[str, bool]

    @property
    def enabled(self) -> bool:
        return bool(self.provider and self.api_key)


def _load_env_manual(path: Path) -> None:
    """Carrega .env sem sobrescrever variáveis já presentes."""
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def load_config() -> AIConfig:
    """Lê o .env e escolhe o provedor de IA pela prioridade solicitada."""
    if load_dotenv:
        load_dotenv(ENV_PATH)
    else:
        _load_env_manual(ENV_PATH)

    key_names = [
        "GROQ_API_KEY",
        "GEMINI_API_KEY",
        "OPENAI_API_KEY",
        "ANTHROPIC_API_KEY",
        "HF_API_KEY",
        "OPENROUTER_API_KEY",
        "CEREBRAS_API_KEY",
        "MISTRAL_API_KEY",
        "DEEPSEEK_API_KEY",
    ]
    available = {name: bool(os.getenv(name)) for name in key_names}

    priority = [
        ("groq", "GROQ_API_KEY", os.getenv("GROQ_MODEL", "llama-3.1-8b-instant")),
        ("gemini", "GEMINI_API_KEY", os.getenv("GEMINI_MODEL", "gemini-1.5-flash")),
        ("openai", "OPENAI_API_KEY", os.getenv("OPENAI_MODEL", "gpt-4o-mini")),
        ("anthropic", "ANTHROPIC_API_KEY", os.getenv("ANTHROPIC_MODEL", "claude-3-5-haiku-latest")),
    ]
    for provider, key_name, model in priority:
        key = os.getenv(key_name)
        if key:
            return AIConfig(provider, key, model, available)
    return AIConfig(None, None, None, available)


CONFIG = load_config()
