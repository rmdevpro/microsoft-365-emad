"""
microsoft_365_emad.inference — LLM resolution.

Reads LLM config from the host's TE config (/config/te.yml) or
from config injected via set_config(). Falls back to environment-based
Gemini endpoint if no config is available.
"""

import hashlib
import logging
import os
import threading
from pathlib import Path

import yaml
from langchain_openai import ChatOpenAI

_log = logging.getLogger("microsoft_365_emad")

# Host TE config (mounted in container)
_TE_CONFIG_PATH = Path("/config/te.yml")

_cache_lock = threading.Lock()
_llm_cache: dict[str, ChatOpenAI] = {}
_current_config: dict | None = None


def set_config(config: dict) -> None:
    global _current_config
    _current_config = config


def _get_config() -> dict:
    if _current_config is not None:
        return _current_config
    if _TE_CONFIG_PATH.exists():
        return yaml.safe_load(_TE_CONFIG_PATH.read_text(encoding="utf-8")) or {}
    return {}


def get_llm(role: str = "fast") -> ChatOpenAI:
    """Get a ChatOpenAI instance for the given role."""
    config = _get_config()
    provider_config = config.get(role, {})
    if not provider_config:
        provider_config = config.get("imperator", {})
    if not provider_config:
        # Fall back to Gemini via environment
        provider_config = {
            "base_url": "https://generativelanguage.googleapis.com/v1beta/openai",
            "model": "gemini-2.5-flash",
            "api_key_env": "GOOGLE_API_KEY",
        }

    base_url = provider_config.get("base_url")
    model = provider_config.get("model", "gpt-4o-mini")
    api_key_env = provider_config.get("api_key_env", "")
    api_key = os.environ.get(api_key_env) if api_key_env else "not-needed"
    temperature = provider_config.get("temperature")
    max_tokens = provider_config.get("max_tokens")

    cache_key = (
        f"{role}:{base_url}:{model}:"
        f"{hashlib.sha256((api_key or 'none').encode()).hexdigest()[:16]}"
    )

    with _cache_lock:
        if cache_key not in _llm_cache:
            kwargs = {
                "base_url": base_url,
                "model": model,
                "api_key": api_key or "not-needed",
            }
            if temperature is not None:
                kwargs["temperature"] = temperature
            if max_tokens is not None:
                kwargs["max_tokens"] = max_tokens
            _llm_cache[cache_key] = ChatOpenAI(**kwargs)
        return _llm_cache[cache_key]
