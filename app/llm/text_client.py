"""Pluggable text-generation backend.

``OpenAICompatibleClient`` talks to any server implementing the OpenAI
``/chat/completions`` schema — that covers VK's inference gateway, a
self-hosted Qwen server behind vLLM/TGI, or OpenAI itself. Point
``LLM_BASE_URL`` / ``LLM_API_KEY`` / ``LLM_MODEL`` at whatever the team
is using (see .env.example); nothing else in the pipeline needs to change.

``OfflineStubClient`` is the network-free fallback used when no LLM is
configured, so the whole pipeline stays runnable for local dev / CI /
this repo's test suite. Callers (``content_generator``, ``audit``) must
treat its output as low quality and are expected to have their own
rule-based fallback for when it (or a real but misbehaving model)
doesn't return usable JSON — see ``content_generator._fallback_plan``.
"""
from __future__ import annotations

import json
import logging
from abc import ABC, abstractmethod
from typing import Any

from ..config import settings

logger = logging.getLogger(__name__)


class LLMClient(ABC):
    @abstractmethod
    def complete_json(self, system: str, user: str, *, temperature: float = 0.7) -> dict[str, Any]:
        """Ask the model for a single JSON object and parse it.

        Must raise ``ValueError`` on malformed/non-JSON output rather than
        guessing — callers decide how to fall back.
        """

    @abstractmethod
    def complete_text(self, system: str, user: str, *, temperature: float = 0.7) -> str:
        ...


class OpenAICompatibleClient(LLMClient):
    """Talks to the endpoint with a plain HTTP POST (via ``requests``)
    rather than the ``openai`` SDK. Not a style choice: as of mid-2026 the
    SDK's own HTTP client trips Cloudflare's bot-fingerprint block on at
    least one popular OpenAI-compatible gateway (OpenRouter) — the exact
    same request succeeds via curl or ``requests`` and fails only through
    the SDK, which points at the TLS ClientHello fingerprint of its HTTP
    stack rather than anything about the account, key, or request itself
    (see https://github.com/SillyTavern/SillyTavern/issues/5825 for the
    same failure mode from a different language's HTTP client). Plain
    ``requests`` sidesteps that without switching providers."""

    def __init__(self, base_url: str, api_key: str | None, model: str, timeout: int):
        self._base_url = base_url.rstrip("/")
        self._api_key = api_key
        self._model = model
        self._timeout = timeout

    def _chat(self, system: str, user: str, temperature: float, json_mode: bool) -> str:
        import requests

        headers = {"Content-Type": "application/json"}
        if self._api_key:
            headers["Authorization"] = f"Bearer {self._api_key}"
        payload: dict[str, Any] = {
            "model": self._model,
            "temperature": temperature,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        }
        if json_mode:
            payload["response_format"] = {"type": "json_object"}

        resp = requests.post(f"{self._base_url}/chat/completions", headers=headers, json=payload, timeout=self._timeout)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"] or ""

    def complete_json(self, system: str, user: str, *, temperature: float = 0.7) -> dict[str, Any]:
        raw = self._chat(system, user, temperature, json_mode=True)
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError(f"LLM did not return valid JSON: {exc}\nRaw (truncated): {raw[:500]}") from exc

    def complete_text(self, system: str, user: str, *, temperature: float = 0.7) -> str:
        return self._chat(system, user, temperature, json_mode=False).strip()


class OfflineStubClient(LLMClient):
    """No network calls. Echoes trivial content so callers' fallback paths
    (which are the real content source in this mode) kick in deterministically."""

    def complete_json(self, system: str, user: str, *, temperature: float = 0.7) -> dict[str, Any]:
        logger.warning("OfflineStubClient: no LLM_BASE_URL configured — returning an empty stub, caller should fall back")
        return {}

    def complete_text(self, system: str, user: str, *, temperature: float = 0.7) -> str:
        logger.warning("OfflineStubClient: no LLM_BASE_URL configured — returning a trivial echo")
        first_line = next((line.strip() for line in user.splitlines() if line.strip()), "")
        return first_line[:200]


def build_llm_client() -> LLMClient:
    if settings.llm_base_url:
        return OpenAICompatibleClient(
            base_url=settings.llm_base_url,
            api_key=settings.llm_api_key,
            model=settings.llm_model,
            timeout=settings.llm_timeout_seconds,
        )
    return OfflineStubClient()
