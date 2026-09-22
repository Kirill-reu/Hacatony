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
    def __init__(self, base_url: str, api_key: str | None, model: str, timeout: int):
        from openai import OpenAI  # lazy import: keeps the offline path dependency-free

        self._client = OpenAI(base_url=base_url, api_key=api_key or "unused", timeout=timeout)
        self._model = model

    def _chat(self, system: str, user: str, temperature: float, json_mode: bool) -> str:
        kwargs: dict[str, Any] = {}
        if json_mode:
            kwargs["response_format"] = {"type": "json_object"}
        resp = self._client.chat.completions.create(
            model=self._model,
            temperature=temperature,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            **kwargs,
        )
        return resp.choices[0].message.content or ""

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
