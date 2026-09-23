"""Central configuration, read from environment variables.

Everything here has a sane default so the service boots and runs
end-to-end (with the offline stub LLM / placeholder image generator)
without any secrets configured. Set the LLM_* / IMAGE_* variables to
point at real inference before a demo — see .env.example.
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def _bool(name: str, default: str) -> bool:
    return os.getenv(name, default).strip().lower() in {"1", "true", "yes", "on"}


@dataclass
class Settings:
    # --- storage -----------------------------------------------------
    data_dir: Path = Path(os.getenv("APP_DATA_DIR", "./data"))

    # --- generation limits --------------------------------------------
    # Hard cap from the brief is 5 minutes per deck; we budget under that
    # so slow variants still finish and we can package results.
    max_generation_seconds: int = int(os.getenv("MAX_GENERATION_SECONDS", "280"))
    default_slide_count: int = int(os.getenv("DEFAULT_SLIDE_COUNT", "12"))
    min_slide_count: int = int(os.getenv("MIN_SLIDE_COUNT", "6"))
    max_slide_count: int = int(os.getenv("MAX_SLIDE_COUNT", "20"))
    variant_count: int = int(os.getenv("VARIANT_COUNT", "3"))

    # --- text LLM -------------------------------------------------------
    # Any OpenAI-compatible /chat/completions endpoint: VK inference
    # gateway, a self-hosted Qwen server, vLLM, OpenAI itself, etc.
    llm_base_url: str | None = os.getenv("LLM_BASE_URL") or None
    llm_api_key: str | None = os.getenv("LLM_API_KEY") or None
    llm_model: str = os.getenv("LLM_MODEL", "qwen3-27b-instruct")
    llm_timeout_seconds: int = int(os.getenv("LLM_TIMEOUT_SECONDS", "60"))

    # --- text-to-image ----------------------------------------------
    # OpenAI-compatible /images/generations endpoint. Open-weights,
    # Apache-2.0/MIT, <=35B (<=20B for the qualifying round) per the ТЗ.
    image_base_url: str | None = os.getenv("IMAGE_BASE_URL") or None
    image_api_key: str | None = os.getenv("IMAGE_API_KEY") or None
    image_model: str = os.getenv("IMAGE_MODEL", "sdxl-base-1.0")

    # --- web data provider ---------------------------------------------
    web_search_enabled: bool = _bool("WEB_SEARCH_ENABLED", "true")
    web_search_max_results: int = int(os.getenv("WEB_SEARCH_MAX_RESULTS", "5"))
    # DuckDuckGo region code (see https://duckduckgo.com/params), used for
    # both fact search (per-request, from the deck's own `language`) and
    # image search (built once at process start, so it can't see a
    # per-request language — this is its fixed default).
    web_search_region: str = os.getenv("WEB_SEARCH_REGION", "ru-ru")

    # --- export -----------------------------------------------------
    soffice_binary: str = os.getenv("SOFFICE_BINARY", "soffice")
    pdf_export_enabled: bool = _bool("PDF_EXPORT_ENABLED", "true")
    pdf_export_timeout_seconds: int = int(os.getenv("PDF_EXPORT_TIMEOUT_SECONDS", "60"))

    # --- api ----------------------------------------------------------
    cors_origins: list[str] | None = None

    def __post_init__(self) -> None:
        self.data_dir = Path(self.data_dir)
        raw = os.getenv("CORS_ORIGINS", "*")
        self.cors_origins = [o.strip() for o in raw.split(",")] if raw else ["*"]


settings = Settings()
