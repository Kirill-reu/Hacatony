"""Filesystem layout and the job registry.

Single-process, in-memory job state + files on local disk. That is the
right amount of infrastructure for a hackathon demo. If this needs to
survive process restarts or run behind more than one worker, swap
``JobStore`` for Redis/Postgres and ``data/`` for S3-compatible storage —
every call site goes through this module, so that's a localized change.
"""
from __future__ import annotations

import threading
import uuid
from pathlib import Path
from typing import Any, Optional

from .config import settings

TEMPLATES_DIR = settings.data_dir / "templates"
JOBS_DIR = settings.data_dir / "jobs"


def init_storage() -> None:
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    JOBS_DIR.mkdir(parents=True, exist_ok=True)


def new_id() -> str:
    return uuid.uuid4().hex[:12]


def template_path(template_id: str) -> Path:
    return TEMPLATES_DIR / f"{template_id}.pptx"


def job_dir(job_id: str) -> Path:
    d = JOBS_DIR / job_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def variant_dir(job_id: str, variant_id: str) -> Path:
    d = job_dir(job_id) / variant_id
    d.mkdir(parents=True, exist_ok=True)
    return d


class JobStore:
    """Thread-safe in-memory job registry keyed by job_id."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._jobs: dict[str, dict[str, Any]] = {}

    def create(self, job_id: str, payload: dict[str, Any]) -> None:
        with self._lock:
            self._jobs[job_id] = payload

    def update(self, job_id: str, **kwargs: Any) -> None:
        with self._lock:
            if job_id in self._jobs:
                self._jobs[job_id].update(kwargs)

    def append_variant(self, job_id: str, variant: dict[str, Any]) -> None:
        with self._lock:
            self._jobs[job_id].setdefault("variants", []).append(variant)

    def get(self, job_id: str) -> Optional[dict[str, Any]]:
        with self._lock:
            job = self._jobs.get(job_id)
            return dict(job) if job is not None else None


job_store = JobStore()
