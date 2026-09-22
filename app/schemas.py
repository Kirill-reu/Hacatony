"""Pydantic models exchanged over the HTTP API. Internal pipeline data
structures (design system, slide plan, ...) live in app/pipeline and are
deliberately kept separate from these — the API contract should be able
to evolve without dragging the whole pipeline's internals with it.
"""
from __future__ import annotations

from enum import Enum
from typing import Any, Optional

from pydantic import BaseModel, Field


class Purpose(str, Enum):
    feature = "feature"       # фича
    product = "product"       # продукт
    project = "project"       # проект
    initiative = "initiative"  # инициатива


class ColorToken(BaseModel):
    role: str
    hex: str


class LayoutSummary(BaseModel):
    index: int
    name: str
    role: str
    placeholder_count: int


class DesignSystemSummary(BaseModel):
    """What we hand back to the client after decomposing an uploaded template —
    the full DesignSystem (app.pipeline.template_parser) has more detail and is
    reused internally, but this is the human-readable subset."""

    slide_width_emu: int
    slide_height_emu: int
    aspect_ratio: str
    colors: list[ColorToken]
    major_font: str
    minor_font: str
    type_scale_pt: list[float]
    layouts: list[LayoutSummary]
    layout_roles: dict[str, list[int]]
    warnings: list[str] = Field(default_factory=list)


class TemplateUploadResponse(BaseModel):
    template_id: str
    filename: str
    design_system: DesignSystemSummary


class GenerateRequest(BaseModel):
    template_id: str
    brief: str = Field(..., min_length=10, description="Краткий бриф: что за фича/продукт/проект/инициатива и что должно быть в презентации")
    purpose: Purpose = Purpose.feature
    slide_count: Optional[int] = Field(None, ge=4, le=30, description="Если не задано — сервис выбирает сам в диапазоне 10-15")
    language: str = "ru"
    audience: Optional[str] = Field(None, description="Например: 'совет директоров', 'инженерная команда'")


class JobStatus(str, Enum):
    queued = "queued"
    running = "running"
    done = "done"
    failed = "failed"
    timed_out = "timed_out"


class AuditIssue(BaseModel):
    slide_index: int
    category: str          # verstka | template | density | integrity | content
    code: str
    message: str
    deterministic: bool
    severity: str = "warning"  # info | warning | error


class VariantAudit(BaseModel):
    issues: list[AuditIssue] = Field(default_factory=list)
    deterministic_issue_count: int = 0
    content_issue_count: int = 0


class VariantResult(BaseModel):
    variant_id: str
    axis: str               # what makes this variant distinct, e.g. "content_density"
    axis_value: str         # e.g. "concise" | "detailed" | "balanced"
    slide_count: int
    pptx_file: str
    pdf_file: Optional[str] = None
    html_file: str
    generation_seconds: float
    audit: VariantAudit


class JobResponse(BaseModel):
    job_id: str
    status: JobStatus
    progress: float = 0.0
    message: str = ""
    template_id: Optional[str] = None
    request: Optional[dict[str, Any]] = None
    variants: list[VariantResult] = Field(default_factory=list)
    error: Optional[str] = None
    elapsed_seconds: float = 0.0
