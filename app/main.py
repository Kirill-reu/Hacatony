"""FastAPI backend for the "Цифровой дизайнер презентаций" service.

Endpoints:
  POST /templates                          upload a .pptx template -> template_id + decomposed design system
  POST /generate                           start a 3-variant generation job -> job_id (async, poll for status)
  GET  /jobs/{job_id}                      job status / progress / results
  GET  /jobs/{job_id}/download/{variant}/{pptx|pdf|html}   fetch a generated file
  GET  /health

Generation runs in a background thread per job (see storage.JobStore) so the
HTTP call returns immediately — decks take up to a few minutes to build.
"""
from __future__ import annotations

import logging
import threading
import time

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from . import schemas, storage
from .config import settings
from .llm.image_client import build_image_client
from .llm.text_client import build_llm_client
from .pipeline.template_parser import DesignSystem, parse_template
from .pipeline.variant_engine import VariantOutcome, generate_variants

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("app")

app = FastAPI(
    title="Цифровой дизайнер презентаций — backend",
    description="Генерация презентаций по шаблону .pptx и краткому брифу: 3 различимых варианта, экспорт в pptx/pdf/html.",
    version="0.1.0",
)
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins or ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

storage.init_storage()

_llm = build_llm_client()
_image_client = build_image_client()

# Parsed design systems are cheap to recompute but no reason to re-parse the
# .pptx on every /generate call within the same process.
_design_systems: dict[str, DesignSystem] = {}


# --------------------------------------------------------------------------
# Templates
# --------------------------------------------------------------------------


@app.post("/templates", response_model=schemas.TemplateUploadResponse)
async def upload_template(file: UploadFile = File(...)) -> schemas.TemplateUploadResponse:
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        raise HTTPException(400, "Ожидается файл .pptx")

    template_id = storage.new_id()
    dest = storage.template_path(template_id)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(await file.read())

    try:
        design_system = parse_template(dest)
    except Exception as exc:  # noqa: BLE001
        dest.unlink(missing_ok=True)
        raise HTTPException(400, f"Не удалось разобрать шаблон: {exc}") from exc

    _design_systems[template_id] = design_system
    return schemas.TemplateUploadResponse(
        template_id=template_id,
        filename=file.filename,
        design_system=_design_system_summary(design_system),
    )


def _design_system_summary(ds: DesignSystem) -> schemas.DesignSystemSummary:
    ratio = f"{ds.slide_width / ds.slide_height:.2f}:1" if ds.slide_height else "?"
    return schemas.DesignSystemSummary(
        slide_width_emu=ds.slide_width,
        slide_height_emu=ds.slide_height,
        aspect_ratio=ratio,
        colors=[schemas.ColorToken(role=c.role, hex=c.hex) for c in ds.colors],
        major_font=ds.major_font,
        minor_font=ds.minor_font,
        type_scale_pt=ds.type_scale_pt,
        layouts=[
            schemas.LayoutSummary(index=l.index, name=l.name, role=l.role, placeholder_count=len(l.placeholders))
            for l in ds.layouts
        ],
        layout_roles=ds.layout_roles,
        warnings=ds.warnings,
    )


def _get_design_system(template_id: str) -> DesignSystem:
    ds = _design_systems.get(template_id)
    if ds is not None:
        return ds
    path = storage.template_path(template_id)
    if not path.exists():
        raise HTTPException(404, "Шаблон не найден. Сначала загрузите его через POST /templates.")
    ds = parse_template(path)
    _design_systems[template_id] = ds
    return ds


# --------------------------------------------------------------------------
# Generation jobs
# --------------------------------------------------------------------------


@app.post("/generate", response_model=schemas.JobResponse)
async def generate(req: schemas.GenerateRequest) -> schemas.JobResponse:
    design_system = _get_design_system(req.template_id)

    slide_count = req.slide_count or settings.default_slide_count
    slide_count = max(settings.min_slide_count, min(settings.max_slide_count, slide_count))

    job_id = storage.new_id()
    storage.job_store.create(
        job_id,
        {
            "job_id": job_id,
            "status": schemas.JobStatus.queued.value,
            "progress": 0.0,
            "message": "В очереди",
            "template_id": req.template_id,
            "request": req.model_dump(mode="json"),
            "variants": [],
            "error": None,
            "elapsed_seconds": 0.0,
        },
    )

    threading.Thread(target=_run_job, args=(job_id, req, design_system, slide_count), daemon=True).start()
    return _job_response(storage.job_store.get(job_id))


def _run_job(job_id: str, req: schemas.GenerateRequest, design_system: DesignSystem, slide_count: int) -> None:
    storage.job_store.update(job_id, status=schemas.JobStatus.running.value, message="Генерация…")
    started = time.monotonic()
    deadline = started + settings.max_generation_seconds

    def progress_cb(done: int, total: int, message: str) -> None:
        storage.job_store.update(job_id, progress=round(done / max(total, 1), 2), message=message)

    try:
        outcomes = generate_variants(
            template_path=storage.template_path(req.template_id),
            design_system=design_system,
            llm=_llm,
            image_client=_image_client,
            out_dir=storage.job_dir(job_id),
            brief=req.brief,
            purpose=req.purpose.value,
            slide_count=slide_count,
            audience=req.audience,
            language=req.language,
            title_for_export=req.brief[:60],
            deadline_ts=deadline,
            progress_cb=progress_cb,
        )
        elapsed = round(time.monotonic() - started, 1)
        variants = [_outcome_to_result(job_id, o).model_dump(mode="json") for o in outcomes]
        timed_out = time.monotonic() >= deadline and len(outcomes) < settings.variant_count
        status = (
            schemas.JobStatus.done.value
            if outcomes and not timed_out
            else schemas.JobStatus.timed_out.value
            if outcomes
            else schemas.JobStatus.failed.value
        )
        storage.job_store.update(
            job_id,
            status=status,
            progress=1.0,
            message="Готово" if status == schemas.JobStatus.done.value else "Успели не всё — сгенерированы не все варианты" if outcomes else "Не удалось сгенерировать ни одного варианта",
            variants=variants,
            elapsed_seconds=elapsed,
        )
    except Exception as exc:  # noqa: BLE001 - a job must fail cleanly, never crash the worker thread silently
        logger.exception("job %s failed", job_id)
        storage.job_store.update(
            job_id,
            status=schemas.JobStatus.failed.value,
            message="Ошибка генерации",
            error=str(exc),
            elapsed_seconds=round(time.monotonic() - started, 1),
        )


def _outcome_to_result(job_id: str, o: VariantOutcome) -> schemas.VariantResult:
    issues = [
        schemas.AuditIssue(
            slide_index=i.slide_index,
            category=i.category,
            code=i.code,
            message=i.message,
            deterministic=i.deterministic,
            severity=i.severity,
        )
        for i in o.audit.issues
    ]
    return schemas.VariantResult(
        variant_id=o.variant_id,
        axis=o.axis,
        axis_value=o.axis_value,
        slide_count=o.slide_count,
        pptx_file=f"/jobs/{job_id}/download/{o.variant_id}/pptx",
        pdf_file=f"/jobs/{job_id}/download/{o.variant_id}/pdf" if o.pdf_path else None,
        html_file=f"/jobs/{job_id}/download/{o.variant_id}/html",
        generation_seconds=o.generation_seconds,
        audit=schemas.VariantAudit(
            issues=issues,
            deterministic_issue_count=o.audit.deterministic_count,
            content_issue_count=o.audit.content_count,
        ),
    )


def _job_response(job: dict) -> schemas.JobResponse:
    return schemas.JobResponse(
        job_id=job["job_id"],
        status=job["status"],
        progress=job.get("progress", 0.0),
        message=job.get("message", ""),
        template_id=job.get("template_id"),
        request=job.get("request"),
        variants=job.get("variants", []),
        error=job.get("error"),
        elapsed_seconds=job.get("elapsed_seconds", 0.0),
    )


@app.get("/jobs/{job_id}", response_model=schemas.JobResponse)
async def get_job(job_id: str) -> schemas.JobResponse:
    job = storage.job_store.get(job_id)
    if job is None:
        raise HTTPException(404, "Задача не найдена")
    return _job_response(job)


_FORMAT_EXT = {"pptx": "pptx", "pdf": "pdf", "html": "html"}
_FORMAT_MEDIA = {
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "pdf": "application/pdf",
    "html": "text/html",
}


@app.get("/jobs/{job_id}/download/{variant_id}/{fmt}")
async def download(job_id: str, variant_id: str, fmt: str) -> FileResponse:
    if fmt not in _FORMAT_EXT:
        raise HTTPException(400, "Неизвестный формат. Доступно: pptx, pdf, html")
    path = storage.variant_dir(job_id, variant_id) / f"{variant_id}.{_FORMAT_EXT[fmt]}"
    if not path.exists():
        raise HTTPException(404, "Файл не найден (возможно, экспорт в этот формат недоступен для данного задания)")
    return FileResponse(path, media_type=_FORMAT_MEDIA[fmt], filename=path.name)


@app.get("/health")
async def health() -> dict:
    return {"status": "ok"}
