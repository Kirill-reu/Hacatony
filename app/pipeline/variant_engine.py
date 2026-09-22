"""ТЗ §2 item 5: "Представление результата генерации в трёх вариантах
верстки для одного и того же шаблона и контента." The axis this
implementation picked (documented per the ТЗ's own "обосновывает свой
выбор" requirement):

  1. ``concise``   — content density axis: few, short bullets; leans on
                      charts/icon-rows over prose; prefers single-column layouts.
  2. ``detailed``  — content density axis (the other end): fuller bullets,
                      a more narrative/explanatory tone; leans on tables and
                      two-column/comparison layouts where the template has them.
  3. ``visual``    — visualization axis: same content budget as a middle
                      ground between the two above, but actively prefers
                      pulling in images and icon-row diagrams over plain text,
                      and prefers whichever multi-column layout the template
                      offers so slides *read* differently even on identical input.

All three are built from the same template (no layout is invented, only
selected from what template_parser found) and the same retrieved facts, so
differences come only from the choices ТЗ allows: wording, density, layout
selection, and visualization approach.
"""
from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from pathlib import Path

from ..config import settings
from ..llm.image_client import ImageClient
from ..llm.text_client import LLMClient
from . import audit as audit_mod
from . import content_generator as cg
from . import exporters
from .slide_builder import build_variant_deck
from .template_parser import DesignSystem

logger = logging.getLogger(__name__)


VARIANT_DEFINITIONS: list[cg.VariantParams] = [
    cg.VariantParams(
        variant_id="concise",
        axis="content_density_and_layout",
        axis_value="concise",
        max_bullets=3,
        max_bullet_words=8,
        tone_instruction="кратко, по делу, формулировки-выводы (не длиннее заголовка новости)",
        preferred_visuals=("chart", "icon_row", "image"),
        prefer_multi_column_layouts=False,
        temperature=0.5,
    ),
    cg.VariantParams(
        variant_id="detailed",
        axis="content_density_and_layout",
        axis_value="detailed",
        max_bullets=6,
        max_bullet_words=15,
        tone_instruction="подробно и повествовательно, с пояснением контекста и причин",
        preferred_visuals=("table", "chart"),
        prefer_multi_column_layouts=True,
        temperature=0.7,
    ),
    cg.VariantParams(
        variant_id="visual",
        axis="visualization_approach",
        axis_value="visual_first",
        max_bullets=4,
        max_bullet_words=10,
        tone_instruction="нейтрально и наглядно, текст поддерживает визуализацию, а не наоборот",
        preferred_visuals=("image", "icon_row", "chart", "table"),
        prefer_multi_column_layouts=True,
        temperature=0.8,
    ),
]


@dataclass
class VariantOutcome:
    variant_id: str
    axis: str
    axis_value: str
    slide_count: int
    pptx_path: Path
    pdf_path: Path | None
    html_path: Path
    generation_seconds: float
    audit: audit_mod.AuditReport


def generate_variants(
    *,
    template_path: str | Path,
    design_system: DesignSystem,
    llm: LLMClient,
    image_client: ImageClient,
    out_dir: str | Path,
    brief: str,
    purpose: str,
    slide_count: int,
    audience: str | None,
    language: str,
    title_for_export: str,
    deadline_ts: float,
    progress_cb=None,
) -> list[VariantOutcome]:
    """Builds every configured variant, stopping gracefully (returning what
    finished) if ``deadline_ts`` (a ``time.monotonic()`` value) is reached —
    the ТЗ's 5-minute-per-deck budget is a *per generation job* budget, and
    with three variants that means each one gets a slice of it, not 5
    minutes each."""
    out_dir = Path(out_dir)
    outcomes: list[VariantOutcome] = []
    variants = VARIANT_DEFINITIONS[: settings.variant_count]

    for i, variant in enumerate(variants):
        remaining = deadline_ts - time.monotonic()
        if remaining <= 5:
            logger.warning("variant_engine: time budget exhausted, stopping after %d/%d variants", i, len(variants))
            break

        started = time.monotonic()
        if progress_cb:
            progress_cb(i, len(variants), f"Генерация варианта «{variant.axis_value}»…")

        plan, facts = cg.generate_plan(
            llm,
            design_system,
            _provider_for(),
            brief=brief,
            purpose=purpose,
            slide_count=slide_count,
            audience=audience,
            language=language,
            variant=variant,
        )

        prs = build_variant_deck(template_path, design_system, plan, facts, llm=llm, image_client=image_client, language=language)

        variant_dir = out_dir / variant.variant_id
        pptx_path = exporters.save_pptx(prs, variant_dir / f"{variant.variant_id}.pptx")
        pdf_path = exporters.export_pdf(pptx_path, variant_dir)
        html_path = exporters.export_html(prs, design_system, variant_dir / f"{variant.variant_id}.html", title=title_for_export, language=language)

        facts_text = "\n".join(f"{f.title}: {f.snippet}" for f in facts)
        report = audit_mod.run_full_audit(llm, str(pptx_path), design_system, facts_text)

        outcomes.append(
            VariantOutcome(
                variant_id=variant.variant_id,
                axis=variant.axis,
                axis_value=variant.axis_value,
                slide_count=len(plan.slides),
                pptx_path=pptx_path,
                pdf_path=pdf_path,
                html_path=html_path,
                generation_seconds=round(time.monotonic() - started, 1),
                audit=report,
            )
        )
        if progress_cb:
            progress_cb(i + 1, len(variants), f"Вариант «{variant.axis_value}» готов")

    return outcomes


def _provider_for():
    # Imported lazily to avoid a hard import-time dependency cycle and to make
    # it trivial to swap providers per call site/tests without touching this module.
    from .data_provider import build_data_provider

    return build_data_provider()
