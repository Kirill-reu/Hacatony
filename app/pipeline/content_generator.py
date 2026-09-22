"""ТЗ §2 item 2: "Генерация структуры и текстового наполнения будущей
презентации по краткому брифу и назначению" — this module turns
(brief, purpose, slide_count, grounded facts, available template layouts)
into a ``SlidePlan``: an ordered list of slides with a title, bullets, a
speaker note, and a requested visual (chart/table/image/none), each tagged
with the *layout_role* (see template_parser) it should be built on.

Three things make the 3 required output variants differ (ТЗ: "Ось различий
команда определяет сама"), all expressed as a ``VariantParams`` the caller
supplies: content density (bullet count/length), wording/tone, and which
visual types + layout roles are preferred. slide_builder later resolves
layout_role -> an actual layout from the *same* template — the template
itself never changes between variants.

If the LLM is unavailable or returns something unusable, ``_fallback_plan``
builds a plan deterministically from the brief and facts so the pipeline
never dead-ends without an API key configured.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

from ..llm.text_client import LLMClient
from .data_provider import DataProvider, SearchResult, gather_facts
from .template_parser import DesignSystem

logger = logging.getLogger(__name__)

_NUMBER_RE = re.compile(r"\d[\d\s.,]*\d|\d")


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------


@dataclass
class VisualSpec:
    type: str = "none"          # none | chart | table | image | icon_row
    chart_type: str = "bar"     # bar | line | pie (only used when type == "chart")
    hint: str = ""              # what data/image is needed, for visual_generator


@dataclass
class SlidePlanItem:
    layout_role: str
    title: str
    bullets: list[str] = field(default_factory=list)
    speaker_note: str = ""
    visual: VisualSpec = field(default_factory=VisualSpec)


@dataclass
class SlidePlan:
    slides: list[SlidePlanItem] = field(default_factory=list)


@dataclass
class VariantParams:
    variant_id: str
    axis: str                       # human label for what differs, e.g. "content_density"
    axis_value: str                 # e.g. "concise" | "detailed" | "balanced"
    max_bullets: int
    max_bullet_words: int
    tone_instruction: str
    preferred_visuals: tuple[str, ...]   # ordered preference, e.g. ("chart", "table", "image")
    prefer_multi_column_layouts: bool
    temperature: float = 0.7


# --------------------------------------------------------------------------
# Prompting
# --------------------------------------------------------------------------

_PURPOSE_RU = {
    "feature": "новая функциональность (фича)",
    "product": "продукт",
    "project": "проект",
    "initiative": "инициатива",
}


def _system_prompt(language: str) -> str:
    if language.startswith("ru"):
        return (
            "Ты — ассистент, который планирует структуру корпоративной презентации. "
            "Отвечай ТОЛЬКО валидным JSON без пояснений и без markdown-разметки. "
            "Используй в качестве фактов и цифр ТОЛЬКО то, что дано в разделе 'Факты' — "
            "никогда не выдумывай статистику. Если фактов не хватает, пиши качественно, без цифр."
        )
    return (
        "You plan the structure of a corporate slide deck. Respond ONLY with valid JSON, "
        "no prose, no markdown. Use numbers/statistics ONLY from the 'Facts' section — never invent data. "
        "If facts are insufficient, write qualitatively, without numbers."
    )


def _user_prompt(
    brief: str,
    purpose: str,
    slide_count: int,
    audience: str | None,
    facts: list[SearchResult],
    available_roles: list[str],
    variant: VariantParams,
    language: str,
) -> str:
    facts_block = "\n".join(f"- {f.title}: {f.snippet}" for f in facts[:12]) or "(факты не найдены — опирайся только на бриф)"
    roles_block = ", ".join(available_roles)
    purpose_label = _PURPOSE_RU.get(purpose, purpose)
    schema = (
        '{"slides": [{"layout_role": "<одна из: ' + roles_block + '>", '
        '"title": "...", "bullets": ["...", "..."], "speaker_note": "...", '
        '"visual": {"type": "none|chart|table|image|icon_row", "chart_type": "bar|line|pie", "hint": "..."}}]}'
    )
    return (
        f"Назначение: {purpose_label}.\n"
        f"Бриф: {brief}\n"
        f"Аудитория: {audience or 'не указана'}\n"
        f"Язык слайдов: {language}\n"
        f"Нужно слайдов: {slide_count} (первый — титульный, последний — выводы/следующие шаги).\n\n"
        f"Факты (используй как источник цифр и утверждений):\n{facts_block}\n\n"
        f"Требования к этому варианту презентации:\n"
        f"- не больше {variant.max_bullets} буллетов на слайде, каждый буллет — не больше {variant.max_bullet_words} слов;\n"
        f"- тон и формулировки: {variant.tone_instruction};\n"
        f"- для визуализации данных предпочитай в таком порядке: {', '.join(variant.preferred_visuals)};\n"
        f"- каждый слайд должен использовать одну из доступных ролей макета: {roles_block}.\n\n"
        f"Верни JSON строго по схеме:\n{schema}"
    )


# --------------------------------------------------------------------------
# LLM-backed generation with a deterministic fallback
# --------------------------------------------------------------------------


def generate_plan(
    llm: LLMClient,
    design_system: DesignSystem,
    provider: DataProvider,
    *,
    brief: str,
    purpose: str,
    slide_count: int,
    audience: str | None,
    language: str,
    variant: VariantParams,
) -> tuple[SlidePlan, list[SearchResult]]:
    facts = gather_facts(provider, _search_topics(brief, language))
    available_roles = sorted(design_system.layout_roles.keys()) or ["generic"]

    plan: SlidePlan | None = None
    try:
        raw = llm.complete_json(
            _system_prompt(language),
            _user_prompt(brief, purpose, slide_count, audience, facts, available_roles, variant, language),
            temperature=variant.temperature,
        )
        plan = _parse_llm_plan(raw, available_roles)
    except Exception as exc:  # noqa: BLE001 - any LLM/parsing failure falls back, it must never crash the job
        logger.warning("content_generator: LLM plan generation failed (%s), using fallback", exc)

    if not plan or len(plan.slides) < 3:
        plan = _fallback_plan(brief, purpose, slide_count, facts, available_roles, variant)

    plan = _enforce_variant_limits(plan, variant)
    plan = _fit_slide_count(plan, slide_count)
    return plan, facts


def _search_topics(brief: str, language: str) -> list[str]:
    topics = [brief]
    suffix = "статистика данные 2026" if language.startswith("ru") else "statistics data 2026"
    topics.append(f"{brief} {suffix}")
    return topics


def _parse_llm_plan(raw: dict, available_roles: list[str]) -> SlidePlan | None:
    slides_raw = raw.get("slides") if isinstance(raw, dict) else None
    if not isinstance(slides_raw, list) or not slides_raw:
        return None

    items: list[SlidePlanItem] = []
    for s in slides_raw:
        if not isinstance(s, dict) or not s.get("title"):
            continue
        role = s.get("layout_role") if s.get("layout_role") in available_roles else _closest_role(available_roles)
        visual_raw = s.get("visual") or {}
        visual = VisualSpec(
            type=visual_raw.get("type", "none") if isinstance(visual_raw, dict) else "none",
            chart_type=visual_raw.get("chart_type", "bar") if isinstance(visual_raw, dict) else "bar",
            hint=visual_raw.get("hint", "") if isinstance(visual_raw, dict) else "",
        )
        bullets = [str(b).strip() for b in s.get("bullets", []) if str(b).strip()]
        items.append(
            SlidePlanItem(
                layout_role=role,
                title=str(s["title"]).strip(),
                bullets=bullets,
                speaker_note=str(s.get("speaker_note", "")).strip(),
                visual=visual,
            )
        )
    return SlidePlan(slides=items) if items else None


def _closest_role(available_roles: list[str]) -> str:
    for preferred in ("content_bullets", "generic", "two_content"):
        if preferred in available_roles:
            return preferred
    return available_roles[0]


# --------------------------------------------------------------------------
# Deterministic fallback (no LLM configured, or the model misbehaved)
# --------------------------------------------------------------------------


def _split_sentences(text: str) -> list[str]:
    parts = re.split(r"(?<=[.!?])\s+", text.strip())
    return [p.strip() for p in parts if p.strip()]


def _fallback_plan(
    brief: str,
    purpose: str,
    slide_count: int,
    facts: list[SearchResult],
    available_roles: list[str],
    variant: VariantParams,
) -> SlidePlan:
    logger.info("content_generator: building fallback (rule-based) plan for variant=%s", variant.variant_id)
    purpose_label = _PURPOSE_RU.get(purpose, purpose)
    sentences = _split_sentences(brief) or [brief]

    slides: list[SlidePlanItem] = []

    title_role = "title_slide" if "title_slide" in available_roles else _closest_role(available_roles)
    slides.append(
        SlidePlanItem(
            layout_role=title_role,
            title=sentences[0][:90],
            bullets=[purpose_label.capitalize()],
            speaker_note="Титульный слайд.",
        )
    )

    body_role = "content_bullets" if "content_bullets" in available_roles else _closest_role(available_roles)
    two_col_role = "two_content" if "two_content" in available_roles else body_role
    image_role = "content_image" if "content_image" in available_roles else body_role

    n_content = max(slide_count - 2, 1)
    chunk_size = max(len(sentences) // n_content, 1)
    fact_pool = list(facts)

    section_titles_ru = ["Проблема", "Подход", "Как это работает", "Данные", "Риски и ограничения", "Дорожная карта"]

    for i in range(n_content):
        chunk = sentences[i * chunk_size : (i + 1) * chunk_size] or [f"{purpose_label.capitalize()}: пункт {i + 1}"]
        bullets = chunk[: variant.max_bullets]
        role = two_col_role if (variant.prefer_multi_column_layouts and i % 3 == 1) else body_role

        visual = VisualSpec(type="none")
        if fact_pool and any(v in variant.preferred_visuals for v in ("chart", "table")):
            f = fact_pool.pop(0)
            numbers = _NUMBER_RE.findall(f.snippet)
            if numbers and "chart" in variant.preferred_visuals:
                visual = VisualSpec(type="chart", chart_type="bar", hint=f.snippet[:160])
            elif "table" in variant.preferred_visuals:
                visual = VisualSpec(type="table", hint=f.snippet[:160])
            bullets = bullets + [f"{f.title}: {f.snippet[:120]}"]
        elif "image" in variant.preferred_visuals and i == 0 and image_role != body_role:
            role = image_role
            visual = VisualSpec(type="image", hint=f"{purpose_label} — {sentences[0][:80]}")

        title = section_titles_ru[i] if i < len(section_titles_ru) else f"{purpose_label.capitalize()}: часть {i + 1}"
        slides.append(
            SlidePlanItem(
                layout_role=role,
                title=title,
                bullets=bullets[: variant.max_bullets],
                speaker_note=chunk[0] if chunk else "",
                visual=visual,
            )
        )

    closing_role = "section_header" if "section_header" in available_roles else body_role
    slides.append(
        SlidePlanItem(
            layout_role=closing_role,
            title="Выводы и следующие шаги",
            bullets=["Обсудить план внедрения", "Собрать обратную связь", "Назначить ответственных"][: variant.max_bullets],
            speaker_note="Закрывающий слайд.",
        )
    )
    return SlidePlan(slides=slides)


# --------------------------------------------------------------------------
# Post-processing shared by both the LLM and fallback paths
# --------------------------------------------------------------------------


def _enforce_variant_limits(plan: SlidePlan, variant: VariantParams) -> SlidePlan:
    for slide in plan.slides:
        slide.bullets = slide.bullets[: variant.max_bullets]
        trimmed = []
        for b in slide.bullets:
            words = b.split()
            trimmed.append(" ".join(words[: variant.max_bullet_words]))
        slide.bullets = trimmed
    return plan


def _fit_slide_count(plan: SlidePlan, slide_count: int) -> SlidePlan:
    if len(plan.slides) > slide_count:
        # Keep the first (title) and last (closing) slides, trim from the middle.
        head, tail = plan.slides[0], plan.slides[-1]
        middle = plan.slides[1:-1][: max(slide_count - 2, 0)]
        plan.slides = [head, *middle, tail]
    return plan
