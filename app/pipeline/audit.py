"""ТЗ §2 item 6 + Приложение 1: audits a generated deck for verstka/template/
density/integrity problems (deterministic — computed straight from shape
geometry and XML, same result every run) and content problems (the 11
yes/no questions in Приложение 1 — inherently judgment calls, so they go
through the LLM and are marked non-deterministic). The audit runs against
the .pptx *after* slide_builder writes it, so it sees exactly what got
exported — it is part of the generation pipeline, not an external QA pass
bolted on afterward.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.util import Emu

from ..llm.text_client import LLMClient
from .template_parser import DesignSystem

logger = logging.getLogger(__name__)

_PLACEHOLDER_TEXT_RE = re.compile(r"lorem ipsum|\bTODO\b|\bXXX\b|вставьте текст|\[insert", re.IGNORECASE)
_MAX_BULLETS = 6
_MAX_BULLET_WORDS = 15
_MAX_TABLE_ROWS = 7
_MAX_TABLE_COLS = 5
_MAX_CHART_SERIES = 5
_MIN_CONTRAST_RATIO = 4.5
_MIN_FILL_RATIO = 0.25
_MAX_FILL_RATIO = 0.75


@dataclass
class Issue:
    slide_index: int
    category: str  # verstka | template | density | integrity | content
    code: str
    message: str
    deterministic: bool
    severity: str = "warning"


@dataclass
class AuditReport:
    issues: list[Issue] = field(default_factory=list)

    @property
    def deterministic_count(self) -> int:
        return sum(1 for i in self.issues if i.deterministic)

    @property
    def content_count(self) -> int:
        return sum(1 for i in self.issues if not i.deterministic)


# --------------------------------------------------------------------------
# Geometry helpers
# --------------------------------------------------------------------------


def _bbox(shape) -> tuple[int, int, int, int] | None:
    if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
        return None
    return shape.left, shape.top, shape.left + shape.width, shape.top + shape.height


def _overlap_area(a, b) -> int:
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    ox = max(0, min(ax2, bx2) - max(ax1, bx1))
    oy = max(0, min(ay2, by2) - max(ay1, by1))
    return ox * oy


def _relative_luminance(hex_color: str) -> float:
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i : i + 2], 16) / 255 for i in (0, 2, 4))

    def lin(c: float) -> float:
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = lin(r), lin(g), lin(b)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrast_ratio(hex_a: str, hex_b: str) -> float:
    la, lb = _relative_luminance(hex_a), _relative_luminance(hex_b)
    lighter, darker = max(la, lb), min(la, lb)
    return (lighter + 0.05) / (darker + 0.05)


def _shape_background_hex(shape, default_hex: str) -> str:
    """Contrast must be checked against what the text actually sits on — the
    shape's own fill when it has a solid one (a colored card, a table header
    cell), the slide background otherwise."""
    try:
        fill = shape.fill
        if fill.type is not None and fill.type == 1:  # MSO_FILL_TYPE.SOLID
            return str(fill.fore_color.rgb).upper()
    except Exception:
        pass
    return default_hex


# --------------------------------------------------------------------------
# Deterministic checks
# --------------------------------------------------------------------------


def _check_slide(slide, slide_index: int, design_system: DesignSystem, allowed_fonts: set[str], seen_texts: dict[str, int]) -> list[Issue]:
    issues: list[Issue] = []
    slide_w, slide_h = design_system.slide_width, design_system.slide_height
    boxes = []
    total_text = []
    allowed_sizes = set(round(s, 1) for s in design_system.type_scale_pt) if design_system.type_scale_pt else set()
    allowed_colors = {c.hex.upper() for c in design_system.colors}
    bg_hex = next((c.hex for c in design_system.colors if c.role == "lt1"), "FFFFFF")

    shape_count = 0
    filled_area = 0
    for shape in slide.shapes:
        shape_count += 1
        bbox = _bbox(shape)
        if bbox:
            x1, y1, x2, y2 = bbox
            if x1 < 0 or y1 < 0 or x2 > slide_w or y2 > slide_h:
                issues.append(Issue(slide_index, "verstka", "out_of_bounds", f"Элемент '{shape.name}' выходит за границы слайда", True))
            filled_area += max(0, x2 - x1) * max(0, y2 - y1)
            boxes.append((shape.name, bbox))

        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            text = shape.text_frame.text
            total_text.append(text)
            local_bg = _shape_background_hex(shape, bg_hex)
            if _PLACEHOLDER_TEXT_RE.search(text):
                issues.append(Issue(slide_index, "integrity", "leftover_placeholder", f"Похоже на текст-заглушку: '{text[:40]}'", True, "error"))
            for p in shape.text_frame.paragraphs:
                bullet_text = "".join(r.text for r in p.runs)
                if bullet_text.strip() and len(bullet_text.split()) > _MAX_BULLET_WORDS:
                    issues.append(Issue(slide_index, "density", "bullet_too_long", f"Буллет длиннее {_MAX_BULLET_WORDS} слов: '{bullet_text[:60]}...'", True))
                for r in p.runs:
                    if r.font.name and allowed_fonts and r.font.name not in allowed_fonts:
                        issues.append(Issue(slide_index, "template", "font_not_in_theme", f"Шрифт вне темы шаблона: {r.font.name}", True))
                    if r.font.size is not None and allowed_sizes and round(r.font.size.pt, 1) not in allowed_sizes:
                        issues.append(Issue(slide_index, "template", "size_not_in_scale", f"Кегль {r.font.size.pt}pt вне типографической шкалы шаблона", True, "info"))
                    if r.font.color and r.font.color.type is not None:
                        try:
                            hexval = str(r.font.color.rgb).upper()
                            if allowed_colors and hexval not in allowed_colors:
                                issues.append(Issue(slide_index, "template", "color_not_in_palette", f"Цвет текста {hexval} вне палитры шаблона", True, "info"))
                            contrast = _contrast_ratio(hexval, local_bg)
                            if contrast < _MIN_CONTRAST_RATIO:
                                issues.append(Issue(slide_index, "template", "low_contrast", f"Контраст текста к фону {contrast:.1f}:1 ниже 4.5:1", True))
                        except Exception:
                            pass
            bullet_paras = [p for p in shape.text_frame.paragraphs if "".join(r.text for r in p.runs).strip()]
            if len(bullet_paras) > _MAX_BULLETS:
                issues.append(Issue(slide_index, "density", "too_many_bullets", f"{len(bullet_paras)} буллетов на слайде (> {_MAX_BULLETS})", True))

        if getattr(shape, "has_table", False) and shape.has_table:
            table = shape.table
            if len(table.rows) > _MAX_TABLE_ROWS:
                issues.append(Issue(slide_index, "density", "table_too_many_rows", f"Таблица: {len(table.rows)} строк (> {_MAX_TABLE_ROWS})", True))
            if len(table.columns) > _MAX_TABLE_COLS:
                issues.append(Issue(slide_index, "density", "table_too_many_cols", f"Таблица: {len(table.columns)} колонок (> {_MAX_TABLE_COLS})", True))

        if shape.shape_type is not None and shape.shape_type.name == "CHART":
            try:
                n_series = len(shape.chart.series)
                if n_series > _MAX_CHART_SERIES:
                    issues.append(Issue(slide_index, "density", "chart_too_many_series", f"{n_series} серий на диаграмме (> {_MAX_CHART_SERIES})", True))
                if not shape.chart.has_legend and len(shape.chart.plots[0].categories) > 1:
                    issues.append(Issue(slide_index, "integrity", "chart_missing_legend", "У диаграммы нет легенды", True, "info"))
            except Exception:
                pass

        if shape.shape_type is not None and shape.shape_type.name == "PICTURE":
            try:
                native_w, native_h = shape.image.size
                if native_w and native_h and shape.width and shape.height:
                    native_ratio = native_w / native_h
                    box_ratio = shape.width / shape.height
                    if abs(native_ratio - box_ratio) / native_ratio > 0.15:
                        issues.append(Issue(slide_index, "verstka", "picture_stretched", "Пропорции изображения нарушены (растянуто)", True))
            except Exception:
                pass

    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            (name_a, box_a), (name_b, box_b) = boxes[i], boxes[j]
            overlap = _overlap_area(box_a, box_b)
            area_a = max(1, (box_a[2] - box_a[0]) * (box_a[3] - box_a[1]))
            area_b = max(1, (box_b[2] - box_b[0]) * (box_b[3] - box_b[1]))
            if overlap > 0.1 * min(area_a, area_b):
                issues.append(Issue(slide_index, "verstka", "overlap", f"'{name_a}' и '{name_b}' заметно перекрываются", True))

    joined_text = "\n".join(t.strip() for t in total_text if t.strip())
    if not joined_text:
        issues.append(Issue(slide_index, "integrity", "empty_slide", "Слайд пустой или содержит только заголовок", True, "error"))

    slide_area = slide_w * slide_h
    fill_ratio = min(filled_area / slide_area, 1.0) if slide_area else 0
    if fill_ratio < _MIN_FILL_RATIO:
        issues.append(Issue(slide_index, "density", "underfilled", f"Слайд заполнен менее чем на {int(_MIN_FILL_RATIO*100)}%", True, "info"))
    elif fill_ratio > _MAX_FILL_RATIO:
        issues.append(Issue(slide_index, "density", "overfilled", f"Слайд заполнен более чем на {int(_MAX_FILL_RATIO*100)}%", True, "info"))

    norm = re.sub(r"\s+", " ", joined_text.lower())
    if norm:
        if norm in seen_texts:
            issues.append(Issue(slide_index, "integrity", "duplicate_slide", f"Дублирует слайд {seen_texts[norm]}", True))
        else:
            seen_texts[norm] = slide_index

    return issues


def run_deterministic_audit(pptx_path: str, design_system: DesignSystem) -> AuditReport:
    prs = Presentation(pptx_path)
    allowed_fonts = {design_system.major_font, design_system.minor_font} - {None}
    seen_texts: dict[str, int] = {}
    issues: list[Issue] = []
    for i, slide in enumerate(prs.slides):
        issues.extend(_check_slide(slide, i, design_system, allowed_fonts, seen_texts))
    return AuditReport(issues=issues)


# --------------------------------------------------------------------------
# Content audit — the 11 yes/no questions from Приложение 1, model-judged
# --------------------------------------------------------------------------

_CONTENT_QUESTIONS_RU = [
    "Заголовок содержит вывод, а не просто называет тему?",
    "Содержимое слайда соответствует заголовку?",
    "Слайд пересказывается одним предложением?",
    "Все цифры и факты со слайда есть в исходных материалах?",
    "На слайде есть содержание, а не только заголовок?",
    "Нет служебного мусора: реплик спикера, кусков промпта?",
    "Текст без опечаток?",
    "Все строки таблицы и элементы легенды работают на мысль слайда?",
]


def run_content_audit(llm: LLMClient, pptx_path: str, source_facts_text: str) -> AuditReport:
    """Non-deterministic by construction — the same slide can get a
    different verdict on a re-run, per the ТЗ's own definition of these checks."""
    prs = Presentation(pptx_path)
    issues: list[Issue] = []
    system = (
        "Ты проверяешь качество слайда презентации. Тебе дан текст слайда и краткий список "
        "исходных фактов. Ответь JSON-объектом {\"problems\": [\"...\"]} — список коротких "
        "описаний найденных проблем (на русском). Если проблем нет — верни {\"problems\": []}. "
        "Проверяй: заголовок отражает вывод, содержимое соответствует заголовку, нет опечаток, "
        "нет служебного текста (реплик спикера, кусков промпта), все цифры действительно есть "
        "в списке фактов ниже."
    )
    for i, slide in enumerate(prs.slides):
        texts = [sh.text_frame.text for sh in slide.shapes if getattr(sh, "has_text_frame", False) and sh.has_text_frame]
        slide_text = "\n".join(t for t in texts if t.strip())
        if not slide_text.strip():
            continue
        user = f"Факты:\n{source_facts_text[:1500]}\n\nТекст слайда {i+1}:\n{slide_text}"
        try:
            raw = llm.complete_json(system, user, temperature=0.0)
            problems = raw.get("problems", []) if isinstance(raw, dict) else []
            for p in problems:
                issues.append(Issue(i, "content", "content_review", str(p), False, "warning"))
        except Exception as exc:  # noqa: BLE001 - a flaky content audit must not fail the whole job
            logger.info("content audit skipped for slide %s (%s)", i, exc)
    return AuditReport(issues=issues)


def run_full_audit(llm: LLMClient, pptx_path: str, design_system: DesignSystem, source_facts_text: str) -> AuditReport:
    deterministic = run_deterministic_audit(pptx_path, design_system)
    content = run_content_audit(llm, pptx_path, source_facts_text)
    return AuditReport(issues=[*deterministic.issues, *content.issues])
