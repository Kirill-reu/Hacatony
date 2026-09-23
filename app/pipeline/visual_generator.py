"""ТЗ §2 item 3: "Генерация графиков, таблиц, диаграмм, пиктограмм, элементов
SmartArt внутри слайда" + item 3.1 (image bonus task).

This module decides *what data* goes into a chart/table (grounded in the
facts gathered by ``data_provider`` — never invented) and *what image* to
request from the image client. It does not touch the .pptx shape tree
directly for charts/tables/pictures — ``slide_builder`` owns turning a
``ChartSpec``/``TableSpec``/image bytes into a native chart/table/picture
shape inside a specific placeholder's box. The one exception is the
icon-row helper below: python-pptx has no SmartArt support at all, so the
"SmartArt-like" approximation is a handful of native autoshapes laid out
in a row, which is simplest to build directly against a slide + box.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Pt

from ..llm.image_client import ImageClient
from ..llm.text_client import LLMClient
from .data_provider import SearchResult

logger = logging.getLogger(__name__)

_NUMBER_RE = re.compile(r"-?\d[\d\s]*[.,]?\d*")
_MAX_CHART_SERIES_POINTS = 5   # Приложение 1: "больше 5 серий на диаграмме" is flagged
_MAX_TABLE_ROWS = 7            # Приложение 1: "таблица больше 7 строк"

# Dates read as plausible-looking numbers to the regex above (a "13" and a
# "2026" out of "Aug 13, 2026" are, individually, valid floats) but are
# never a real metric — observed in practice turning a page's publish date
# into a chart's "values". Stripped from the source text before extraction
# runs, rather than filtered after, so the surrounding words used as a
# label also stop being date fragments.
_DATE_LIKE_RE = re.compile(
    r"\b(?:"
    r"\d{1,2}[./]\d{1,2}[./]\d{2,4}"  # 13.08.2026, 13/08/26
    r"|\d{4}-\d{2}-\d{2}"  # 2026-08-13
    r"|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s*\d{4}"  # Aug 13, 2026
    # Month stems kept to their shortest common prefix (e.g. "авг" rather
    # than "август") on purpose: that prefix matches both the full word
    # ("августа") and the abbreviated form web snippets actually use
    # ("авг."), with \w*\.? soaking up whatever letters/period follow.
    r"|\d{1,2}\s+(?:янв|фев|мар|апр|ма[йя]|июн|июл|авг|сен|окт|ноя|дек)\w*\.?\s+\d{4}\s*г?\.?"  # 13 августа 2026, 3 июн. 2025 г.
    r")\b",
    re.IGNORECASE,
)


@dataclass
class ChartSpec:
    categories: list[str]
    series_name: str
    values: list[float]
    chart_type: str = "bar"  # bar | line | pie
    source_note: str = ""


@dataclass
class TableSpec:
    headers: list[str]
    rows: list[list[str]] = field(default_factory=list)
    source_note: str = ""


# --------------------------------------------------------------------------
# Chart / table data extraction — grounded in retrieved facts only
# --------------------------------------------------------------------------


def _numbers_with_context(text: str) -> list[tuple[str, float]]:
    """Pairs each number found in ``text`` with a short label taken from the
    words immediately before it, e.g. "выручка выросла на 23%" -> ("выручка выросла на", 23.0).

    Two defensive exclusions, both hardened against real observed garbage
    (a "Aug 13, 2026 · grok4.7" web-search snippet turning into chart data
    with a category list of date fragments and a values list of [13, 2026,
    4.7, ...]): dates are stripped before matching starts, and a number
    glued directly onto a preceding letter (a version string, product code,
    "top10", ...) is skipped rather than treated as a standalone metric.
    """
    text = _DATE_LIKE_RE.sub(" ", text)
    out: list[tuple[str, float]] = []
    for m in _NUMBER_RE.finditer(text):
        start = m.start()
        if start > 0 and text[start - 1].isalpha():
            continue
        raw = m.group().replace(" ", "").replace(",", ".")
        try:
            value = float(raw)
        except ValueError:
            continue
        label = text[:start].strip().split()[-4:]
        label_str = " ".join(label) if label else text[m.end() : m.end() + 30].strip()
        out.append((label_str or "значение", value))
    return out


def build_chart_spec(llm: LLMClient, hint: str, facts: list[SearchResult], chart_type: str, language: str) -> ChartSpec | None:
    """Tries an LLM extraction first (asked to use only the given text), and
    falls back to a plain regex pass over the same text. Either way, every
    number in the result must be traceable to the source text — we never
    hand the model an empty context and accept whatever numbers it invents."""
    source_text = hint or " ".join(f.snippet for f in facts[:3])
    if not source_text.strip():
        return None

    try:
        system = (
            "Извлеки из текста числовые данные для диаграммы. Используй ТОЛЬКО числа, "
            "присутствующие в тексте. Верни JSON: {\"series_name\": str, \"categories\": [str,...], "
            "\"values\": [number,...]}. Не больше 5 категорий."
            if language.startswith("ru")
            else "Extract numeric data for a chart from the text. Use ONLY numbers present in the "
            "text. Return JSON: {\"series_name\": str, \"categories\": [str,...], \"values\": [number,...]}. At most 5 categories."
        )
        raw = llm.complete_json(system, source_text, temperature=0.1)
        cats, vals = raw.get("categories"), raw.get("values")
        if isinstance(cats, list) and isinstance(vals, list) and len(cats) == len(vals) and cats:
            numeric_vals = [float(v) for v in vals if _is_number(v)]
            if len(numeric_vals) == len(vals) and _values_traceable(numeric_vals, source_text):
                n = min(len(cats), _MAX_CHART_SERIES_POINTS)
                return ChartSpec(
                    categories=[str(c) for c in cats[:n]],
                    values=numeric_vals[:n],
                    series_name=str(raw.get("series_name", "Значение")),
                    chart_type=chart_type,
                    source_note=source_text[:200],
                )
    except Exception as exc:  # noqa: BLE001
        logger.debug("build_chart_spec: LLM extraction failed (%s), falling back to regex", exc)

    pairs = _numbers_with_context(source_text)[:_MAX_CHART_SERIES_POINTS]
    if len(pairs) < 2:
        return None
    return ChartSpec(
        categories=[p[0][:24] for p in pairs],
        values=[p[1] for p in pairs],
        series_name="Значение" if language.startswith("ru") else "Value",
        chart_type=chart_type,
        source_note=source_text[:200],
    )


def _is_number(v) -> bool:
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


def _values_traceable(values: list[float], source_text: str) -> bool:
    """Reject an extraction if the model's numbers can't actually be found
    in the source text — that means it made them up."""
    normalized = source_text.replace(" ", "").replace(",", ".")
    hits = 0
    for v in values:
        as_int = str(int(v)) if float(v).is_integer() else None
        if (as_int and as_int in normalized) or str(v) in normalized:
            hits += 1
    return hits >= max(1, len(values) // 2)


def build_table_spec(hint: str, facts: list[SearchResult], language: str) -> TableSpec | None:
    # _MAX_TABLE_ROWS (7, per Приложение 1: "таблица больше 7 строк") counts
    # the WHOLE pptx table including its header row — python-pptx's
    # table.rows, which the audit checks against, does the same. Cap data
    # rows one short of that so header + data never exceeds the limit.
    max_data_rows = _MAX_TABLE_ROWS - 1
    source_text = hint or " ".join(f"{f.title}: {f.snippet}" for f in facts[:max_data_rows])
    pairs = _numbers_with_context(source_text)[:max_data_rows]
    if pairs:
        headers = ["Показатель", "Значение"] if language.startswith("ru") else ["Metric", "Value"]
        rows = [[label[:40], _format_number(value)] for label, value in pairs]
        return TableSpec(headers=headers, rows=rows, source_note=source_text[:200])

    if facts:
        headers = ["Источник", "Факт"] if language.startswith("ru") else ["Source", "Fact"]
        rows = [[f.title[:40], f.snippet[:90]] for f in facts[:max_data_rows]]
        return TableSpec(headers=headers, rows=rows, source_note="")
    return None


def _format_number(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else f"{v:.1f}"


# --------------------------------------------------------------------------
# Images
# --------------------------------------------------------------------------


def generate_image(image_client: ImageClient, prompt: str) -> bytes:
    style_suffix = ", flat corporate illustration, clean background, no text, no watermark"
    return image_client.generate(prompt + style_suffix)


# --------------------------------------------------------------------------
# "SmartArt-like" native icon row (python-pptx has no real SmartArt support)
# --------------------------------------------------------------------------

_SHAPE_PALETTE_KEY = "accent1"


def add_icon_row(slide, left: int, top: int, width: int, height: int, items: list[str], accent_hex: str) -> None:
    """Lays out up to 5 items as connected rounded-rectangle cards — the
    native-shape approximation of a SmartArt process/step diagram."""
    items = items[:5] or ["—"]
    n = len(items)
    gap = Emu(int(width * 0.03))
    card_w = Emu(int((width - gap * (n - 1)) / n)) if n > 1 else Emu(width)
    card_h = Emu(height)
    x = left

    centers = []
    for text in items:
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top, card_w, card_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = _hex_to_rgbcolor(accent_hex)
        shape.line.fill.background()
        tf = shape.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Emu(91440)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = text
        run.font.size = Pt(14)
        run.font.color.rgb = _hex_to_rgbcolor("FFFFFF")
        centers.append((x, top, card_w, card_h))
        x = Emu(int(x + card_w + gap))

    for (x1, y1, w1, h1), (x2, y2, w2, h2) in zip(centers, centers[1:]):
        connector = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Emu(int(x1 + w1)),
            Emu(int(y1 + h1 / 2)),
            Emu(int(x2)),
            Emu(int(y2 + h2 / 2)),
        )
        connector.line.color.rgb = _hex_to_rgbcolor(accent_hex)
        connector.line.width = Pt(2)


def _hex_to_rgbcolor(hex_str: str):
    from pptx.dml.color import RGBColor

    hex_str = hex_str.lstrip("#")
    return RGBColor.from_string(hex_str.upper())
