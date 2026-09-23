"""Decomposes an uploaded .pptx template into a machine-usable design system:
color tokens, typography (fonts + type scale), composition patterns
(classified slide layouts with placeholder geometry), margins, and any
logo/decoration on the slide masters.

This is ТЗ §2 item 1 ("Парсинг шаблона презентации для последующей
воспроизводимости при генерации") and the "выделить токены дизайн-системы"
deliverable. Everything downstream (content_generator, slide_builder, audit)
consumes the ``DesignSystem`` this module returns instead of touching the
.pptx file directly.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------


@dataclass
class PlaceholderSpec:
    idx: int
    type: str  # PP_PLACEHOLDER_TYPE name, e.g. "TITLE", "OBJECT", "PICTURE"
    name: str
    left: int
    top: int
    width: int
    height: int


@dataclass
class LayoutSpec:
    index: int             # global id across ALL masters — unique key into DesignSystem.layouts
    master_index: int      # which prs.slide_masters[...] this layout actually belongs to
    layout_idx_in_master: int  # its index within that master's own .slide_layouts collection
    name: str
    role: str  # classified composition pattern, see _classify_layout
    placeholders: list[PlaceholderSpec] = field(default_factory=list)

    def placeholder_by_role(self, roles: set[str]) -> PlaceholderSpec | None:
        for p in self.placeholders:
            if p.type in roles:
                return p
        return None


@dataclass
class ColorToken:
    role: str  # dk1 | lt1 | dk2 | lt2 | accent1..6 | hlink | folHlink
    hex: str


@dataclass
class LogoBox:
    left: int
    top: int
    width: int
    height: int


@dataclass
class DesignSystem:
    slide_width: int
    slide_height: int
    colors: list[ColorToken]
    major_font: str
    minor_font: str
    type_scale_pt: list[float]           # distinct point sizes actually used, ascending
    layouts: list[LayoutSpec]
    layout_roles: dict[str, list[int]]   # role -> layout indices, best-scoring first
    margins_emu: dict[str, int]          # left/top/right/bottom safe-content margins
    logo_boxes: list[LogoBox]
    warnings: list[str] = field(default_factory=list)

    def palette_hexes(self) -> list[str]:
        return [c.hex for c in self.colors]

    def best_layout_for(self, role: str, fallback_roles: tuple[str, ...] = ()) -> LayoutSpec | None:
        for r in (role, *fallback_roles):
            indices = self.layout_roles.get(r)
            if indices:
                return self.layouts[indices[0]]
        return None


# --------------------------------------------------------------------------
# Layout classification
# --------------------------------------------------------------------------

CENTER_TITLE_TYPES = {"CENTER_TITLE"}
TITLE_TYPES = {"TITLE"}
SUBTITLE_TYPES = {"SUBTITLE"}
CONTENT_TYPES = {"BODY", "OBJECT"}
PICTURE_TYPES = {"PICTURE"}
CHART_TYPES = {"CHART"}
TABLE_TYPES = {"TABLE"}
DECORATIVE_TYPES = {"DATE", "FOOTER", "SLIDE_NUMBER"}


def _classify_layout(placeholders: list[PlaceholderSpec], name_hint: str, slide_height: int) -> str:
    """Geometry-first classification (so it works regardless of what language
    or naming convention the template's own layout names use), with layout
    names used only as a tie-breaker between structurally-similar layouts."""
    content = [p for p in placeholders if p.type not in DECORATIVE_TYPES]
    if not content:
        return "blank"

    name_l_early = name_hint.lower()
    if any(k in name_l_early for k in ("спасибо", "thank you", "thanks")):
        # A closing/"thank you" slide must never be picked for ordinary
        # mid-deck content — without this override it can structurally look
        # just like any other title+body(+body) layout and get reused for
        # regular sections, which is a real semantic bug, not a cosmetic one.
        return "closing"

    if any(k in name_l_early for k in ("титульн", "title slide", "cover slide")):
        # Decks authored in Google Slides and exported to .pptx don't carry
        # PowerPoint's native CENTER_TITLE placeholder type, so the
        # structural check below never fires for them — but "Титульный
        # слайд" is an unambiguous, deliberate naming convention worth
        # trusting outright, even when the layout also has a richer bio-card
        # composition (photo + name + role) that would otherwise look like
        # a two_content layout.
        return "title_slide"

    center_titles = [p for p in content if p.type in CENTER_TITLE_TYPES]
    titles = [p for p in content if p.type in TITLE_TYPES]
    subtitles = [p for p in content if p.type in SUBTITLE_TYPES]
    pictures = [p for p in content if p.type in PICTURE_TYPES]
    bodies = [p for p in content if p.type in CONTENT_TYPES]
    all_titles = titles + center_titles
    others = [p for p in content if p not in center_titles + titles + subtitles + pictures + bodies]

    name_l = name_hint.lower()

    if center_titles:
        return "title_slide"

    if pictures and all_titles:
        return "content_image"

    # A lone title (± one short subtitle line) is either a section/divider
    # slide (vertically centered, i.e. its top sits well below the slide's
    # top edge) or a plain "title only" content slide.
    if all_titles and not bodies and not pictures and not others and len(content) <= 2:
        t = all_titles[0]
        if slide_height and t.top and t.top > slide_height * 0.25:
            return "section_header"
        return "title_only"

    if len(bodies) >= 2:
        # Group by x position with a tolerance, so placeholders that are
        # merely stacked vertically (their left edges differ only by a
        # rounding artefact of a few EMU — common in decks exported from
        # Google Slides) aren't mistaken for genuine side-by-side columns.
        tolerance = max(int((slide_height or 0) * 0.02), 50_000)
        distinct_columns: list[int] = []
        for p in sorted(bodies, key=lambda b: b.left or 0):
            if not distinct_columns or (p.left or 0) - distinct_columns[-1] > tolerance:
                distinct_columns.append(p.left or 0)
        if len(distinct_columns) >= 2:
            return "comparison" if any(k in name_l for k in ("compar", "сравнен")) else "two_content"

    if all_titles and bodies:
        # Title + a single short supporting line, both sitting low on the
        # slide, is a section-divider composition even though it technically
        # has a "body" placeholder (PowerPoint's own "Section Header" layout
        # is built exactly this way).
        if len(content) == 2 and len(bodies) == 1:
            t, b = all_titles[0], bodies[0]
            both_low = slide_height and min(t.top or 0, b.top or 0) > slide_height * 0.3
            if both_low:
                return "section_header"
        return "content_bullets"

    return "generic"


def _extract_layouts(prs: Presentation) -> tuple[list[LayoutSpec], dict[str, list[int]]]:
    """Walks EVERY slide master's own .slide_layouts collection.

    ``Presentation.slide_layouts`` in python-pptx is a shortcut for
    ``slide_masters[0].slide_layouts`` only — a template with more than one
    master (common in decks originally built in Google Slides and exported
    to .pptx, which is exactly this project's target templates) silently
    loses every layout on its other masters if you iterate that shortcut
    instead of the masters themselves.
    """
    layouts: list[LayoutSpec] = []
    idx = 0
    for master_index, master in enumerate(prs.slide_masters):
        for layout_idx_in_master, layout in enumerate(master.slide_layouts):
            phs = []
            for ph in layout.placeholders:
                pf = ph.placeholder_format
                type_name = pf.type.name if pf.type is not None else "UNKNOWN"
                phs.append(
                    PlaceholderSpec(
                        idx=pf.idx,
                        type=type_name,
                        name=ph.name,
                        left=ph.left or 0,
                        top=ph.top or 0,
                        width=ph.width or 0,
                        height=ph.height or 0,
                    )
                )
            role = _classify_layout(phs, layout.name or "", prs.slide_height)
            layouts.append(
                LayoutSpec(
                    index=idx,
                    master_index=master_index,
                    layout_idx_in_master=layout_idx_in_master,
                    name=layout.name or f"Layout {idx}",
                    role=role,
                    placeholders=phs,
                )
            )
            idx += 1

    roles: dict[str, list[int]] = defaultdict(list)
    for l in layouts:
        roles[l.role].append(l.index)
    return layouts, dict(roles)


# --------------------------------------------------------------------------
# Theme (colors + fonts)
# --------------------------------------------------------------------------

_COLOR_ROLE_ORDER = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]


def _extract_theme(prs: Presentation, warnings: list[str]) -> tuple[list[ColorToken], str, str]:
    try:
        master = prs.slide_masters[0]
        theme_part = master.part.part_related_by(RT.THEME)
        theme_elm = parse_xml(theme_part.blob)
        theme_elements = theme_elm.find(qn("a:themeElements"))
        clr_scheme = theme_elements.find(qn("a:clrScheme"))

        colors: list[ColorToken] = []
        by_tag = {}
        for child in clr_scheme:
            tag = child.tag.split("}")[-1]
            srgb = child.find(qn("a:srgbClr"))
            sysclr = child.find(qn("a:sysClr"))
            hexval = srgb.get("val") if srgb is not None else (sysclr.get("lastClr") if sysclr is not None else None)
            if hexval:
                by_tag[tag] = hexval.upper()
        for role in _COLOR_ROLE_ORDER:
            if role in by_tag:
                colors.append(ColorToken(role=role, hex=by_tag[role]))

        font_scheme = theme_elements.find(qn("a:fontScheme"))
        major = font_scheme.find(qn("a:majorFont")).find(qn("a:latin")).get("typeface") or "Calibri Light"
        minor = font_scheme.find(qn("a:minorFont")).find(qn("a:latin")).get("typeface") or "Calibri"
        return colors, major, minor
    except Exception as exc:  # pragma: no cover - defensive; malformed/odd themes shouldn't crash parsing
        warnings.append(f"Could not fully read theme (colors/fonts fall back to defaults): {exc}")
        return (
            [ColorToken(role="dk1", hex="000000"), ColorToken(role="lt1", hex="FFFFFF"), ColorToken(role="accent1", hex="4F81BD")],
            "Calibri Light",
            "Calibri",
        )


# --------------------------------------------------------------------------
# Typography scale (observed point sizes across masters/layouts/sample slides)
# --------------------------------------------------------------------------


def _extract_type_scale(prs: Presentation) -> list[float]:
    sizes: Counter[float] = Counter()

    def walk_text_frame(tf) -> None:
        for p in tf.paragraphs:
            if p.font.size is not None:
                sizes[p.font.size.pt] += 1
            for r in p.runs:
                if r.font.size is not None:
                    sizes[r.font.size.pt] += 1

    all_layouts = [layout for master in prs.slide_masters for layout in master.slide_layouts]
    containers = list(prs.slide_masters) + all_layouts + list(prs.slides)
    for container in containers:
        for shape in container.shapes:
            if shape.has_text_frame:
                walk_text_frame(shape.text_frame)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        walk_text_frame(cell.text_frame)

    if not sizes:
        return [18.0, 24.0, 32.0, 44.0]
    return sorted(sizes.keys())


# --------------------------------------------------------------------------
# Margins & logos
# --------------------------------------------------------------------------


def _estimate_margins(layouts: list[LayoutSpec], slide_w: int, slide_h: int) -> dict[str, int]:
    lefts, tops, rights, bottoms = [], [], [], []
    for l in layouts:
        for p in l.placeholders:
            if p.type in DECORATIVE_TYPES or p.width == 0:
                continue
            lefts.append(p.left)
            tops.append(p.top)
            rights.append(slide_w - (p.left + p.width))
            bottoms.append(slide_h - (p.top + p.height))
    if not lefts:
        # ~5% margins as a reasonable default
        return {"left": slide_w // 20, "top": slide_h // 20, "right": slide_w // 20, "bottom": slide_h // 20}
    return {
        "left": max(min(lefts), 0),
        "top": max(min(tops), 0),
        "right": max(min(rights), 0),
        "bottom": max(min(bottoms), 0),
    }


def _extract_logo_boxes(prs: Presentation) -> list[LogoBox]:
    boxes: list[LogoBox] = []
    for master in prs.slide_masters:
        for shape in master.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                boxes.append(LogoBox(left=shape.left or 0, top=shape.top or 0, width=shape.width or 0, height=shape.height or 0))
    if boxes:
        return boxes

    # Some decks (notably ones authored in Google Slides and exported to
    # .pptx) leave the masters themselves empty and instead repeat brand
    # elements — a logo, a footer mark — on every individual layout. A
    # decorative shape recurring at (roughly) the same position across
    # several layouts is almost certainly one of those, as opposed to
    # one-off content that happens to live at that spot on a single layout.
    slide_w, slide_h = prs.slide_width, prs.slide_height
    position_counts: Counter[tuple[int, int, int, int]] = Counter()
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            for shape in layout.shapes:
                if shape.shape_type not in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP):
                    continue
                if not (shape.width and shape.height):
                    continue
                # a shape covering (almost) the whole slide is a repeated
                # background, not a logo — exclude it here
                if shape.width >= slide_w * 0.9 and shape.height >= slide_h * 0.9:
                    continue
                # round to tolerate tiny sub-EMU drift between otherwise-identical copies
                key = (round((shape.left or 0) / 10000), round((shape.top or 0) / 10000), shape.width, shape.height)
                position_counts[key] += 1

    recurring = [key for key, count in position_counts.items() if count >= 3]
    for left_r, top_r, width, height in recurring:
        boxes.append(LogoBox(left=left_r * 10000, top=top_r * 10000, width=width, height=height))
    return boxes


# --------------------------------------------------------------------------
# Public entry point
# --------------------------------------------------------------------------


def parse_template(path: str | Path) -> DesignSystem:
    prs = Presentation(str(path))
    warnings: list[str] = []

    layouts, layout_roles = _extract_layouts(prs)
    if not any(r in layout_roles for r in ("content_bullets", "generic", "two_content")):
        warnings.append("Template has no obvious text+content layout; generation will fall back to the closest available layout.")

    colors, major_font, minor_font = _extract_theme(prs, warnings)
    type_scale = _extract_type_scale(prs)
    margins = _estimate_margins(layouts, prs.slide_width, prs.slide_height)
    logo_boxes = _extract_logo_boxes(prs)

    return DesignSystem(
        slide_width=prs.slide_width,
        slide_height=prs.slide_height,
        colors=colors,
        major_font=major_font,
        minor_font=minor_font,
        type_scale_pt=type_scale,
        layouts=layouts,
        layout_roles=layout_roles,
        margins_emu=margins,
        logo_boxes=logo_boxes,
        warnings=warnings,
    )
