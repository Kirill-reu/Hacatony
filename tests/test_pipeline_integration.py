"""End-to-end pipeline test using the offline stubs (no network, no API keys)
so it runs anywhere, including CI. It exercises the same path a real request
takes: parse template -> plan 3 variants -> build decks -> export -> audit."""
from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation

from app.llm.image_client import PlaceholderImageClient
from app.llm.text_client import OfflineStubClient
from app.pipeline.template_parser import parse_template
from app.pipeline.variant_engine import generate_variants


@pytest.fixture()
def sample_template(tmp_path: Path) -> Path:
    prs = Presentation()
    path = tmp_path / "template.pptx"
    prs.save(str(path))
    return path


_LONG_BRIEF = (
    "Мы запускаем новую фичу автоматической генерации презентаций для внутренних команд. "
    "Она разбирает загруженный шаблон и извлекает дизайн-систему. "
    "Сервис сам пишет текст, подбирает визуализацию и собирает три варианта верстки. "
    "Это экономит часы работы дизайнеров и обеспечивает единый корпоративный стиль. "
    "Команда сможет проверить результат и выбрать подходящий вариант перед публикацией. "
    "Планируется пилотный запуск на внутренних командах в течение квартала."
)


def test_generates_three_distinguishable_variants(sample_template: Path, tmp_path: Path, monkeypatch):
    # Patch the shared settings singleton directly — it's read once at import
    # time, so setting the env var after the fact wouldn't take effect.
    from app.config import settings as app_settings

    monkeypatch.setattr(app_settings, "web_search_enabled", False)
    import time

    ds = parse_template(sample_template)
    outcomes = generate_variants(
        template_path=sample_template,
        design_system=ds,
        llm=OfflineStubClient(),
        image_client=PlaceholderImageClient(),
        out_dir=tmp_path / "out",
        brief=_LONG_BRIEF,
        purpose="feature",
        slide_count=6,
        audience=None,
        language="ru",
        title_for_export="Тест",
        deadline_ts=time.monotonic() + 60,
    )

    assert len(outcomes) == 3
    ids = {o.variant_id for o in outcomes}
    assert ids == {"concise", "detailed", "visual"}

    for o in outcomes:
        assert o.pptx_path.exists()
        assert o.html_path.exists()
        assert o.slide_count >= 4

    # The whole point of the 3-variant requirement: they must differ. With a
    # short/offline-fallback source, slide-by-slide bullet *count* can tie
    # (there's only ever one sentence available per content slide here), but
    # word count (max_bullet_words truncation) and shape composition (only
    # "concise"/"visual" pull in an image) must not.
    word_counts = {o.variant_id: _total_words(o.pptx_path) for o in outcomes}
    shape_counts = {o.variant_id: _total_shapes(o.pptx_path) for o in outcomes}
    assert len(set(word_counts.values())) > 1 or len(set(shape_counts.values())) > 1, (
        f"variants produced identical content: words={word_counts} shapes={shape_counts}"
    )


def _total_words(pptx_path: Path) -> int:
    prs = Presentation(str(pptx_path))
    total = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                total += len(shape.text_frame.text.split())
    return total


def _total_shapes(pptx_path: Path) -> int:
    prs = Presentation(str(pptx_path))
    return sum(len(slide.shapes) for slide in prs.slides)


def test_respects_time_budget_by_stopping_early(sample_template: Path, tmp_path: Path, monkeypatch):
    from app.config import settings as app_settings

    monkeypatch.setattr(app_settings, "web_search_enabled", False)
    import time

    ds = parse_template(sample_template)
    outcomes = generate_variants(
        template_path=sample_template,
        design_system=ds,
        llm=OfflineStubClient(),
        image_client=PlaceholderImageClient(),
        out_dir=tmp_path / "out",
        brief=_LONG_BRIEF,
        purpose="project",
        slide_count=6,
        audience=None,
        language="ru",
        title_for_export="Тест",
        deadline_ts=time.monotonic() - 1,  # already expired
    )
    assert outcomes == []
