"""Retrieves supporting facts from the web for a brief, so generated slides
can cite real numbers/quotes instead of the model inventing them (ТЗ audit
item "Все цифры и факты со слайда есть в исходных материалах").

Kept as a small, swappable interface: ``DataProvider.search`` is the only
thing ``content_generator`` calls. The default implementation
(``DuckDuckGoDataProvider``) needs no API key, which matters for a
hackathon judged on a laptop with no time to provision search-API
credentials; swap in Bing/SerpAPI/whatever the team already has a key for
by implementing the same interface.
"""
from __future__ import annotations

import logging
from abc import ABC, abstractmethod
from dataclasses import dataclass

from ..config import settings

logger = logging.getLogger(__name__)


@dataclass
class SearchResult:
    title: str
    snippet: str
    url: str


class DataProvider(ABC):
    @abstractmethod
    def search(self, query: str, max_results: int, language: str = "ru") -> list[SearchResult]:
        ...


_DDG_REGION_BY_LANGUAGE = {"ru": "ru-ru", "en": "us-en"}


class DuckDuckGoDataProvider(DataProvider):
    def search(self, query: str, max_results: int, language: str = "ru") -> list[SearchResult]:
        try:
            from duckduckgo_search import DDGS
        except ImportError:
            logger.warning("duckduckgo_search not installed — skipping web grounding for %r", query)
            return []

        # Without a region hint, DuckDuckGo happily returns whatever-language
        # results rank highest globally — for a brief in Russian that has
        # actually surfaced Chinese-language pages, whose raw snippets then
        # end up quoted verbatim on a Russian-language slide. Bias toward the
        # deck's own language; "wt-wt" (DDG's own code for "no region") is
        # the fallback for languages we don't have a mapping for.
        region = _DDG_REGION_BY_LANGUAGE.get(language.split("-")[0].lower(), "wt-wt")

        try:
            with DDGS() as ddgs:
                hits = list(ddgs.text(query, region=region, max_results=max_results))
        except Exception as exc:  # network errors, rate limits, etc. must not kill a generation job
            logger.warning("web search failed for %r: %s", query, exc)
            return []

        results = [
            SearchResult(title=h.get("title", ""), snippet=h.get("body", ""), url=h.get("href", ""))
            for h in hits
            if h.get("body")
        ]
        return [r for r in results if not _is_wrong_script(r.title + " " + r.snippet, language)]


class NullDataProvider(DataProvider):
    """Used when WEB_SEARCH_ENABLED=false. Content generation still works —
    it just can't ground claims in fresh web facts, only the user's brief."""

    def search(self, query: str, max_results: int, language: str = "ru") -> list[SearchResult]:
        return []


def _is_wrong_script(text: str, language: str, threshold: float = 0.2) -> bool:
    """Defensive filter behind the region hint above, not a replacement for
    it: even region-scoped DuckDuckGo results occasionally include a page in
    a third script entirely (observed: Chinese results for a Russian-region
    query). A snippet that's mostly CJK is unusable on a Russian or English
    slide regardless of how it got returned, so drop it outright rather than
    let it get quoted into slide text or mined for "chart data" (a date like
    "Aug 13, 2026" embedded in such a snippet has, in practice, been picked
    up by the numeric-fact extractor as if it were a real metric)."""
    if not text.strip():
        return False
    cjk_chars = sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff" or "\u3040" <= ch <= "\u30ff")
    letters = sum(1 for ch in text if ch.isalpha())
    if letters == 0:
        return False
    return (cjk_chars / letters) > threshold


def build_data_provider() -> DataProvider:
    if not settings.web_search_enabled:
        return NullDataProvider()
    return DuckDuckGoDataProvider()


def gather_facts(provider: DataProvider, topics: list[str], language: str = "ru", per_topic: int | None = None) -> list[SearchResult]:
    """Search each topic and return a deduplicated (by URL) flat list."""
    per_topic = per_topic or settings.web_search_max_results
    seen_urls: set[str] = set()
    results: list[SearchResult] = []
    for topic in topics:
        for r in provider.search(topic, per_topic, language):
            if r.url and r.url in seen_urls:
                continue
            seen_urls.add(r.url)
            results.append(r)
    return results
