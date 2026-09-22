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
    def search(self, query: str, max_results: int) -> list[SearchResult]:
        ...


class DuckDuckGoDataProvider(DataProvider):
    def search(self, query: str, max_results: int) -> list[SearchResult]:
        try:
            from duckduckgo_search import DDGS
        except ImportError:
            logger.warning("duckduckgo_search not installed — skipping web grounding for %r", query)
            return []

        try:
            with DDGS() as ddgs:
                hits = list(ddgs.text(query, max_results=max_results))
        except Exception as exc:  # network errors, rate limits, etc. must not kill a generation job
            logger.warning("web search failed for %r: %s", query, exc)
            return []

        return [
            SearchResult(title=h.get("title", ""), snippet=h.get("body", ""), url=h.get("href", ""))
            for h in hits
            if h.get("body")
        ]


class NullDataProvider(DataProvider):
    """Used when WEB_SEARCH_ENABLED=false. Content generation still works —
    it just can't ground claims in fresh web facts, only the user's brief."""

    def search(self, query: str, max_results: int) -> list[SearchResult]:
        return []


def build_data_provider() -> DataProvider:
    if not settings.web_search_enabled:
        return NullDataProvider()
    return DuckDuckGoDataProvider()


def gather_facts(provider: DataProvider, topics: list[str], per_topic: int | None = None) -> list[SearchResult]:
    """Search each topic and return a deduplicated (by URL) flat list."""
    per_topic = per_topic or settings.web_search_max_results
    seen_urls: set[str] = set()
    results: list[SearchResult] = []
    for topic in topics:
        for r in provider.search(topic, per_topic):
            if r.url and r.url in seen_urls:
                continue
            seen_urls.add(r.url)
            results.append(r)
    return results
