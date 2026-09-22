"""Pluggable text-to-image backend for the "картинки внутри слайда" bonus task.

The ТЗ requires an open-weights (Apache-2.0/MIT), <=20B model on the
qualifying round, served either from the team's own GPU or an external
API provider. ``OpenAICompatibleImageClient`` covers both: point it at
any server implementing the OpenAI ``/images/generations`` schema
(a local Automatic1111/ComfyUI-in-front-of-an-OpenAI-shim, a vLLM image
endpoint, a hosted API — whatever the team lands on).

``PlaceholderImageClient`` is the network-free fallback: a generated
gradient card with the prompt text on it, so slides that call for an
image still render something in-place during local dev instead of a
broken picture placeholder.
"""
from __future__ import annotations

import hashlib
import io
import logging
from abc import ABC, abstractmethod

from ..config import settings

logger = logging.getLogger(__name__)


class ImageClient(ABC):
    @abstractmethod
    def generate(self, prompt: str, *, width: int = 1024, height: int = 768) -> bytes:
        """Return PNG bytes for the given prompt."""


class OpenAICompatibleImageClient(ImageClient):
    def __init__(self, base_url: str, api_key: str | None, model: str):
        from openai import OpenAI

        self._client = OpenAI(base_url=base_url, api_key=api_key or "unused")
        self._model = model

    def generate(self, prompt: str, *, width: int = 1024, height: int = 768) -> bytes:
        import base64

        size = f"{_round_to_supported(width)}x{_round_to_supported(height)}"
        resp = self._client.images.generate(model=self._model, prompt=prompt, size=size, n=1)
        b64 = resp.data[0].b64_json
        if not b64:
            raise ValueError("image endpoint returned no image data")
        return base64.b64decode(b64)


def _round_to_supported(dim: int) -> int:
    # Most SD-family endpoints want multiples of 64 and reject arbitrary sizes.
    return max(512, (dim // 64) * 64)


class PlaceholderImageClient(ImageClient):
    """Deterministic gradient-card placeholder — used when IMAGE_BASE_URL
    is not configured, so the pipeline still produces a complete deck."""

    _PALETTE = [
        ((79, 129, 189), (155, 187, 89)),
        ((192, 80, 77), (128, 100, 162)),
        ((75, 172, 198), (247, 150, 70)),
    ]

    def generate(self, prompt: str, *, width: int = 1024, height: int = 768) -> bytes:
        from PIL import Image, ImageDraw, ImageFont

        idx = int(hashlib.sha1(prompt.encode("utf-8")).hexdigest(), 16) % len(self._PALETTE)
        c1, c2 = self._PALETTE[idx]
        img = Image.new("RGB", (width, height), c1)
        draw = ImageDraw.Draw(img)
        for y in range(height):
            t = y / max(height - 1, 1)
            r = int(c1[0] + (c2[0] - c1[0]) * t)
            g = int(c1[1] + (c2[1] - c1[1]) * t)
            b = int(c1[2] + (c2[2] - c1[2]) * t)
            draw.line([(0, y), (width, y)], fill=(r, g, b))

        label = (prompt[:80] + "…") if len(prompt) > 80 else prompt
        try:
            font = ImageFont.load_default()
        except Exception:  # pragma: no cover - Pillow always ships a default font
            font = None
        draw.rectangle([(0, height - 90), (width, height)], fill=(0, 0, 0))
        draw.text((20, height - 65), f"[placeholder image] {label}", fill=(255, 255, 255), font=font)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()


def build_image_client() -> ImageClient:
    if settings.image_base_url:
        return OpenAICompatibleImageClient(
            base_url=settings.image_base_url,
            api_key=settings.image_api_key,
            model=settings.image_model,
        )
    return PlaceholderImageClient()
