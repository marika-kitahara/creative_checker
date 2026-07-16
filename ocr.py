from __future__ import annotations

from io import BytesIO
from typing import Any

import streamlit as st
from PIL import Image

from models import OCRBlock, OCRResult
from utils import clean_text


def get_ocr_package_status() -> tuple[bool, str]:
    try:
        from rapidocr import RapidOCR  # noqa: F401
        return True, "rapidocr"
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"


@st.cache_resource(show_spinner=False)
def get_ocr_engine():
    from rapidocr import RapidOCR
    return RapidOCR()


def _to_list(value: Any) -> list:
    """NumPy配列・tuple・list・Noneを安全にlist化する。"""
    if value is None:
        return []
    if hasattr(value, "tolist"):
        value = value.tolist()
    if isinstance(value, tuple):
        return list(value)
    if isinstance(value, list):
        return value
    return []


def _normalize_box(raw_box: Any) -> list[list[float]]:
    points = _to_list(raw_box)
    normalized: list[list[float]] = []
    for point in points:
        point_values = _to_list(point)
        if len(point_values) >= 2:
            normalized.append(
                [float(point_values[0]), float(point_values[1])]
            )
    return normalized


@st.cache_data(show_spinner=False)
def run_ocr(image_bytes: bytes) -> OCRResult:
    """
    OCRは1回だけ実行し、全文と座標を同時に返す。
    RapidOCR 3系と旧形式の双方に対応する。
    """
    image = Image.open(BytesIO(image_bytes)).convert("RGB")
    engine = get_ocr_engine()
    output = engine(image)

    blocks: list[OCRBlock] = []

    # RapidOCR 3.x
    if hasattr(output, "txts"):
        texts = _to_list(getattr(output, "txts", None))
        boxes = _to_list(getattr(output, "boxes", None))
        scores = _to_list(getattr(output, "scores", None))

        for index, value in enumerate(texts):
            text = clean_text(value)
            if not text:
                continue
            box = _normalize_box(boxes[index]) if index < len(boxes) else []
            score = None
            if index < len(scores) and scores[index] is not None:
                try:
                    score = float(scores[index])
                except (TypeError, ValueError):
                    score = None
            blocks.append(OCRBlock(text=text, box=box, score=score))

        return OCRResult(
            text="\n".join(block.text for block in blocks),
            blocks=blocks,
        )

    # 旧形式：(result, elapsed)
    result = output[0] if isinstance(output, tuple) else output
    rows = _to_list(result)

    for item in rows:
        item_values = _to_list(item)
        if len(item_values) < 2:
            continue
        text = clean_text(item_values[1])
        if not text:
            continue
        box = _normalize_box(item_values[0])
        score = None
        if len(item_values) >= 3 and item_values[2] is not None:
            try:
                score = float(item_values[2])
            except (TypeError, ValueError):
                score = None
        blocks.append(OCRBlock(text=text, box=box, score=score))

    return OCRResult(
        text="\n".join(block.text for block in blocks),
        blocks=blocks,
    )
