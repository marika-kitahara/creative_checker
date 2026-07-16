from __future__ import annotations

import re
from io import BytesIO

from PIL import Image, ImageDraw, ImageFont

from models import CheckResult, OCRBlock
from utils import clean_text, normalized_annotation_text


def _result_terms(result: CheckResult) -> list[str]:
    if result.rule_type in {"必須文言", "必須文言・設定", "システム"}:
        return []

    terms = []
    if result.matched_text:
        if not (
            result.rule_type == "媒体・形式条件"
            and ":" in result.matched_text
        ):
            terms.append(result.matched_text)

    prefix = "検出した注釈："
    if result.check_detail.startswith(prefix):
        terms.append(result.check_detail[len(prefix):].strip())

    return [term for term in dict.fromkeys(terms) if term]


def _matches(block_text: str, term: str) -> bool:
    block = normalized_annotation_text(block_text)
    target = normalized_annotation_text(term)
    return bool(target and target in block)


def _color(judgment: str) -> tuple[int, int, int]:
    return {
        "NG": (220, 30, 30),
        "要確認": (255, 140, 0),
        "OK": (40, 160, 70),
    }.get(judgment, (80, 80, 80))


def create_highlighted_image(
    image_bytes: bytes,
    ocr_blocks: list[OCRBlock],
    results: list[CheckResult],
) -> tuple[bytes, int]:
    image = Image.open(BytesIO(image_bytes)).convert("RGB")
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()
    count = 0
    drawn = set()

    for result in results:
        terms = _result_terms(result)
        if not terms:
            continue

        for block in ocr_blocks:
            if len(block.box) < 4:
                continue
            if not any(_matches(block.text, term) for term in terms):
                continue

            points = [
                (int(round(point[0])), int(round(point[1])))
                for point in block.box
                if len(point) >= 2
            ]
            if len(points) < 4:
                continue

            key = (result.rule_id, tuple(points))
            if key in drawn:
                continue
            drawn.add(key)

            color = _color(result.judgment)
            draw.line(points + [points[0]], fill=color, width=5)

            left = min(point[0] for point in points)
            top = min(point[1] for point in points)
            label = f"{result.judgment} {result.rule_id}"
            bbox = draw.textbbox((0, 0), label, font=font)
            width = bbox[2] - bbox[0] + 10
            height = bbox[3] - bbox[1] + 8
            label_top = max(0, top - height)
            draw.rectangle(
                [left, label_top, min(image.width, left + width), top],
                fill=color,
            )
            draw.text(
                (left + 5, label_top + 3),
                label,
                fill=(255, 255, 255),
                font=font,
            )
            count += 1

    output = BytesIO()
    image.save(output, format="PNG")
    return output.getvalue(), count
