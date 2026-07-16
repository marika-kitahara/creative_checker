from __future__ import annotations

import re
import unicodedata
from pathlib import PurePosixPath
from typing import Any

import pandas as pd


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def normalized_text(value: str) -> str:
    return re.sub(r"[\s\u3000]+", "", clean_text(value)).lower()


def normalized_annotation_text(value: str) -> str:
    text = unicodedata.normalize("NFKC", clean_text(value)).lower()
    text = re.sub(
        r"[\s\u3000※*＊・。、「」『』【】\[\]（）()：:;；!！?？]+",
        "",
        text,
    )
    for ending in ("です", "ます"):
        if text.endswith(ending):
            text = text[:-len(ending)]
    return text


def annotation_matches(
    ocr_text: str,
    required_annotation: str,
    match_method: str,
) -> bool:
    source = normalized_annotation_text(ocr_text)
    target = normalized_annotation_text(required_annotation)
    if not target:
        return False

    method = clean_text(match_method)
    if method == "完全一致":
        return source == target
    if method == "正規表現":
        try:
            return re.search(
                required_annotation,
                ocr_text,
                flags=re.IGNORECASE | re.MULTILINE,
            ) is not None
        except re.error:
            return False
    return target in source


def matches_rule(text: str, pattern: str, match_method: str) -> bool:
    source = normalized_text(text)
    target = normalized_text(pattern)
    if not target:
        return False

    method = clean_text(match_method)
    if method == "完全一致":
        return source == target

    if method == "正規表現":
        try:
            return re.search(
                pattern,
                text,
                flags=re.IGNORECASE | re.MULTILINE,
            ) is not None
        except re.error:
            return False

    if method == "数値付き":
        escaped = re.escape(clean_text(pattern))
        number_pattern = (
            r"(?:最大|約|およそ)?\s*"
            r"\d+(?:[.,]\d+)*(?:万|千|億)?"
            r"\s*" + escaped
        )
        return re.search(
            number_pattern,
            text,
            flags=re.IGNORECASE | re.MULTILINE,
        ) is not None

    return target in source


def split_exclusion_texts(value: str) -> list[str]:
    raw = clean_text(value)
    if not raw:
        return []
    return [
        item.strip()
        for item in re.split(r"[\r\n|｜]+", raw)
        if item.strip()
    ]


def safe_file_stem(file_name: str) -> str:
    stem = PurePosixPath(file_name.replace("\\", "/")).stem
    return re.sub(r'[\\/:*?"<>|]+', "_", stem).strip() or "creative"


def relative_upload_path(uploaded_file: Any) -> str:
    raw_name = clean_text(getattr(uploaded_file, "name", "uploaded_image"))
    parts = [
        part
        for part in raw_name.replace("\\", "/").split("/")
        if part not in {"", ".", ".."}
    ]
    return "/".join(parts) if parts else "uploaded_image"


def extension_of(file_name: str) -> str:
    return PurePosixPath(file_name).suffix.lower().lstrip(".")
