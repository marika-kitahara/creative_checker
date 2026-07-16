from __future__ import annotations

from dataclasses import dataclass


@dataclass
class CheckResult:
    file_name: str
    relative_path: str
    rule_id: str
    rule_type: str
    category: str
    judgment: str
    matched_text: str
    message: str
    check_detail: str
    ai_check_required: str
    ai_check_category: str
    ai_check_question: str
    ai_priority: str
    detected_by: str = "Python"


@dataclass
class FileSummary:
    file_name: str
    relative_path: str
    extension: str
    width_px: int | None
    height_px: int | None
    aspect_ratio: str
    color_mode: str
    file_size_mb: float
    sha256: str
    ocr_text: str
    overall_judgment: str
    ng_count: int
    review_count: int


@dataclass
class AICheckItem:
    file_name: str
    relative_path: str
    source: str
    rule_id: str
    category: str
    priority: str
    question: str
    answer_required: str


@dataclass
class OCRBlock:
    text: str
    box: list[list[float]]
    score: float | None = None


@dataclass
class OCRResult:
    text: str
    blocks: list[OCRBlock]
