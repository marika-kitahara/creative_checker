from __future__ import annotations

from typing import Any

import pandas as pd

from models import AICheckItem, CheckResult
from utils import (
    annotation_matches,
    clean_text,
    extension_of,
    matches_rule,
    normalized_text,
    split_exclusion_texts,
)


PRIORITY_ORDER = {"高": 1, "中": 2, "低": 3}
JUDGMENT_ORDER = {"NG": 1, "要確認": 2, "OK": 3}


def is_active(row: pd.Series) -> bool:
    return clean_text(row.get("ステータス")) == "有効"


def make_result(
    file_name: str,
    relative_path: str,
    row: pd.Series,
    rule_type: str,
    matched_text: str,
    message: str,
    check_detail: str = "",
) -> CheckResult:
    return CheckResult(
        file_name=file_name,
        relative_path=relative_path,
        rule_id=clean_text(row.get("rule_id")),
        rule_type=rule_type,
        category=clean_text(row.get("カテゴリ")),
        judgment=clean_text(row.get("判定結果")) or "要確認",
        matched_text=matched_text,
        message=message,
        check_detail=check_detail,
        ai_check_required=clean_text(row.get("AI確認要否")),
        ai_check_category=clean_text(row.get("AI確認カテゴリ")),
        ai_check_question=clean_text(row.get("AIへの確認事項")),
        ai_priority=clean_text(row.get("AI確認優先度")) or "中",
    )


def sort_results(results: list[CheckResult]) -> list[CheckResult]:
    return sorted(
        results,
        key=lambda item: (
            JUDGMENT_ORDER.get(item.judgment, 99),
            PRIORITY_ORDER.get(item.ai_priority, 99),
            item.rule_id,
        ),
    )


def overall_judgment(results: list[CheckResult]) -> str:
    judgments = {result.judgment for result in results}
    if "NG" in judgments:
        return "NG"
    if "要確認" in judgments:
        return "要確認"
    return "OK"


def check_format_rules(
    file_name: str,
    relative_path: str,
    file_size_mb: float,
    width_px: int,
    height_px: int,
    master_df: pd.DataFrame,
) -> list[CheckResult]:
    results: list[CheckResult] = []
    extension = extension_of(file_name)
    actual_values: dict[str, Any] = {
        "拡張子": extension,
        "横幅px": width_px,
        "縦幅px": height_px,
        "ファイルサイズMB": file_size_mb,
    }

    for _, row in master_df.iterrows():
        if not is_active(row) or clean_text(row.get("媒体種別")) != "画像":
            continue
        item = clean_text(row.get("対象項目"))
        standard = clean_text(row.get("許可形式/基準値"))
        comparison = clean_text(row.get("大小区分"))
        actual = actual_values.get(item)
        if actual is None or not standard:
            continue

        violated = False
        if item == "拡張子":
            allowed = {
                value.strip().lower().lstrip(".")
                for value in standard.split(",")
                if value.strip()
            }
            violated = extension not in allowed
        else:
            try:
                actual_num = float(actual)
                standard_num = float(standard)
            except (TypeError, ValueError):
                continue
            if comparison == "以下":
                violated = actual_num > standard_num
            elif comparison == "以上":
                violated = actual_num < standard_num
            else:
                violated = actual_num != standard_num

        if violated:
            results.append(
                make_result(
                    file_name, relative_path, row, "媒体・形式条件",
                    f"{item}: {actual}",
                    clean_text(row.get("NG理由として出力する文言")),
                    f"基準値：{standard}",
                )
            )
    return results


def check_ng_words(
    file_name: str,
    relative_path: str,
    ocr_text: str,
    master_df: pd.DataFrame,
) -> list[CheckResult]:
    results = []
    for _, row in master_df.iterrows():
        if not is_active(row):
            continue
        expression = clean_text(row.get("NG表現"))
        if matches_rule(ocr_text, expression, clean_text(row.get("一致方法"))):
            results.append(
                make_result(
                    file_name, relative_path, row, "NGワード", expression,
                    clean_text(row.get("NG理由として出力する文言")),
                    f"推奨表現：{clean_text(row.get('OK/推奨表現'))}",
                )
            )
    return results


def condition_is_applicable(code: str, selected: set[str]) -> bool:
    code = clean_text(code).upper()
    return not code or code == "ALL" or code in selected


def check_required_words(
    file_name: str,
    relative_path: str,
    ocr_text: str,
    selected_conditions: set[str],
    master_df: pd.DataFrame,
) -> list[CheckResult]:
    results = []
    for _, row in master_df.iterrows():
        if not is_active(row):
            continue
        if not condition_is_applicable(
            clean_text(row.get("適用条件コード")),
            selected_conditions,
        ):
            continue

        target = clean_text(row.get("対象"))
        required_text = clean_text(row.get("必須文言"))

        if target == "人間確認":
            results.append(
                make_result(
                    file_name, relative_path, row, "必須文言・設定",
                    required_text,
                    clean_text(row.get("未検出時の出力文言")),
                    clean_text(row.get("適用条件メモ")),
                )
            )
            continue

        if not matches_rule(
            ocr_text,
            required_text,
            clean_text(row.get("一致方法")),
        ):
            results.append(
                make_result(
                    file_name, relative_path, row, "必須文言", required_text,
                    clean_text(row.get("未検出時の出力文言")),
                    clean_text(row.get("適用条件メモ")),
                )
            )
    return results


def check_warning_words(
    file_name: str,
    relative_path: str,
    ocr_text: str,
    master_df: pd.DataFrame,
) -> list[CheckResult]:
    results = []
    for _, row in master_df.iterrows():
        if not is_active(row):
            continue

        warning_word = clean_text(row.get("注意ワード"))
        exclusions = split_exclusion_texts(
            clean_text(row.get("除外文言"))
        )

        if any(
            annotation_matches(ocr_text, exclusion, "部分一致")
            for exclusion in exclusions
        ):
            continue

        if matches_rule(
            ocr_text,
            warning_word,
            clean_text(row.get("一致方法")),
        ):
            results.append(
                make_result(
                    file_name, relative_path, row, "注意ワード", warning_word,
                    clean_text(row.get("出力文言")),
                    clean_text(row.get("確認内容")),
                )
            )
    return results


def check_required_annotations(
    file_name: str,
    relative_path: str,
    ocr_text: str,
    master_df: pd.DataFrame,
) -> list[CheckResult]:
    results = []
    for _, row in master_df.iterrows():
        if not is_active(row):
            continue

        trigger = clean_text(row.get("検出ワード"))
        if not matches_rule(
            ocr_text,
            trigger,
            clean_text(row.get("一致方法")),
        ):
            continue

        annotation = clean_text(row.get("必須注釈"))
        found = annotation_matches(
            ocr_text,
            annotation,
            clean_text(row.get("注釈一致方法")),
        )

        if found:
            results.append(
                CheckResult(
                    file_name=file_name,
                    relative_path=relative_path,
                    rule_id=clean_text(row.get("rule_id")),
                    rule_type="訴求別必須注釈",
                    category=clean_text(row.get("カテゴリ")),
                    judgment="OK",
                    matched_text=trigger,
                    message="必須注釈を検出しました。",
                    check_detail=f"検出した注釈：{annotation}",
                    ai_check_required="不要",
                    ai_check_category=clean_text(row.get("AI確認カテゴリ")),
                    ai_check_question="",
                    ai_priority="低",
                )
            )
        else:
            results.append(
                make_result(
                    file_name, relative_path, row, "訴求別必須注釈", trigger,
                    clean_text(row.get("未検出時の出力文言")),
                    f"必要な注釈：{annotation}",
                )
            )
    return results


def run_checks(
    file_name: str,
    relative_path: str,
    image_bytes: bytes,
    image_info: dict[str, Any],
    ocr_text: str,
    selected_conditions: set[str],
    master: dict[str, pd.DataFrame],
) -> list[CheckResult]:
    size_mb = len(image_bytes) / (1024 * 1024)
    results = []
    results += check_format_rules(
        file_name, relative_path, size_mb,
        int(image_info["width_px"]), int(image_info["height_px"]),
        master["05_媒体_形式条件"],
    )
    results += check_ng_words(
        file_name, relative_path, ocr_text, master["01_NGワード"]
    )
    results += check_required_words(
        file_name, relative_path, ocr_text, selected_conditions,
        master["02_必須文言"],
    )
    results += check_warning_words(
        file_name, relative_path, ocr_text, master["03_注意ワード"]
    )
    results += check_required_annotations(
        file_name, relative_path, ocr_text,
        master["04_訴求別_必須注釈"],
    )
    return sort_results(results)


def build_ai_items(
    file_name: str,
    relative_path: str,
    python_results: list[CheckResult],
    common_ai_df: pd.DataFrame,
    selected_conditions: set[str],
) -> list[AICheckItem]:
    items: list[AICheckItem] = []

    for _, row in common_ai_df.iterrows():
        if not is_active(row):
            continue
        target = clean_text(row.get("対象"))
        if target not in {"", "ALL", "画像", "画像・動画"}:
            continue
        condition = clean_text(row.get("適用条件")).upper()
        if condition not in {"", "ALL"} and condition not in selected_conditions:
            continue
        items.append(
            AICheckItem(
                file_name=file_name,
                relative_path=relative_path,
                source="共通AI確認",
                rule_id=clean_text(row.get("ai_check_id")),
                category=clean_text(row.get("カテゴリ")),
                priority=clean_text(row.get("優先度")) or "中",
                question=clean_text(row.get("確認事項")),
                answer_required=clean_text(row.get("回答必須")),
            )
        )

    for result in python_results:
        if result.ai_check_required not in {"必要", "常時"}:
            continue
        if not result.ai_check_question:
            continue
        items.append(
            AICheckItem(
                file_name=file_name,
                relative_path=relative_path,
                source="Python検出結果",
                rule_id=result.rule_id,
                category=result.ai_check_category or result.category,
                priority=result.ai_priority or "中",
                question=result.ai_check_question,
                answer_required="はい",
            )
        )

    unique = []
    seen = set()
    for item in items:
        key = normalized_text(item.question)
        if key and key not in seen:
            seen.add(key)
            unique.append(item)

    return sorted(
        unique,
        key=lambda item: (
            PRIORITY_ORDER.get(item.priority, 99),
            item.category,
            item.rule_id,
        ),
    )
