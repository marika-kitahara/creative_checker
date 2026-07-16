from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st

from utils import clean_text


REQUIRED_MASTER_COLUMNS: dict[str, list[str]] = {
    "01_NGワード": [
        "rule_id", "ステータス", "カテゴリ", "対象", "NG表現", "一致方法",
        "判定結果", "NG理由として出力する文言",
        "AI確認要否", "AI確認カテゴリ", "AIへの確認事項", "AI確認優先度",
    ],
    "02_必須文言": [
        "rule_id", "ステータス", "カテゴリ", "対象", "適用条件コード",
        "必須文言", "一致方法", "判定結果", "未検出時の出力文言",
        "AI確認要否", "AI確認カテゴリ", "AIへの確認事項", "AI確認優先度",
    ],
    "03_注意ワード": [
        "rule_id", "ステータス", "カテゴリ", "対象", "注意ワード",
        "一致方法", "除外文言", "判定結果", "確認内容", "出力文言",
        "AI確認要否", "AI確認カテゴリ", "AIへの確認事項", "AI確認優先度",
    ],
    "04_訴求別_必須注釈": [
        "rule_id", "ステータス", "カテゴリ", "対象", "検出ワード",
        "一致方法", "必須注釈", "注釈一致方法", "判定結果",
        "未検出時の出力文言", "AI確認要否", "AI確認カテゴリ",
        "AIへの確認事項", "AI確認優先度",
    ],
    "05_媒体_形式条件": [
        "rule_id", "ステータス", "媒体/用途", "媒体種別", "対象項目",
        "許可形式/基準値", "大小区分", "判定結果",
        "NG理由として出力する文言", "AI確認要否",
        "AI確認カテゴリ", "AIへの確認事項", "AI確認優先度",
    ],
    "06_AI確認項目": [
        "ai_check_id", "ステータス", "対象", "カテゴリ", "適用条件",
        "確認事項", "優先度", "回答必須", "備考",
    ],
}


@st.cache_data(show_spinner=False)
def load_master_from_bytes(master_bytes: bytes) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(BytesIO(master_bytes), engine="openpyxl")
    missing_sheets = [
        name for name in REQUIRED_MASTER_COLUMNS if name not in xls.sheet_names
    ]
    if missing_sheets:
        raise ValueError("必要なシートがありません：" + "、".join(missing_sheets))

    master: dict[str, pd.DataFrame] = {}
    errors: list[str] = []

    for sheet_name, required_columns in REQUIRED_MASTER_COLUMNS.items():
        df = pd.read_excel(
            BytesIO(master_bytes),
            sheet_name=sheet_name,
            dtype=str,
            engine="openpyxl",
        )
        df.columns = [clean_text(column) for column in df.columns]
        df = df.dropna(how="all").fillna("")

        missing_columns = [
            column for column in required_columns if column not in df.columns
        ]
        if missing_columns:
            errors.append(
                f"{sheet_name}：不足列 " + "、".join(missing_columns)
            )

        id_column = "ai_check_id" if sheet_name == "06_AI確認項目" else "rule_id"
        if id_column in df.columns:
            ids = df[id_column].map(clean_text)
            duplicated = ids[
                (ids != "") & ids.duplicated(keep=False)
            ].unique().tolist()
            if duplicated:
                errors.append(
                    f"{sheet_name}：ID重複 " + "、".join(map(str, duplicated))
                )

        master[sheet_name] = df

    if errors:
        raise ValueError("\n".join(errors))
    return master
