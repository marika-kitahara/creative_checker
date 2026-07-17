from __future__ import annotations

import hashlib
from datetime import datetime
from io import BytesIO
from pathlib import PurePosixPath
from typing import Any

import pandas as pd
import streamlit as st
from PIL import Image, UnidentifiedImageError

from checker import (
    build_ai_items,
    overall_judgment,
    run_checks,
)
from excel_writer import create_prompt_zip, create_result_excel
from highlight import create_highlighted_image
from master import load_master_from_bytes
from models import CheckResult, FileSummary
from ocr import get_ocr_package_status, run_ocr
from utils import (
    clean_text,
    extension_of,
    relative_upload_path,
    safe_file_stem,
)


APP_TITLE = "クリエイティブチェッカー"
SUPPORTED = {"jpg", "jpeg", "png"}
JUDGMENT_ORDER = {"NG": 1, "要確認": 2, "OK": 3}
MASTER_URL = "https://rak.box.com/s/0bw4j0ceukpt2is4wbjvwjskaohxejn5"


def inspect_image(image_bytes: bytes) -> dict[str, Any]:
    try:
        image = Image.open(BytesIO(image_bytes))
        image.verify()
        image = Image.open(BytesIO(image_bytes))
        width, height = image.size
        return {
            "width_px": width,
            "height_px": height,
            "aspect_ratio": f"{width / height:.4f}" if height else "",
            "color_mode": clean_text(image.mode),
        }
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise ValueError("画像ファイルを正常に読み込めません。") from exc


def build_prompt(
    summary: FileSummary,
    results: list[CheckResult],
    ai_items,
    condition_labels: list[str],
) -> str:
    result_text = "\n".join(
        f"- [{item.judgment}] {item.rule_id}／{item.category}\n"
        f"  検出・不足内容：{item.matched_text or '―'}\n"
        f"  出力内容：{item.message}\n"
        f"  確認補足：{item.check_detail or '―'}"
        for item in results
    ) or "明確なNG・要確認項目は検出されませんでした。"

    ai_text = "\n".join(
        f"{number}. [{item.priority}] [{item.category}] {item.question}"
        for number, item in enumerate(ai_items, 1)
    ) or "クリエイティブ全体の視認性と誤認リスクを確認してください。"

    conditions = "、".join(condition_labels) if condition_labels else "共通条件のみ"

    return f"""添付された広告クリエイティブを確認してください。

この確認は掲載可否を決定する最終審査ではありません。
Pythonによる機械判定結果を参考に、人間が確認すべき箇所を抽出する一次チェックとして回答してください。

【対象ファイル】
ファイル名：{summary.file_name}
相対パス：{summary.relative_path}
画像サイズ：{summary.width_px} × {summary.height_px}px
適用条件：{conditions}

【OCR抽出テキスト】
{summary.ocr_text or "OCRテキストを取得できませんでした。画像を直接確認してください。"}

【Pythonによる判定結果】
{result_text}

【重点確認事項】
{ai_text}

【確認時の注意】
- 法務上の適法性を断定しないでください。
- 掲載可能・掲載不可の最終判断は行わないでください。
- 判断できない内容は推測せず「要確認」としてください。
- OCR結果より添付画像を優先してください。

【回答形式】
総合判定候補：問題なし／要確認／修正推奨

確認結果：
1.
- ルールID：
- 確認項目：
- 判定：
- 該当箇所：
- 理由：
- 人間が確認すべき内容：
""".strip()


def render_card(
    summary,
    results,
    ai_items,
    prompt,
    image_bytes,
    highlighted_bytes,
    highlighted_count,
):
    icon = {"NG": "🔴", "要確認": "🟡", "OK": "🟢"}.get(
        summary.overall_judgment, "⚪"
    )
    with st.expander(
        f"{icon} {summary.relative_path}｜{summary.overall_judgment}",
        expanded=summary.overall_judgment != "OK",
    ):
        image_col, info_col = st.columns([1, 1])
        with image_col:
            original_tab, highlight_tab = st.tabs(["元画像", "指摘箇所を強調"])
            with original_tab:
                st.image(image_bytes, caption=summary.relative_path)
            with highlight_tab:
                st.image(
                    highlighted_bytes,
                    caption=f"強調箇所 {highlighted_count}件",
                )
                st.caption("赤：NG　オレンジ：要確認　緑：OK")
                if highlighted_count == 0:
                    st.info("画像上で位置を特定できる指摘はありませんでした。")

        with info_col:
            st.write(f"**サイズ**：{summary.width_px} × {summary.height_px}px")
            st.write(f"**容量**：{summary.file_size_mb:.3f}MB")
            st.write(f"**NG**：{summary.ng_count}件")
            st.write(f"**要確認**：{summary.review_count}件")

        result_tab, ocr_tab, ai_tab, prompt_tab = st.tabs(
            ["Python判定", "OCR結果", "AI確認事項", "AIプロンプト"]
        )
        with result_tab:
            if results:
                st.dataframe(pd.DataFrame([{
                    "判定": r.judgment, "ルールID": r.rule_id,
                    "カテゴリ": r.category, "検出・不足内容": r.matched_text,
                    "結果": r.message, "確認補足": r.check_detail,
                } for r in results]), width="stretch", hide_index=True)
            else:
                st.success("明確なNG・要確認項目はありませんでした。")
        with ocr_tab:
            st.text_area(
                "抽出テキスト",
                summary.ocr_text,
                height=260,
                key=f"ocr_{summary.sha256}",
            )
        with ai_tab:
            st.dataframe(pd.DataFrame([{
                "優先度": i.priority, "ルールID": i.rule_id,
                "カテゴリ": i.category, "確認事項": i.question,
                "回答必須": i.answer_required,
            } for i in ai_items]), width="stretch", hide_index=True)
        with prompt_tab:
            st.code(prompt, language=None)
            st.download_button(
                "このAIプロンプトをダウンロード",
                prompt.encode("utf-8-sig"),
                f"{safe_file_stem(summary.file_name)}_AI確認プロンプト.txt",
                "text/plain",
                key=f"prompt_{summary.sha256}",
            )


def main() -> None:
    if "uploader_version" not in st.session_state:
        st.session_state.uploader_version = 0

    st.set_page_config(page_title=APP_TITLE, page_icon="🔎", layout="wide")
    st.title(APP_TITLE)
    st.caption(
        "Pythonでルール判定し、AI確認事項と確認用プロンプトを生成します。"
        "最終判断は人間が行ってください。"
    )

    with st.sidebar:
        st.header("1. 判定マスタ")
        st.info("📄 マスタのフォーマットはこちら")
        st.link_button("マスタフォーマットを開く", MASTER_URL)
        master_file = st.file_uploader(
            "マスタExcel", type=["xlsx"], key="master_uploader"
        )

        st.header("2. 追加ルール（任意）")
        platforms = st.multiselect(
            "掲載媒体", ["X", "Instagram", "YouTube"]
        )
        benefit = st.checkbox("画像内で特典・ポイントを訴求している")

        conditions = set()
        labels = []
        mapping = {"X": "X", "Instagram": "INSTAGRAM", "YouTube": "YOUTUBE"}
        for platform in platforms:
            conditions.add(mapping[platform])
            labels.append(platform)
        if benefit:
            conditions.add("BENEFIT")
            labels.append("特典・ポイント記載あり")

        st.header("3. OCR")
        available, status = get_ocr_package_status()
        if available:
            st.success(f"OCR利用可能：{status}")
        else:
            st.error(f"OCR初期化エラー：{status}")
        use_ocr = st.checkbox(
            "OCRを実行する",
            value=available,
            disabled=not available,
        )

    st.subheader("チェック対象画像")
    upload_col, clear_col = st.columns([5, 1])
    with upload_col:
        selected = st.file_uploader(
            "画像ファイルまたはフォルダ",
            type=sorted(SUPPORTED),
            accept_multiple_files=True,
            key=f"images_{st.session_state.uploader_version}",
        )
    with clear_col:
        st.write("")
        st.write("")
        if st.button("全クリア", disabled=not bool(selected), width="stretch"):
            st.session_state.uploader_version += 1
            st.rerun()

    uploaded = []
    seen = set()
    for item in selected or []:
        path = relative_upload_path(item)
        data = item.getvalue()
        key = (path.lower(), hashlib.sha256(data).hexdigest())
        if key not in seen:
            seen.add(key)
            uploaded.append(item)

    run = st.button(
        "クリエイティブチェックを実行",
        type="primary",
        disabled=not (master_file and uploaded),
        width="stretch",
    )
    if not run:
        return

    try:
        master = load_master_from_bytes(master_file.getvalue())
    except Exception as exc:
        st.error(f"マスタを読み込めませんでした。\n\n{exc}")
        return

    summaries = []
    all_results = []
    all_ai_items = []
    prompts = {}
    errors = []
    original_map = {}
    highlight_map = {}
    highlight_count_map = {}
    results_map = {}
    ai_map = {}

    progress = st.progress(0, text="チェックを開始します。")

    for index, item in enumerate(uploaded, 1):
        path = relative_upload_path(item)
        file_name = PurePosixPath(path).name
        data = item.getvalue()
        progress.progress((index - 1) / len(uploaded), text=f"{path} を処理中...")

        try:
            ext = extension_of(file_name)
            if ext not in SUPPORTED:
                raise ValueError(f"対応外形式：{ext}")

            info = inspect_image(data)
            ocr_warning = ""
            if use_ocr:
                try:
                    ocr_result = run_ocr(data)
                except Exception as exc:
                    ocr_result = None
                    ocr_warning = str(exc)
            else:
                ocr_result = None

            ocr_text = ocr_result.text if ocr_result else ""
            blocks = ocr_result.blocks if ocr_result else []

            results = run_checks(
                file_name, path, data, info, ocr_text, conditions, master
            )
            if ocr_warning:
                results.append(CheckResult(
                    file_name, path, "SYS_OCR", "システム", "OCR",
                    "要確認", "OCR未実行",
                    "OCRを実行できなかったため文字ルールは未判定です。",
                    ocr_warning, "不要", "OCR", "", "高",
                ))

            ai_items = build_ai_items(
                file_name, path, results,
                master["06_AI確認項目"], conditions,
            )
            highlighted, highlighted_count = create_highlighted_image(
                data, blocks, results
            )

            summary = FileSummary(
                file_name=file_name,
                relative_path=path,
                extension=ext,
                width_px=info["width_px"],
                height_px=info["height_px"],
                aspect_ratio=info["aspect_ratio"],
                color_mode=info["color_mode"],
                file_size_mb=round(len(data) / (1024 * 1024), 4),
                sha256=hashlib.sha256(data).hexdigest(),
                ocr_text=ocr_text,
                overall_judgment=overall_judgment(results),
                ng_count=sum(r.judgment == "NG" for r in results),
                review_count=sum(r.judgment == "要確認" for r in results),
            )
            prompt = build_prompt(summary, results, ai_items, labels)

            summaries.append(summary)
            all_results.extend(results)
            all_ai_items.extend(ai_items)
            prompts[path] = prompt
            original_map[path] = data
            highlight_map[path] = highlighted
            highlight_count_map[path] = highlighted_count
            results_map[path] = results
            ai_map[path] = ai_items

        except Exception as exc:
            errors.append({"relative_path": path, "エラー内容": str(exc)})

    progress.progress(1.0, text="チェックが完了しました。")
    if not summaries:
        st.error("正常に処理できた画像がありませんでした。")
        return

    st.balloons()
    st.divider()
    st.subheader("集計")
    cols = st.columns(4)
    cols[0].metric("処理済み", len(summaries))
    cols[1].metric("NG", sum(r.judgment == "NG" for r in all_results))
    cols[2].metric("要確認", sum(r.judgment == "要確認" for r in all_results))

    excel = create_result_excel(
        summaries, all_results, all_ai_items,
        prompts, errors, original_map,
    )
    prompt_zip = create_prompt_zip(prompts)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    col1, col2 = st.columns(2)
    col1.download_button(
        "結果Excelをダウンロード",
        excel,
        f"creative_check_result_{stamp}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    col2.download_button(
        "AIプロンプト一式をZIPでダウンロード",
        prompt_zip,
        f"ai_prompts_{stamp}.zip",
        "application/zip",
        width="stretch",
    )

    st.divider()
    st.subheader("ファイル別結果")
    for summary in sorted(
        summaries,
        key=lambda item: (
            JUDGMENT_ORDER.get(item.overall_judgment, 99),
            item.relative_path.lower(),
        ),
    ):
        path = summary.relative_path
        render_card(
            summary, results_map[path], ai_map[path], prompts[path],
            original_map[path], highlight_map[path], highlight_count_map[path],
        )


if __name__ == "__main__":
    main()
