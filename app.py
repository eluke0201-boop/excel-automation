"""
EC事業者向け 受注データ統合・集計ツール（Streamlit版）

ブラウザで動くローカルWebアプリ。クライアントはCSVをドラッグ&ドロップするだけで
統合済みExcelレポートをダウンロードできる。

起動方法:
    streamlit run app.py

ブラウザが自動で開いて http://localhost:8501 でアプリが表示される。
"""

import base64
import json
import math
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import plotly.graph_objects as go

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


# ============================================================
# DEMO MODE フラグ
# True: ポートフォリオ用のサンプルタイル UI を表示（Streamlit Cloud デプロイ版）
# False: 通常版（Windows バンドル等の配布版） — build_windows_bundle.py で自動的に False に書き換えられる
# ============================================================
DEMO_MODE = True

import pandas as pd
import streamlit as st

from merge_reports import (
    FALLBACK_PALETTE,
    _adjust_color_by_rank,
    build_daily_summary,
    build_platform_color_map,
    build_platform_summary,
    build_product_summary,
    detect_platform,
    get_platform_color,
    load_config,
    load_platform_csv,
    remove_duplicates,
    suggest_column_mapping,
    write_excel,
)


# ============================================================
# ページ設定
# ============================================================
st.set_page_config(
    page_title="EC受注データ統合ツール",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",  # サイドバーはデフォルトで閉じる
)


# ============================================================
# カスタムCSS（1画面に収まるようコンパクト化 + アップロード欄を見やすく）
# ============================================================
st.markdown(
    """
    <style>
    /* メインコンテナの上下余白を圧縮 */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1rem !important;
        max-width: 1100px !important;
    }
    /* Streamlit のデフォルトヘッダー帯を最小化 */
    header[data-testid="stHeader"] {
        height: 0;
        background: transparent;
    }
    /* ファイルアップロードのドロップゾーン */
    [data-testid="stFileUploaderDropzone"] {
        min-height: 180px;
        padding: 32px 24px;
        border: 2px dashed #4F46E5;
        background-color: #F5F3FF;
        transition: background-color 0.2s, border-color 0.2s;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        background-color: #EDE9FE;
        border-color: #6D28D9;
    }
    /* サンプルタイルからドラッグ中のハイライト（iframe 越しに JS でクラスを付与） */
    [data-testid="stFileUploaderDropzone"].tile-drag-active {
        background-color: #DDD6FE !important;
        border-color: #7C3AED !important;
        box-shadow: 0 0 0 3px rgba(124, 58, 237, 0.25) !important;
        transform: scale(1.01);
    }
    /* 親フレーム側のマーキー（範囲選択）の四角 */
    .marquee-rect {
        position: fixed;
        background: rgba(37, 99, 235, 0.12);
        border: 1.5px solid #2563EB;
        pointer-events: none;
        z-index: 10000;
        border-radius: 1px;
    }
    /* 親フレーム側のマルチドラッグ件数バッジ */
    .multi-drag-badge {
        position: fixed;
        background: #2563EB;
        color: white;
        font-weight: 700;
        font-size: 0.7rem;
        padding: 3px 8px;
        border-radius: 10px;
        pointer-events: none;
        z-index: 10001;
        box-shadow: 0 2px 6px rgba(0,0,0,0.25);
    }
    [data-testid="stFileUploaderDropzoneInstructions"] > div > span {
        font-size: 1.1rem;
        font-weight: 600;
    }
    [data-testid="stFileUploaderDropzone"] svg {
        width: 40px;
        height: 40px;
    }
    /* タイトル下のマージン削減 */
    h1, h2, h3 {
        margin-top: 0 !important;
        padding-top: 0 !important;
    }
    /* hr の上下余白を圧縮 */
    hr {
        margin: 0.8rem 0 !important;
    }
    /* alert系（success / info / warning）のpadding圧縮 */
    [data-testid="stAlert"] {
        padding: 0.5rem 1rem !important;
    }
    /* サイドバー本体と展開トグルを非表示（設定はダイアログで開く） */
    [data-testid="stSidebar"],
    [data-testid="stSidebarCollapsedControl"] {
        display: none !important;
    }
    /* 歯車ボタンを少しコンパクトに、目立たせない */
    [data-testid="stHorizontalBlock"] [data-testid="stButton"] button[kind="secondary"] {
        font-size: 1.2rem;
    }
    /* エキスパンダーをコンパクトに（特に設定ダイアログ内のリスト感を出す） */
    [data-testid="stExpander"] {
        margin-bottom: 4px !important;
        border-radius: 6px;
        border-color: #E5E7EB !important;
    }
    [data-testid="stExpander"] details > summary {
        padding: 6px 12px !important;
        min-height: auto !important;
    }
    [data-testid="stExpander"] details > summary p {
        font-size: 0.92rem !important;
        font-weight: 500;
    }
    /* エキスパンダー展開時の内側余白も詰める */
    [data-testid="stExpanderDetails"] {
        padding: 10px 14px !important;
    }
    /* プライマリボタンを目に優しい緑系に（デフォルトの赤を上書き） */
    button[kind="primary"],
    button[kind="primaryFormSubmit"] {
        background-color: #16A34A !important;
        border-color: #16A34A !important;
        color: white !important;
    }
    button[kind="primary"]:hover,
    button[kind="primaryFormSubmit"]:hover {
        background-color: #15803D !important;
        border-color: #15803D !important;
    }
    button[kind="primary"]:active,
    button[kind="primaryFormSubmit"]:active {
        background-color: #14532D !important;
        border-color: #14532D !important;
    }
    /* st.pills（チップ）─ さわやかなスカイブルー */
    [data-testid="stPills"] button,
    [data-testid="stPillsButton"],
    [data-testid="stButtonGroup"] button,
    [data-baseweb="button-group"] button {
        background-color: #BAE6FD !important;  /* sky-200 さわやかな水色 */
        border: 1px solid #7DD3FC !important;  /* sky-300 */
        color: #075985 !important;             /* sky-800 */
        transition: all 0.15s !important;
    }
    [data-testid="stPills"] button:hover,
    [data-testid="stPillsButton"]:hover,
    [data-testid="stButtonGroup"] button:hover,
    [data-baseweb="button-group"] button:hover {
        background-color: #7DD3FC !important;  /* sky-300 */
        border-color: #38BDF8 !important;      /* sky-400 */
        color: #0C4A6E !important;             /* sky-900 */
    }
    [data-testid="stPills"] button[aria-pressed="true"],
    [data-testid="stPillsButton"][aria-pressed="true"],
    [data-testid="stButtonGroup"] button[aria-pressed="true"] {
        background-color: #0284C7 !important;  /* sky-600 濃いスカイブルー */
        border-color: #0284C7 !important;
        color: white !important;
    }
    /* セカンダリボタン（「← 一覧に戻る」など）─ 極薄グレー */
    [data-testid="stButton"] button[kind="secondary"] {
        background-color: #F3F4F6 !important;  /* gray-100 とても薄く */
        border: 1px solid #E5E7EB !important;  /* gray-200 */
        color: #4B5563 !important;             /* gray-600 */
    }
    [data-testid="stButton"] button[kind="secondary"]:hover {
        background-color: #E5E7EB !important;
        border-color: #D1D5DB !important;
        color: #1F2937 !important;
    }
    /* ターシャリボタン（「🗑 削除」など）─ 優しい赤 */
    [data-testid="stButton"] button[kind="tertiary"] {
        background-color: #FEE2E2 !important;
        border: 1px solid #FCA5A5 !important;
        color: #991B1B !important;
    }
    [data-testid="stButton"] button[kind="tertiary"]:hover {
        background-color: #FECACA !important;
        border-color: #F87171 !important;
        color: #7F1D1D !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 自動判別ロジック
# ============================================================
def detect_platform_from_headers(headers: list[str], config: dict) -> str | None:
    best_match = None
    best_count = 0
    actual = set(headers)
    for platform_key, platform_config in config["platforms"].items():
        expected = set(platform_config["columns"].keys())
        match_count = len(expected & actual)
        if match_count > best_count:
            best_count = match_count
            best_match = platform_key
    return best_match if best_count >= 2 else None


def detect_platform_for_file(uploaded_file, config: dict) -> str | None:
    platform_key = detect_platform(uploaded_file.name, config)
    if platform_key is not None:
        return platform_key
    try:
        uploaded_file.seek(0)
        df_peek = pd.read_csv(uploaded_file, nrows=0)
        uploaded_file.seek(0)
        return detect_platform_from_headers(list(df_peek.columns), config)
    except Exception:
        return None


# ============================================================
# 設定ファイル読み込み・保存
# ============================================================
CONFIG_PATH = Path("config.json")
CANONICAL_COLS = ["注文番号", "注文日", "商品名", "個数", "金額"]


def save_config(cfg: dict) -> None:
    """config.json に書き戻す（既存の '_comment' などのキーも保持）。"""
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


try:
    config = load_config(CONFIG_PATH)
except Exception as e:
    st.error(f"⚠️ 設定ファイル読み込みエラー: {e}")
    st.stop()


# ============================================================
# 設定ダイアログ（歯車ボタンで開くモーダル、ビュー切替型）
# ============================================================
def _on_settings_dismiss():
    """ダイアログが X で閉じられたときにフラグをリセット"""
    st.session_state["settings_dialog_open"] = False


@st.dialog(
    "⚙️ 設定 — 対応プラットフォーム管理",
    width="large",
    on_dismiss=_on_settings_dismiss,
)
def show_settings_dialog():
    # ビュー状態（"list" / "edit_<key>" / "add" / "raw"）
    current_view = st.session_state.get("settings_view", "list")

    # ===== 編集ビュー =====
    if current_view.startswith("edit_"):
        plat_key = current_view[5:]
        if plat_key not in config["platforms"]:
            st.session_state["settings_view"] = "list"
            st.rerun()
        _render_edit_view(plat_key)
        return

    # ===== 新規追加ビュー =====
    if current_view == "add":
        _render_add_view()
        return

    # ===== JSON直接編集ビュー =====
    if current_view == "raw":
        _render_raw_view()
        return

    # ===== 一覧ビュー（デフォルト） =====
    st.caption("対応プラットフォームの一覧。クリックで編集画面へ進みます。")
    st.subheader("📦 対応プラットフォーム")

    platform_keys = list(config["platforms"].keys())
    labels = {k: config["platforms"][k]["label"] for k in platform_keys}

    # プラットフォーム選択チップ（文字数フィット・横並び）
    st.markdown("**プラットフォーム選択**")
    selected = st.pills(
        "プラットフォーム",
        options=platform_keys,
        format_func=lambda k: labels.get(k, k),
        selection_mode="single",
        label_visibility="collapsed",
        key="settings_list_chips",
    )
    if selected:
        st.session_state["settings_view"] = f"edit_{selected}"
        # 次回戻ってきたときに自動再選択されないようウィジェット状態をクリア
        st.session_state.pop("settings_list_chips", None)
        st.rerun()

    st.markdown("---")
    st.markdown("**詳細操作**")
    action_mode = st.pills(
        "詳細操作",
        options=["➕ 新規追加", "🛠 JSON直接編集"],
        selection_mode="single",
        label_visibility="collapsed",
        key="settings_action_chips",
    )
    if action_mode == "➕ 新規追加":
        st.session_state["settings_view"] = "add"
        st.session_state.pop("settings_action_chips", None)
        st.rerun()
    elif action_mode == "🛠 JSON直接編集":
        st.session_state["settings_view"] = "raw"
        st.session_state.pop("settings_action_chips", None)
        st.rerun()


def _render_edit_view(plat_key: str) -> None:
    """編集ビュー：選択されたプラットフォームを編集"""
    plat = config["platforms"][plat_key]

    # 保留中の自動推測を反映
    pending_key = f"pending_suggestions_{plat_key}"
    show_autofill_msg = False
    if pending_key in st.session_state:
        sugg = st.session_state.pop(pending_key)
        for canon, header in sugg.items():
            if header:
                st.session_state[f"col_{plat_key}_{canon}"] = header
        show_autofill_msg = True

    # ヘッダー：戻るボタン + 自動入力通知（同じ行に右寄せ）
    head_left, head_right = st.columns([1, 3])
    with head_left:
        if st.button("← 一覧に戻る", key=f"back_edit_{plat_key}"):
            st.session_state["settings_view"] = "list"
            st.rerun()
    with head_right:
        if show_autofill_msg:
            st.success("✨ 列マッピングを自動入力しました")

    st.markdown(f"##### ✏️ {plat['label']} の編集")

    # 🔍 CSVドロップで自動入力
    auto_csv = st.file_uploader(
        "🔍 CSVをドロップで列マッピングを自動推測（オプション）",
        type=["csv"],
        key=f"auto_csv_{plat_key}",
    )
    last_csv_key = f"last_auto_csv_{plat_key}"
    if auto_csv is not None and st.session_state.get(last_csv_key) != auto_csv.name:
        try:
            df_peek = pd.read_csv(auto_csv, nrows=0)
            headers = list(df_peek.columns)
            suggestions = suggest_column_mapping(headers)
            if any(suggestions.values()):
                st.session_state[pending_key] = suggestions
                st.session_state[last_csv_key] = auto_csv.name
                st.rerun()
            else:
                st.warning("⚠️ どの列も推測できませんでした。列名を確認してください。")
        except Exception as e:
            st.error(f"CSV読み込みエラー: {e}")

    col_a, col_b, col_c = st.columns([2, 2, 1])
    with col_a:
        new_label = st.text_input("表示名", value=plat["label"], key=f"label_{plat_key}")
    with col_b:
        new_prefix = st.text_input(
            "ファイル名の先頭（CSV判別用）",
            value=plat["filename_prefix"],
            key=f"prefix_{plat_key}",
        )
    with col_c:
        # 現在の色を取得（config > 既定 > FALLBACK の順）
        current_color = get_platform_color(plat["label"], position=0, platform_config=plat)
        new_color = st.color_picker(
            "色",
            value=current_color,
            key=f"color_{plat_key}",
            help="円グラフでの色（順位が下のものは薄めに自動調整）",
        )

    st.markdown("**列名マッピング** — CSV側の列名 → 統一スキーマ")
    inverse_map = {v: k for k, v in plat["columns"].items()}
    new_columns: dict = {}
    map_cols = st.columns(5)
    for i, canon in enumerate(CANONICAL_COLS):
        with map_cols[i]:
            default = inverse_map.get(canon, "")
            new_src = st.text_input(
                f"`{canon}` ←", value=default, key=f"col_{plat_key}_{canon}",
            )
            if new_src.strip():
                new_columns[new_src.strip()] = canon

    st.caption("💡 自動入力した内容に間違いがなければ「💾 更新」を押してください")

    btn_save, btn_del = st.columns(2)
    with btn_save:
        if st.button("💾 更新", key=f"save_{plat_key}", use_container_width=True, type="primary"):
            if len(new_columns) < 5:
                st.error("5つの列マッピングがすべて必要です")
            else:
                config["platforms"][plat_key]["label"] = new_label
                config["platforms"][plat_key]["filename_prefix"] = new_prefix.strip()
                config["platforms"][plat_key]["columns"] = new_columns
                config["platforms"][plat_key]["color"] = new_color
                save_config(config)
                st.success(f"{new_label} を更新しました ✓")
                st.session_state["settings_view"] = "list"
                st.rerun()
    with btn_del:
        if st.button(
            "🗑 削除",
            key=f"del_{plat_key}",
            use_container_width=True,
            type="tertiary",
        ):
            del config["platforms"][plat_key]
            save_config(config)
            st.session_state["settings_view"] = "list"
            st.rerun()


def _render_add_view() -> None:
    """新規プラットフォーム追加ビュー"""
    # 保留中の自動推測を反映
    show_autofill_msg = False
    if "pending_suggestions_new" in st.session_state:
        sugg = st.session_state.pop("pending_suggestions_new")
        for canon, header in sugg.items():
            if header:
                st.session_state[f"new_col_{canon}"] = header
        if "pending_prefix_new" in st.session_state:
            st.session_state["new_prefix"] = st.session_state.pop("pending_prefix_new")
        show_autofill_msg = True

    # ヘッダー：戻るボタン + 自動入力通知
    head_left, head_right = st.columns([1, 3])
    with head_left:
        if st.button("← 一覧に戻る", key="back_add"):
            st.session_state["settings_view"] = "list"
            st.rerun()
    with head_right:
        if show_autofill_msg:
            st.success("✨ 列マッピング・ファイル名先頭を自動入力しました")

    st.markdown("##### ➕ 新規プラットフォームの追加")

    # [Demo only] 左側にサンプル CSV タイルを並べる
    if DEMO_MODE:
        _c_tiles, _c_form = st.columns([1, 5])
        with _c_tiles:
            from demo_features import render_dialog_sample_tiles
            render_dialog_sample_tiles()
        _container = _c_form
    else:
        _container = st.container()

    with _container:

        auto_csv_new = st.file_uploader(
            "🔍 CSVをドロップで列マッピングを自動推測（オプション）",
            type=["csv"],
            key="auto_csv_new",
        )
        if (
            auto_csv_new is not None
            and st.session_state.get("last_auto_csv_new") != auto_csv_new.name
        ):
            try:
                df_peek = pd.read_csv(auto_csv_new, nrows=0)
                headers = list(df_peek.columns)
                suggestions = suggest_column_mapping(headers)
                if any(suggestions.values()):
                    st.session_state["pending_suggestions_new"] = suggestions
                    basename = auto_csv_new.name.rsplit(".", 1)[0]
                    parts = re.split(r"[_\-\.]", basename)
                    for p in parts:
                        if p.isalpha() and len(p) >= 3:
                            st.session_state["pending_prefix_new"] = p.lower()
                            break
                    st.session_state["last_auto_csv_new"] = auto_csv_new.name
                    st.rerun()
                else:
                    st.warning("⚠️ どの列も推測できませんでした。列名を確認してください。")
            except Exception as e:
                st.error(f"CSV読み込みエラー: {e}")

        with st.form("add_platform_form", clear_on_submit=False):
            col_y, col_z, col_color = st.columns([2, 2, 1])
            with col_y:
                new_label = st.text_input(
                    "表示名", placeholder="例: Yahoo!ショッピング", key="new_display_label",
                )
            with col_z:
                new_prefix = st.text_input(
                    "ファイル名の先頭",
                    placeholder="例: yahoo",
                    key="new_prefix",
                    help="この値が内部IDとしても使われます（英小文字推奨）",
                )
            with col_color:
                # 次に追加されるプラットフォームのデフォルト色（FALLBACK_PALETTE から循環）
                default_new_color = FALLBACK_PALETTE[
                    len(config["platforms"]) % len(FALLBACK_PALETTE)
                ]
                new_color = st.color_picker(
                    "色",
                    value=default_new_color,
                    key="new_color",
                    help="円グラフでの色（後から変更可）",
                )

            st.markdown("**列名マッピング** — CSV側の列名を入力")
            col_src_inputs: dict = {}
            map_cols = st.columns(5)
            for i, canon in enumerate(CANONICAL_COLS):
                with map_cols[i]:
                    col_src_inputs[canon] = st.text_input(
                        f"`{canon}` ←",
                        placeholder=f"例: {canon}",
                        key=f"new_col_{canon}",
                    )

            st.caption("💡 自動入力した内容に問題なければ「➕ 追加する」を押してください")

            submitted = st.form_submit_button(
                "➕ 追加する", use_container_width=True, type="primary"
            )
            if submitted:
                if not new_label or not new_prefix:
                    st.error("表示名・ファイル名先頭は必須です")
                elif new_prefix.strip() in config["platforms"]:
                    st.error(
                        f"ファイル名先頭 '{new_prefix.strip()}' は既に使われています。別の名前を入れてください。"
                    )
                elif any(not v.strip() for v in col_src_inputs.values()):
                    st.error("5つの列名すべてを入力してください")
                else:
                    # ファイル名先頭をそのまま内部キーとして使用
                    internal_key = new_prefix.strip()
                    config["platforms"][internal_key] = {
                        "filename_prefix": new_prefix.strip(),
                        "label": new_label.strip(),
                        "color": new_color,
                        "columns": {v.strip(): k for k, v in col_src_inputs.items()},
                    }
                    save_config(config)
                    st.success(f"{new_label} を追加しました ✓")
                    # クリア
                    for canon in CANONICAL_COLS:
                        st.session_state.pop(f"new_col_{canon}", None)
                    st.session_state.pop("new_display_label", None)
                    st.session_state.pop("new_prefix", None)
                    st.session_state.pop("new_color", None)
                    st.session_state.pop("last_auto_csv_new", None)
                    # ダイアログを閉じて、メイン画面に戻る（次回開いたら一覧から）
                    st.session_state["settings_view"] = "list"
                    st.session_state["settings_dialog_open"] = False
                    st.rerun()


def _render_raw_view() -> None:
    """JSON直接編集ビュー"""
    if st.button("← 一覧に戻る", key="back_raw"):
        st.session_state["settings_view"] = "list"
        st.rerun()

    st.markdown("##### 🛠 config.json を直接編集（上級者向け）")
    current_json = json.dumps(config, ensure_ascii=False, indent=2)
    edited_json = st.text_area(
        "設定ファイル全体",
        value=current_json,
        height=300,
        key="raw_config_editor",
        help="JSON形式で直接編集できます。商品名翻訳辞書なども設定可能。",
    )
    if st.button("💾 JSON を保存", use_container_width=True, type="primary"):
        try:
            new_cfg = json.loads(edited_json)
            save_config(new_cfg)
            st.success("保存しました ✓")
            st.session_state["settings_view"] = "list"
            st.rerun()
        except json.JSONDecodeError as e:
            st.error(f"JSON形式が正しくありません: {e}")


# ============================================================
# ヘッダー（コンパクト版・1行で済ませる）
# ============================================================
supported_labels = " / ".join([p["label"] for p in config["platforms"].values()])

# タイトル | 説明 | 歯車ボタン
col_title, col_caption, col_gear = st.columns([2, 2.5, 0.4])
with col_title:
    st.markdown("### 📊 EC受注データ統合・集計ツール")
with col_caption:
    st.caption(
        f"対応: **{supported_labels}** ／ "
        "データはこのPC内だけで処理（クラウド送信なし）"
    )
with col_gear:
    if st.button("⚙️", help="設定を開く", key="settings_btn", use_container_width=True):
        # ダイアログを開くフラグを立てる（ビュー初期化も）
        st.session_state["settings_dialog_open"] = True
        st.session_state["settings_view"] = "list"
        st.rerun()

# ダイアログ表示制御：フラグがTrueならダイアログを呼び出す（毎再実行時にチェック）
if st.session_state.get("settings_dialog_open", False):
    show_settings_dialog()


# ============================================================
# サンプルデータタイル（Demo only） — demo_features.py からインポート
# ============================================================
if DEMO_MODE:
    from demo_features import render_main_sample_tiles
    render_main_sample_tiles(config)


# ============================================================
# CSVアップロード
# ============================================================
# キーにカウンタを含めることで「全クリア」を実現
if "uploader_counter" not in st.session_state:
    st.session_state["uploader_counter"] = 0

uploaded_files = st.file_uploader(
    "CSVファイル（複数選択可、種類は自動判別）",
    type=["csv"],
    accept_multiple_files=True,
    key=f"main_file_uploader_{st.session_state['uploader_counter']}",
)

if not uploaded_files:
    st.info("👆 各ECショップから出力した受注CSVをまとめてドロップしてください。ファイル名や中身の列名から自動判別します。")
    st.stop()

# ファイルがある場合は「全クリア」ボタンを右上に小さく表示
_, col_clear = st.columns([5, 1])
with col_clear:
    if st.button(
        "🗑 全クリア",
        key="clear_all_files",
        use_container_width=True,
        help="アップロード済みファイルをすべて削除",
    ):
        st.session_state["uploader_counter"] += 1
        st.rerun()


# ============================================================
# 判別＋実行を1ブロックに集約
# ============================================================
files_by_platform: dict = {}
unknown_file_objs: list = []  # 未登録ファイル（オブジェクト）

for f in uploaded_files:
    platform_key = detect_platform_for_file(f, config)
    if platform_key:
        files_by_platform.setdefault(platform_key, []).append(f)
    else:
        unknown_file_objs.append(f)


# 判別結果と実行ボタンを横並びに
col_detect, col_run = st.columns([3, 1])

with col_detect:
    if files_by_platform:
        detection_lines = []
        for platform_key, files in files_by_platform.items():
            label = config["platforms"][platform_key]["label"]
            names = "、".join(f.name for f in files)
            detection_lines.append(f"✅ **{label}**: {names}")
        for f in unknown_file_objs:
            detection_lines.append(f"⚠️ 未登録形式: {f.name}")
        st.markdown("  \n".join(detection_lines))
    elif not unknown_file_objs:
        st.error("有効なCSVが認識できませんでした。")

with col_run:
    run = st.button("🚀 集計を実行", type="primary", use_container_width=True, disabled=not files_by_platform)


# 未登録ファイルがある場合：すべて一覧表示し、各行に登録ボタン
if unknown_file_objs:
    st.markdown("---")
    st.markdown(f"##### 🆕 未登録のCSVが {len(unknown_file_objs)} 件あります")
    st.caption("各ファイルの「🔧 設定に追加」を押して、推測結果を確認しながら登録してください。")

    for _idx, _f in enumerate(unknown_file_objs):
        try:
            _f.seek(0)
            _df_peek = pd.read_csv(_f, nrows=0)
            _f.seek(0)
            _headers = list(_df_peek.columns)
            _suggestions = suggest_column_mapping(_headers)
            _matched_count = sum(1 for v in _suggestions.values() if v)

            col_msg, col_reg = st.columns([3, 1])
            with col_msg:
                st.markdown(
                    f"📄 **{_f.name}** — 自動推測: "
                    f"**{_matched_count} / 5 項目**"
                )
            with col_reg:
                if st.button(
                    "🔧 設定に追加",
                    type="primary",
                    use_container_width=True,
                    key=f"register_unknown_{_idx}_{_f.name}",
                ):
                    st.session_state["pending_suggestions_new"] = _suggestions
                    _basename = _f.name.rsplit(".", 1)[0]
                    for _p in re.split(r"[_\-\.]", _basename):
                        if _p.isalpha() and len(_p) >= 3:
                            st.session_state["pending_prefix_new"] = _p.lower()
                            break
                    st.session_state["settings_dialog_open"] = True
                    st.session_state["settings_view"] = "add"
                    st.rerun()
        except Exception as e:
            st.warning(f"⚠️ `{_f.name}` の解析エラー: {e}")

if not run:
    st.stop()


# ============================================================
# 処理本体（ファイルはディスクに自動保存しない、メモリのみで保持）
# ============================================================
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_filename = f"統合レポート_{timestamp}.xlsx"

with st.spinner("処理中..."):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        frames = []

        for platform_key, files in files_by_platform.items():
            platform_config = config["platforms"][platform_key]
            for idx, uploaded_file in enumerate(files):
                save_name = f"{platform_config['filename_prefix']}_{idx}_{uploaded_file.name}"
                save_path = tmp_path / save_name
                with open(save_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                df = load_platform_csv(save_path, platform_key, config)
                frames.append(df)

        df = pd.concat(frames, ignore_index=True)
        df = remove_duplicates(df)
        product_summary = build_product_summary(df)
        platform_summary = build_platform_summary(df)
        daily_summary = build_daily_summary(df)

        # 一時ファイルに書き出してbytesを取得（メモリ上のみ、ユーザーがドラッグするまでディスクに保存しない）
        tmp_output = tmp_path / output_filename
        write_excel(df, product_summary, platform_summary, daily_summary, tmp_output, config=config)
        excel_bytes = tmp_output.read_bytes()


# ============================================================
# 結果表示（KPI + ダウンロードを横並びに）
# ============================================================
st.markdown("---")

col1, col2, col3 = st.columns(3)
col1.metric("総注文数", f"{len(df):,} 件")
col2.metric("総販売個数", f"{int(df['個数'].sum()):,} 個")
col3.metric("売上合計", f"¥{int(df['金額'].sum()):,}")


# ============================================================
# 出力ファイルをドラッグ可能なタイルとして表示
# ============================================================
st.markdown("---")
st.markdown("### 🎉 集計完了！出力ファイルを取り出してください")

mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
b64_excel = base64.b64encode(excel_bytes).decode()
file_size_kb = len(excel_bytes) / 1024
size_str = f"{file_size_kb:.1f}KB" if file_size_kb < 1024 else f"{file_size_kb/1024:.1f}MB"

file_tile_html = f"""
<!DOCTYPE html>
<html>
<head>
<style>
  body {{
    margin: 0;
    padding: 10px 8px 24px 8px;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', sans-serif;
  }}
  @keyframes bounce {{
    0%, 100% {{ transform: translateY(0); }}
    50% {{ transform: translateY(-4px); }}
  }}
  @keyframes pulse-shadow {{
    0%, 100% {{ box-shadow: 0 4px 12px rgba(22, 163, 74, 0.15); }}
    50% {{ box-shadow: 0 6px 20px rgba(22, 163, 74, 0.30); }}
  }}
  .file-tile {{
    display: flex;
    align-items: center;
    gap: 18px;
    padding: 20px 24px;
    background: linear-gradient(135deg, #FFFFFF 0%, #F0FDF4 100%);
    border: 2px solid #16A34A;
    border-radius: 12px;
    cursor: grab;
    user-select: none;
    transition: all 0.2s ease;
    animation: pulse-shadow 2.4s ease-in-out infinite;
  }}
  .file-tile:hover {{
    border-color: #15803D;
    transform: translateY(-2px) scale(1.01);
    animation: none;
    box-shadow: 0 8px 24px rgba(22, 163, 74, 0.35);
  }}
  .file-tile:active {{
    cursor: grabbing;
    transform: translateY(0) scale(1);
  }}
  .icon-box {{
    display: flex;
    align-items: center;
    justify-content: center;
    width: 56px;
    height: 56px;
    background: #16A34A;
    border-radius: 10px;
    flex-shrink: 0;
    box-shadow: 0 2px 8px rgba(22, 163, 74, 0.25);
  }}
  .icon-box svg {{
    width: 32px;
    height: 32px;
    fill: white;
  }}
  .details {{
    display: flex;
    flex-direction: column;
    min-width: 0;
    flex: 1;
  }}
  .name {{
    font-size: 1.05rem;
    font-weight: 700;
    color: #14532D;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }}
  .meta {{
    font-size: 0.8rem;
    color: #6B7280;
    margin-top: 4px;
  }}
  .drag-cta {{
    display: flex;
    align-items: center;
    gap: 8px;
    background: #16A34A;
    color: white;
    padding: 10px 16px;
    border-radius: 8px;
    font-weight: 700;
    font-size: 0.9rem;
    flex-shrink: 0;
    animation: bounce 1.8s ease-in-out infinite;
  }}
</style>
</head>
<body>
<div class="file-tile"
     draggable="true"
     ondragstart="event.dataTransfer.setData('DownloadURL', '{mime}:{output_filename}:data:{mime};base64,{b64_excel}');">
  <div class="icon-box">
    <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zm-1 7V3.5L18.5 9H13z"/></svg>
  </div>
  <div class="details">
    <div class="name">{output_filename}</div>
    <div class="meta">{size_str} · Excelファイル</div>
  </div>
  <div class="drag-cta">
    🖱 ここをドラッグ →
  </div>
</div>
</body>
</html>
"""

# ファイルタイル(左：メイン) + ダウンロードカード(右：サブ) を横並びに
col_tile, col_dl = st.columns([2, 1])

with col_tile:
    st.components.v1.html(file_tile_html, height=144)
    st.caption("👆 このタイルを掴んで、デスクトップやエクスプローラーにドロップしてください")

with col_dl:
    # カード型のダウンロードリンク（HTML5のdownload属性を使った直接ダウンロード）
    download_card_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
      body {{
        margin: 0;
        padding: 10px 8px 24px 8px;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      }}
      .download-card {{
        display: flex;
        align-items: center;
        gap: 14px;
        padding: 18px 20px;
        background: white;
        border: 2px solid #E5E7EB;
        border-radius: 12px;
        text-decoration: none;
        color: inherit;
        cursor: pointer;
        transition: all 0.2s ease;
      }}
      .download-card:hover {{
        border-color: #4F46E5;
        background: linear-gradient(135deg, #FFFFFF 0%, #F5F3FF 100%);
        box-shadow: 0 6px 18px rgba(79, 70, 229, 0.18);
        transform: translateY(-2px);
      }}
      .download-card:active {{
        transform: translateY(0);
      }}
      .dl-icon-box {{
        width: 48px;
        height: 48px;
        background: #EDE9FE;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
        transition: background 0.2s;
      }}
      .download-card:hover .dl-icon-box {{
        background: #DDD6FE;
      }}
      .dl-icon-box svg {{
        width: 26px;
        height: 26px;
        fill: #4F46E5;
      }}
      .dl-details {{
        display: flex;
        flex-direction: column;
        min-width: 0;
      }}
      .dl-title {{
        font-size: 0.95rem;
        font-weight: 700;
        color: #1F2937;
      }}
      .dl-sub {{
        font-size: 0.75rem;
        color: #6B7280;
        margin-top: 3px;
      }}
    </style>
    </head>
    <body>
    <a class="download-card" download="{output_filename}"
       href="data:{mime};base64,{b64_excel}">
      <div class="dl-icon-box">
        <svg viewBox="0 0 24 24"><path d="M5 20h14v-2H5v2zM19 9h-4V3H9v6H5l7 7 7-7z"/></svg>
      </div>
      <div class="dl-details">
        <div class="dl-title">クリックでダウンロード</div>
        <div class="dl-sub">ドラッグできない場合はこちら</div>
      </div>
    </a>
    </body>
    </html>
    """
    st.components.v1.html(download_card_html, height=144)


# プラットフォーム色は merge_reports.get_platform_color() を使用
# （登録済み: PLATFORM_COLORS_MPL の固定色 / 未登録: FALLBACK_PALETTE から自動割当）

tab1, tab2, tab3, tab4 = st.tabs(
    ["🏪 プラットフォーム別", "📦 商品別", "📅 日別推移", "📋 全注文一覧"]
)

with tab1:
    # 各プラットフォームの色: config の "color" を優先 → 推奨色 → FALLBACK_PALETTE
    n_platforms = len(platform_summary)
    color_map = build_platform_color_map(config, list(platform_summary["プラットフォーム"]))
    adjusted_colors = [
        _adjust_color_by_rank(
            color_map[p_label],
            rank=i,
            total=n_platforms,
        )
        for i, p_label in enumerate(platform_summary["プラットフォーム"])
    ]

    fig = go.Figure(
        data=[
            go.Pie(
                labels=platform_summary["プラットフォーム"],
                values=platform_summary["売上合計"],
                hole=0.62,
                textinfo="none",
                hovertemplate=(
                    "<b>%{label}</b><br>"
                    "売上: ¥%{value:,}<br>"
                    "構成比: %{percent}<extra></extra>"
                ),
                marker=dict(
                    colors=adjusted_colors,
                    line=dict(color="white", width=2),
                ),
                sort=False,
                direction="clockwise",
                rotation=0,
                domain=dict(x=[0.30, 0.70], y=[0.10, 0.90]),
            )
        ]
    )

    # 各スライスから引き出し線でラベルを描画する annotations を構築
    total_value = platform_summary["売上合計"].sum()
    center_x, center_y = 0.5, 0.5
    # 実際のパイの外周（縦長レイアウトでの実測値）
    pie_edge_radius_x = 0.19
    pie_edge_radius_y = 0.34
    # ラベル位置（パイから少し離す）
    label_radius_x = 0.30
    label_radius_y = 0.46

    cumulative = 0.0
    annotations = []
    shapes = []  # 引き出し線（2本の直線セグメント）をshapesで描く
    for i, row in platform_summary.iterrows():
        label = row["プラットフォーム"]
        value = int(row["売上合計"])
        fraction = value / total_value
        mid_fraction = cumulative + fraction / 2
        cumulative += fraction

        # 12時を起点に時計回り
        angle_rad = math.pi / 2 - mid_fraction * 2 * math.pi
        cos_a = math.cos(angle_rad)
        sin_a = math.sin(angle_rad)

        # ① パイ外周の点（引き出し線の起点）
        arrow_x = center_x + pie_edge_radius_x * cos_a
        arrow_y = center_y + pie_edge_radius_y * sin_a

        # ② 屈折点（ヒジ）：全スライスで同じくらいの曲がり角度になるよう固定オフセット
        # 図のアスペクト比を考慮（横長キャンバスなので x方向は控えめに、y方向は大きめに）
        horizontal_dir = 1 if cos_a >= 0 else -1
        vertical_dir = 1 if sin_a >= 0 else -1
        bend_offset_x = 0.035 * horizontal_dir
        bend_offset_y = 0.07 * vertical_dir
        bend_x = arrow_x + bend_offset_x
        bend_y = arrow_y + bend_offset_y

        # ③ ラベル位置（屈折点から水平に延長した先）
        horizontal_extension = 0.06
        label_x = bend_x + horizontal_extension * horizontal_dir
        label_y = bend_y  # 水平セグメントなので bend と同じ y

        # 線セグメント1: パイ外周 → ヒジ（放射方向の斜線）
        shapes.append(
            dict(
                type="line",
                xref="paper",
                yref="paper",
                x0=arrow_x,
                y0=arrow_y,
                x1=bend_x,
                y1=bend_y,
                line=dict(color="#6B7280", width=1),
            )
        )

        # 線セグメント2: ヒジ → ラベル（水平線）
        shapes.append(
            dict(
                type="line",
                xref="paper",
                yref="paper",
                x0=bend_x,
                y0=bend_y,
                x1=label_x,
                y1=label_y,
                line=dict(color="#6B7280", width=1),
            )
        )

        # テキストラベル（社名と%、控えめな階層）
        # スライスと同じ調整済み色を使用（小さいスライスのラベルは控えめに）
        slice_color = adjusted_colors[i] if i < len(adjusted_colors) else "#1F2937"
        annotations.append(
            dict(
                x=label_x,
                y=label_y,
                xref="paper",
                yref="paper",
                showarrow=False,
                text=(
                    f"<span style='font-size:12px;color:#374151;font-weight:600;'>{label}</span>"
                    f"<br>"
                    f"<span style='font-size:13px;font-weight:700;color:{slice_color};'>"
                    f"{fraction*100:.1f}%</span>"
                ),
                font=dict(family="-apple-system, sans-serif"),
                align="left" if cos_a >= 0 else "right",
                xanchor="left" if cos_a >= 0 else "right",
                yanchor="middle",
            )
        )

    # ドーナツの中央に「合計売上」を表示
    annotations.append(
        dict(
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
            text=(
                f"<span style='font-size:10px;color:#4B5563;letter-spacing:0.05em;font-weight:600;'>売上合計</span>"
                f"<br>"
                f"<span style='font-size:17px;font-weight:700;color:#111827;'>"
                f"¥{int(total_value):,}</span>"
                f"<br>"
                f"<span style='font-size:10px;color:#4B5563;font-weight:500;'>{int(platform_summary['注文件数'].sum()):,}件</span>"
            ),
            font=dict(family="-apple-system, sans-serif"),
            align="center",
        )
    )

    fig.update_layout(
        title=dict(
            text="<b>プラットフォーム別 売上構成</b>",
            x=0.5,
            font=dict(size=16, color="#111827"),
        ),
        height=500,
        margin=dict(t=70, b=40, l=20, r=20),
        showlegend=False,
        annotations=annotations,
        shapes=shapes,
        paper_bgcolor="white",
        plot_bgcolor="white",
    )
    st.plotly_chart(fig, use_container_width=True)
    with st.expander("📋 元データを見る"):
        st.dataframe(platform_summary, use_container_width=True, hide_index=True)

with tab2:
    df_sorted = product_summary.sort_values("売上合計", ascending=True)
    fig = go.Figure(
        go.Bar(
            x=df_sorted["売上合計"],
            y=df_sorted["商品名"],
            orientation="h",
            text=[f"¥{int(v):,}" for v in df_sorted["売上合計"]],
            textposition="outside",
            textfont=dict(size=12, color="#1F2937"),
            marker=dict(
                color=df_sorted["売上合計"],
                colorscale=[[0, "#C7D2FE"], [1, "#4F46E5"]],
                showscale=False,
                line=dict(width=0),
            ),
            customdata=df_sorted[["注文件数", "販売個数"]].values,
            hovertemplate=(
                "<b>%{y}</b><br>"
                "売上: ¥%{x:,}<br>"
                "注文件数: %{customdata[0]} 件<br>"
                "販売個数: %{customdata[1]} 個<extra></extra>"
            ),
        )
    )
    max_value = df_sorted["売上合計"].max()
    fig.update_layout(
        title=dict(
            text=f"<b>商品別 売上ランキング</b>（{len(product_summary)}商品）",
            x=0.5,
            font=dict(size=16),
        ),
        xaxis_title="売上合計（円）",
        yaxis_title="",
        height=max(420, 80 + 40 * len(product_summary)),
        margin=dict(t=70, b=60, l=20, r=180),
        plot_bgcolor="white",
        xaxis=dict(
            showgrid=True,
            gridcolor="#F3F4F6",
            automargin=True,
            # 金額ラベル分のヘッドルームを20%確保（"¥21,000" などが見切れないように）
            range=[0, max_value * 1.20],
        ),
        yaxis=dict(showgrid=False, automargin=True),
    )
    st.plotly_chart(fig, use_container_width=True)
    with st.expander("📋 元データを見る"):
        st.dataframe(product_summary, use_container_width=True, hide_index=True)

with tab3:
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=daily_summary["注文日"],
            y=daily_summary["売上合計"],
            mode="lines+markers",
            fill="tozeroy",
            line=dict(color="#4F46E5", width=2.5, shape="spline"),
            marker=dict(size=8, color="#4F46E5", line=dict(width=2, color="white")),
            fillcolor="rgba(79, 70, 229, 0.12)",
            customdata=daily_summary[["注文件数"]].values,
            hovertemplate=(
                "<b>%{x}</b><br>"
                "売上: ¥%{y:,}<br>"
                "注文件数: %{customdata[0]} 件<extra></extra>"
            ),
            name="売上",
        )
    )
    fig.update_layout(
        title=dict(
            text=f"<b>日別 売上推移</b>（{len(daily_summary)}日分）",
            x=0.5,
            font=dict(size=16),
        ),
        xaxis_title="注文日",
        yaxis_title="売上合計（円）",
        height=460,
        # 両端の日付ラベル（斜め45度）が切れないよう左右余白を大きく
        margin=dict(t=70, b=110, l=100, r=80),
        hovermode="x unified",
        plot_bgcolor="white",
        xaxis=dict(
            showgrid=True,
            gridcolor="#F3F4F6",
            tickangle=-45,
            automargin=True,
            type="category",
            # 両端のカテゴリの内側にパディング（最初と最後の点を端から離す）
            range=[-0.7, len(daily_summary) - 0.3],
        ),
        yaxis=dict(showgrid=True, gridcolor="#F3F4F6", automargin=True),
    )
    st.plotly_chart(fig, use_container_width=True)
    with st.expander("📋 元データを見る"):
        st.dataframe(daily_summary, use_container_width=True, hide_index=True)

with tab4:
    st.dataframe(df, use_container_width=True, hide_index=True, height=480)
