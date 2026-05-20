"""
EC事業者向け 受注データ統合・集計ツール

複数プラットフォーム（楽天市場・Amazon・自社BASE）から出力された
受注CSVを自動で統合し、商品別・プラットフォーム別の集計レポートを
グラフ付きのExcelで出力する。

使い方:
    python merge_reports.py <入力フォルダ> <出力ファイル名>

例:
    python merge_reports.py ./sample_data ./output.xlsx

ファイル名規則:
    rakuten_*.csv  → 楽天市場形式として読み込み
    amazon_*.csv   → Amazon形式として読み込み（商品名は日本語に変換）
    base_*.csv     → 自社BASE形式として読み込み

機能:
    - 各プラットフォームの列名のゆれを統一スキーマに変換
    - Amazon英語商品名を日本語に自動マッピング
    - 重複行の自動削除
    - 「全注文」「商品別集計」「プラットフォーム別集計」「日別売上推移」シート出力
    - 商品別売上の棒グラフをExcelに埋め込み
"""

import sys
from pathlib import Path

# Windows のターミナル（CP932）でも日本語を文字化けせず表示するためにUTF-8に切り替え
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
# プラットフォーム別の列名マッピング
# 各プラットフォームの列名 → 統一スキーマ
# ============================================================

PLATFORM_COLUMN_MAPS = {
    "rakuten": {
        "注文番号": "注文番号",
        "注文日": "注文日",
        "商品名": "商品名",
        "個数": "個数",
        "金額": "金額",
    },
    "amazon": {
        "order-id": "注文番号",
        "purchase-date": "注文日",
        "product-name": "商品名",
        "quantity": "個数",
        "item-total": "金額",
    },
    "base": {
        "order_number": "注文番号",
        "order_date": "注文日",
        "item_name": "商品名",
        "qty": "個数",
        "total_price": "金額",
    },
}

# ============================================================
# Amazon英語商品名 → 日本語商品名のマッピング
# 商品別集計のために統一する
# ============================================================

AMAZON_PRODUCT_NAME_MAP = {
    "Natural Stone Pierced Earrings Rose Quartz": "天然石ピアス（ローズクォーツ）",
    "Natural Stone Pierced Earrings Amethyst": "天然石ピアス（アメジスト）",
    "Silver Chain Necklace": "シルバーチェーンネックレス",
    "14kgf Pearl Necklace": "14kgf一粒淡水パールネックレス",
    "Leather Bracelet": "レザーブレスレット",
    "Natural Stone Bracelet Turquoise": "天然石ブレスレット（ターコイズ）",
    "Silver Ring Simple": "シルバーリング（シンプル）",
    "Natural Stone Ring Garnet": "天然石リング（ガーネット）",
    "Hair Barrette Marble Pattern": "バレッタ（大理石柄）",
    "Hair Clip Pearl": "ヘアクリップ（パール）",
}

PLATFORM_LABEL = {
    "rakuten": "楽天市場",
    "amazon": "Amazon",
    "base": "自社BASE",
}

CANONICAL_COLUMNS = ["注文番号", "注文日", "商品名", "個数", "金額", "プラットフォーム"]


def detect_platform(filename: str) -> str | None:
    """ファイル名先頭からプラットフォームを判定する。"""
    lower = filename.lower()
    for key in PLATFORM_COLUMN_MAPS:
        if lower.startswith(key):
            return key
    return None


def load_platform_csv(file: Path, platform: str) -> pd.DataFrame:
    """プラットフォーム別のCSVを読み込み、統一スキーマに変換する。"""
    df = pd.read_csv(file)

    # 列名を統一スキーマにリネーム
    column_map = PLATFORM_COLUMN_MAPS[platform]
    df = df.rename(columns=column_map)

    # 統一スキーマに必要な列だけ残す（余分な列はカット）
    keep_columns = [c for c in column_map.values() if c in df.columns]
    df = df[keep_columns].copy()

    # Amazonのみ: 英語商品名 → 日本語に変換
    if platform == "amazon":
        df["商品名"] = df["商品名"].map(AMAZON_PRODUCT_NAME_MAP).fillna(df["商品名"])

    # 注文日をdatetime型に統一（文字列でもパースして揃える）
    df["注文日"] = pd.to_datetime(df["注文日"]).dt.strftime("%Y-%m-%d")

    # プラットフォーム列を付与
    df["プラットフォーム"] = PLATFORM_LABEL[platform]

    return df[CANONICAL_COLUMNS]


def load_all(folder: Path) -> pd.DataFrame:
    """フォルダ内の対象CSVをすべて統一スキーマで読み込み、統合する。"""
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"フォルダが見つかりません: {folder}")

    frames = []
    for file in sorted(folder.iterdir()):
        if file.suffix.lower() != ".csv":
            continue

        platform = detect_platform(file.name)
        if platform is None:
            print(f"  ⚠ スキップ（不明な形式）: {file.name}")
            continue

        print(f"  読み込み: {file.name} → {PLATFORM_LABEL[platform]} として処理")
        df = load_platform_csv(file, platform)
        frames.append(df)

    if not frames:
        raise ValueError(
            "対象CSVが見つかりませんでした。\n"
            "ファイル名は rakuten_*.csv / amazon_*.csv / base_*.csv の形式で配置してください。"
        )

    return pd.concat(frames, ignore_index=True)


def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """注文番号ベースで重複削除（万一の二重エクスポート対策）。"""
    before = len(df)
    df = df.drop_duplicates(subset=["注文番号"]).reset_index(drop=True)
    after = len(df)
    if before != after:
        print(f"  重複削除: {before} 行 → {after} 行 ({before - after} 行削除)")
    return df


def build_product_summary(df: pd.DataFrame) -> pd.DataFrame:
    """商品別の売上集計を作る。"""
    summary = (
        df.groupby("商品名", as_index=False)
        .agg(注文件数=("注文番号", "count"), 販売個数=("個数", "sum"), 売上合計=("金額", "sum"))
        .sort_values("売上合計", ascending=False)
        .reset_index(drop=True)
    )
    return summary


def build_platform_summary(df: pd.DataFrame) -> pd.DataFrame:
    """プラットフォーム別の売上集計を作る。"""
    summary = (
        df.groupby("プラットフォーム", as_index=False)
        .agg(注文件数=("注文番号", "count"), 販売個数=("個数", "sum"), 売上合計=("金額", "sum"))
        .sort_values("売上合計", ascending=False)
        .reset_index(drop=True)
    )
    total = summary["売上合計"].sum()
    summary["売上構成比"] = (summary["売上合計"] / total * 100).round(1).astype(str) + "%"
    return summary


def build_daily_summary(df: pd.DataFrame) -> pd.DataFrame:
    """日別の売上推移を作る。"""
    summary = (
        df.groupby("注文日", as_index=False)
        .agg(注文件数=("注文番号", "count"), 売上合計=("金額", "sum"))
        .sort_values("注文日")
        .reset_index(drop=True)
    )
    return summary


def style_header(ws, color: str = "4F46E5") -> None:
    """ヘッダー行に色付けと白文字を適用する。"""
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")


def autosize_columns(ws) -> None:
    """列幅を内容に合わせてざっくり自動調整する。"""
    for column in ws.columns:
        max_length = max(
            (len(str(cell.value)) for cell in column if cell.value is not None),
            default=10,
        )
        # 日本語は1文字で約2幅扱いなので少し広めに
        ws.column_dimensions[column[0].column_letter].width = min(max_length * 1.5 + 2, 40)


def write_excel(
    df: pd.DataFrame,
    product_summary: pd.DataFrame,
    platform_summary: pd.DataFrame,
    daily_summary: pd.DataFrame,
    output: Path,
) -> None:
    """4シート構成のExcelを出力する。"""
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="全注文", index=False)
        product_summary.to_excel(writer, sheet_name="商品別集計", index=False)
        platform_summary.to_excel(writer, sheet_name="プラットフォーム別集計", index=False)
        daily_summary.to_excel(writer, sheet_name="日別売上推移", index=False)

    # 装飾とグラフを追加
    wb = load_workbook(output)

    for sheet_name in ["全注文", "商品別集計", "プラットフォーム別集計", "日別売上推移"]:
        ws = wb[sheet_name]
        style_header(ws)
        autosize_columns(ws)

    # 商品別 売上グラフ
    ws_product = wb["商品別集計"]
    n_products = len(product_summary)
    if n_products > 0:
        chart = BarChart()
        chart.type = "bar"
        chart.style = 11
        chart.title = "商品別 売上合計（円）"
        chart.x_axis.title = "売上合計"
        chart.y_axis.title = "商品"

        data_ref = Reference(ws_product, min_col=4, min_row=1, max_row=1 + n_products)
        cats_ref = Reference(ws_product, min_col=1, min_row=2, max_row=1 + n_products)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.height = 12
        chart.width = 22
        ws_product.add_chart(chart, "F2")

    # 日別売上 推移グラフ
    ws_daily = wb["日別売上推移"]
    n_days = len(daily_summary)
    if n_days > 0:
        chart = LineChart()
        chart.style = 12
        chart.title = "日別 売上推移（円）"
        chart.x_axis.title = "日付"
        chart.y_axis.title = "売上合計"

        data_ref = Reference(ws_daily, min_col=3, min_row=1, max_row=1 + n_days)
        cats_ref = Reference(ws_daily, min_col=1, min_row=2, max_row=1 + n_days)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.height = 10
        chart.width = 20
        ws_daily.add_chart(chart, "E2")

    wb.save(output)


def print_summary_to_console(
    df: pd.DataFrame,
    product_summary: pd.DataFrame,
    platform_summary: pd.DataFrame,
) -> None:
    """ターミナルにも結果のサマリーを表示する。"""
    total_orders = len(df)
    total_units = int(df["個数"].sum())
    total_revenue = int(df["金額"].sum())

    print()
    print("=" * 50)
    print("[集計サマリー]")
    print("=" * 50)
    print(f"  総注文数:   {total_orders} 件")
    print(f"  総販売個数: {total_units} 個")
    print(f"  売上合計:   ¥{total_revenue:,}")
    print()
    print("  プラットフォーム別:")
    for _, row in platform_summary.iterrows():
        print(
            f"    {row['プラットフォーム']:<10} "
            f"{row['注文件数']:>3}件 / ¥{int(row['売上合計']):>8,} ({row['売上構成比']})"
        )
    print()
    print("  売上TOP3:")
    for i, (_, row) in enumerate(product_summary.head(3).iterrows(), start=1):
        print(f"    {i}. {row['商品名']:<30} ¥{int(row['売上合計']):>7,}")
    print("=" * 50)


def main() -> None:
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    input_folder = Path(sys.argv[1])
    output_file = Path(sys.argv[2])

    print(f"[1/5] ファイル読み込み: {input_folder}")
    df = load_all(input_folder)
    print(f"  → 合計 {len(df)} 行")

    print("[2/5] データクリーニング")
    df = remove_duplicates(df)

    print("[3/5] 集計シート生成")
    product_summary = build_product_summary(df)
    platform_summary = build_platform_summary(df)
    daily_summary = build_daily_summary(df)

    print(f"[4/5] Excel出力: {output_file}")
    write_excel(df, product_summary, platform_summary, daily_summary, output_file)

    print("[5/5] 完了！")
    print_summary_to_console(df, product_summary, platform_summary)
    print(f"\n  Excelファイル: {output_file.resolve()}")


if __name__ == "__main__":
    main()
