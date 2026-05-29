"""
Windows 用の配布パッケージを作る（Python 同梱版）。

実行:
    python build_windows_bundle.py

出力:
    windows_bundle/  ← このフォルダごと zip にして配布
        python/         (Python 同梱)
        app.py 等
        run.bat         (ダブルクリックで起動)
        使い方.txt
"""

import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

# ---- 設定 ----
ROOT = Path(__file__).resolve().parent
DIST = ROOT / "windows_bundle"
PY_VERSION = "3.12.7"
PY_EMBED_URL = f"https://www.python.org/ftp/python/{PY_VERSION}/python-{PY_VERSION}-embed-amd64.zip"
GET_PIP_URL = "https://bootstrap.pypa.io/get-pip.py"

REQUIRED_PACKAGES = [
    "streamlit",
    "pandas",
    "openpyxl",
    "matplotlib",
    "plotly",
    "tabulate",
]

APP_FILES = ["app.py", "merge_reports.py", "config.json", "requirements.txt"]
SAMPLE_DIRS = ["sample_data", "sample_data_extended", "sample_data_for_test", "docs"]


def step(n, msg):
    print(f"\n[{n}] {msg}")


def main():
    # ---- 1. クリーン ----
    step(1, "既存の windows_bundle/ を削除")
    if DIST.exists():
        shutil.rmtree(DIST)
    DIST.mkdir(parents=True)
    py_dir = DIST / "python"
    py_dir.mkdir()

    # ---- 2. Python embeddable をDL ----
    step(2, f"Python {PY_VERSION} embeddable をダウンロード（約10MB）")
    embed_zip = DIST / "_python_embed.zip"
    urllib.request.urlretrieve(PY_EMBED_URL, embed_zip)
    print(f"   サイズ: {embed_zip.stat().st_size:,} バイト")

    # ---- 3. 展開 ----
    step(3, "展開")
    with zipfile.ZipFile(embed_zip) as zf:
        zf.extractall(py_dir)
    embed_zip.unlink()

    # ---- 4. _pth ファイルを修正（site-packages を有効化） ----
    step(4, "_pth ファイルを修正（site-packages を有効化）")
    pth_files = list(py_dir.glob("python*._pth"))
    if not pth_files:
        sys.exit("❌ _pth ファイルが見つかりません")
    pth = pth_files[0]
    content = pth.read_text()
    if "#import site" in content:
        content = content.replace("#import site", "import site")
    elif "import site" not in content:
        content += "\nimport site\n"
    pth.write_text(content)
    print(f"   修正後の {pth.name}:\n{content}")

    # ---- 5. get-pip.py をDLしてpipをセットアップ ----
    step(5, "pip をブートストラップ")
    get_pip = py_dir / "get-pip.py"
    urllib.request.urlretrieve(GET_PIP_URL, get_pip)
    subprocess.run(
        [str(py_dir / "python.exe"), str(get_pip)],
        check=True,
    )
    get_pip.unlink()

    # ---- 6. 依存パッケージをインストール（バンドル内に確実に入れる） ----
    step(6, f"パッケージをインストール: {', '.join(REQUIRED_PACKAGES)}")
    python_exe = py_dir / "python.exe"
    target_dir = py_dir / "Lib" / "site-packages"
    # --target で明示的にバンドル内にインストール（システムのを参照しない）
    subprocess.run(
        [
            str(python_exe), "-m", "pip", "install",
            "--target", str(target_dir),
            "--upgrade",
            "--no-warn-script-location",
        ] + REQUIRED_PACKAGES,
        check=True,
    )

    # ---- 7. アプリファイルをコピー ----
    step(7, "アプリファイルをコピー")
    for f in APP_FILES:
        src = ROOT / f
        if src.exists():
            if f == "app.py":
                # 配布版は DEMO_MODE = False に書き換える（サンプルタイル機能を無効化）
                content = src.read_text(encoding="utf-8")
                if "DEMO_MODE = True" not in content:
                    sys.exit(
                        "❌ app.py に 'DEMO_MODE = True' が見つかりません。"
                        "build_windows_bundle.py の前提が崩れています。"
                    )
                content = content.replace("DEMO_MODE = True", "DEMO_MODE = False", 1)
                (DIST / f).write_text(content, encoding="utf-8")
                print(f"   {f} (DEMO_MODE = False に書き換え済み)")
            else:
                shutil.copy(src, DIST / f)
                print(f"   {f}")
    # demo_features.py はバンドルにコピーしない（DEMO_MODE=False なので import されない）
    print("   [skip] demo_features.py (配布版では不要)")

    # サンプルデータフォルダ
    for d in SAMPLE_DIRS:
        src = ROOT / d
        if src.exists():
            shutil.copytree(src, DIST / d)
            print(f"   {d}/")

    # ---- 8. run.bat を作成 ----
    step(8, "run.bat を作成")
    run_bat_content = (
        "@echo off\r\n"
        'cd /d "%~dp0"\r\n'
        "echo ============================================\r\n"
        "echo EC受注データ統合ツール を起動しています...\r\n"
        "echo ブラウザが自動で開きます。少しお待ちください。\r\n"
        "echo ============================================\r\n"
        'python\\python.exe -m streamlit run app.py\r\n'
        "pause\r\n"
    )
    (DIST / "run.bat").write_text(run_bat_content, encoding="cp932")

    # ---- 9. 使い方.txt を作成 ----
    step(9, "使い方.txt を作成")
    readme = (
        "EC受注データ統合ツール - 使い方\r\n"
        "================================\r\n"
        "\r\n"
        "【起動方法】\r\n"
        "  run.bat をダブルクリックしてください。\r\n"
        "  → 黒いウィンドウが開き、その後 ブラウザが自動で開きます。\r\n"
        "  → ブラウザでアプリ画面が表示されたら準備OK。\r\n"
        "\r\n"
        "【使い方】\r\n"
        "  1. 各ECショップの受注CSVをドラッグ&ドロップ\r\n"
        "  2. 「集計を実行」をクリック\r\n"
        "  3. 結果のExcelをダウンロード（タイルをドラッグ）\r\n"
        "\r\n"
        "【終了方法】\r\n"
        "  ブラウザを閉じてから、黒いウィンドウも閉じてください。\r\n"
        "\r\n"
        "【注意】\r\n"
        "  * 初回起動は少し時間がかかります（30秒〜1分）\r\n"
        "  * Python のインストールは不要です（このフォルダ内に同梱）\r\n"
        "  * フォルダごとどこへ移動してもOKです（USBメモリでも動きます）\r\n"
        "  * このフォルダを丸ごと削除すれば、アンインストール完了\r\n"
        "\r\n"
        "【サンプルデータ】\r\n"
        "  sample_data/        ... 標準サンプル（3プラットフォーム・33件）\r\n"
        "  sample_data_extended/ ... 拡張サンプル（7プラットフォーム・59件）\r\n"
        "  sample_data_for_test/ ... 未登録CSVテスト用\r\n"
        "\r\n"
        "【設定変更】\r\n"
        "  アプリ画面右上の歯車アイコン⚙️から、対応プラットフォームの\r\n"
        "  追加・編集ができます。\r\n"
    )
    # UTF-8 with BOM で書き出し（Windowsメモ帳でも文字化けせず開ける）
    (DIST / "使い方.txt").write_text(readme, encoding="utf-8-sig")

    # ---- 10. サイズ計算と完了報告 ----
    step(10, "完成")
    total_size = sum(f.stat().st_size for f in DIST.rglob("*") if f.is_file())
    print(f"   出力先: {DIST}")
    print(f"   合計サイズ: {total_size / 1024 / 1024:.1f} MB")
    print()
    print("【次のステップ】")
    print(f"  1. {DIST / 'run.bat'} をダブルクリックして動作確認")
    print(f"  2. 問題なければ windows_bundle/ フォルダを zip 圧縮")
    print(f"  3. クライアントに配布")


if __name__ == "__main__":
    main()
