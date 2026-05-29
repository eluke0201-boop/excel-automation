"""
[Demo only] ポートフォリオ展示用のサンプルタイル機能。

このモジュールは Streamlit Cloud にデプロイした「触れるデモ版」用。
Windows バンドル等の通常配布版では:
  - app.py の DEMO_MODE フラグが False
  - build_windows_bundle.py がこのファイルをバンドルから除外
  どちらの仕組みでも実行されない。

提供関数:
  render_main_sample_tiles(config) — メイン画面上部のサンプルタイル群
  render_dialog_sample_tiles()     — 設定ダイアログ「新規追加」ビューの側パネル
"""

import base64
from pathlib import Path

import streamlit as st

from merge_reports import get_platform_color


SAMPLE_DIR_REGISTERED = Path("sample_data_extended")
SAMPLE_DIR_UNREGISTERED = Path("sample_data_for_test")


# ============================================================
# 共通: 1枚のタイル HTML
# ============================================================
def _tile_html(file_path: Path, color: str) -> str:
    csv_bytes = file_path.read_bytes()
    b64 = base64.b64encode(csv_bytes).decode()
    size_kb = len(csv_bytes) / 1024
    size_str = f"{size_kb:.1f} KB"
    return f"""
    <div class="sample-tile" draggable="true"
         data-filename="{file_path.name}"
         data-b64="{b64}"
         title="{file_path.name}">
      <svg class="file-icon" viewBox="0 0 32 40" xmlns="http://www.w3.org/2000/svg">
        <path d="M4 2 L22 2 L30 10 L30 38 L4 38 Z"
              fill="#FFFFFF" stroke="#9CA3AF" stroke-width="1.2"
              stroke-linejoin="round"/>
        <path d="M22 2 L22 10 L30 10"
              fill="#F3F4F6" stroke="#9CA3AF" stroke-width="1.2"
              stroke-linejoin="round"/>
        <line x1="8" y1="18" x2="26" y2="18" stroke="#D1D5DB" stroke-width="1"/>
        <line x1="8" y1="22" x2="26" y2="22" stroke="#D1D5DB" stroke-width="1"/>
        <line x1="8" y1="26" x2="22" y2="26" stroke="#D1D5DB" stroke-width="1"/>
        <rect x="4" y="32" width="26" height="6" fill="{color}" rx="0"/>
        <text x="17" y="37" text-anchor="middle"
              font-family="-apple-system,sans-serif"
              font-size="4.5" font-weight="700" fill="#FFFFFF"
              letter-spacing="0.5">CSV</text>
      </svg>
      <div class="tile-name">{file_path.name}</div>
      <div class="tile-meta">{size_str}</div>
    </div>
    """


# ============================================================
# 共通: 親フレームに drop ヘルパーを注入する JS スニペット
# （iframe 越しの drag-drop は dataTransfer.files が伝搬しないため、
#  親フレームのコンテキストで drop イベントを生成する必要がある）
# ============================================================
_PARENT_HELPER_JS = """
function ensureParentHelper() {
  const pw = window.parent;
  if (pw.__tileDropHelper) return true;
  try {
    const script = pw.document.createElement('script');
    script.textContent = `
      window.__tileDropHelper = function(fileList, clientX, clientY, targetSelector) {
        const dz = document.querySelector(targetSelector || '[data-testid="stFileUploaderDropzone"]');
        if (!dz) return 'no-dropzone';
        const dt = new DataTransfer();
        for (const f of fileList) {
          const file = new File([f.bytes], f.filename, { type: f.mime });
          dt.items.add(file);
        }
        const rect = dz.getBoundingClientRect();
        const cx = (typeof clientX === 'number' && clientX) ? clientX : (rect.left + rect.width / 2);
        const cy = (typeof clientY === 'number' && clientY) ? clientY : (rect.top + rect.height / 2);
        const dropEvent = new DragEvent('drop', {
          bubbles: true, cancelable: true,
          clientX: cx, clientY: cy, dataTransfer: dt,
        });
        dz.dispatchEvent(dropEvent);
        return 'ok:' + fileList.length;
      };
    `;
    pw.document.head.appendChild(script);
    return !!pw.__tileDropHelper;
  } catch (e) {
    console.error('ensureParentHelper failed:', e);
    return false;
  }
}
"""


# ============================================================
# メイン画面: 設定済み7件 + 未設定3件の二段タイル
# ============================================================
def render_main_sample_tiles(config: dict) -> None:
    """[Demo only] メイン画面のサンプルタイル群（CSVアップロード欄の上）。

    - 設定済みプラットフォーム（sample_data_extended/）
    - 未設定CSV（sample_data_for_test/）
    クリック/Ctrl+クリック/マーキー範囲選択 → ドラッグ複数同時投入対応。
    投入先はメイン画面の file_uploader。
    """
    registered_files = (
        sorted(SAMPLE_DIR_REGISTERED.glob("*.csv"))
        if SAMPLE_DIR_REGISTERED.exists() else []
    )
    unregistered_files = (
        sorted(SAMPLE_DIR_UNREGISTERED.glob("*.csv"))
        if SAMPLE_DIR_UNREGISTERED.exists() else []
    )
    if not registered_files and not unregistered_files:
        return

    # 登録済みタイル
    registered_tiles_html = ""
    for f in registered_files:
        basename = f.stem.split("_")[0].lower()
        platform_color = "#4F46E5"
        for _k, _v in config["platforms"].items():
            if _v.get("filename_prefix") == basename:
                platform_color = get_platform_color(
                    _v["label"], position=0, platform_config=_v
                )
                break
        registered_tiles_html += _tile_html(f, platform_color)

    # 未設定タイル
    unregistered_tiles_html = ""
    for f in unregistered_files:
        unregistered_tiles_html += _tile_html(f, "#9CA3AF")

    section_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
      body {{
        margin: 0;
        padding: 4px 2px;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      }}
      .sample-section {{
        background: #FAFAFA;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 10px 12px;
        margin-bottom: 8px;
        user-select: none;
      }}
      .section-title {{
        font-size: 0.85rem;
        font-weight: 700;
        color: #374151;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
      }}
      .section-sub {{
        font-size: 0.72rem;
        color: #6B7280;
        font-weight: 400;
      }}
      .tiles-container {{
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
      }}
      .sample-tile {{
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 4px;
        padding: 8px 6px 6px 6px;
        background: transparent;
        border: 1.5px solid transparent;
        border-radius: 6px;
        cursor: grab;
        user-select: none;
        transition: all 0.12s ease;
        width: 92px;
      }}
      .sample-tile:hover {{
        background: #EFF6FF;
        border-color: #93C5FD;
      }}
      .sample-tile:active {{
        cursor: grabbing;
        background: #DBEAFE;
      }}
      .sample-tile.selected {{
        background: #DBEAFE;
        border-color: #2563EB;
      }}
      .sample-tile.selected:hover {{
        background: #BFDBFE;
        border-color: #1D4ED8;
      }}
      .marquee-rect {{
        position: fixed;
        background: rgba(37, 99, 235, 0.12);
        border: 1.5px solid #2563EB;
        pointer-events: none;
        z-index: 10000;
        border-radius: 1px;
      }}
      .multi-drag-badge {{
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
      }}
      .file-icon {{
        width: 40px;
        height: 50px;
        flex-shrink: 0;
        filter: drop-shadow(0 1px 1px rgba(0,0,0,0.06));
      }}
      .tile-name {{
        font-size: 0.68rem;
        color: #1F2937;
        text-align: center;
        line-height: 1.2;
        overflow: hidden;
        text-overflow: ellipsis;
        display: -webkit-box;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
        word-break: break-all;
        max-width: 100%;
      }}
      .tile-meta {{
        font-size: 0.62rem;
        color: #9CA3AF;
        line-height: 1;
      }}
    </style>
    </head>
    <body>
      <div class="sample-section">
        <div class="section-title">
          ✅ 設定済みプラットフォーム（{len(registered_files)}件）
          <span class="section-sub">クリックで選択 / Ctrl+クリックで複数選択 / 余白ドラッグで範囲選択 → 下のゾーンへドロップ</span>
        </div>
        <div class="tiles-container">{registered_tiles_html}</div>
      </div>
      <div class="sample-section">
        <div class="section-title">
          🆕 未設定CSV（{len(unregistered_files)}件）
          <span class="section-sub">自動推測 → 設定追加フローを体験できます</span>
        </div>
        <div class="tiles-container">{unregistered_tiles_html}</div>
      </div>
      <script>
        const TARGET_SELECTOR = '[data-testid="stFileUploaderDropzone"]';
        {_PARENT_HELPER_JS}

        function getParentDropzone() {{
          try {{
            return window.parent.document.querySelector(TARGET_SELECTOR);
          }} catch (e) {{ return null; }}
        }}

        function findMyIframeEl() {{
          try {{
            for (const f of window.parent.document.querySelectorAll('iframe')) {{
              try {{ if (f.contentWindow === window) return f; }} catch (e) {{}}
            }}
          }} catch (e) {{}}
          return null;
        }}

        // ========================
        // 選択状態
        // ========================
        const selectedFilenames = new Set();
        function updateTileSelection() {{
          document.querySelectorAll('.sample-tile').forEach(t => {{
            if (selectedFilenames.has(t.dataset.filename)) {{
              t.classList.add('selected');
            }} else {{
              t.classList.remove('selected');
            }}
          }});
        }}
        function getTileFileData(tile) {{
          const b64 = tile.dataset.b64;
          const filename = tile.dataset.filename;
          const binary = atob(b64);
          const bytes = new Uint8Array(binary.length);
          for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
          return {{ bytes, filename, mime: 'text/csv' }};
        }}

        // タイルクリック選択
        document.querySelectorAll('.sample-tile').forEach(tile => {{
          tile.addEventListener('mousedown', (event) => {{
            if (event.button !== 0) return;
            const filename = tile.dataset.filename;
            if (event.ctrlKey || event.metaKey) {{
              if (selectedFilenames.has(filename)) selectedFilenames.delete(filename);
              else selectedFilenames.add(filename);
            }} else {{
              if (!selectedFilenames.has(filename)) {{
                selectedFilenames.clear();
                selectedFilenames.add(filename);
              }}
            }}
            updateTileSelection();
          }});
        }});

        // ========================
        // マーキー範囲選択（iframe 内＋親フレーム白余白の両対応）
        // ========================
        let marqueeStart = null;
        let marqueeEl = null;
        let marqueeAddMode = false;

        function startMarquee(parentX, parentY, addMode) {{
          marqueeStart = {{ x: parentX, y: parentY }};
          marqueeAddMode = addMode;
          if (!addMode) {{
            selectedFilenames.clear();
            updateTileSelection();
          }}
          try {{
            marqueeEl = window.parent.document.createElement('div');
            marqueeEl.className = 'marquee-rect';
            window.parent.document.body.appendChild(marqueeEl);
          }} catch (e) {{
            marqueeEl = document.createElement('div');
            marqueeEl.className = 'marquee-rect';
            document.body.appendChild(marqueeEl);
          }}
        }}

        function updateMarquee(parentX, parentY) {{
          if (!marqueeStart || !marqueeEl) return;
          const x1 = Math.min(marqueeStart.x, parentX);
          const y1 = Math.min(marqueeStart.y, parentY);
          const x2 = Math.max(marqueeStart.x, parentX);
          const y2 = Math.max(marqueeStart.y, parentY);
          marqueeEl.style.left = x1 + 'px';
          marqueeEl.style.top = y1 + 'px';
          marqueeEl.style.width = (x2 - x1) + 'px';
          marqueeEl.style.height = (y2 - y1) + 'px';

          const iframeEl = findMyIframeEl();
          if (!iframeEl) return;
          const iframeRect = iframeEl.getBoundingClientRect();
          document.querySelectorAll('.sample-tile').forEach(t => {{
            const r = t.getBoundingClientRect();
            const tx1 = iframeRect.left + r.left;
            const ty1 = iframeRect.top + r.top;
            const tx2 = iframeRect.left + r.right;
            const ty2 = iframeRect.top + r.bottom;
            const intersects = tx2 >= x1 && tx1 <= x2 && ty2 >= y1 && ty1 <= y2;
            if (intersects) selectedFilenames.add(t.dataset.filename);
            else if (!marqueeAddMode) selectedFilenames.delete(t.dataset.filename);
          }});
          updateTileSelection();
        }}

        function endMarquee() {{
          if (marqueeEl) {{ marqueeEl.remove(); marqueeEl = null; }}
          marqueeStart = null;
        }}

        document.addEventListener('mousedown', (event) => {{
          if (event.button !== 0 || marqueeStart) return;
          if (event.target.closest('.sample-tile')) return;
          const iframeEl = findMyIframeEl();
          if (!iframeEl) return;
          const iframeRect = iframeEl.getBoundingClientRect();
          startMarquee(
            iframeRect.left + event.clientX,
            iframeRect.top + event.clientY,
            event.ctrlKey || event.metaKey,
          );
          event.preventDefault();
        }});
        document.addEventListener('mousemove', (event) => {{
          if (!marqueeStart) return;
          const iframeEl = findMyIframeEl();
          if (!iframeEl) return;
          const iframeRect = iframeEl.getBoundingClientRect();
          updateMarquee(iframeRect.left + event.clientX, iframeRect.top + event.clientY);
        }});
        document.addEventListener('mouseup', () => endMarquee());

        function ensureParentMarquee() {{
          try {{
            const pw = window.parent;
            if (pw.__tileMarqueeInstalled) {{
              pw.removeEventListener('mousedown', pw.__tileMarqueeMD, true);
              pw.removeEventListener('mousemove', pw.__tileMarqueeMM, true);
              pw.removeEventListener('mouseup',   pw.__tileMarqueeMU, true);
            }}
            pw.__tileMarqueeInstalled = true;
            pw.__tileMarqueeMD = (event) => {{
              if (event.button !== 0 || marqueeStart) return;
              if (event.target.closest(
                'button, input, textarea, select, a, label, [role="button"], iframe, ' +
                '[data-testid="stFileUploaderDropzone"], [data-testid="stFileChip"], ' +
                '[data-testid="stTabs"], [data-testid="stExpander"], [data-baseweb], ' +
                '[data-testid="stDataFrame"], [data-testid="stMarkdownContainer"] a'
              )) return;
              startMarquee(event.clientX, event.clientY, event.ctrlKey || event.metaKey);
              event.preventDefault();
            }};
            pw.__tileMarqueeMM = (event) => {{
              if (!marqueeStart) return;
              updateMarquee(event.clientX, event.clientY);
            }};
            pw.__tileMarqueeMU = () => endMarquee();
            pw.addEventListener('mousedown', pw.__tileMarqueeMD, true);
            pw.addEventListener('mousemove', pw.__tileMarqueeMM, true);
            pw.addEventListener('mouseup',   pw.__tileMarqueeMU, true);
          }} catch (e) {{
            console.warn('ensureParentMarquee failed:', e);
          }}
        }}
        ensureParentMarquee();

        // ========================
        // ドラッグ & ドロップ
        // ========================
        let activeFiles = [];
        let parentDropHandler = null;
        let parentDragOverHandler = null;
        let multiDragBadge = null;
        let multiDragMoveHandler = null;

        function cleanup() {{
          const dz = getParentDropzone();
          if (dz) {{
            dz.classList.remove('tile-drag-active');
            if (parentDropHandler) {{
              dz.removeEventListener('drop', parentDropHandler, true);
              parentDropHandler = null;
            }}
            if (parentDragOverHandler) {{
              dz.removeEventListener('dragover', parentDragOverHandler, true);
              dz.removeEventListener('dragenter', parentDragOverHandler, true);
              parentDragOverHandler = null;
            }}
          }}
          activeFiles = [];
          if (multiDragBadge) {{ multiDragBadge.remove(); multiDragBadge = null; }}
          if (multiDragMoveHandler) {{
            document.removeEventListener('dragover', multiDragMoveHandler);
            try {{ window.parent.document.removeEventListener('dragover', multiDragMoveHandler); }} catch(e) {{}}
            multiDragMoveHandler = null;
          }}
        }}

        document.querySelectorAll('.sample-tile').forEach(tile => {{
          tile.addEventListener('dragstart', (event) => {{
            try {{
              cleanup();
              ensureParentHelper();
              const filename = tile.dataset.filename;
              if (!selectedFilenames.has(filename)) {{
                selectedFilenames.clear();
                selectedFilenames.add(filename);
                updateTileSelection();
              }}
              activeFiles = Array.from(selectedFilenames).map(fn => {{
                const t = document.querySelector('.sample-tile[data-filename="' + CSS.escape(fn) + '"]');
                return t ? getTileFileData(t) : null;
              }}).filter(Boolean);

              event.dataTransfer.effectAllowed = 'copy';
              event.dataTransfer.setData('text/plain', activeFiles.map(f => f.filename).join(', '));

              const dz = getParentDropzone();
              if (!dz) return;
              dz.classList.add('tile-drag-active');

              if (activeFiles.length > 1) {{
                multiDragBadge = window.parent.document.createElement('div');
                multiDragBadge.className = 'multi-drag-badge';
                multiDragBadge.textContent = activeFiles.length + ' 個のファイル';
                window.parent.document.body.appendChild(multiDragBadge);
                multiDragMoveHandler = (e) => {{
                  if (!multiDragBadge) return;
                  multiDragBadge.style.left = (e.clientX + 12) + 'px';
                  multiDragBadge.style.top = (e.clientY + 12) + 'px';
                }};
                window.parent.document.addEventListener('dragover', multiDragMoveHandler);
              }}

              parentDropHandler = (dropEvent) => {{
                if (activeFiles.length === 0) return;
                dropEvent.preventDefault();
                dropEvent.stopImmediatePropagation();
                const files = activeFiles;
                const cx = dropEvent.clientX;
                const cy = dropEvent.clientY;
                cleanup();
                window.parent.__tileDropHelper(files, cx, cy, TARGET_SELECTOR);
              }};
              parentDragOverHandler = (e) => {{
                e.preventDefault();
                if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
              }};
              dz.addEventListener('dragenter', parentDragOverHandler, true);
              dz.addEventListener('dragover', parentDragOverHandler, true);
              dz.addEventListener('drop', parentDropHandler, true);
            }} catch (e) {{
              console.error('tile dragstart failed:', e);
            }}
          }});

          tile.addEventListener('dragend', () => {{ setTimeout(cleanup, 100); }});
        }});
      </script>
    </body>
    </html>
    """
    reg_rows = max(1, (len(registered_files) + 9) // 10)
    unreg_rows = max(1, (len(unregistered_files) + 9) // 10)
    height = 50 + reg_rows * 95 + 50 + unreg_rows * 95 + 30
    st.components.v1.html(section_html, height=height)


# ============================================================
# 設定ダイアログ「新規追加」ビュー: 未設定サンプル CSV の側パネル
# ============================================================
def render_dialog_sample_tiles() -> None:
    """[Demo only] ダイアログ用の小さい縦並びタイル。

    sampleA/B/C をドラッグして、ダイアログ内の「自動推測」アップローダへ投入。
    マーキー・複数選択なし。シンプルに1枚ずつ drag 想定。
    """
    sample_dir = SAMPLE_DIR_UNREGISTERED
    if not sample_dir.exists():
        return
    files = sorted(sample_dir.glob("*.csv"))
    if not files:
        return

    # ダイアログ内のドロップゾーン
    target_selector = '[role="dialog"] [data-testid="stFileUploaderDropzone"]'

    tiles_html = ""
    for f in files:
        tiles_html += _tile_html(f, "#9CA3AF")

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
      body {{
        margin: 0;
        padding: 4px;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      }}
      .dlg-panel {{
        background: #FAFAFA;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 10px 8px;
        user-select: none;
      }}
      .dlg-title {{
        font-size: 0.75rem;
        font-weight: 700;
        color: #374151;
        text-align: center;
        margin-bottom: 4px;
      }}
      .dlg-hint {{
        font-size: 0.62rem;
        color: #6B7280;
        text-align: center;
        margin-bottom: 8px;
        line-height: 1.25;
      }}
      .tiles-container {{
        display: flex;
        flex-direction: column;
        gap: 4px;
        align-items: center;
      }}
      .sample-tile {{
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 3px;
        padding: 6px 4px;
        background: transparent;
        border: 1.5px solid transparent;
        border-radius: 6px;
        cursor: grab;
        user-select: none;
        transition: all 0.12s ease;
        width: 80px;
      }}
      .sample-tile:hover {{
        background: #EFF6FF;
        border-color: #93C5FD;
      }}
      .sample-tile:active {{ cursor: grabbing; background: #DBEAFE; }}
      .file-icon {{
        width: 34px; height: 42px; flex-shrink: 0;
        filter: drop-shadow(0 1px 1px rgba(0,0,0,0.06));
      }}
      .tile-name {{
        font-size: 0.65rem;
        color: #1F2937;
        text-align: center;
        line-height: 1.2;
        word-break: break-all;
      }}
      .tile-meta {{ font-size: 0.58rem; color: #9CA3AF; line-height: 1; }}
    </style>
    </head>
    <body>
      <div class="dlg-panel">
        <div class="dlg-title">📁 サンプルCSV</div>
        <div class="dlg-hint">→ ドラッグして<br/>自動推測を試す</div>
        <div class="tiles-container">{tiles_html}</div>
      </div>
      <script>
        const TARGET_SELECTOR = '{target_selector}';
        {_PARENT_HELPER_JS}

        function getParentDropzone() {{
          try {{ return window.parent.document.querySelector(TARGET_SELECTOR); }}
          catch (e) {{ return null; }}
        }}

        let activeFile = null;
        let parentDropHandler = null;
        let parentDragOverHandler = null;

        function cleanup() {{
          const dz = getParentDropzone();
          if (dz) {{
            dz.classList.remove('tile-drag-active');
            if (parentDropHandler) {{
              dz.removeEventListener('drop', parentDropHandler, true);
              parentDropHandler = null;
            }}
            if (parentDragOverHandler) {{
              dz.removeEventListener('dragover', parentDragOverHandler, true);
              dz.removeEventListener('dragenter', parentDragOverHandler, true);
              parentDragOverHandler = null;
            }}
          }}
          activeFile = null;
        }}

        document.querySelectorAll('.sample-tile').forEach(tile => {{
          tile.addEventListener('dragstart', (event) => {{
            try {{
              cleanup();
              ensureParentHelper();
              const b64 = tile.dataset.b64;
              const filename = tile.dataset.filename;
              const binary = atob(b64);
              const bytes = new Uint8Array(binary.length);
              for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
              activeFile = {{ bytes, filename, mime: 'text/csv' }};
              event.dataTransfer.effectAllowed = 'copy';
              event.dataTransfer.setData('text/plain', filename);

              const dz = getParentDropzone();
              if (!dz) return;
              dz.classList.add('tile-drag-active');

              parentDropHandler = (dropEvent) => {{
                if (!activeFile) return;
                dropEvent.preventDefault();
                dropEvent.stopImmediatePropagation();
                const f = activeFile;
                const cx = dropEvent.clientX;
                const cy = dropEvent.clientY;
                cleanup();
                window.parent.__tileDropHelper([f], cx, cy, TARGET_SELECTOR);
              }};
              parentDragOverHandler = (e) => {{
                e.preventDefault();
                if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
              }};
              dz.addEventListener('dragenter', parentDragOverHandler, true);
              dz.addEventListener('dragover', parentDragOverHandler, true);
              dz.addEventListener('drop', parentDropHandler, true);
            }} catch (e) {{
              console.error('dialog tile dragstart failed:', e);
            }}
          }});

          tile.addEventListener('dragend', () => {{ setTimeout(cleanup, 100); }});
        }});
      </script>
    </body>
    </html>
    """
    height = 80 + len(files) * 78
    st.components.v1.html(html, height=height)
