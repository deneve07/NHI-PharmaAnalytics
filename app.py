"""
藥品資料庫分析系統 (Streamlit 拖曳版 - 廠商分析)
----------------------------------------------
沿用「華典streamlit-0731」的拖曳樞紐分析引擎（streamlit-sortables），
但僅開放「成分、劑型、劑量、規格、廠商」五個欄位供拖曳排列，
且只有一種分析功能（廠商申報量排名），不需要「選擇分析功能」這一步。

步驟：
1. 選擇成分
2. 拖曳欄位到「報表欄位」並排序
3. 逐欄篩選 / 是否合計
4. 選擇要加總的數值欄位（2023~2026年數量），可加占比(%)、成長率(%)
5. 產生 Excel 報表與網頁即時預覽（沿用原本深藍色主題與 A4 橫式列印設定）

需要的套件（requirements.txt）：
    streamlit
    pandas
    openpyxl
    streamlit-sortables
"""

import re
import io
import base64
from datetime import datetime

try:
    from zoneinfo import ZoneInfo
    TAIPEI_TZ = ZoneInfo("Asia/Taipei")
except Exception:
    TAIPEI_TZ = None

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from streamlit_sortables import sort_items
    HAS_SORTABLES = True
except ImportError:
    HAS_SORTABLES = False


# ============================================================
# 基礎設定與資料載入
# ============================================================

st.set_page_config(page_title="健保資料庫分析系統", page_icon="📊", layout="wide")

DATA_FILE = "data2023-2026.csv"

# 只開放這五個欄位供拖曳排列（依使用者要求，不像 0731 版本開放「所有欄位」）
FIXED_FIELDS = ["成分", "劑型", "劑量", "規格", "廠商"]

# 原始資料的四個年度數量欄 -> 內部工作用欄名（沿用「申報量」關鍵字，讓占比/成長率引擎能自動辨識）
QTY_COL_MAP = {
    "2023年數量": "2023年申報量",
    "2024年數量": "2024年申報量",
    "2025年數量": "2025年申報量",
    "2026年(1-5月)數量": "2026年申報量",
}
# 內部工作用欄名 -> 顯示用 (年份標籤, 欄位標籤)，2026 特別標示 (1-5月)
QTY_DISPLAY_LABEL = {
    "2023年申報量": ("2023年", "數量"),
    "2024年申報量": ("2024年", "數量"),
    "2025年申報量": ("2025年", "數量"),
    "2026年申報量": ("2026年(1-5月)", "數量"),
}
BLANK_TOKENS = {"nan", "none", "<na>", "null", ""}


@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    try:
        df = pd.read_csv(path)
    except Exception as e:
        st.error(f"⚠️ 讀取資料檔案「{path}」時發生錯誤：{e}")
        return pd.DataFrame()

    df.columns = df.columns.str.strip()

    # 清洗五個維度欄位：去除空白，並把各種「無資料」寫法統一成空字串
    for col in FIXED_FIELDS:
        if col not in df.columns:
            df[col] = ""
            continue
        cleaned = df[col].astype(str).str.strip()
        if col == "劑量":
            # 移除劑量字串內所有空白（含中間空格），避免 "250mg" 與 "250 mg" 被當成不同劑量
            cleaned = cleaned.str.replace(r"\s+", "", regex=True)
        cleaned = cleaned.apply(lambda x: "" if str(x).lower() in BLANK_TOKENS else x)
        df[col] = cleaned

    # 四個年度數量欄：清除千分位逗號/空白後轉數字，轉出內部工作欄名
    for raw_col, internal_col in QTY_COL_MAP.items():
        if raw_col not in df.columns:
            df[raw_col] = 0
        cleaned = df[raw_col].astype(str).str.replace(",", "", regex=False).str.strip()
        df[internal_col] = pd.to_numeric(cleaned, errors="coerce").fillna(0)

    return df


def parse_dosage(d):
    m = re.search(r"[\d.]+", str(d))
    return float(m.group()) if m else 0.0


def order_display_cols(value_cols, pct_cols, growth_cols):
    """占比要接在數量後面，而不是放在所有數值欄位最後"""
    qty_cols = [c for c in value_cols if "申報量" in c]
    other_cols = [c for c in value_cols if c not in qty_cols]
    return qty_cols + pct_cols + other_cols + growth_cols


def pretty_header(col: str):
    """把內部欄名轉成 (年份標籤, 欄位標籤) 供標題換行顯示；2026 年一律標示 (1-5月)"""
    if col in QTY_DISPLAY_LABEL:
        return QTY_DISPLAY_LABEL[col]

    m = re.match(r"^(\d{4})年占比\(%\)$", col)
    if m:
        y = m.group(1)
        y_disp = f"{y}年(1-5月)" if y == "2026" else f"{y}年"
        return y_disp, "占比(%)"

    m = re.match(r"^(\d{4})-(\d{4})年成長率\(%\)$", col)
    if m:
        y1, y2 = m.group(1), m.group(2)
        y1d = f"{y1}年(1-5月)" if y1 == "2026" else f"{y1}年"
        y2d = f"{y2}年(1-5月)" if y2 == "2026" else f"{y2}年"
        return f"{y1d}-{y2d}", "成長率(%)"

    return None, col


# ============================================================
# 通用樞紐分析引擎：任意欄位順序 + 任意層級巢狀合計 + 占比/成長率
# （沿用 0731 版本的引擎邏輯，因資料本身已是乾淨數字，移除原本的
#   「無資料(-)/異常標記文字」保留顯示機制，只保留核心的巢狀分組/合計）
# ============================================================

def build_nested_rows(df: pd.DataFrame, row_fields: list, subtotal_fields: list, value_cols: list,
                       pct_years: list = None, add_growth: bool = False):
    # 先依「報表欄位」的組合彙總數值，確保同一個欄位組合只會出現一列
    df = df.groupby(row_fields, as_index=False)[value_cols].sum()

    qty_cols = [c for c in value_cols if "申報量" in c]
    qty_years = sorted(set(c[:4] for c in qty_cols))
    pct_years = pct_years or []

    pct_cols = [f"{y}年占比(%)" for y in qty_years if y in pct_years]
    growth_cols = [f"{qty_years[i]}-{qty_years[i+1]}年成長率(%)" for i in range(len(qty_years) - 1)] if add_growth else []

    def compute_extra_for_row(sums, top_totals):
        extra = {}
        for c in pct_cols:
            y = c[:4]
            base = f"{y}年申報量"
            extra[c] = sums.get(base, 0) / top_totals.get(base, 0) if top_totals.get(base, 0) else 0
        for c in growth_cols:
            y1, y2 = c[:4], c[5:9]
            b1, b2 = f"{y1}年申報量", f"{y2}年申報量"
            extra[c] = (sums.get(b2, 0) - sums.get(b1, 0)) / sums.get(b1, 0) if sums.get(b1, 0) else 0
        return extra

    def compute_extra_for_group(totals):
        extra = {c: 1.0 for c in pct_cols}  # 小計/總計列本身佔比恆為 100%
        for c in growth_cols:
            y1, y2 = c[:4], c[5:9]
            b1, b2 = f"{y1}年申報量", f"{y2}年申報量"
            extra[c] = (totals.get(b2, 0) - totals.get(b1, 0)) / totals.get(b1, 0) if totals.get(b1, 0) else 0
        return extra

    year_cols = sorted(qty_cols, reverse=True)

    sort_cols, sort_asc, temp_sort_cols = [], [], []
    for col in subtotal_fields:
        if col == "劑量":
            # 劑量欄需依實際數值排序 (例如 12mg/ml < 24mg/ml < 100mg/ml)，而非字母排序
            tmp_col = "__sortkey_劑量"
            df[tmp_col] = df[col].map(parse_dosage)
            sort_cols.append(tmp_col)
            temp_sort_cols.append(tmp_col)
        else:
            sort_cols.append(col)
        sort_asc.append(True)
    sort_cols += year_cols
    sort_asc += [False] * len(year_cols)

    df_sorted = df.sort_values(by=sort_cols, ascending=sort_asc) if sort_cols else df
    if temp_sort_cols:
        df_sorted = df_sorted.drop(columns=temp_sort_cols)

    if subtotal_fields:
        min_level_idx = min(row_fields.index(c) for c in subtotal_fields)
    else:
        min_level_idx = len(row_fields)

    group_cols_to_blank = []
    for c in row_fields:
        if c in subtotal_fields:
            continue
        if df[c].nunique() <= 1 or row_fields.index(c) < min_level_idx:
            group_cols_to_blank.append(c)

    rows = []
    last_vals = {}
    first_flags = {c: True for c in subtotal_fields}

    def emit_row(row, top_totals):
        rec_vals = {}
        for f in row_fields:
            val = row[f]
            show = val
            if f in group_cols_to_blank:
                if val != last_vals.get(f):
                    last_vals[f] = val
                else:
                    show = ""
            elif f in first_flags:
                if not first_flags[f]:
                    show = ""
            rec_vals[f] = show
        sums = {c: row[c] for c in value_cols}
        sums.update(compute_extra_for_row(sums, top_totals))
        rows.append({"type": "data", "values": rec_vals, "sums": sums})
        for c in first_flags:
            first_flags[c] = False

    def emit_subtotal(field, label, totals):
        rec_vals = {f: "" for f in row_fields}
        rec_vals[field] = f"{label} 合計"
        full = dict(totals)
        full.update(compute_extra_for_group(totals))
        rows.append({"type": "subtotal", "values": rec_vals, "sums": full})

    def recurse(sub_df, level_idx, top_totals):
        if level_idx >= len(subtotal_fields):
            for _, row in sub_df.iterrows():
                emit_row(row, top_totals)
            return
        col = subtotal_fields[level_idx]
        for _, grp in sub_df.groupby(col, sort=False):
            totals = {c: grp[c].sum() for c in value_cols}
            next_top = totals if level_idx == len(subtotal_fields) - 1 else top_totals
            label = grp[col].iloc[0]
            first_flags[col] = True
            recurse(grp, level_idx + 1, next_top)
            emit_subtotal(col, label, totals)

    grand_totals = {c: df[c].sum() for c in value_cols}
    if subtotal_fields:
        recurse(df_sorted, 0, grand_totals)
    else:
        for _, row in df_sorted.iterrows():
            emit_row(row, grand_totals)

    total_vals = {f: "" for f in row_fields}
    if row_fields:
        total_vals[row_fields[0]] = "總計"
    full_grand = dict(grand_totals)
    full_grand.update(compute_extra_for_group(grand_totals))
    rows.append({"type": "total", "values": total_vals, "sums": full_grand})

    return rows, pct_cols, growth_cols


def get_report_timestamp() -> str:
    now = datetime.now(TAIPEI_TZ) if TAIPEI_TZ else datetime.now()
    return f"{now.year}/{now.month}/{now.day} {now.hour:02d}:{now.minute:02d}"


def safe_numeric(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            if pd.isna(v):
                return 0.0, True
        except (TypeError, ValueError):
            pass
        return float(v), True
    try:
        n = pd.to_numeric(v)
        return float(n), True
    except (TypeError, ValueError):
        return v, False


# ============================================================
# 樣式化 HTML 預覽（沿用原 Gradio 版本的深藍色主題：#1F497D 標題、
# #DCE6F1 小計、#B8CCE4 總計）
# ============================================================

FONT_FAMILY = "'微軟正黑體', 'Microsoft JhengHei', sans-serif"
HEADER_COLOR = "#1F497D"
SUBTOTAL_COLOR = "#DCE6F1"
TOTAL_COLOR = "#B8CCE4"


def build_html_table(rows, row_fields, value_cols, pct_cols, growth_cols, report_title):
    extra_cols = set(pct_cols + growth_cols)
    display_cols = order_display_cols(value_cols, pct_cols, growth_cols)
    headers = row_fields + display_cols

    th_style = (
        f"color:#FFFFFF;background-color:{HEADER_COLOR};text-align:center;"
        f"border:1px solid #D9D9D9;padding:10px 12px;font-weight:bold;font-family:{FONT_FAMILY};"
    )
    td_style = f"border:1px solid #D9D9D9;padding:8px 12px;font-family:{FONT_FAMILY};white-space:nowrap;"

    html = [
        "<div class='report-preview-wrapper'><div class='report-container'>",
        f"<h2 class='report-title'>{report_title}</h2>",
        "<div class='table-responsive'>",
        f"<table style='border-collapse:collapse;width:max-content;min-width:100%;font-size:15px;font-family:{FONT_FAMILY};'>",
        "<tr>",
    ]
    for f in row_fields:
        html.append(f"<th style='{th_style}'>{f}</th>")
    for c in display_cols:
        year, label = pretty_header(c)
        if year:
            html.append(f"<th style='{th_style}'>{year}<br>{label}</th>")
        else:
            html.append(f"<th style='{th_style}'>{label}</th>")
    html.append("</tr>")

    for r in rows:
        bg = ""
        fw = "font-weight:normal;"
        if r["type"] == "subtotal":
            bg = f"background-color:{SUBTOTAL_COLOR};"
            fw = "font-weight:bold;"
        elif r["type"] == "total":
            bg = f"background-color:{TOTAL_COLOR};"
            fw = "font-weight:bold;"
        html.append(f"<tr style='{bg}{fw}'>")
        for i, f in enumerate(row_fields):
            v = r["values"].get(f, "")
            align = "left" if i == 0 else "center"
            html.append(f"<td style='{td_style}text-align:{align};'>{v}</td>")
        for c in display_cols:
            if c in extra_cols:
                v, ok = safe_numeric(r["sums"].get(c, 0))
                html.append(f"<td style='{td_style}text-align:right;'>{v:.1%}</td>" if ok else f"<td style='{td_style}text-align:right;'>{v}</td>")
            else:
                raw = r["sums"].get(c, "")
                if raw == "":
                    html.append(f"<td style='{td_style}'></td>")
                else:
                    v, ok = safe_numeric(raw)
                    html.append(f"<td style='{td_style}text-align:right;'>{v:,.0f}</td>" if ok else f"<td style='{td_style}text-align:right;'>{v}</td>")
        html.append("</tr>")

    html.append("</table></div>")
    html.append(
        "<div class='report-footer'>"
        "<span>中央健康保險署  政府資料開放平台 2026年資料</span>"
        "<span>https://data.gov.tw/dataset/22131</span>"
        "</div>"
    )
    html.append("</div></div>")

    style_block = f"""
    <style>
        .report-preview-wrapper {{ padding: 15px 0; }}
        .report-container {{
            background-color: #FFFFFF !important;
            color: #333333 !important;
            padding: 20px !important;
            width: 100%;
            box-sizing: border-box;
            border-radius: 12px;
            box-shadow: 0 8px 24px rgba(0,0,0,0.15);
            border: 1px solid #E0E0E0;
        }}
        .table-responsive {{ width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; padding-bottom: 10px; }}
        .report-title {{
            color: #000000 !important; text-align: center; font-family: {FONT_FAMILY};
            font-weight: bold; margin-top: 5px; margin-bottom: 20px;
            white-space: normal; word-break: break-word; overflow-wrap: break-word;
            /* 讓標題不要撐大外層容器的寬度：容器寬度由表格(max-content)決定，
               標題本身用 width:0; min-width:100% 的技巧強制在容器現有寬度內換行 */
            display: block; width: 0; min-width: 100%; box-sizing: border-box;
        }}
        .report-footer {{
            display: flex; flex-wrap: wrap; justify-content: space-between;
            margin-top: 25px; padding-top: 15px; border-top: 1px solid #D9D9D9;
            font-size: 13px !important; font-family: {FONT_FAMILY}; font-weight: bold; color: #333333 !important;
        }}
    </style>
    """
    return style_block + "".join(html)


CAPTURE_HTML_TEMPLATE = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<div style="text-align:center;">
  <button id="export-png-btn-{uid}" style="background-color:#1F497D;color:white;border:none;padding:10px 20px;
    border-radius:6px;font-size:14px;cursor:pointer;width:100%;">🖼️ 匯出 PNG 圖片</button>
</div>
<div id="result-area-{uid}" style="margin-top:10px;font-size:13px;color:#1F497D;text-align:center;"></div>
<div id="capture-wrap-{uid}" style="position:absolute; left:-99999px; top:0; width:max-content;">{table_html}</div>
<script>
(function() {{
    const resultArea = document.getElementById('result-area-{uid}');
    const pngBtn = document.getElementById('export-png-btn-{uid}');
    let busy = false;

    async function captureCanvas() {{
        const target = document.getElementById('capture-wrap-{uid}');
        if (document.fonts && document.fonts.ready) {{ await document.fonts.ready; }}
        if (typeof html2canvas === 'undefined') {{ throw new Error('html2canvas 尚未載入完成，請確認網路連線後再試一次'); }}
        return await html2canvas(target, {{
            scale: 2, backgroundColor: '#ffffff',
            windowWidth: target.scrollWidth, windowHeight: target.scrollHeight,
            width: target.scrollWidth, height: target.scrollHeight, useCORS: true
        }});
    }}

    async function shareOrDownload(blob, filename, mime) {{
        const file = new File([blob], filename, {{type: mime}});
        const isMobile = /iPhone|iPad|iPod|Android/i.test(navigator.userAgent) ||
            (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
        if (isMobile) {{
            if (navigator.canShare && navigator.canShare({{files: [file]}})) {{
                try {{ await navigator.share({{files: [file]}}); resultArea.innerText = '✅ 已開啟分享選單'; }}
                catch (shareErr) {{ }}
            }} else {{
                const url = URL.createObjectURL(blob);
                resultArea.innerHTML = "目前瀏覽器不支援原生分享，請長按下方圖片選擇「儲存影像」：<br/>" +
                    "<img src='" + url + "' style='max-width:100%;border-radius:8px;border:1px solid #eee;margin-top:6px;' />";
            }}
        }} else {{
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = filename;
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            setTimeout(function() {{ URL.revokeObjectURL(url); }}, 5000);
            resultArea.innerText = '✅ 已開始下載';
        }}
    }}

    pngBtn.addEventListener('click', async function() {{
        if (busy) return;
        busy = true; pngBtn.disabled = true;
        resultArea.innerText = '🔄 產生圖片中，請稍候...';
        try {{
            const canvas = await captureCanvas();
            const blob = await new Promise(function(resolve) {{ canvas.toBlob(resolve, 'image/png'); }});
            if (!blob) {{ throw new Error('圖片產生失敗 (canvas 轉換為空)'); }}
            await shareOrDownload(blob, '{filename}.png', 'image/png');
        }} catch (err) {{
            resultArea.innerText = '❌ 匯出失敗：' + (err && err.message ? err.message : err);
        }} finally {{ busy = false; pngBtn.disabled = false; }}
    }});
}})();
</script>
"""

EXCEL_SHARE_HTML_TEMPLATE = """
<div style="text-align:center;">
  <button id="excel-share-btn-{uid}" style="background-color:#1F497D;color:white;border:none;padding:10px 20px;
    border-radius:6px;font-size:14px;cursor:pointer;width:100%;">📄 下載 / 分享 Excel 報表</button>
</div>
<script>
(function() {{
    const btn = document.getElementById('excel-share-btn-{uid}');
    let busy = false;
    btn.addEventListener('click', async function() {{
        if (busy) return;
        busy = true; btn.disabled = true;
        try {{
            const b64 = "{b64data}";
            const byteChars = atob(b64);
            const byteNumbers = new Array(byteChars.length);
            for (let i = 0; i < byteChars.length; i++) {{ byteNumbers[i] = byteChars.charCodeAt(i); }}
            const byteArray = new Uint8Array(byteNumbers);
            const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
            const blob = new Blob([byteArray], {{type: mime}});
            const file = new File([blob], '{filename}.xlsx', {{type: mime}});
            const isMobile = /iPhone|iPad|iPod|Android/i.test(navigator.userAgent) ||
                (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
            if (isMobile && navigator.canShare && navigator.canShare({{files: [file]}})) {{
                try {{ await navigator.share({{files: [file]}}); }} catch (shareErr) {{ }}
            }} else {{
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = '{filename}.xlsx';
                document.body.appendChild(a); a.click(); document.body.removeChild(a);
                setTimeout(function() {{ URL.revokeObjectURL(url); }}, 5000);
            }}
        }} finally {{ busy = false; btn.disabled = false; }}
    }});
}})();
</script>
"""


def render_excel_share(excel_bytes: bytes, filename: str, uid: str):
    b64data = base64.b64encode(excel_bytes).decode()
    safe_uid = re.sub(r"[^0-9A-Za-z_]", "_", uid)
    components.html(EXCEL_SHARE_HTML_TEMPLATE.format(b64data=b64data, filename=filename, uid=safe_uid), height=50)


# ============================================================
# Excel 匯出（沿用原 Gradio 版本的深藍色主題、A4 橫式、列印設定）
# ============================================================

def generate_excel_bytes(rows, row_fields, value_cols, pct_cols, growth_cols, report_title):
    extra_cols = set(pct_cols + growth_cols)
    display_cols = order_display_cols(value_cols, pct_cols, growth_cols)
    headers = row_fields + display_cols

    wb = Workbook()
    ws = wb.active
    ws.title = "廠商分析"

    ws.views.sheetView[0].showGridLines = False
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.orientation = "landscape"
    ws.print_options.horizontalCentered = True
    ws.page_setup.scaleWithDoc = True
    ws.page_setup.alignWithMargins = True
    # 沒有這一行，Excel 會忽略 fitToWidth/fitToHeight，直接照 100% 實際大小分頁，
    # 導致欄位很多時被切成好幾頁 —— 必須明確開啟 fitToPage 才會套用「調整為 1 頁寬」
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins.top = 2.7 / 2.54
    ws.page_margins.bottom = 2.5 / 2.54
    ws.page_margins.left = 1.5 / 2.54
    ws.page_margins.right = 1.5 / 2.54
    ws.page_margins.header = 1.5 / 2.54
    ws.page_margins.footer = 1.0 / 2.54

    ws.oddHeader.center.text = f'&"微軟正黑體,Bold"&16{report_title}'
    ws.oddFooter.left.text = '&"微軟正黑體,Regular"&12中央健康保險署  政府資料開放平台 2026年資料'
    ws.oddFooter.right.text = '&"微軟正黑體,Regular"&12https://data.gov.tw/dataset/22131'

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    subtotal_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    total_fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")

    font_family = "微軟正黑體"
    header_font = Font(name=font_family, size=12, bold=True, color="FFFFFF")
    data_font = Font(name=font_family, size=12)
    bold_font = Font(name=font_family, size=12, bold=True)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center")

    thin_side = Side(border_style="thin", color="D9D9D9")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # 標題列
    header_texts = list(row_fields)
    for c in display_cols:
        year, label = pretty_header(c)
        header_texts.append(f"{year}\n{label}" if year else label)
    ws.append(header_texts)
    ws.row_dimensions[1].height = 60
    ws.print_title_rows = "1:1"

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = cell_border

    current_row = 2
    for r in rows:
        row_vals = []
        for f in row_fields:
            row_vals.append(r["values"].get(f, ""))
        for c in display_cols:
            if c in extra_cols:
                v, ok = safe_numeric(r["sums"].get(c, 0))
                row_vals.append(v if ok else str(v))
            else:
                raw = r["sums"].get(c, "")
                if raw == "":
                    row_vals.append("")
                else:
                    v, ok = safe_numeric(raw)
                    row_vals.append(v if ok else str(v))
        ws.append(row_vals)
        ws.row_dimensions[current_row].height = 35

        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=i)
            if i <= len(row_fields):
                cell.alignment = left_align if i == 1 else center_align
                cell.font = bold_font if r["type"] != "data" else data_font
            elif h in extra_cols:
                cell.number_format = "0.0%"
                cell.alignment = right_align
                cell.font = bold_font if r["type"] != "data" else data_font
            else:
                cell.number_format = "#,##0"
                cell.alignment = right_align
                cell.font = bold_font if r["type"] != "data" else data_font
            cell.border = cell_border
            if r["type"] == "subtotal":
                cell.fill = subtotal_fill
            elif r["type"] == "total":
                cell.fill = total_fill
        current_row += 1

    for col_idx, col_name in enumerate(headers, 1):
        col_letter = get_column_letter(col_idx)
        if col_idx == 1:
            ws.column_dimensions[col_letter].width = 30
        elif col_idx <= len(row_fields):
            ws.column_dimensions[col_letter].width = 18
        elif headers[col_idx - 1] in ("占比(%)", "成長率(%)"):
            ws.column_dimensions[col_letter].width = 16
        else:
            ws.column_dimensions[col_letter].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# Streamlit 主畫面（單一分析功能：廠商分析，不需選擇分析功能）
# ============================================================

st.title("📊 健保資料庫分析工具")
st.caption("先選擇成分，接著把需要的欄位拖到「報表欄位」，即時預覽會隨著您的設定更新，效果如同 Excel 樞紐分析。")

df_raw = load_data(DATA_FILE)

if df_raw.empty:
    st.warning(f"找不到資料檔案「{DATA_FILE}」，請確認檔案已放置於工作目錄。")
elif "成分" not in df_raw.columns:
    st.error(f"資料檔案缺少「成分」欄位，實際欄位為：{list(df_raw.columns)}")
else:
    comp_options = sorted([c for c in df_raw["成分"].dropna().unique() if c])

    if "comps_selected_persist" not in st.session_state:
        st.session_state["comps_selected_persist"] = []

    # 搜尋關鍵字用獨立的 text_input 元件保存，不會因為在下方勾選/取消勾選成分而被清空，
    # 這樣選了一項之後，下拉式清單仍會停留在同一組關鍵字的搜尋結果，不必重新輸入
    search_kw = st.text_input(
        "第1步：輸入關鍵字搜尋成分",
        key="comp_search_kw",
        placeholder="輸入關鍵字，例如 Ezetimibe",
    )

    if search_kw.strip():
        kw = search_kw.strip().lower()

        def _relevance_rank(name):
            n = name.lower()
            if n == kw:
                return 0
            if n.startswith(kw):
                return 1
            return 2

        filtered_options = [c for c in comp_options if kw in c.lower()]
        filtered_options.sort(key=lambda c: (_relevance_rank(c), c.lower()))
    else:
        filtered_options = []

    persisted_selected = [c for c in st.session_state["comps_selected_persist"] if c in comp_options]
    # 合併「搜尋結果」與「目前已勾選項目」並去重，確保已勾選成分不會因為搜尋字變動而從清單消失，
    # 也避免同一個成分同時出現在兩邊來源時被重複列出
    combined_options = list(dict.fromkeys(filtered_options + persisted_selected))

    comps_selected = st.multiselect(
        "第2步：勾選成分品項 (可多選)",
        options=combined_options,
        default=persisted_selected,
        key="comps_select_widget",
    )
    st.session_state["comps_selected_persist"] = comps_selected

    if not comps_selected:
        st.info("請先選擇至少一個成分，才會顯示後續的欄位設定。")
    else:
        df_comp = df_raw[df_raw["成分"].isin(comps_selected)]

        st.markdown("### 🧩 第2步：拖曳欄位到「報表欄位」，並排序")
        st.caption("僅開放「成分、劑型、劑量、規格、廠商」五個欄位可拖曳排列。")

        dnd_key = f"dnd_{'_'.join(sorted(comps_selected))}"
        state_key = f"pivot_state_{dnd_key}"
        DEFAULT_SELECTED_FIELDS = ["成分", "劑型", "劑量", "廠商"]
        if state_key not in st.session_state:
            default_selected = [c for c in DEFAULT_SELECTED_FIELDS if c in FIXED_FIELDS]
            default_available = [c for c in FIXED_FIELDS if c not in default_selected]
            st.session_state[state_key] = {"available": default_available, "selected": default_selected}
        else:
            prev = st.session_state[state_key]
            avail = [c for c in prev["available"] if c in FIXED_FIELDS]
            sel = [c for c in prev["selected"] if c in FIXED_FIELDS]
            known = set(avail) | set(sel)
            avail += [c for c in FIXED_FIELDS if c not in known]
            st.session_state[state_key] = {"available": avail, "selected": sel}

        use_fallback = st.checkbox(
            "⚠️ 拖曳排序出現錯誤時，改用勾選方式選擇欄位 (勾選順序＝欄位順序)",
            value=False, key=f"use_fallback_{dnd_key}",
        )

        if HAS_SORTABLES and not use_fallback:
            containers = sort_items(
                [
                    {"header": "📋 可用欄位 (拖曳到下方使用)", "items": st.session_state[state_key]["available"]},
                    {"header": "📊 報表欄位 (由左到右排列)", "items": st.session_state[state_key]["selected"]},
                ],
                multi_containers=True, direction="horizontal", key=dnd_key,
            )
            if containers and len(containers) > 1:
                st.session_state[state_key] = {"available": containers[0]["items"], "selected": containers[1]["items"]}
            row_fields = st.session_state[state_key]["selected"]
        elif HAS_SORTABLES:
            sel = st.multiselect(
                "選擇報表欄位 (依勾選順序排列)",
                options=FIXED_FIELDS, default=st.session_state[state_key]["selected"],
                key=f"fallback_fields_{dnd_key}",
            )
            st.session_state[state_key] = {"available": [c for c in FIXED_FIELDS if c not in sel], "selected": sel}
            row_fields = sel
        else:
            st.error("尚未安裝 streamlit-sortables 套件，暫以勾選方式呈現，請於 requirements.txt 加入 streamlit-sortables 後重新部署即可拖曳。")
            row_fields = st.multiselect("選擇報表欄位", options=FIXED_FIELDS, key="fallback_fields_static")

        if not row_fields:
            st.info("請至少拖曳一個欄位到「報表欄位」區塊。")
        else:
            st.markdown("### 🔍 第3步：逐欄篩選與合計 (點欄位下方的「🔽 篩選 / 合計」展開設定)")
            st.caption("💡 篩選會彼此連動：某欄位選擇後，其他欄位只會列出仍有對應資料的選項。")

            current_filters = {}
            for f in row_fields:
                key = f"filt_{dnd_key}_{f}"
                if key in st.session_state and st.session_state[key]:
                    current_filters[f] = st.session_state[key]

            filter_ui_cols = st.columns(len(row_fields))
            filters = {}
            subtotal_fields_selected = []
            for i, f in enumerate(row_fields):
                with filter_ui_cols[i]:
                    st.markdown(f"**{f}**")
                    with st.expander("🔽 篩選 / 合計", expanded=False):
                        df_scope = df_comp
                        for other_col, other_sel in current_filters.items():
                            if other_col != f and other_sel:
                                df_scope = df_scope[df_scope[other_col].isin(other_sel)]
                        options = sorted(
                            [v for v in df_scope[f].dropna().unique() if str(v).strip() != ""],
                            key=parse_dosage if f == "劑量" else None,
                        )

                        key = f"filt_{dnd_key}_{f}"
                        if key in st.session_state:
                            st.session_state[key] = [v for v in st.session_state[key] if v in options]

                        sel = st.multiselect(f"篩選「{f}」(不選代表全選)", options=options, key=key)
                        if sel:
                            filters[f] = sel
                        if st.checkbox("Σ 此欄要合計", key=f"sub_{dnd_key}_{f}"):
                            subtotal_fields_selected.append(f)
            subtotal_fields = [f for f in row_fields if f in subtotal_fields_selected]

            df_filtered = df_comp.copy()
            for col, sel in filters.items():
                df_filtered = df_filtered[df_filtered[col].isin(sel)]

            st.caption(f"篩選後共 **{len(df_filtered):,}** 筆資料")

            st.markdown("### 🧮 第4步：選擇要加總的數值欄位 (預設帶入 2023~2026年 全部數量)")

            qty_options = list(QTY_COL_MAP.values())
            with st.expander("🔧 如需排除年度請展開勾選 (預設全選)", expanded=False):
                sel_v = st.multiselect(
                    "選擇要加總的數值欄位", options=qty_options, default=qty_options, key=f"vals_{dnd_key}",
                )

            order_key = f"vals_order_{dnd_key}"
            if order_key not in st.session_state:
                st.session_state[order_key] = list(sel_v)
            else:
                prev_order = [c for c in st.session_state[order_key] if c in sel_v]
                prev_order += [c for c in sel_v if c not in set(prev_order)]
                st.session_state[order_key] = prev_order

            if HAS_SORTABLES and sel_v:
                sel_sig = "_".join(sorted(sel_v))
                ordered_v = sort_items(
                    st.session_state[order_key], direction="horizontal", key=f"vals_sort_{dnd_key}_{sel_sig}",
                )
                st.session_state[order_key] = ordered_v
                value_cols = ordered_v
                st.caption("💡 可直接拖曳上方欄位調整加總／顯示順序。")
            elif not sel_v:
                value_cols = []
            else:
                value_cols = sel_v

            qty_years_avail = sorted(set(c[:4] for c in value_cols if "申報量" in c))
            has_qty = len(qty_years_avail) > 0
            default_pct_years = [y for y in ["2025", "2026"] if y in qty_years_avail]
            pct_years = st.multiselect(
                "➕ 加入年度占比(%) (可只選需要的年份)",
                options=qty_years_avail, default=default_pct_years, disabled=not has_qty, key=f"pct_{dnd_key}",
            )
            add_growth = st.checkbox("➕ 加入年度成長率(%)", value=False, disabled=not has_qty, key=f"growth_{dnd_key}")

            filename_parts = list(comps_selected)
            for col in row_fields:
                if col in filters:
                    filename_parts.append("_".join(filters[col]))
            filename_parts.append("廠商申報量排名")
            report_title = "、".join(comps_selected) + "廠商申報量排名"

            raw_filename = re.sub(r'[\\/*?:"<>|]', "_", "_".join(filename_parts))
            MAX_FILENAME_BYTES = 200  # 留安全餘裕，避免超過檔案系統上限 (通常 255 bytes)，
            # 也避免選太多成分/篩選條件時檔名過長導致下載失敗
            ext_reserve = 5  # 保留給 ".xlsx" / ".png" 副檔名的空間
            if len(raw_filename.encode("utf-8")) + ext_reserve <= MAX_FILENAME_BYTES:
                safe_filename = raw_filename
            else:
                budget = MAX_FILENAME_BYTES - ext_reserve - 3  # 預留「...」空間
                truncated = raw_filename
                while len(truncated.encode("utf-8")) > budget and len(truncated) > 0:
                    truncated = truncated[:-1]
                safe_filename = truncated + "..."

            st.divider()

            if value_cols and not df_filtered.empty:
                rows, pct_cols, growth_cols = build_nested_rows(
                    df_filtered, row_fields, subtotal_fields, value_cols, pct_years, add_growth
                )
                st.markdown("### 📄 報表即時預覽")
                table_html = build_html_table(rows, row_fields, value_cols, pct_cols, growth_cols, report_title)
                st.markdown(table_html, unsafe_allow_html=True)

                st.markdown("### 📥 下載")
                dl_col1, dl_col2 = st.columns(2)
                safe_dnd_key = re.sub(r"[^0-9A-Za-z_]", "_", dnd_key)
                with dl_col1:
                    excel_bytes = generate_excel_bytes(rows, row_fields, value_cols, pct_cols, growth_cols, report_title)
                    render_excel_share(excel_bytes, safe_filename, uid=f"pivot_{safe_dnd_key}")
                with dl_col2:
                    components.html(
                        CAPTURE_HTML_TEMPLATE.format(table_html=table_html, filename=safe_filename, uid=f"pivot_{safe_dnd_key}"),
                        height=90,
                    )
            elif not value_cols:
                st.warning("⚠️ 請至少選擇一個要加總的數值欄位。")
            else:
                st.warning("❌ 篩選後無資料，請放寬篩選條件。")
