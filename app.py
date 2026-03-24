# app.py
# 放到：excel_dashboard02/app.py
# 執行：streamlit run app.py

import io
import urllib.parse
import streamlit as st
import pandas as pd
import altair as alt
from datetime import date
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Weblink M365 續約精準行銷｜企業儀表板", page_icon="📊", layout="wide")

# ---------------------------
# 🎨 全域專業樣式注入
# ---------------------------
_GLOBAL_CSS = """
<style>
/* === 全域字體與背景 === */
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700;900&family=Inter:wght@300;400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Noto Sans TC', 'Inter', sans-serif !important;
}

/* 主背景 */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #f0f4f8 0%, #e8edf5 50%, #f0f4f8 100%) !important;
    min-height: 100vh;
}

[data-testid="stApp"] {
    background: transparent !important;
}

.main .block-container {
    padding: 1.5rem 2rem 2rem 2rem !important;
    max-width: 1600px !important;
}

/* === 頁面標題 === */
h1 {
    font-size: 1.8rem !important;
    font-weight: 900 !important;
    background: linear-gradient(135deg, #1a237e 0%, #283593 50%, #1565c0 100%);
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
    background-clip: text !important;
    letter-spacing: -0.5px;
    padding-bottom: 0.3rem;
}

/* === subheader 樣式 === */
h2, [data-testid="stHeadingWithActionElements"] h2 {
    font-size: 1.1rem !important;
    font-weight: 700 !important;
    color: #1a237e !important;
    border-left: 4px solid #1565c0;
    padding: 0.4rem 0 0.4rem 0.8rem !important;
    margin-top: 1.2rem !important;
    margin-bottom: 0.5rem !important;
    background: rgba(21, 101, 192, 0.06);
    border-radius: 0 6px 6px 0;
}

/* === 側邊欄美化 === */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a237e 0%, #283593 40%, #1565c0 100%) !important;
    border-right: none !important;
    box-shadow: 4px 0 20px rgba(0,0,0,0.15) !important;
}

/* 側邊欄標籤、說明文字 */
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stCaption,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] small {
    color: rgba(255,255,255,0.88) !important;
}

[data-testid="stSidebar"] hr {
    border-color: rgba(255,255,255,0.2) !important;
}

/* 側邊欄 h1/h2/h3 */
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: white !important;
    -webkit-text-fill-color: white !important;
}

/* 輸入框容器改白底深字，確保可讀 */
[data-testid="stSidebar"] [data-baseweb="input"],
[data-testid="stSidebar"] [data-baseweb="textarea"] {
    background: rgba(255,255,255,0.95) !important;
    border: 1px solid rgba(255,255,255,0.5) !important;
    border-radius: 8px !important;
}
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] textarea {
    color: #1a237e !important;
    -webkit-text-fill-color: #1a237e !important;
    background: transparent !important;
}
[data-testid="stSidebar"] input::placeholder {
    color: #9e9e9e !important;
    -webkit-text-fill-color: #9e9e9e !important;
}

/* select 選擇框 */
[data-testid="stSidebar"] [data-baseweb="select"] > div:first-child {
    background: rgba(255,255,255,0.95) !important;
    border: 1px solid rgba(255,255,255,0.5) !important;
    border-radius: 8px !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] span,
[data-testid="stSidebar"] [data-baseweb="select"] div {
    color: #1a237e !important;
    -webkit-text-fill-color: #1a237e !important;
}
/* select 內部箭頭保持可見 */
[data-testid="stSidebar"] [data-baseweb="select"] svg {
    fill: #1a237e !important;
}

/* multiselect 已選 tag */
[data-testid="stSidebar"] [data-baseweb="tag"] {
    background: #1565c0 !important;
    border: 1px solid rgba(255,255,255,0.4) !important;
    border-radius: 6px !important;
}
[data-testid="stSidebar"] [data-baseweb="tag"] span,
[data-testid="stSidebar"] [data-baseweb="tag"] div {
    color: white !important;
    -webkit-text-fill-color: white !important;
}

/* 日期輸入框 */
[data-testid="stSidebar"] [data-testid="stDateInput"] [data-baseweb="input"] {
    background: rgba(255,255,255,0.95) !important;
}
[data-testid="stSidebar"] [data-testid="stDateInput"] input {
    color: #1a237e !important;
    -webkit-text-fill-color: #1a237e !important;
}

/* 上傳元件 */
[data-testid="stSidebar"] [data-testid="stFileUploader"] {
    background: rgba(255,255,255,0.12) !important;
    border: 2px dashed rgba(255,255,255,0.4) !important;
    border-radius: 10px !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploader"] * {
    color: rgba(255,255,255,0.9) !important;
    -webkit-text-fill-color: rgba(255,255,255,0.9) !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
    background: rgba(255,255,255,0.2) !important;
    border: 1px solid rgba(255,255,255,0.4) !important;
    color: white !important;
    -webkit-text-fill-color: white !important;
    border-radius: 6px !important;
}

/* checkbox */
[data-testid="stSidebar"] [data-testid="stCheckbox"] label,
[data-testid="stSidebar"] [data-testid="stCheckbox"] span {
    color: rgba(255,255,255,0.9) !important;
    -webkit-text-fill-color: rgba(255,255,255,0.9) !important;
}

/* 側邊欄按鈕 */
[data-testid="stSidebar"] .stButton button {
    background: rgba(255,255,255,0.15) !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
    color: white !important;
    -webkit-text-fill-color: white !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    transition: all 0.2s ease !important;
}
[data-testid="stSidebar"] .stButton button:hover {
    background: rgba(255,255,255,0.25) !important;
    border-color: rgba(255,255,255,0.5) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2) !important;
}

/* subheader 在 sidebar */
[data-testid="stSidebar"] h2 {
    color: rgba(255,255,255,0.95) !important;
    -webkit-text-fill-color: rgba(255,255,255,0.95) !important;
    background: rgba(255,255,255,0.12) !important;
    border-left: 3px solid rgba(255,255,255,0.6) !important;
}

/* === KPI 指標卡片 === */
[data-testid="stMetric"] {
    background: white !important;
    border-radius: 14px !important;
    padding: 1rem 1.2rem !important;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07), 0 0 0 1px rgba(0,0,0,0.04) !important;
    transition: transform 0.2s ease, box-shadow 0.2s ease !important;
    border-top: 3px solid #1565c0 !important;
}

[data-testid="stMetric"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 20px rgba(21,101,192,0.15), 0 0 0 1px rgba(21,101,192,0.1) !important;
}

[data-testid="stMetricLabel"] {
    color: #546e7a !important;
    font-size: 0.75rem !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.5px !important;
}

[data-testid="stMetricValue"] {
    color: #1a237e !important;
    font-size: 1.6rem !important;
    font-weight: 800 !important;
    font-family: 'Inter', sans-serif !important;
    letter-spacing: -0.5px !important;
}

/* === 資料表美化 === */
[data-testid="stDataFrame"],
[data-testid="stDataEditor"] {
    border-radius: 12px !important;
    overflow: visible !important;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06) !important;
    border: 1px solid rgba(0,0,0,0.06) !important;
}

/* 保留資料表右上角 Show chart / Show data 控制列 */
[data-testid="stDataFrame"] > div,
[data-testid="stDataEditor"] > div {
    overflow: visible !important;
}

/* === 圖表容器：保留 Show chart / Show data 工具列可見 === */
[data-testid="stVegaLiteChart"],
[data-testid="stVegaLiteChart"] > div,
[data-testid="stVegaLiteChart"] [data-testid="stToolbar"] {
    overflow: visible !important;
}

[data-testid="stVegaLiteChart"] .vega-embed,
[data-testid="stVegaLiteChart"] .vega-embed > div {
    max-width: 100% !important;
    overflow: visible !important;
}

[data-testid="stVegaLiteChart"] {
    padding-top: 0.35rem !important;
    padding-bottom: 0.5rem !important;
    min-height: 420px !important;
}

/* === 分隔線 === */
hr {
    border: none !important;
    border-top: 2px solid rgba(21,101,192,0.1) !important;
    margin: 1.5rem 0 !important;
}

/* === caption 樣式 === */
[data-testid="stCaptionContainer"],
.stCaption {
    color: #78909c !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.2px;
}

/* === status/progress === */
[data-testid="stStatus"] {
    border-radius: 10px !important;
    border: 1px solid rgba(21,101,192,0.15) !important;
    background: white !important;
}

[data-testid="stProgressBar"] > div {
    background: linear-gradient(90deg, #1565c0, #42a5f5) !important;
    border-radius: 4px !important;
}

/* === 上傳器 === */
[data-testid="stFileUploader"] {
    background: white !important;
    border-radius: 12px !important;
    padding: 0.5rem !important;
    border: 2px dashed rgba(255,255,255,0.3) !important;
}

/* === info/warning boxes === */
[data-testid="stInfo"] {
    background: rgba(21,101,192,0.06) !important;
    border: 1px solid rgba(21,101,192,0.15) !important;
    border-radius: 10px !important;
    border-left: 4px solid #1565c0 !important;
}

/* === 圖表容器 === */
[data-testid="stVegaLiteChart"] {
    background: white !important;
    border-radius: 14px !important;
    padding: 0.8rem !important;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06) !important;
    border: 1px solid rgba(0,0,0,0.05) !important;
    overflow: visible !important;
    max-width: 100% !important;
}

/* 修正 Altair canvas/svg 超出容器問題，同時保留 Show chart 工具列 */
[data-testid="stVegaLiteChart"] canvas,
[data-testid="stVegaLiteChart"] svg {
    max-width: 100% !important;
    height: auto !important;
    display: block !important;
}

/* === multiselect 標籤 === */
[data-baseweb="tag"] {
    background: rgba(21,101,192,0.12) !important;
    border: 1px solid rgba(21,101,192,0.25) !important;
    border-radius: 6px !important;
}

/* === checkbox === */
[data-testid="stCheckbox"] label {
    font-weight: 500 !important;
    font-size: 0.88rem !important;
}

/* === selectbox === */
[data-baseweb="select"] {
    border-radius: 8px !important;
}

/* === 版面底部留白 === */
footer { display: none !important; }
#MainMenu { display: none !important; }

/* overflow 修正 */
html, body { overflow-x: hidden !important; max-width: 100% !important; }
.stApp, .main, section.main, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
    overflow-x: hidden !important;
    max-width: 100% !important;
}

/* === 圖表 toolbar：Show data / Show chart 按鈕 === */
[data-testid="stVegaLiteChart"] [data-testid="StyledFullScreenButton"],
[data-testid="stVegaLiteChart"] button[title="View fullscreen"],
[data-testid="stVegaLiteChart"] [aria-label="Show data"],
[data-testid="stVegaLiteChart"] [aria-label="Show chart"] {
    color: #1565c0 !important;
    background: white !important;
    border: 1px solid rgba(21,101,192,0.2) !important;
    border-radius: 6px !important;
}
/* 強制圖表 toolbar 按鈕可見 */
[data-testid="stVegaLiteChart"] .vega-embed .chart-wrapper ~ details summary,
[data-testid="stVegaLiteChart"] .vega-embed summary,
[data-testid="stVegaLiteChart"] details summary {
    color: #1565c0 !important;
}
/* Show data / Show chart 切換按鈕 (Streamlit wraps in data-testid) */
button[data-testid="StyledFullScreenButton"] { color: #546e7a !important; }
/* 圖表上方 toolbar 的所有按鈕 */
[data-testid="stVegaLiteChart"] > div > div:last-child button,
[data-testid="stVegaLiteChart"] > div > div:last-child a {
    color: #546e7a !important;
    background: rgba(255,255,255,0.9) !important;
}

/* === Vega toolbar (Show data / Show chart) === */
.vega-embed .vega-actions a,
.vega-embed details > summary {
    color: #1565c0 !important;
}
.vega-embed .chart-wrapper {
    overflow: visible !important;
    max-width: 100% !important;
}

/* === 側邊欄上傳器 — 白色方塊修正 === */
[data-testid="stSidebar"] [data-testid="stFileUploader"] > div {
    background: rgba(255,255,255,0.08) !important;
    border: 2px dashed rgba(255,255,255,0.4) !important;
    border-radius: 10px !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploader"] > div > div {
    background: transparent !important;
}
/* 上傳區域拖放框本體 */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background: rgba(255,255,255,0.08) !important;
    border: 2px dashed rgba(255,255,255,0.35) !important;
    border-radius: 10px !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * {
    color: rgba(255,255,255,0.9) !important;
    -webkit-text-fill-color: rgba(255,255,255,0.9) !important;
}
/* 上傳按鈕 */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInput"] + div button,
[data-testid="stSidebar"] [data-testid="baseButton-secondary"] {
    background: rgba(255,255,255,0.18) !important;
    color: white !important;
    -webkit-text-fill-color: white !important;
    border: 1px solid rgba(255,255,255,0.4) !important;
    border-radius: 7px !important;
}
</style>
"""

# ---------------------------
# 常數設定
# ---------------------------
#DEFAULT_XLSX_PATH = r"D:\DeskT\Austin 自動化\新增資料夾 (2)\CSP訂單資料_raw.xlsx"

DEFAULT_XLSX_PATH = "CSP訂單資料_raw.xlsx"

DROP_ORDER_STATUS = {"下單異常", "已取消", "已退貨"}

KEEP_COLUMNS = [
    "經銷商",
    "資格",
    "最終客戶",
    "訂閱動作",
    "產品分類",
    "商品名稱",
    "數量",
    "展碁COST單價未稅",
    "展碁COST未稅小計",
    "成交單價未稅",
    "成交價未稅小計",
    "訂閱到期日",
    "訂單狀態",
    "訂單下單日",
    "客戶微軟網域",
    "展碁業務",
]

NUMERIC_COLUMNS = [
    "數量",
    "成交單價未稅",
    "成交價未稅小計",
    "展碁COST單價未稅",
    "展碁COST未稅小計",
]

DATE_COLUMNS = ["訂閱到期日", "訂單下單日"]
ANNUAL_COL = "展碁COST未稅小計(年)"
EXPIRY_YEAR_COL = "訂閱到期日年度"
ANALYSIS_VALUE_COL = "成交價未稅小計"

# ✅ 差異欄位/數值棕色（你的指定）
BROWN_COLOR = "#8B4513"  # saddlebrown

# ✅ 12-B 今年度分組明細表標題紫色
PURPLE_COLOR = "#6A0DAD"

WARNING_SPECS = [
    (15, "🔴", "#D32F2F"),
    (30, "🟠", "#F57C00"),
    (45, "🟡", "#C9A227"),
    (60, "🔵", "#1976D2"),
    (90, "🟣", "#7B1FA2"),
]


# ---------------------------
# 工具函數
# ---------------------------
def _coerce_numeric(series: pd.Series) -> pd.Series:
    """將欄位轉成數字：移除逗號/空白/常見貨幣符號，允許 NaN。"""
    if series is None:
        return series
    s = series.astype(str).str.replace(",", "", regex=False)
    s = s.str.replace(" ", "", regex=False)
    for sym in ["$", "NT$", "NTD", "TWD", "元", "＄"]:
        s = s.str.replace(sym, "", regex=False)
    s = s.str.replace("—", "", regex=False)
    return pd.to_numeric(s, errors="coerce")


def _safe_thousands_formatter(v) -> str:
    """Styler.format 用的安全 formatter：空值/字串回空白；數值才做千分位。"""
    try:
        if v is None or pd.isna(v):
            return ""
    except Exception:
        pass

    if isinstance(v, str):
        return ""

    try:
        fv = float(v)
    except Exception:
        return ""

    return f"{fv:,.0f}"


def _safe_date_formatter(v) -> str:
    """日期顯示安全 formatter：空值回空白；日期轉 YYYY-MM-DD。"""
    try:
        if v is None or pd.isna(v):
            return ""
    except Exception:
        pass

    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")

    try:
        return pd.to_datetime(v).strftime("%Y-%m-%d")
    except Exception:
        return ""


def _style_header_color_for_column(df: pd.DataFrame, col_name: str, color: str):
    """
    給 st.dataframe(Styler) 用：指定欄位的「表頭」改色。
    Streamlit 會保留 pandas Styler 的 header style。
    """
    if df is None or df.empty or col_name not in df.columns:
        return []

    idx = list(df.columns).index(col_name)
    # pandas Styler header selector：th.col_heading.level0.col{idx}
    return [
        {
            "selector": f"th.col_heading.level0.col{idx}",
            "props": [("color", color), ("font-weight", "700")],
        }
    ]


@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: bytes | None, use_upload: bool, default_path: str) -> pd.DataFrame:
    """加速：避免每次 rerun 都重讀 Excel。"""
    if use_upload:
        return pd.read_excel(io.BytesIO(file_bytes))
    return pd.read_excel(default_path)


@st.cache_data(show_spinner=False)
def clean_transform_cached(df_raw: pd.DataFrame) -> pd.DataFrame:
    """加速：清洗/轉型做快取。"""
    if "訂單狀態" in df_raw.columns:
        df_raw = df_raw[~df_raw["訂單狀態"].astype(str).isin(DROP_ORDER_STATUS)].copy()

    missing = [c for c in KEEP_COLUMNS if c not in df_raw.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要欄位：{missing}")

    df = df_raw[KEEP_COLUMNS].copy()

    for c in NUMERIC_COLUMNS:
        df[c] = _coerce_numeric(df[c])

    for c in DATE_COLUMNS:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    # ✅ 訂閱到期日年度
    df[EXPIRY_YEAR_COL] = df["訂閱到期日"].dt.year

    # ✅ 年小計欄位（維持既有邏輯）
    df[ANNUAL_COL] = (df["數量"].fillna(0) * df["展碁COST單價未稅"].fillna(0) * 12)

    # 欄位順序調整：年度欄位緊接在「訂閱到期日」後；年小計緊接在「展碁COST未稅小計」後
    cols = df.columns.tolist()

    if EXPIRY_YEAR_COL in cols and "訂閱到期日" in cols:
        cols.remove(EXPIRY_YEAR_COL)
        cols.insert(cols.index("訂閱到期日") + 1, EXPIRY_YEAR_COL)

    if ANNUAL_COL in cols and "展碁COST未稅小計" in cols:
        cols.remove(ANNUAL_COL)
        cols.insert(cols.index("展碁COST未稅小計") + 1, ANNUAL_COL)

    df = df[cols]
    return df


def apply_filters(df: pd.DataFrame, ui: dict, base_today: date | None = None) -> pd.DataFrame:
    """
    - 勾選「未來 N 月內到期」後：忽略「訂閱到期日區間」與「訂單下單日區間」
    - 其他多條件篩選仍適用
    - base_today：僅供「未來 N 月內到期」模式使用（預設 date.today()），用於產生明年度範圍時不影響既有行為
    """
    out = df.copy()

    if ui["future_expiry_enabled"]:
        today = base_today or date.today()
        end_date = today + relativedelta(months=ui["future_expiry_months"])
        out = out[(out["訂閱到期日"].dt.date >= today) & (out["訂閱到期日"].dt.date <= end_date)]
    else:
        exp_from, exp_to = ui["expiry_range"]
        if exp_from and exp_to:
            out = out[(out["訂閱到期日"].dt.date >= exp_from) & (out["訂閱到期日"].dt.date <= exp_to)]

        order_from, order_to = ui["order_range"]
        if order_from and order_to:
            out = out[(out["訂單下單日"].dt.date >= order_from) & (out["訂單下單日"].dt.date <= order_to)]

    multi_filters = [
        ("經銷商", ui["dealer"]),
        ("資格", ui["qual"]),
        ("最終客戶", ui["customer"]),
        ("訂閱動作", ui["action"]),
        ("商品名稱", ui["product"]),
        ("展碁業務", ui["sales"]),
    ]
    for col, selected in multi_filters:
        if selected:
            out = out[out[col].astype(str).isin(selected)]

    return out


def safe_sum(s: pd.Series) -> float:
    return float(pd.to_numeric(s, errors="coerce").fillna(0).sum()) if s is not None else 0.0


def format_money(v: float) -> str:
    try:
        return f"{float(v):,.0f}"
    except Exception:
        return str(v)


def format_signed_int(v: float | int) -> str:
    try:
        iv = int(round(float(v)))
    except Exception:
        return str(v)
    return f"{iv:+,}"


def format_signed_money(v: float) -> str:
    try:
        fv = float(v)
    except Exception:
        return str(v)
    sign = "+" if fv >= 0 else "-"
    return f"{sign}{abs(fv):,.0f}"


def get_group_warning_meta(min_expiry_dt) -> tuple[str, str, int | None]:
    """依距今最近到期日回傳（顯示文字, 顏色, 門檻天數）。"""
    if min_expiry_dt is None or pd.isna(min_expiry_dt):
        return "", "", None

    try:
        min_dt = pd.to_datetime(min_expiry_dt)
    except Exception:
        return "", "", None

    days_left = (min_dt.date() - date.today()).days

    # ✅ 新增：若尚未超過「到期日後 30 天」，且明年度尚無續約資訊，顯示寬限期
    # -30 ~ -1：代表已過到期日，但仍在 30 天內
    if -30 <= days_left < 0:
        return "寬限期", "#8B4513", -30

    if days_left < -30:
        return "🔴 已到期", "#D32F2F", 0

    for threshold, symbol, color in WARNING_SPECS:
        if days_left <= threshold:
            return f"{symbol} {threshold}", color, threshold

    return "", "", None


def build_group_renewal_lookup(df_next: pd.DataFrame | None) -> dict:
    """
    建立明年度續約資訊查詢表：
    key = (最終客戶, 訂閱到期日年度)
    value = {"amount": 明年度訂閱總金額, "count": 筆數}
    """
    if df_next is None or df_next.empty:
        return {}

    required_cols = {"最終客戶", EXPIRY_YEAR_COL, ANALYSIS_VALUE_COL}
    if not required_cols.issubset(set(df_next.columns)):
        return {}

    d = df_next[["最終客戶", EXPIRY_YEAR_COL, ANALYSIS_VALUE_COL]].copy()
    d[ANALYSIS_VALUE_COL] = pd.to_numeric(d[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)

    grouped = (
        d.groupby(["最終客戶", EXPIRY_YEAR_COL], dropna=False)
        .agg(
            明年度訂閱總金額=(ANALYSIS_VALUE_COL, "sum"),
            筆數=(ANALYSIS_VALUE_COL, "size"),
        )
        .reset_index()
    )

    lookup: dict = {}
    for _, row in grouped.iterrows():
        cust = str(row.get("最終客戶", "") or "")
        yr = row.get(EXPIRY_YEAR_COL, pd.NA)

        try:
            yr_key = int(float(yr))
        except Exception:
            yr_key = str(yr)

        lookup[(cust, yr_key)] = {
            "amount": float(row.get("明年度訂閱總金額", 0) or 0),
            "count": int(row.get("筆數", 0) or 0),
        }

    return lookup


def get_group_warning_meta_with_renewal(
    min_expiry_dt,
    current_total: float,
    customer,
    expiry_year,
    renewal_lookup: dict | None = None,
) -> tuple[str, str, int | None]:
    """
    今年度分組明細表警示邏輯：
    1. 若明年度已有續約資訊，優先顯示「已續約」
       - 明年度金額 <= 今年度：綠色
       - 明年度金額 > 今年度：紅色
    2. 若無明年度續約資訊，沿用原本 15/30/45/60/90 與已到期邏輯
    """
    renewal_lookup = renewal_lookup or {}
    cust_key = str(customer or "")

    yr_key_candidates = []
    try:
        yr_int = int(float(expiry_year))
        yr_key_candidates.append(yr_int + 1)
    except Exception:
        pass

    if expiry_year is not None and not pd.isna(expiry_year):
        yr_key_candidates.append(str(expiry_year))

    for next_year_key in yr_key_candidates:
        renewal_info = renewal_lookup.get((cust_key, next_year_key))
        if renewal_info and int(renewal_info.get("count", 0) or 0) > 0:
            next_amount = float(renewal_info.get("amount", 0) or 0)
            if next_amount > float(current_total or 0):
                return "已續約", "#D32F2F", None
            return "已續約", "#2E7D32", None

    return get_group_warning_meta(min_expiry_dt)


def build_warning_style_map(values: pd.Series | list) -> dict:
    """給 Styler 套 cell 顏色。"""
    items = pd.Series(values).fillna("").astype(str).tolist()
    style_map = {}
    for v in items:
        vv = v.strip()
        if not vv:
            continue
        if vv in {"已續約", "🟢 已續約", "🔴 已續約"}:
            # 已續約顏色需由隱藏欄 _warning_color 決定，這裡先不寫死。
            style_map[v] = "font-weight:700;"
            continue
        if vv == "寬限期" or "寬限期" in vv:
            style_map[v] = "color:#8B4513; font-weight:700;"
            continue
        if "已到期" in vv:
            style_map[v] = "color:#D32F2F; font-weight:700;"
            continue
        for threshold, symbol, color in WARNING_SPECS:
            if vv.startswith(symbol) or vv.endswith(str(threshold)) or f" {threshold}" in vv:
                style_map[v] = f"color:{color}; font-weight:700;"
                break
    return style_map


def format_warning_display_text(warning_text, warning_color="", warning_threshold=None) -> str:
    """將未續約示警轉為適合 data_editor 顯示的圖示文字。"""
    if pd.isna(warning_text) or str(warning_text).strip() == "":
        return ""

    text_val = str(warning_text).strip()
    color_val = "" if pd.isna(warning_color) else str(warning_color).strip().lower()

    if text_val == "已續約":
        if color_val == "#2e7d32":
            return "🟢 已續約"
        if color_val == "#d32f2f":
            return "🔴 已續約"
        return "✅ 已續約"

    if text_val == "寬限期":
        return "🟤 寬限期"

    if "已到期" in text_val:
        return "🔴 已到期"

    if warning_threshold is not None and not pd.isna(warning_threshold):
        try:
            th = int(float(warning_threshold))
            icon_map = {15: "🟡", 30: "🟠", 45: "🟣", 60: "🟤", 90: "⚫"}
            return f"{icon_map.get(th, '🟨')} {th}天"
        except Exception:
            pass

    return text_val


def _merge_selected_into_options(raw_opts: list[str], selected_raw: list[str] | None) -> list[str]:
    """保留目前已選值，避免時間條件變更時 multiselect 自動把已選項目清掉。"""
    out = list(raw_opts or [])
    for v in (selected_raw or []):
        vv = str(v).strip()
        if vv and vv not in out:
            out.append(vv)
    return out


def _normalize_customer_year_key(customer, expiry_year):
    cust = str(customer or "")
    try:
        yr = int(float(expiry_year))
    except Exception:
        yr = str(expiry_year)
    return cust, yr


def apply_warning_filter_to_datasets(
    df_this: pd.DataFrame,
    df_next: pd.DataFrame | None,
    warning_selected: list[str] | None,
) -> tuple[pd.DataFrame, pd.DataFrame | None]:
    """
    將 sidebar 的「未續約示警篩選」同步套用到右側所有資訊。
    依今年度分組 Header 的示警狀態挑出 (最終客戶, 到期年度)，
    再過濾今年度與其對應的明年度資料。
    """
    selected = [str(x).strip() for x in (warning_selected or []) if str(x).strip()]
    if not selected:
        return df_this, df_next

    if df_this is None or df_this.empty:
        return df_this, df_next

    grouped = build_grouped_detail_report_v2(df_this, df_next=df_next)
    if grouped is None or grouped.empty or '_row_type' not in grouped.columns:
        return df_this.iloc[0:0].copy(), (df_next.iloc[0:0].copy() if df_next is not None else df_next)

    headers = grouped[grouped['_row_type'].astype(str) == 'HEADER'].copy()
    if headers.empty:
        return df_this.iloc[0:0].copy(), (df_next.iloc[0:0].copy() if df_next is not None else df_next)

    headers['未續約示警顯示'] = headers.apply(
        lambda r: format_warning_display_text(
            r.get('未續約示警', ''),
            r.get('_warning_color', ''),
            r.get('_warning_threshold', pd.NA),
        ),
        axis=1,
    )

    headers = headers[headers['未續約示警顯示'].astype(str).str.strip().isin(selected)].copy()
    if headers.empty:
        return df_this.iloc[0:0].copy(), (df_next.iloc[0:0].copy() if df_next is not None else df_next)

    current_keys = set()
    next_keys = set()
    customers = set()
    for _, row in headers.iterrows():
        cust, yr = _normalize_customer_year_key(row.get('最終客戶', ''), row.get(EXPIRY_YEAR_COL, pd.NA))
        current_keys.add((cust, yr))
        customers.add(cust)
        if isinstance(yr, int):
            next_keys.add((cust, yr + 1))

    def _filter_by_keys(df_in: pd.DataFrame | None, valid_keys: set[tuple]) -> pd.DataFrame | None:
        if df_in is None:
            return df_in
        if df_in.empty:
            return df_in.copy()
        if not valid_keys:
            return df_in.iloc[0:0].copy()
        keys = df_in.apply(lambda r: _normalize_customer_year_key(r.get('最終客戶', ''), r.get(EXPIRY_YEAR_COL, pd.NA)), axis=1)
        mask = keys.isin(valid_keys)
        return df_in[mask].copy()

    df_this_out = _filter_by_keys(df_this, current_keys)

    if df_next is None:
        return df_this_out, df_next

    if next_keys:
        df_next_out = _filter_by_keys(df_next, next_keys)
    else:
        df_next_out = df_next[df_next['最終客戶'].astype(str).isin(customers)].copy() if '最終客戶' in df_next.columns else df_next.copy()

    return df_this_out, df_next_out


def init_filter_defaults_from_data(df: pd.DataFrame):
    """依資料初始化日期預設值：
    - 訂單下單日（起）= 前一年 1/1
    - 訂單下單日（迄）= 匯入資料最後一日
    - 訂閱到期日維持本年度 1/1~12/31
    僅在 session_state 尚未存在時初始化，避免覆蓋使用者操作。
    """
    today_d = date.today()
    expiry_default_from = date(today_d.year, 1, 1)
    expiry_default_to = date(today_d.year, 12, 31)

    order_max = None
    if df is not None and not df.empty and "訂單下單日" in df.columns:
        order_max_ts = pd.to_datetime(df["訂單下單日"], errors="coerce").max()
        if pd.notna(order_max_ts):
            order_max = order_max_ts.date()

    order_default_from = date(today_d.year - 1, 1, 1)
    order_default_to = order_max or today_d

    if order_max is not None:
        st.session_state["order_max_date_from_data"] = order_max

    st.session_state.setdefault("expiry_from", expiry_default_from)
    st.session_state.setdefault("expiry_to", expiry_default_to)
    st.session_state.setdefault("order_from", order_default_from)
    st.session_state.setdefault("order_to", order_default_to)


def build_kpis(df: pd.DataFrame) -> dict:
    total = safe_sum(df[ANALYSIS_VALUE_COL])
    row_cnt = int(len(df))
    cust_cnt = int(df["最終客戶"].astype(str).nunique())
    dealer_cnt = int(df["經銷商"].astype(str).nunique())
    avg_per_row = (total / row_cnt) if row_cnt else 0.0
    avg_per_cust = (total / cust_cnt) if cust_cnt else 0.0

    return {
        "筆數": row_cnt,
        "最終客戶數": cust_cnt,
        "經銷商數": dealer_cnt,
        f"{ANALYSIS_VALUE_COL}合計": total,
        f"{ANALYSIS_VALUE_COL}平均每筆": avg_per_row,
        f"{ANALYSIS_VALUE_COL}平均每客戶": avg_per_cust,
    }


def build_quarterly_kpi_df(df_this: pd.DataFrame, df_next: pd.DataFrame) -> pd.DataFrame:
    """KPI 區塊下方四季圖表資料：本年度、明年度、差異（明年度 - 今年度）。"""
    quarter_order = ["Q1", "Q2", "Q3", "Q4"]

    def _quarter_sum(df_in: pd.DataFrame) -> pd.DataFrame:
        if df_in is None or df_in.empty or "訂閱到期日" not in df_in.columns:
            return pd.DataFrame({"季度": quarter_order, "金額": [0, 0, 0, 0]})

        d = df_in.copy()
        d["訂閱到期日"] = pd.to_datetime(d["訂閱到期日"], errors="coerce")
        d = d[pd.notna(d["訂閱到期日"])].copy()
        if d.empty:
            return pd.DataFrame({"季度": quarter_order, "金額": [0, 0, 0, 0]})

        d["季度"] = "Q" + d["訂閱到期日"].dt.quarter.astype(str)
        d["金額"] = pd.to_numeric(d[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)

        q = d.groupby("季度", dropna=False)["金額"].sum().reindex(quarter_order, fill_value=0).reset_index()
        return q

    q_this = _quarter_sum(df_this).rename(columns={"金額": "本年度"})
    q_next = _quarter_sum(df_next).rename(columns={"金額": "明年度"})

    merged = q_this.merge(q_next, on="季度", how="outer").fillna(0)
    merged["本年度"] = pd.to_numeric(merged["本年度"], errors="coerce").fillna(0)
    merged["明年度"] = pd.to_numeric(merged["明年度"], errors="coerce").fillna(0)
    merged["差異"] = merged["明年度"] - merged["本年度"]
    merged["季度"] = pd.Categorical(merged["季度"], categories=quarter_order, ordered=True)
    merged = merged.sort_values("季度").reset_index(drop=True)
    return merged


def top10_by(df: pd.DataFrame, group_col: str, value_col: str, top_n: int = 10, extra_col: str | None = None):
    g = df.groupby(group_col, dropna=False)[value_col].sum().reset_index()
    g[value_col] = pd.to_numeric(g[value_col], errors="coerce").fillna(0)
    g = g.sort_values(value_col, ascending=False).head(top_n).reset_index(drop=True)

    if extra_col:
        tmp = (
            df.groupby(group_col)[extra_col]
            .apply(lambda x: " / ".join(pd.Series(x).dropna().astype(str).unique()[:5]))
            .reset_index()
        )
        g = g.merge(tmp, on=group_col, how="left")
    return g


def uniq_options(df: pd.DataFrame, col: str):
    return sorted([x for x in df[col].dropna().astype(str).unique().tolist() if x != "nan"])


def uniq_options_numeric(df: pd.DataFrame, col: str):
    """側邊欄用：依數字前綴排序，純字串則照字母。回傳格式 '01. 項目名稱'"""
    raw = sorted([x for x in df[col].dropna().astype(str).unique().tolist() if x != "nan"])
    return [f"{i+1:02d}. {v}" for i, v in enumerate(raw)]


def strip_numeric_prefix(labeled: list[str]) -> list[str]:
    """將 '01. 項目名稱' 還原為 '項目名稱'"""
    import re
    return [re.sub(r"^\d+\.\s*", "", s) for s in labeled]


def uniq_options_alpha(df: pd.DataFrame, col: str):
    """右側主區域用：依英文字母 A-Z 排序，回傳格式 'A. 項目名稱'"""
    raw = sorted([x for x in df[col].dropna().astype(str).unique().tolist() if x != "nan"])
    # 以 A~Z 循環標記（若超過 26 項改用 AA, AB...）
    def _label(i):
        import string
        alpha = string.ascii_uppercase
        if i < 26:
            return alpha[i]
        return alpha[(i // 26) - 1] + alpha[i % 26]
    return {_label(i): v for i, v in enumerate(raw)}


def _clear_query_param_overlay_close():
    """相容不同 Streamlit 版本的 query_params 清除方式。"""
    try:
        st.query_params.pop("overlay_close", None)
    except Exception:
        pass
    try:
        if "overlay_close" in st.query_params:
            del st.query_params["overlay_close"]
    except Exception:
        pass


def _clear_all_detail_editor_states():
    """清掉所有 data_editor 的 state，避免 checkbox 勾選殘留。"""
    for k in list(st.session_state.keys()):
        if str(k).startswith("detail_editor_"):
            del st.session_state[k]


def reset_filters_to_defaults():
    """清除所有篩選：同步重設所有 UI 篩選項目（含下拉/日期/多選）。
    ✅ 預設僅回到「本年度」訂閱到期日區間（1/1~12/31），其餘維持原設計。
    """
    today_d = date.today()

    # ✅ 訂閱到期日：預設只看「本年度」
    year_from = date(today_d.year, 1, 1)
    year_to = date(today_d.year, 12, 31)

    # 訂單下單日：預設為前一年 1/1 ~ 匯入資料最後一日（若尚未載入資料則先到今日）
    order_default_from = date(today_d.year - 1, 1, 1)
    order_default_to = st.session_state.get("order_max_date_from_data", today_d)

    st.session_state["future_expiry_enabled"] = False
    st.session_state["future_expiry_months"] = 3

    st.session_state["expiry_from"] = year_from
    st.session_state["expiry_to"] = year_to
    st.session_state["order_from"] = order_default_from
    st.session_state["order_to"] = order_default_to

    st.session_state["dealer"] = []
    st.session_state["qual"] = []
    st.session_state["customer"] = []
    st.session_state["action"] = []
    st.session_state["product"] = []
    st.session_state["sales"] = []

    # 清除帶數字前綴的 labeled 版 session keys
    for k in ["dealer_labeled", "qual_labeled", "customer_labeled", "action_labeled", "product_labeled", "sales_labeled"]:
        st.session_state[k] = []

    st.session_state["top30_dealer_pick"] = "（不套用 Top 30）"
    st.session_state["warning_filter_pick"] = []

    st.session_state["selected_row_key"] = None
    st.session_state["email_text"] = ""
    st.session_state["last_selected_sig"] = None

    # ✅ 分組 Email overlay 相關狀態一併清除
    st.session_state["group_selected_sig"] = None
    st.session_state["group_last_selected_sig"] = None
    st.session_state["group_email_text"] = ""
    st.session_state["group_email_subject"] = ""
    st.session_state["group_mailto_link"] = ""
    # ✅ 分組「電話訪談」overlay 相關狀態一併清除
    st.session_state["group_call_selected_sig"] = None
    st.session_state["group_call_last_selected_sig"] = None
    st.session_state["group_call_text"] = ""

    _clear_all_detail_editor_states()
    st.session_state["detail_editor_version"] = st.session_state.get("detail_editor_version", 0) + 1


def handle_close_overlay_request():
    """相容保留：若仍有 overlay_close=1 進來，仍會清狀態"""
    _clear_query_param_overlay_close()
    st.session_state["selected_row_key"] = None
    st.session_state["email_text"] = ""
    st.session_state["last_selected_sig"] = None
    st.session_state["skip_selection_once"] = True

    # ✅ 分組 Email overlay 相關狀態一併清除
    st.session_state["group_selected_sig"] = None
    st.session_state["group_last_selected_sig"] = None
    st.session_state["group_email_text"] = ""
    st.session_state["group_email_subject"] = ""
    st.session_state["group_mailto_link"] = ""
    # ✅ 分組「電話訪談」overlay 相關狀態一併清除
    st.session_state["group_call_selected_sig"] = None
    st.session_state["group_call_last_selected_sig"] = None
    st.session_state["group_call_text"] = ""

    _clear_all_detail_editor_states()
    st.session_state["detail_editor_version"] = st.session_state.get("detail_editor_version", 0) + 1


def shift_ui_state_one_year(ui: dict) -> dict:
    """把原篩選範圍 +1 年（不改動原 ui 物件）。"""
    ui2 = dict(ui)

    exp_from, exp_to = ui.get("expiry_range", (None, None))
    ord_from, ord_to = ui.get("order_range", (None, None))

    def _shift(d):
        if d is None:
            return None
        try:
            return d + relativedelta(years=1)
        except Exception:
            return d

    ui2["expiry_range"] = (_shift(exp_from), _shift(exp_to))
    ui2["order_range"] = (_shift(ord_from), _shift(ord_to))
    return ui2


# ---------------------------
# ✅ 顯示「本年度/隔年度」篩選時間範圍（依程式邏輯）
# ---------------------------
def _fmt_d(d: date | None) -> str:
    if d is None:
        return "-"
    try:
        return d.strftime("%Y-%m-%d")
    except Exception:
        return str(d)


def _get_effective_ranges(ui_state: dict) -> tuple[tuple[date | None, date | None], tuple[date | None, date | None]]:
    """
    回傳：
    - 本年度有效篩選區間（依「未來N月」或「訂閱到期日(起迄)」決定）
    - 隔年度有效篩選區間（本年度區間 +1 年；未來N月則以 today+1y ~ (today+N月)+1y）
    """
    if ui_state.get("future_expiry_enabled"):
        base_today = date.today()
        end_date = base_today + relativedelta(months=int(ui_state.get("future_expiry_months", 3)))
        this_rng = (base_today, end_date)

        base_today_next = base_today + relativedelta(years=1)
        end_date_next = base_today_next + relativedelta(months=int(ui_state.get("future_expiry_months", 3)))
        next_rng = (base_today_next, end_date_next)
        return this_rng, next_rng

    exp_from, exp_to = ui_state.get("expiry_range", (None, None))
    this_rng = (exp_from, exp_to)

    def _shift(d):
        if d is None:
            return None
        try:
            return d + relativedelta(years=1)
        except Exception:
            return d

    next_rng = (_shift(exp_from), _shift(exp_to))
    return this_rng, next_rng


def _fmt_filter_values(vals) -> str:
    vals = vals or []
    if not vals:
        return "全部"
    vals = [str(v) for v in vals if str(v).strip()]
    if not vals:
        return "全部"
    return "、".join(vals)


def show_filter_ranges_if_enabled(ui_state: dict):
    """每個區塊標題下方顯示（若 sidebar checkbox 勾選）。"""
    if not st.session_state.get("show_filter_ranges", False):
        return
    this_rng, next_rng = _get_effective_ranges(ui_state)
    st.caption(
        f"篩選時間範圍｜本年度：{_fmt_d(this_rng[0])} ~ {_fmt_d(this_rng[1])}　｜　明年度：{_fmt_d(next_rng[0])} ~ {_fmt_d(next_rng[1])}"
    )
    st.caption(
        "目前篩選項目｜"
        f"展碁業務：{_fmt_filter_values(ui_state.get('sales'))}　｜　"
        f"經銷商：{_fmt_filter_values(ui_state.get('dealer'))}　｜　"
        f"最終客戶：{_fmt_filter_values(ui_state.get('customer'))}　｜　"
        f"資格：{_fmt_filter_values(ui_state.get('qual'))}　｜　"
        f"訂閱動作：{_fmt_filter_values(ui_state.get('action'))}　｜　"
        f"商品名稱：{_fmt_filter_values(ui_state.get('product'))}　｜　"
        f"未續約示警：{_fmt_filter_values(st.session_state.get('warning_filter_pick', []))}"
    )



# ---------------------------
# 12 下方區域 #2 分組明細表（大綱模式）
# Group By：最終客戶 > 訂閱到期日年度 > 訂閱動作 > 訂閱到期日
# Header Row：只顯示 最終客戶 + 訂閱到期日年度 + 訂閱總金額（其餘欄位 pd.NA）
# ✅ 依訂閱總金額由高至低排序（你指定）
# ---------------------------
def build_grouped_detail_report_v2(df: pd.DataFrame, df_next: pd.DataFrame | None = None) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    base_cols = [
        "最終客戶",
        EXPIRY_YEAR_COL,
        "訂閱動作",
        "訂閱到期日",
        "商品名稱",
        "數量",
        "成交單價未稅",
        ANALYSIS_VALUE_COL,
        "經銷商",
        "展碁業務",
    ]
    cols = [c for c in base_cols if c in df.columns]
    d = df[cols].copy()

    # 明細排序：最終客戶 > 訂閱到期日年度 > 訂閱動作 > 訂閱到期日（維持原行為）
    sort_cols = [c for c in ["最終客戶", EXPIRY_YEAR_COL, "訂閱動作", "訂閱到期日"] if c in d.columns]
    if sort_cols:
        d = d.sort_values(sort_cols, kind="mergesort")

    # 客戶+年度 總金額（使用成交價未稅小計）+ 最近到期日（用於未續約示警）
    key_cols = ["最終客戶", EXPIRY_YEAR_COL]
    key_cols = [c for c in key_cols if c in d.columns]
    totals = (
        d.groupby(key_cols, dropna=False)
        .agg(
            訂閱總金額=(ANALYSIS_VALUE_COL, lambda s: pd.to_numeric(s, errors="coerce").fillna(0).sum()),
            最近到期日=("訂閱到期日", "min"),
        )
        .reset_index()
    )

    # ✅ 群組排序：依訂閱總金額 DESC
    totals_sorted = totals.copy()
    totals_sorted["訂閱總金額"] = pd.to_numeric(totals_sorted["訂閱總金額"], errors="coerce").fillna(0)
    totals_sorted = totals_sorted.sort_values("訂閱總金額", ascending=False, kind="mergesort")

    renewal_lookup = build_group_renewal_lookup(df_next)

    out_rows: list[dict] = []

    # 依 totals_sorted 的順序輸出 Header + Detail
    for _, trow in totals_sorted.iterrows():
        cust = trow.get("最終客戶", pd.NA)
        yr = trow.get(EXPIRY_YEAR_COL, pd.NA)
        total = float(trow.get("訂閱總金額", 0) or 0)
        warning_text, warning_color, warning_threshold = get_group_warning_meta_with_renewal(
            min_expiry_dt=trow.get("最近到期日", pd.NaT),
            current_total=total,
            customer=cust,
            expiry_year=yr,
            renewal_lookup=renewal_lookup,
        )

        # 對應的明細列
        sub = d.copy()
        if "最終客戶" in sub.columns:
            sub = sub[sub["最終客戶"].astype(str) == str(cust)]
        if EXPIRY_YEAR_COL in sub.columns:
            # 盡量以 int 比對，避免 '2026.0' 之類差異
            try:
                yr_int = int(yr)
                sub = sub[pd.to_numeric(sub[EXPIRY_YEAR_COL], errors="coerce").fillna(-1).astype(int) == yr_int]
            except Exception:
                sub = sub[sub[EXPIRY_YEAR_COL].astype(str) == str(yr)]

        out_rows.append(
            {
                "_row_type": "HEADER",
                "未續約示警": warning_text,
                "_warning_color": warning_color,
                "_warning_threshold": warning_threshold,
                "最終客戶": cust,
                EXPIRY_YEAR_COL: yr,
                "訂閱總金額": total,
                "訂閱動作": pd.NA,
                "訂閱到期日": pd.NA,
                "商品名稱": pd.NA,
                "數量": pd.NA,
                "成交單價未稅": pd.NA,
                ANALYSIS_VALUE_COL: pd.NA,
                "經銷商": pd.NA,
                "展碁業務": pd.NA,
            }
        )

        for _, r in sub.iterrows():
            out_rows.append(
                {
                    "_row_type": "DETAIL",
                    "未續約示警": pd.NA,
                    "_warning_color": "",
                    "_warning_threshold": pd.NA,
                    "最終客戶": pd.NA,
                    EXPIRY_YEAR_COL: pd.NA,
                    "訂閱總金額": pd.NA,
                    "訂閱動作": r.get("訂閱動作", pd.NA),
                    "訂閱到期日": r.get("訂閱到期日", pd.NA),
                    "商品名稱": r.get("商品名稱", pd.NA),
                    "數量": r.get("數量", pd.NA),
                    "成交單價未稅": r.get("成交單價未稅", pd.NA),
                    ANALYSIS_VALUE_COL: r.get(ANALYSIS_VALUE_COL, pd.NA),
                    "經銷商": r.get("經銷商", pd.NA),
                    "展碁業務": r.get("展碁業務", pd.NA),
                }
            )

    out = pd.DataFrame(out_rows)

    ordered_cols = [
        "_row_type",
        "未續約示警",
        "_warning_color",
        "_warning_threshold",
        "最終客戶",
        EXPIRY_YEAR_COL,
        "訂閱總金額",
        "訂閱動作",
        "訂閱到期日",
        "商品名稱",
        "數量",
        "成交單價未稅",
        ANALYSIS_VALUE_COL,
        "經銷商",
        "展碁業務",
    ]
    ordered_cols = [c for c in ordered_cols if c in out.columns]
    out = out[ordered_cols]
    return out


# ---------------------------
# 推薦邏輯（原本存在，保持不動）
# ---------------------------
def _pick_recommendations(product: str, category: str):
    p_raw = (product or "").strip()
    c_raw = (category or "").strip()
    p = p_raw.lower()
    c = c_raw.lower()

    def has_any(*keys: str) -> bool:
        return any(k.lower() in p for k in keys) or any(k.lower() in c for k in keys)

    def has_in_product(*keys: str) -> bool:
        return any(k.lower() in p for k in keys)

    is_exchange_only = has_any("exchange online") and not has_any("microsoft 365", "m365", "office 365", "business", "e3", "e5")
    is_m365_family = has_any("microsoft 365", "m365", "office 365", "business", "e3", "e5", "teams", "sharepoint", "onedrive", "exchange")

    is_basic = has_any("business basic", "商務基本", "基本版")
    is_standard = has_any("business standard", "商務標準", "標準版")
    is_premium = has_any("business premium", "商務進階", "m365 business premium")
    is_e3 = has_any(" microsoft 365 e3", "microsoft 365 e3", " m365 e3", " e3")
    is_e5 = has_any(" microsoft 365 e5", "microsoft 365 e5", " m365 e5", " e5")

    already_copilot = has_any("copilot")
    already_purview = has_any("purview")
    already_security_suite = has_in_product("business security suite")
    already_purview_suite = has_in_product("business purview suite")

    already_defender = has_any("defender", "atp", "mdo")
    already_intune = has_any("intune", "endpoint manager")
    already_entra = has_any("entra", "azure ad", "aad", "p1", "p2")

    copilot_base_ok = (is_standard or is_premium or is_e3 or is_e5) and is_m365_family and (not is_exchange_only)
    copilot_needs_base_upgrade = (is_exchange_only or is_basic) and is_m365_family and (not copilot_base_ok)

    upsell: list[tuple[str, str]] = []
    cross: list[tuple[str, str]] = []

    if is_m365_family:
        if copilot_needs_base_upgrade:
            upsell.append(
                (
                    "Microsoft 365 Business Standard",
                    "先把基礎工作平台補齊（桌面版 Office + Teams/SharePoint/OneDrive 協作），後續才能把 AI、生產力與治理『一起做成流程』。",
                )
            )
            upsell.append(
                (
                    "Microsoft 365 Business Premium",
                    "若更在意『可控、可管、可追』：Premium 在 Standard 基礎上加入更完整的裝置合規與安全能力。",
                )
            )

        if is_standard and not is_premium:
            upsell.append(
                (
                    "Microsoft 365 Business Premium",
                    "把生產力升級成『安全生產力』：可用裝置合規＋端點防護＋條件式存取把風險變成可控。",
                )
            )

        if (is_premium or is_e3) and not is_e5:
            upsell.append(
                (
                    "Microsoft 365 E5（進階安全與治理）",
                    "當公司開始遇到釣魚信、帳號盜用、端點異常、稽核調閱等情境，E5 可把偵測/調查/回應串成同一條事件鏈。",
                )
            )

        if not already_copilot:
            if copilot_base_ok:
                upsell.append(
                    (
                        "Microsoft 365 Copilot Business",
                        "把日常工作變成可複製的 AI 作業線：會議紀要、郵件/文件快速產出與改寫、Excel 從資料到洞察。",
                    )
                )
            else:
                upsell.append(
                    (
                        "Microsoft 365 Copilot Business（需先具備合格 Base SKU）",
                        "要讓 Copilot 發揮價值需先升級到 Business Standard/Business Premium 以上（或 E3/E5）。",
                    )
                )
        else:
            upsell.append(
                (
                    "Microsoft 365 Copilot Business（擴大使用情境）",
                    "把 Copilot 從個人使用推進到部門標準流程：建立範本並用指標量化節省工時。",
                )
            )

    if is_m365_family:
        if not already_security_suite:
            cross.append(
                (
                    "Microsoft 365 Business Security Suite",
                    "用一套組合把身分、端點、信箱、雲端應用的基本防線一次補齊。",
                )
            )
        else:
            cross.append(
                (
                    "Microsoft 365 Business Security Suite（高風險角色加強包）",
                    "針對財會/主管/業務等高風險帳號加強條件式存取與釣魚防護。",
                )
            )

        if not (already_purview or already_purview_suite):
            cross.append(
                (
                    "Microsoft 365 Business Purview Suite",
                    "把敏感資料自動分類、套標籤並防外流，並留下稽核紀錄。",
                )
            )
        else:
            cross.append(
                (
                    "Microsoft 365 Business Purview Suite（稽核與留存制度化）",
                    "把留存/刪除/稽核制度化，遇到調閱可快速找到證據與版本。",
                )
            )

        if not already_entra:
            cross.append(
                (
                    "Entra ID P1（條件式存取）",
                    "把撞庫、異地登入、非受控裝置登入擋在門外。",
                )
            )
        if not already_defender:
            cross.append(
                (
                    "Microsoft Defender for Office 365",
                    "強化釣魚信、惡意連結與附件沙箱。",
                )
            )
        if not already_intune:
            cross.append(
                (
                    "Intune（端點/裝置管理）",
                    "把裝置合規、政策下發與 App 配置自動化。",
                )
            )

        cross.append(
            (
                "45 分鐘情境式 Demo（用真實流程示範）",
                "用實際工作流程做 Demo，讓加購理由更具體、核決更快。",
            )
        )
        cross.append(
            (
                "續約 + 加購一頁式提案（免費版）",
                "把現況、風險、效益、建議組合與預估預算濃縮成一頁。",
            ),
        )

    def dedup(items: list[tuple[str, str]]) -> list[tuple[str, str]]:
        seen = set()
        out = []
        for name, reason in items:
            if name not in seen:
                out.append((name, reason))
                seen.add(name)
        return out

    upsell = dedup(upsell)[:6]
    cross = dedup(cross)[:8]
    upsell = [(n, r) for (n, r) in upsell if "teams phone" not in n.lower()]
    cross = [(n, r) for (n, r) in cross if "teams phone" not in n.lower()]
    return {"upsell": upsell, "cross_sell": cross}


@st.cache_data(show_spinner=False)
def generate_email_cached(
    customer: str,
    reseller: str,
    domain: str,
    product: str,
    action: str,
    qty: float | int | None,
    expiry_str: str,
    unit_price: float | None,
    subtotal_price: float | None,
    category: str,
    sales: str,
) -> str:
    indent = " " * 4
    qty_str = "-" if qty is None or pd.isna(qty) else f"{int(qty)}"
    unit_price_str = "-" if unit_price is None or pd.isna(unit_price) else format_money(unit_price)
    subtotal_str = "-" if subtotal_price is None or pd.isna(subtotal_price) else format_money(subtotal_price)

    recs = _pick_recommendations(product, category)
    upsell = recs.get("upsell", [])
    cross_sell = recs.get("cross_sell", [])

    def _format_items(items: list[tuple[str, str]]) -> str:
        lines = []
        for i, (name, reason) in enumerate(items, 1):
            reason = " ".join((reason or "").strip().split())
            lines.append(f"{indent}{i}. {name}\n{indent}   - 理由：{reason}")
        return "\n\n".join(lines) if lines else f"{indent}（可依客戶現況提供客製化建議）"

    upsell_block = _format_items(upsell)
    cross_block = _format_items(cross_sell)

    return (
        f"{indent}主旨：{customer} 續約提醒與加購建議（到期日：{expiry_str}）\n\n"
        f"{indent}{reseller} 您好，\n\n"
        f"{indent}提醒您以下訂閱即將到期，建議提前安排續約以避免服務中斷：\n"
        f"{indent}- 最終客戶：{customer}\n"
        f"{indent}- 客戶網域：{domain}\n"
        f"{indent}- 商品名稱：{product}\n"
        f"{indent}- 訂閱動作：{action}\n"
        f"{indent}- 數量：{qty_str}\n"
        f"{indent}- 成交單價未稅：{unit_price_str}\n"
        f"{indent}- 成交價未稅小計：{subtotal_str}\n"
        f"{indent}- 訂閱到期日：{expiry_str}\n\n"
        f"{indent}想請您協助確認：\n"
        f"{indent}1) 是否同意續約此訂閱？\n"
        f"{indent}2) 是否需要調整數量（增購/減量）？若需要，請回覆調整後數量與預計生效日。\n\n"
        f"{indent}建議 Upsell（升級/進階）：\n\n"
        f"{upsell_block}\n\n"
        f"{indent}建議 Cross-sell（加購/搭配）：\n\n"
        f"{cross_block}\n\n"
        f"{indent}若您願意，我可協助把「續約 + 加購」整理成一頁式提案（現況、風險、效益、建議組合與預估預算），方便主管快速核決。\n\n"
        f"{indent}—\n"
        f"{indent}（展碁業務）{sales}\n"
    )


def build_email_from_row(row: pd.Series) -> str:
    customer = str(row.get("最終客戶", "") or "")
    reseller = str(row.get("經銷商", "") or "")
    domain = str(row.get("客戶微軟網域", "") or "")
    product = str(row.get("商品名稱", "") or "")
    action = str(row.get("訂閱動作", "") or "")
    qty = row.get("數量", None)
    expiry = row.get("訂閱到期日", None)

    unit_price = row.get("成交單價未稅", None)
    subtotal = row.get("成交價未稅小計", None)

    category = str(row.get("產品分類", "") or "")
    sales = str(row.get("展碁業務", "") or "")
    expiry_str = expiry.strftime("%Y-%m-%d") if pd.notna(expiry) else ""

    return generate_email_cached(
        customer=customer,
        reseller=reseller,
        domain=domain,
        product=product,
        action=action,
        qty=qty,
        expiry_str=expiry_str,
        unit_price=unit_price,
        subtotal_price=subtotal,
        category=category,
        sales=sales,
    )


def build_group_email_from_header(df_filtered_src: pd.DataFrame, customer: str, expiry_year) -> tuple[str, str, str]:
    """依據（最終客戶 + 到期年度）彙整 header 下所有 detail 商品，回傳 (email_text含主旨, subject, mailto_link)。"""
    if df_filtered_src is None or df_filtered_src.empty:
        return "", "", ""

    d = df_filtered_src.copy()
    d = d[d["最終客戶"].astype(str) == str(customer)]
    try:
        yr_int = int(expiry_year)
        d = d[pd.to_numeric(d[EXPIRY_YEAR_COL], errors="coerce").fillna(-1).astype(int) == yr_int]
        yr_str = str(yr_int)
    except Exception:
        yr_str = str(expiry_year)
        d = d[d[EXPIRY_YEAR_COL].astype(str) == yr_str]

    if d.empty:
        return "", "", ""

    reseller = str(d["經銷商"].dropna().astype(str).iloc[0]) if "經銷商" in d.columns and d["經銷商"].notna().any() else ""
    domain = str(d["客戶微軟網域"].dropna().astype(str).iloc[0]) if "客戶微軟網域" in d.columns and d["客戶微軟網域"].notna().any() else ""
    sales = str(d["展碁業務"].dropna().astype(str).iloc[0]) if "展碁業務" in d.columns and d["展碁業務"].notna().any() else ""

    total_amount = float(pd.to_numeric(d[ANALYSIS_VALUE_COL], errors="coerce").fillna(0).sum())

    agg = d.copy()
    agg["數量"] = pd.to_numeric(agg.get("數量"), errors="coerce").fillna(0)
    agg[ANALYSIS_VALUE_COL] = pd.to_numeric(agg.get(ANALYSIS_VALUE_COL), errors="coerce").fillna(0)
    agg["訂閱到期日"] = pd.to_datetime(agg.get("訂閱到期日"), errors="coerce")

    # ✅ 新增：依商品名稱彙總（數量彙總）
    qg = (
        agg.groupby(["商品名稱"], dropna=False)
        .agg(數量=("數量", "sum"))
        .reset_index()
    )
    qg["商品名稱"] = qg["商品名稱"].astype(str).replace({"nan": "未填"})
    qg["數量"] = pd.to_numeric(qg["數量"], errors="coerce").fillna(0)
    qg = qg.sort_values(["數量", "商品名稱"], ascending=[False, True], kind="mergesort")

    indent = " " * 4
    qty_lines_list = []
    for _, r in qg.iterrows():
        prod = str(r.get("商品名稱", "") or "")
        qty = int(round(float(r.get("數量", 0) or 0)))
        qty_lines_list.append(f"{indent}- {prod} × {qty}")
    product_qty_lines = "\n".join(qty_lines_list) if qty_lines_list else f"{indent}（無可彙整品項）"

    # 既有：依「到期日 + 商品」彙總（含金額）
    g = (
        agg.groupby(["訂閱到期日", "商品名稱", "產品分類"], dropna=False)
        .agg(
            數量=("數量", "sum"),
            金額=(ANALYSIS_VALUE_COL, "sum"),
        )
        .reset_index()
    )
    g["訂閱到期日"] = pd.to_datetime(g["訂閱到期日"], errors="coerce")
    g["金額"] = pd.to_numeric(g["金額"], errors="coerce").fillna(0)
    g = g.sort_values(["訂閱到期日", "金額"], ascending=[True, False], kind="mergesort")

    lines = []
    current_date = None
    for _, r in g.iterrows():
        expiry = r.get("訂閱到期日", pd.NaT)
        expiry_str = expiry.strftime("%Y-%m-%d") if pd.notna(expiry) else "未填"

        if expiry_str != current_date:
            lines.append(f"{indent}【到期日：{expiry_str}】")
            current_date = expiry_str

        prod = str(r.get("商品名稱", "") or "")
        qty = int(round(float(r.get("數量", 0) or 0)))
        amt = float(r.get("金額", 0) or 0)

        lines.append(f"{indent}- {prod} × {qty}；小計：{format_money(amt)}")

    items_lines = "\n".join(lines) if lines else f"{indent}（無可彙整品項）"

    upsell_all: list[tuple[str, str]] = []
    cross_all: list[tuple[str, str]] = []
    for _, rr in g.iterrows():
        prod = str(rr.get("商品名稱", "") or "")
        cat = str(rr.get("產品分類", "") or "")
        recs = _pick_recommendations(prod, cat)
        upsell_all += recs.get("upsell", [])
        cross_all += recs.get("cross_sell", [])

    def _dedup(items: list[tuple[str, str]], limit: int):
        seen = set()
        out = []
        for name, reason in items:
            if name not in seen:
                out.append((name, reason))
                seen.add(name)
            if len(out) >= limit:
                break
        return tuple(out)

    rec_upsell = _dedup(upsell_all, 6)
    rec_cross = _dedup(cross_all, 8)

    subject, body = generate_group_email_cached(
        customer=str(customer),
        reseller=reseller,
        domain=domain,
        expiry_year=yr_str,
        total_amount=total_amount,
        product_qty_lines=product_qty_lines,   # ✅ 新增
        items_lines=items_lines,
        sales=sales,
        rec_upsell=rec_upsell,
        rec_cross=rec_cross,
    )

    email_text = (" " * 4) + f"主旨：{subject}\n\n{body}"
    mailto_link = build_mailto_link(subject, body)
    return email_text, subject, mailto_link


def build_group_call_script_from_header(df_filtered_src: pd.DataFrame, customer: str, expiry_year) -> str:
    """
    依據（最終客戶 + 到期年度）彙整 header 下所有 detail 商品，產生「電話訪談」口語化腳本。
    - 目的：提供業務人員電話/線上訪談時可直接照著講（親切且專業，含情境與案例）
    """
    if df_filtered_src is None or df_filtered_src.empty:
        return ""

    d = df_filtered_src.copy()
    d = d[d["最終客戶"].astype(str) == str(customer)]
    try:
        yr_int = int(expiry_year)
        d = d[pd.to_numeric(d[EXPIRY_YEAR_COL], errors="coerce").fillna(-1).astype(int) == yr_int]
        yr_str = str(yr_int)
    except Exception:
        yr_str = str(expiry_year)
        d = d[d[EXPIRY_YEAR_COL].astype(str) == yr_str]

    if d.empty:
        return ""

    reseller = str(d["經銷商"].dropna().astype(str).iloc[0]) if "經銷商" in d.columns and d["經銷商"].notna().any() else ""
    domain = str(d["客戶微軟網域"].dropna().astype(str).iloc[0]) if "客戶微軟網域" in d.columns and d["客戶微軟網域"].notna().any() else ""
    sales = str(d["展碁業務"].dropna().astype(str).iloc[0]) if "展碁業務" in d.columns and d["展碁業務"].notna().any() else ""

    total_amount = float(pd.to_numeric(d[ANALYSIS_VALUE_COL], errors="coerce").fillna(0).sum())

    agg = d.copy()
    agg["數量"] = pd.to_numeric(agg.get("數量"), errors="coerce").fillna(0)
    agg[ANALYSIS_VALUE_COL] = pd.to_numeric(agg.get(ANALYSIS_VALUE_COL), errors="coerce").fillna(0)
    agg["訂閱到期日"] = pd.to_datetime(agg.get("訂閱到期日"), errors="coerce")

    # ✅ 新增：三、到期品項快速對齊下方「依商品名稱彙總數量」
    qg = (
        agg.groupby(["商品名稱"], dropna=False)
        .agg(數量=("數量", "sum"))
        .reset_index()
    )
    qg["商品名稱"] = qg["商品名稱"].astype(str).replace({"nan": "未填"})
    qg["數量"] = pd.to_numeric(qg["數量"], errors="coerce").fillna(0)
    qg = qg.sort_values(["數量", "商品名稱"], ascending=[False, True], kind="mergesort")

    indent = " " * 4
    qty_lines_list = []
    for _, r in qg.iterrows():
        prod = str(r.get("商品名稱", "") or "")
        qty = int(round(float(r.get("數量", 0) or 0)))
        qty_lines_list.append(f"{indent}- {prod} × {qty}")
    product_qty_lines = "\n".join(qty_lines_list) if qty_lines_list else f"{indent}（無可彙整品項）"

    g = (
        agg.groupby(["訂閱到期日", "商品名稱", "產品分類"], dropna=False)
        .agg(
            數量=("數量", "sum"),
            金額=(ANALYSIS_VALUE_COL, "sum"),
        )
        .reset_index()
    )
    g["訂閱到期日"] = pd.to_datetime(g["訂閱到期日"], errors="coerce")
    g["金額"] = pd.to_numeric(g["金額"], errors="coerce").fillna(0)
    g = g.sort_values(["訂閱到期日", "金額"], ascending=[True, False], kind="mergesort")

    # 彙整到期區間（抓最早/最晚；若只有單一日期就顯示單一日期）
    overall_min = agg["訂閱到期日"].min()
    overall_max = agg["訂閱到期日"].max()
    overall_min_s = overall_min.strftime("%Y-%m-%d") if pd.notna(overall_min) else ""
    overall_max_s = overall_max.strftime("%Y-%m-%d") if pd.notna(overall_max) else ""
    overall_range = overall_min_s if (overall_min_s == overall_max_s or not overall_max_s) else f"{overall_min_s} ~ {overall_max_s}"

    # 產品清單（口語化用）：同一天多商品排一起；到期日由近到遠（格式/行距同 Email 視窗）
    items_lines = []
    current_date = None

    for _, r in g.iterrows():
        expiry = r.get("訂閱到期日", pd.NaT)
        expiry_str = expiry.strftime("%Y-%m-%d") if pd.notna(expiry) else "未填"

        if expiry_str != current_date:
            items_lines.append(f"{indent}【到期日：{expiry_str}】")
            current_date = expiry_str

        prod = str(r.get("商品名稱", "") or "")
        qty = int(round(float(r.get("數量", 0) or 0)))
        amt = float(r.get("金額", 0) or 0)

        items_lines.append(f"{indent}- {prod} × {qty}；小計：{format_money(amt)}")

    items_block = "\n".join(items_lines) if items_lines else f"{indent}- （無可彙整品項）"

    # 建議（沿用既有推薦邏輯）
    upsell_all: list[tuple[str, str]] = []
    cross_all: list[tuple[str, str]] = []
    for _, rr in g.iterrows():
        prod = str(rr.get("商品名稱", "") or "")
        cat = str(rr.get("產品分類", "") or "")
        recs = _pick_recommendations(prod, cat)
        upsell_all += recs.get("upsell", [])
        cross_all += recs.get("cross_sell", [])

    def _dedup(items: list[tuple[str, str]], limit: int):
        seen = set()
        out = []
        for name, reason in items:
            if name not in seen:
                out.append((name, reason))
                seen.add(name)
            if len(out) >= limit:
                break
        return out

    rec_upsell = _dedup(upsell_all, 4)
    rec_cross = _dedup(cross_all, 5)

    def _talk_points(items: list[tuple[str, str]]) -> str:
        indent = " " * 4
        if not items:
            return f"{indent}（可依客戶現況客製化）"
        lines = []
        for name, reason in items:
            reason = " ".join((reason or "").strip().split())
            lines.append(f"{indent}- {name}：{reason}")
        return "\n".join(lines)

    upsell_talk = _talk_points(rec_upsell)
    cross_talk = _talk_points(rec_cross)

    # 口語化腳本（可直接照念）
    script = (
        f"{indent}【電話訪談腳本（建議 8~12 分鐘）】\n"
        f"{indent}對象：{reseller or '（經銷商）'} / 最終客戶：{customer}（網域：{domain}）\n"
        f"{indent}到期年度：{yr_str}；到期區間：{overall_range}\n"
        f"{indent}彙總金額（{ANALYSIS_VALUE_COL}）：{format_money(total_amount)}\n\n"
        f"{indent}一、開場（30 秒）\n"
        f"{indent}您好，我是（展碁業務）{sales}。想跟您簡單確認一下 {customer} 今年度到期的訂閱，避免到期後服務中斷，也順便看有沒有需要調整數量或補齊安全/治理。\n\n"
        f"{indent}二、先確認現況（1~2 分鐘）\n"
        f"{indent}1) 這些帳號目前主要用在什麼部門/用途？（辦公協作、郵件、檔案、外勤、產線…）\n"
        f"{indent}2) 近 3 個月有沒有遇到：信件被釣魚、帳號被盜、外部分享檔案失控、裝置遺失、離職交接資料混亂？\n"
        f"{indent}3) 這次續約您希望：維持現狀 / 增購 / 減量 / 升級？（我可以幫您把預算抓一個範圍）\n\n"
        f"{indent}三、到期品項快速對齊（1 分鐘）\n"
        f"{indent}我這邊先把到期品項彙整給您，您聽聽看有沒有需要調整：\n"
        f"{indent}（依商品名稱彙總數量）\n"
        f"{product_qty_lines}\n\n"
        f"{items_block}\n\n"
        f"{indent}四、情境式說明（用案例講價值，2~4 分鐘）\n"
        f"{indent}情境 A：『人員異動/離職交接』\n"
        f"{indent}- 常見痛點：交接檔案散在個人 OneDrive/郵件，找不到版本或缺證據。\n"
        f"{indent}- 我們常見做法：用留存/稽核把關鍵資料留下來，交接就不會靠人工找。\n\n"
        f"{indent}情境 B：『釣魚信/帳號盜用』\n"
        f"{indent}- 常見痛點：只要一個主管帳號被盜，外部轉帳或資料外洩風險很高。\n"
        f"{indent}- 我們常見做法：條件式存取 + 端點合規，把『不安全的登入』直接擋掉。\n\n"
        f"{indent}情境 C：『外部分享與資料外流』\n"
        f"{indent}- 常見痛點：檔案分享出去後無法控管轉傳與下載。\n"
        f"{indent}- 我們常見做法：自動分類/標籤 + DLP 防外流，並保留稽核紀錄。\n\n"
        f"{indent}五、自然帶到升級/加購建議（1~2 分鐘）\n"
        f"{indent}如果您覺得上述情境有中到，我會建議您優先看這幾個組合（我可以依您現況再縮小範圍）：\n\n"
        f"{indent}【Upsell（升級/進階）】\n{upsell_talk}\n\n"
        f"{indent}【Cross-sell（加購/搭配）】\n{cross_talk}\n\n"
        f"{indent}六、收斂與下一步（30~60 秒）\n"
        f"{indent}我先跟您確認兩件事：\n"
        f"{indent}1) 續約：這批訂閱是『全數維持』還是要『調整數量』？\n"
        f"{indent}2) 若要補強：您比較在意『安全防護』還是『資料治理/稽核』？我可以幫您整理一頁式提案（現況、風險、效益、建議組合與預估預算），讓您內部核決更快。\n\n"
        f"{indent}【備註】\n"
        f"{indent}- 這份腳本可直接照念；若您告訴我客戶產業/人數/痛點，我也能再把案例換成更貼近的版本。\n"
    )
    return script


def _html_escape(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _mailto_quote(s: str) -> str:
    # mailto URL 需要 percent-encoding；保留換行（%0A）以利貼上閱讀
    return urllib.parse.quote((s or ""), safe="")


def build_mailto_link(subject: str, body: str) -> str:
    subj = _mailto_quote(subject)
    bdy = _mailto_quote(body)
    # recipient 由操作者在 Outlook 中填寫，因此留空
    return f"mailto:?subject={subj}&body={bdy}"


@st.cache_data(show_spinner=False)
def generate_group_email_cached(
    customer: str,
    reseller: str,
    domain: str,
    expiry_year: str,
    total_amount: float,
    product_qty_lines: str,  # ✅ 新增：依商品名稱彙總（數量）
    items_lines: str,
    sales: str,
    rec_upsell: tuple[tuple[str, str], ...],
    rec_cross: tuple[tuple[str, str], ...],
) -> tuple[str, str]:
    """回傳 (subject, body)。使用 cache_data 避免重複生成。"""
    indent = " " * 4

    def _format_items(items: tuple[tuple[str, str], ...]) -> str:
        if not items:
            return f"{indent}（可依客戶現況提供客製化建議）"
        lines = []
        for i, (name, reason) in enumerate(items, 1):
            reason = " ".join((reason or "").strip().split())
            lines.append(f"{indent}{i}. {name}\n{indent}   - 理由：{reason}")
        return "\n\n".join(lines)

    upsell_block = _format_items(rec_upsell)
    cross_block = _format_items(rec_cross)

    subject = f"{customer} 續約提醒與加購建議（到期年度：{expiry_year}）"
    body = (
        f"{indent}{reseller} 您好，\n\n"
        f"{indent}以下為 {customer} 於 {expiry_year} 年度到期之訂閱彙整，建議提前安排續約以避免服務中斷：\n\n"
        f"{indent}- 最終客戶：{customer}\n"
        f"{indent}- 客戶網域：{domain}\n"
        f"{indent}- 到期年度：{expiry_year}\n"
        # ✅ 需求 2：彙總金額 與 到期品項（依商品名稱彙總數量）之間空一行
        f"{indent}- 彙總金額（{ANALYSIS_VALUE_COL}）：{format_money(total_amount)}\n\n"
        f"{indent}到期品項（依商品名稱彙總數量）：\n"
        f"{product_qty_lines}\n\n"
        f"{indent}到期品項（彙總）：\n"
        f"{items_lines}\n\n"
        f"{indent}想請您協助確認：\n"
        f"{indent}1) 是否同意續約上述訂閱？\n"
        f"{indent}2) 是否需要調整數量（增購/減量）？若需要，請回覆調整後數量與預計生效日。\n\n"
        f"{indent}建議 Upsell（升級/進階）：\n\n"
        f"{upsell_block}\n\n"
        f"{indent}建議 Cross-sell（加購/搭配）：\n\n"
        f"{cross_block}\n\n"
        f"{indent}若您願意，我可協助把「續約 + 加購」整理成一頁式提案（現況、風險、效益、建議組合與預估預算），方便主管快速核決。\n\n"
        f"{indent}—\n"
        f"{indent}（展碁業務）{sales}\n"
    )
    return subject, body


# ---------------------------
# 明細表格共用：欄位順序（保持一致）
# ---------------------------
DETAIL_DESIRED_COLS = [
    "_row_key",
    "經銷商",
    "資格",
    "最終客戶",
    "訂閱動作",
    "產品分類",
    "商品名稱",
    "數量",
    "訂閱到期日",
    EXPIRY_YEAR_COL,
    "成交單價未稅",
    "成交價未稅小計",
    "展碁COST單價未稅",
    "展碁COST未稅小計",
    ANNUAL_COL,
    "訂單狀態",
    "訂單下單日",
    "客戶微軟網域",
    "展碁業務",
]


def render_detail_table(df_source: pd.DataFrame, editor_key: str, selectable: bool) -> tuple[pd.DataFrame | None, int | None]:
    """
    回傳 (edited_df, selected_row_key)
    - selectable=True：含「選取」checkbox（用於今年度，維持既有 overlay 行為）
    - selectable=False：仍顯示同欄位/格示，但 checkbox 不提供互動（避免影響既有功能）
    """
    if df_source is None or df_source.empty:
        st.info("無資料可顯示（此區塊視為 0）")
        return None, None

    df_view_local = df_source.reset_index(drop=False).rename(columns={"index": "_row_key"})
    show_df_local = df_view_local.copy()

    # 維持與既有一致：第一欄為選取
    show_df_local.insert(0, "選取", False)

    ordered = ["選取"] + [c for c in DETAIL_DESIRED_COLS if c in show_df_local.columns]
    ordered += [c for c in show_df_local.columns if c not in set(ordered)]
    show_df_local = show_df_local[ordered]

    if selectable:
        disabled_cols = [c for c in show_df_local.columns if c != "選取"]
        column_config = {"選取": st.column_config.CheckboxColumn(help="勾選一筆以顯示 Email", default=False)}
    else:
        # 仍用 data_editor 保持格示一致，但全部 disabled（包含選取）
        disabled_cols = list(show_df_local.columns)
        column_config = {"選取": st.column_config.CheckboxColumn(help="（此表不提供勾選）", default=False)}

    edited_local = st.data_editor(
        show_df_local,
        use_container_width=True,
        hide_index=True,
        disabled=disabled_cols,
        column_config=column_config,
        key=editor_key,
    )

    selected_row_key_local = None
    if selectable and edited_local is not None:
        try:
            sel = edited_local.loc[edited_local["選取"] == True, "_row_key"].tolist()
            if sel:
                selected_row_key_local = sel[0]
        except Exception:
            selected_row_key_local = None

    return edited_local, selected_row_key_local


# ---------------------------
# 分組明細表共用：渲染（保持一致）
# ---------------------------
def render_grouped_table(df_source: pd.DataFrame, df_next: pd.DataFrame | None = None):
    grouped_df_local = build_grouped_detail_report_v2(df_source, df_next=df_next)

    if grouped_df_local.empty:
        st.info("無資料可顯示（此區塊視為 0）")
        return

    def _row_style(row: pd.Series):
        if str(row.get("_row_type", "")) == "HEADER":
            return ["font-weight:700; background-color: rgba(0,0,0,0.04);"] * len(row)
        return [""] * len(row)

    def _warning_color_style(df_in: pd.DataFrame) -> pd.DataFrame:
        styles = pd.DataFrame("", index=df_in.index, columns=df_in.columns)
        if "未續約示警" not in df_in.columns:
            return styles

        for idx, row in df_in.iterrows():
            raw_text_val = row.get("未續約示警", "")
            raw_color_val = row.get("_warning_color", "")

            text_val = "" if pd.isna(raw_text_val) else str(raw_text_val).strip()
            color_val = "" if pd.isna(raw_color_val) else str(raw_color_val).strip()

            if text_val and color_val:
                styles.at[idx, "未續約示警"] = f"color:{color_val}; font-weight:700;"
        return styles

    num_cols = [c for c in ["訂閱總金額", "數量", "成交單價未稅", ANALYSIS_VALUE_COL] if c in grouped_df_local.columns]
    fmt_map = {c: _safe_thousands_formatter for c in num_cols}
    if "訂閱到期日" in grouped_df_local.columns:
        fmt_map["訂閱到期日"] = _safe_date_formatter

    styled_grouped = grouped_df_local.style.apply(_row_style, axis=1).apply(_warning_color_style, axis=None).format(fmt_map)

    hide_cols = [c for c in ["_row_type", "_warning_color", "_warning_threshold"] if c in grouped_df_local.columns]

    st.dataframe(
        styled_grouped.hide(axis="columns", subset=hide_cols),
        use_container_width=True,
        hide_index=True,
    )


# ---------------------------
# 側邊欄
# ---------------------------
st.sidebar.markdown("""
<div style="text-align:center; padding: 0.8rem 0 1rem 0; border-bottom: 1px solid rgba(255,255,255,0.2); margin-bottom: 0.8rem;">
    <div style="font-size: 1.5rem; margin-bottom: 0.3rem;">⚙️</div>
    <div style="color: white; font-size: 0.95rem; font-weight: 800; letter-spacing: 1.5px; text-transform: uppercase;">篩選與匯入</div>
    <div style="color: rgba(255,255,255,0.55); font-size: 0.7rem; margin-top: 0.2rem; letter-spacing: 0.5px;">Filter &amp; Import</div>
</div>
""", unsafe_allow_html=True)

# ✅ 新增：顯示篩選時間範圍（放在左邊篩選與匯入最上方）
st.session_state.setdefault("show_filter_ranges", False)
st.sidebar.checkbox(
    "顯示各區塊篩選時間範圍（今年度/明年度）",
    value=st.session_state.get("show_filter_ranges", False),
    key="show_filter_ranges",
)

uploaded = st.sidebar.file_uploader(
    "1) 上傳 Excel（CSP訂單資料_raw.xlsx）",
    type=["xlsx"],
    help="上傳後將自動清洗：刪除指定訂單狀態、保留欄位、數字轉型、新增年小計欄位。",
)

col_sb1, col_sb2 = st.sidebar.columns([1, 1])
with col_sb1:
    if st.button("🧹 清除所有篩選（恢復全部）", use_container_width=True):
        reset_filters_to_defaults()
        st.rerun()
with col_sb2:
    st.caption("（清除後回到本年度預設篩選）")

st.sidebar.divider()

future_expiry_enabled = st.sidebar.checkbox(
    "6) 未來 1~12 月內訂閱到期",
    value=st.session_state.get("future_expiry_enabled", False),
    key="future_expiry_enabled",
)
default_months = st.session_state.get("future_expiry_months", 3)
future_expiry_months = st.sidebar.selectbox(
    "下拉 1~12（月）",
    options=list(range(1, 13)),
    index=int(default_months) - 1,
    key="future_expiry_months",
    disabled=not future_expiry_enabled,
)

st.sidebar.divider()

# 先佔位，待資料讀取完成後依資料內容建立真正預設值
expiry_from = st.session_state.get("expiry_from", None)
expiry_to = st.session_state.get("expiry_to", None)
order_from = st.session_state.get("order_from", None)
order_to = st.session_state.get("order_to", None)

if st.query_params.get("overlay_close") == "1":
    handle_close_overlay_request()

# ---------------------------
# 匯入 + 清洗（含進度）
# ---------------------------
st.markdown(_GLOBAL_CSS, unsafe_allow_html=True)

# 頂部品牌 Header
st.markdown("""
<div style="
    background: linear-gradient(135deg, #1a237e 0%, #1565c0 60%, #0288d1 100%);
    border-radius: 16px;
    padding: 1.5rem 2rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 24px rgba(21,101,192,0.25);
    display: flex;
    align-items: center;
    gap: 1rem;
">
    <div style="font-size: 2.5rem; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.3));">📊</div>
    <div>
        <div style="color: white; font-size: 1.5rem; font-weight: 900; letter-spacing: -0.5px; line-height: 1.2; font-family: 'Noto Sans TC', sans-serif;">
            Weblink M365 續約精準行銷
        </div>
        <div style="color: rgba(255,255,255,0.75); font-size: 0.85rem; font-weight: 400; margin-top: 0.2rem; letter-spacing: 0.5px;">
            Enterprise Subscription Dashboard &nbsp;·&nbsp; CSP 訂單分析平台
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

status = st.status("等待匯入 Excel…", expanded=True)
progress = st.progress(0)

try:
    progress.progress(10)
    status.update(label="讀取 Excel…", state="running", expanded=True)

    if uploaded is not None:
        df_raw = load_excel_cached(file_bytes=uploaded.getvalue(), use_upload=True, default_path=DEFAULT_XLSX_PATH)
    else:
        df_raw = load_excel_cached(file_bytes=None, use_upload=False, default_path=DEFAULT_XLSX_PATH)

    progress.progress(40)
    status.update(label="清洗資料：刪除訂單狀態 / 保留欄位 / 轉型 / 新增年小計…", state="running", expanded=True)
    df = clean_transform_cached(df_raw)

    # ✅ 依匯入資料初始化日期預設值
    init_filter_defaults_from_data(df)

    expiry_from = st.sidebar.date_input("7) 訂閱到期日（起）", value=st.session_state.get("expiry_from"), key="expiry_from")
    expiry_to = st.sidebar.date_input("7) 訂閱到期日（迄）", value=st.session_state.get("expiry_to"), key="expiry_to")
    order_from = st.sidebar.date_input("7) 訂單下單日（起）", value=st.session_state.get("order_from"), key="order_from")
    order_to = st.sidebar.date_input("7) 訂單下單日（迄）", value=st.session_state.get("order_to"), key="order_to")

    progress.progress(70)
    status.update(label="建立篩選選項…", state="running", expanded=True)

    import re as _re

    # ══════════════════════════════════════════════════════════════════════════
    # 工具函式
    # ══════════════════════════════════════════════════════════════════════════
    def _make_num_opts(raw_list):
        return [f"{i+1:02d}. {v}" for i, v in enumerate(raw_list)]

    def _strip_num(labeled_list):
        return [_re.sub(r"^\d+\.\s*", "", s) for s in labeled_list]

    def _to_labeled(raw_sel, raw_opts, num_opts):
        """原始值清單 → 帶編號標籤清單（供 multiselect default= 使用）"""
        lmap = {r: n for r, n in zip(raw_opts, num_opts)}
        return [lmap[r] for r in (raw_sel or []) if r in lmap]

    def _safe_list(v):
        if v is None:
            return []
        if isinstance(v, list):
            return v
        return [v]

    def _sync_valid_selection(raw_key: str, labeled_key: str, raw_opts: list, num_opts: list):
        """將 session_state 既有選擇修正為目前 options 內仍合法的值。"""
        raw_selected = _safe_list(st.session_state.get(raw_key, []))
        valid_raw = [x for x in raw_selected if x in raw_opts]
        st.session_state[raw_key] = valid_raw

        lmap = {r: n for r, n in zip(raw_opts, num_opts)}
        st.session_state[labeled_key] = [lmap[x] for x in valid_raw if x in lmap]

    def _sync_valid_labeled(raw_key: str, labeled_key: str, raw_opts: list, num_opts: list):
        """保留目前已選的 raw 值並重建新標籤，避免時間範圍/未來月份切換後因編號改變而被清空。"""
        labeled_selected = _safe_list(st.session_state.get(labeled_key, []))
        raw_from_labeled = _strip_num(labeled_selected)
        raw_selected = _safe_list(st.session_state.get(raw_key, []))

        merged_raw = []
        for v in raw_from_labeled + raw_selected:
            vv = str(v).strip()
            if vv and vv not in merged_raw:
                merged_raw.append(vv)

        valid_raw = [x for x in merged_raw if x in raw_opts]
        lmap = {r: n for r, n in zip(raw_opts, num_opts)}
        st.session_state[raw_key] = valid_raw
        st.session_state[labeled_key] = [lmap[x] for x in valid_raw if x in lmap]

    def _reset_top30_if_invalid(pick_key: str, valid_raw_opts: list):
        """若 Top30 目前選到的值已不在新的 linked options 內，則自動重置。"""
        picked_label = st.session_state.get(pick_key, "（不套用 Top 30）")
        if picked_label == "（不套用 Top 30）":
            return

        picked_raw = _re.sub(r"^\d+｜", "", str(picked_label)).strip()
        if picked_raw not in valid_raw_opts:
            st.session_state[pick_key] = "（不套用 Top 30）"

    def _set_single_pick_to_multiselect(raw_key: str, labeled_key: str, picked_raw, raw_opts: list, num_opts: list):
        """將 Top30 selectbox 的單選結果同步到 multiselect state。"""
        if not picked_raw:
            st.session_state[raw_key] = []
            st.session_state[labeled_key] = []
            return

        lmap = {r: n for r, n in zip(raw_opts, num_opts)}
        st.session_state[raw_key] = [picked_raw]
        st.session_state[labeled_key] = [lmap[picked_raw]] if picked_raw in lmap else []

    # ──────────────────────────────────────────────────────────────────────────
    # 篩選選項邏輯
    # 1) 多條件篩選清單：僅依「時間範圍」列出選項，不互相連動
    # 2) Top 30 經銷商 / 最終客戶：僅依「時間範圍 + 展碁業務」排序變動
    # ──────────────────────────────────────────────────────────────────────────
    def _time_scope_state() -> dict:
        return {
            "future_expiry_enabled": bool(st.session_state.get("future_expiry_enabled", False)),
            "future_expiry_months": int(st.session_state.get("future_expiry_months", 3)),
            "expiry_range": (st.session_state.get("expiry_from"), st.session_state.get("expiry_to")),
            "order_range": (st.session_state.get("order_from"), st.session_state.get("order_to")),
            "dealer": [],
            "qual": [],
            "customer": [],
            "action": [],
            "product": [],
            "sales": [],
        }

    def _time_and_sales_scope_state() -> dict:
        s = _time_scope_state()
        s["sales"] = st.session_state.get("sales", []) or []
        return s

    def _options_by_scope(col: str, ui_scope: dict) -> list[str]:
        if df is None or df.empty or col not in df.columns:
            return []
        df_scope = apply_filters(df, ui_scope)
        if df_scope is None or df_scope.empty or col not in df_scope.columns:
            return []
        g = df_scope.groupby(col, dropna=False)[ANALYSIS_VALUE_COL].sum().reset_index()
        g[col] = g[col].astype(str).replace({"nan": "未填"})
        g[ANALYSIS_VALUE_COL] = pd.to_numeric(g[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)
        return g.sort_values(ANALYSIS_VALUE_COL, ascending=False, kind="mergesort")[col].tolist()

    def _base_opts(col: str) -> list[str]:
        return _options_by_scope(col, _time_scope_state())

    def _top30_opts(col: str) -> list[str]:
        return _options_by_scope(col, _time_and_sales_scope_state())

    # ══════════════════════════════════════════════════════════════════════════
    # Step 1：所有多條件篩選選項只依時間範圍產生，不互相連動
    # ══════════════════════════════════════════════════════════════════════════
    sales_opts = _merge_selected_into_options(_base_opts("展碁業務"), st.session_state.get("sales", []))
    dealer_opts = _merge_selected_into_options(_base_opts("經銷商"), st.session_state.get("dealer", []))
    qual_opts = _merge_selected_into_options(_base_opts("資格"), st.session_state.get("qual", []))
    customer_opts = _merge_selected_into_options(_base_opts("最終客戶"), st.session_state.get("customer", []))
    action_opts = _merge_selected_into_options(_base_opts("訂閱動作"), st.session_state.get("action", []))
    product_opts = _merge_selected_into_options(_base_opts("商品名稱"), st.session_state.get("product", []))

    sales_num = _make_num_opts(sales_opts)
    dealer_num = _make_num_opts(dealer_opts)
    qual_num = _make_num_opts(qual_opts)
    customer_num = _make_num_opts(customer_opts)
    action_num = _make_num_opts(action_opts)
    product_num = _make_num_opts(product_opts)

    _sync_valid_labeled("sales", "sales_labeled", sales_opts, sales_num)
    _sync_valid_labeled("dealer", "dealer_labeled", dealer_opts, dealer_num)
    _sync_valid_labeled("qual", "qual_labeled", qual_opts, qual_num)
    _sync_valid_labeled("customer", "customer_labeled", customer_opts, customer_num)
    _sync_valid_labeled("action", "action_labeled", action_opts, action_num)
    _sync_valid_labeled("product", "product_labeled", product_opts, product_num)

    def _on_change_sales_labeled():
        st.session_state["sales"] = _strip_num(st.session_state.get("sales_labeled", []) or [])

    def _on_change_dealer_labeled():
        dealer_raw = _strip_num(st.session_state.get("dealer_labeled", []) or [])
        st.session_state["dealer"] = dealer_raw
        picked_label = st.session_state.get("top30_dealer_pick", "（不套用 Top 30）")
        picked_raw = _re.sub(r"^\d+｜", "", str(picked_label)).strip() if picked_label != "（不套用 Top 30）" else None
        if picked_raw is not None and dealer_raw != [picked_raw]:
            st.session_state["top30_dealer_pick"] = "（不套用 Top 30）"

    def _on_change_customer_labeled():
        customer_raw = _strip_num(st.session_state.get("customer_labeled", []) or [])
        st.session_state["customer"] = customer_raw
        picked_label = st.session_state.get("top30_customer_pick", "（不套用 Top 30）")
        picked_raw = _re.sub(r"^\d+｜", "", str(picked_label)).strip() if picked_label != "（不套用 Top 30）" else None
        if picked_raw is not None and customer_raw != [picked_raw]:
            st.session_state["top30_customer_pick"] = "（不套用 Top 30）"

    def _on_change_qual_labeled():
        st.session_state["qual"] = _strip_num(st.session_state.get("qual_labeled", []) or [])

    def _on_change_action_labeled():
        st.session_state["action"] = _strip_num(st.session_state.get("action_labeled", []) or [])

    def _on_change_product_labeled():
        st.session_state["product"] = _strip_num(st.session_state.get("product_labeled", []) or [])

    dealer_top30_opts = _top30_opts("經銷商")
    customer_top30_opts = _top30_opts("最終客戶")
    _reset_top30_if_invalid("top30_dealer_pick", dealer_top30_opts)
    _reset_top30_if_invalid("top30_customer_pick", customer_top30_opts)

    # ══════════════════════════════════════════════════════════════════════════
    # Step 2：顯示 sidebar 篩選欄位
    # ══════════════════════════════════════════════════════════════════════════
    st.sidebar.divider()
    st.sidebar.subheader("多條件篩選（可多選）")

    sales_sel_labeled = st.sidebar.multiselect(
        "展碁業務",
        options=sales_num,
        key="sales_labeled",
        on_change=_on_change_sales_labeled,
        help="此欄位只依時間範圍列出選項；變更後會更新 Top 30 經銷商 / Top 30 最終客戶",
    )
    sales_sel = _strip_num(sales_sel_labeled)
    st.session_state["sales"] = sales_sel

    dealer_top30_opts = _top30_opts("經銷商")
    customer_top30_opts = _top30_opts("最終客戶")
    _reset_top30_if_invalid("top30_dealer_pick", dealer_top30_opts)
    _reset_top30_if_invalid("top30_customer_pick", customer_top30_opts)

    _top30_d_lbls = ["（不套用 Top 30）"] + [f"{i+1:02d}｜{v}" for i, v in enumerate(dealer_top30_opts[:30])]
    _top30_d_map = {
        "（不套用 Top 30）": None,
        **{f"{i+1:02d}｜{v}": v for i, v in enumerate(dealer_top30_opts[:30])}
    }
    st.session_state.setdefault("top30_dealer_pick", "（不套用 Top 30）")

    def _on_pick_top30_dealer():
        picked = _top30_d_map.get(st.session_state.get("top30_dealer_pick", "（不套用 Top 30）"))
        _set_single_pick_to_multiselect("dealer", "dealer_labeled", picked, dealer_opts, dealer_num)

    st.sidebar.selectbox(
        "Top 30 經銷商（依篩選範圍金額排序）",
        options=_top30_d_lbls,
        key="top30_dealer_pick",
        on_change=_on_pick_top30_dealer,
        help=f"依 {ANALYSIS_VALUE_COL} 排序，只套用時間範圍與展碁業務。選取後會同步寫入「經銷商」。",
    )

    dealer_sel_labeled = st.sidebar.multiselect(
        "經銷商",
        options=dealer_num,
        key="dealer_labeled",
        on_change=_on_change_dealer_labeled,
        help="此欄位只依時間範圍列出選項，不與其他多選條件互相連動",
    )
    dealer_sel = _strip_num(dealer_sel_labeled)
    st.session_state["dealer"] = dealer_sel

    qual_sel_labeled = st.sidebar.multiselect(
        "資格",
        options=qual_num,
        key="qual_labeled",
        on_change=_on_change_qual_labeled,
        help="此欄位只依時間範圍列出選項，不與其他多選條件互相連動",
    )
    qual_sel = _strip_num(qual_sel_labeled)
    st.session_state["qual"] = qual_sel

    _WARNING_OPTS = [
        "🔴 已到期", "🟤 寬限期",
        "🟡 15天", "🟠 30天", "🟣 45天", "🟤 60天", "⚫ 90天",
        "🟢 已續約", "🔴 已續約",
    ]
    if not isinstance(st.session_state.get("warning_filter_pick"), list):
        st.session_state["warning_filter_pick"] = []
    st.sidebar.multiselect(
        "未續約示警篩選（可多選）",
        options=_WARNING_OPTS,
        key="warning_filter_pick",
        help="篩選 F) 今年度分組明細中特定示警狀態，可多選",
    )

    _top30_c_lbls = ["（不套用 Top 30）"] + [f"{i+1:02d}｜{v}" for i, v in enumerate(customer_top30_opts[:30])]
    _top30_c_map = {
        "（不套用 Top 30）": None,
        **{f"{i+1:02d}｜{v}": v for i, v in enumerate(customer_top30_opts[:30])}
    }
    st.session_state.setdefault("top30_customer_pick", "（不套用 Top 30）")

    def _on_pick_top30_customer():
        picked = _top30_c_map.get(st.session_state.get("top30_customer_pick", "（不套用 Top 30）"))
        _set_single_pick_to_multiselect("customer", "customer_labeled", picked, customer_opts, customer_num)

    st.sidebar.selectbox(
        "Top 30 最終客戶（依篩選範圍金額排序）",
        options=_top30_c_lbls,
        key="top30_customer_pick",
        on_change=_on_pick_top30_customer,
        help=f"依 {ANALYSIS_VALUE_COL} 排序，只套用時間範圍與展碁業務。選取後會同步寫入「最終客戶」。",
    )

    customer_sel_labeled = st.sidebar.multiselect(
        "最終客戶",
        options=customer_num,
        key="customer_labeled",
        on_change=_on_change_customer_labeled,
        help="此欄位只依時間範圍列出選項，不與其他多選條件互相連動",
    )
    customer_sel = _strip_num(customer_sel_labeled)
    st.session_state["customer"] = customer_sel

    action_sel_labeled = st.sidebar.multiselect(
        "訂閱動作",
        options=action_num,
        key="action_labeled",
        on_change=_on_change_action_labeled,
        help="此欄位只依時間範圍列出選項，不與其他多選條件互相連動",
    )
    action_sel = _strip_num(action_sel_labeled)
    st.session_state["action"] = action_sel

    product_sel_labeled = st.sidebar.multiselect(
        "商品名稱",
        options=product_num,
        key="product_labeled",
        on_change=_on_change_product_labeled,
        help="此欄位只依時間範圍列出選項，不與其他多選條件互相連動",
    )
    product_sel = _strip_num(product_sel_labeled)
    st.session_state["product"] = product_sel


    progress.progress(90)
    status.update(label="套用篩選…", state="running", expanded=False)

    ui_state = {
        "future_expiry_enabled": bool(future_expiry_enabled),
        "future_expiry_months":  int(future_expiry_months),
        "expiry_range": (expiry_from, expiry_to),
        "order_range":  (order_from,  order_to),
        "dealer":   dealer_sel,
        "qual":     qual_sel,
        "customer": customer_sel,
        "action":   action_sel,
        "product":  product_sel,
        "sales":    sales_sel,
    }

    # 今年度（原篩選）
    df_filtered = apply_filters(df, ui_state)

    # 明年度（原篩選範圍 +1 年；若無資料視為 0）
    ui_next = shift_ui_state_one_year(ui_state)
    base_today_next = date.today() + relativedelta(years=1)
    df_filtered_next = apply_filters(df, ui_next, base_today=base_today_next)

    # ✅ 未續約示警篩選同步連動右方所有資訊
    df_filtered, df_filtered_next = apply_warning_filter_to_datasets(
        df_filtered,
        df_filtered_next,
        st.session_state.get("warning_filter_pick", []),
    )

    progress.progress(100)
    status.update(label=f"完成 ✅ 目前顯示 {len(df_filtered):,} 筆資料", state="complete", expanded=False)

except FileNotFoundError:
    status.update(label=f"找不到 Excel：請上傳檔案或確認路徑存在：{DEFAULT_XLSX_PATH}", state="error", expanded=True)
    st.stop()
except Exception as e:
    status.update(label=f"處理失敗：{e}", state="error", expanded=True)
    st.stop()

def render_csp_dashboard_v2_architecture(df_this: pd.DataFrame, df_next: pd.DataFrame, ui_state: dict):
    """第二分頁：CSP 續約儀表板 v2 架構（完整可執行版 + 四季 YoY + 統計口徑說明）"""

    st.markdown("""
    <div style="
        background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 55%, #0ea5e9 100%);
        border-radius: 20px;
        padding: 1.55rem 1.9rem;
        margin-bottom: 1rem;
        box-shadow: 0 14px 36px rgba(0,0,0,0.18);
    ">
        <div style="color:white; font-size:1.6rem; font-weight:900; letter-spacing:-0.4px;">
            🚀 CSP 續約儀表板 v2 架構
        </div>
        <div style="color:rgba(255,255,255,0.85); font-size:0.95rem; margin-top:0.35rem; line-height:1.65;">
            決策層 → 分析層 → 行動層｜不是只看報表，而是看風險、抓商機、排優先順序、直接行動。
        </div>
    </div>
    """, unsafe_allow_html=True)

    show_filter_ranges_if_enabled(ui_state)

    if df_this is None:
        df_this = pd.DataFrame()
    if df_next is None:
        df_next = pd.DataFrame()

    def _fmt_pct(v: float) -> str:
        try:
            return f"{float(v):,.1f}%"
        except Exception:
            return "0.0%"

    def _safe_first(x: pd.Series) -> str:
        if x is None:
            return ""
        s = x.dropna().astype(str)
        s = s[s.str.strip() != ""]
        return s.iloc[0] if not s.empty else ""

    def _product_mix(x: pd.Series, limit: int = 3) -> str:
        if x is None:
            return ""
        vals = pd.Series(x).dropna().astype(str)
        vals = vals[vals.str.strip() != ""]
        uniq = vals.unique().tolist()[:limit]
        return " / ".join(uniq)

    def _risk_level(days_left, next_amt: float, current_amt: float) -> str:
        if next_amt > 0:
            return "🟢 Safe"
        if pd.isna(days_left):
            return "⚪ 未知"
        try:
            d = int(days_left)
        except Exception:
            return "⚪ 未知"
        if d < 0:
            return "🔴 High"
        if d <= 30:
            return "🔴 High"
        if d <= 60:
            return "🟠 Medium"
        if d <= 90:
            return "🟡 Low"
        return "🔵 Pipeline"

    def _loss_probability(days_left, next_amt: float, current_amt: float) -> int:
        cur = float(current_amt or 0)
        nxt = float(next_amt or 0)
        if cur <= 0:
            return 0
        if nxt >= cur and nxt > 0:
            return 10
        if nxt > 0 and nxt < cur:
            return 45
        if pd.isna(days_left):
            return 35
        d = int(days_left)
        if d < 0 and nxt == 0:
            return 90
        if d <= 30 and nxt == 0:
            return 80
        if d <= 60 and nxt == 0:
            return 60
        if d <= 90 and nxt == 0:
            return 40
        return 20

    def _health_score(next_amt: float, current_amt: float, product_count: int, security_score: int) -> int:
        score = 0
        cur = float(current_amt or 0)
        nxt = float(next_amt or 0)
        if nxt > 0:
            score += 40
        if cur > 0 and nxt >= cur:
            score += 20
        elif nxt > 0:
            score += 10
        score += min(int(product_count or 0) * 5, 20)
        score += min(int(security_score or 0) * 5, 20)
        return int(min(score, 100))

    def _security_presence(text: str) -> int:
        s = str(text or "").lower()
        score = 0
        for k in ["defender", "purview", "intune", "entra", "azure ad", "aad", "p1", "p2"]:
            if k in s:
                score += 1
        return score

    def _month_sum(df_in: pd.DataFrame) -> pd.DataFrame:
        if df_in is None or df_in.empty or "訂閱到期日" not in df_in.columns:
            return pd.DataFrame({"月份": list(range(1, 13)), "金額": [0] * 12})
        d = df_in.copy()
        d["訂閱到期日"] = pd.to_datetime(d["訂閱到期日"], errors="coerce")
        d = d[pd.notna(d["訂閱到期日"])].copy()
        if d.empty:
            return pd.DataFrame({"月份": list(range(1, 13)), "金額": [0] * 12})
        d["月份"] = d["訂閱到期日"].dt.month
        d["金額"] = pd.to_numeric(d[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)
        return d.groupby("月份", dropna=False)["金額"].sum().reindex(list(range(1, 13)), fill_value=0).reset_index()

    this_total = float(pd.to_numeric(df_this.get(ANALYSIS_VALUE_COL), errors="coerce").fillna(0).sum()) if not df_this.empty and ANALYSIS_VALUE_COL in df_this.columns else 0.0
    next_total = float(pd.to_numeric(df_next.get(ANALYSIS_VALUE_COL), errors="coerce").fillna(0).sum()) if not df_next.empty and ANALYSIS_VALUE_COL in df_next.columns else 0.0
    renewed_amount = min(this_total, next_total)
    renewal_rate = (renewed_amount / this_total * 100) if this_total > 0 else 0.0

    if df_this.empty:
        st.info("目前無資料可供建立 v2 分析。")
        return

    mart = (
        df_this.groupby(["最終客戶", EXPIRY_YEAR_COL], dropna=False)
        .agg(
            今年金額=(ANALYSIS_VALUE_COL, "sum"),
            最近到期日=("訂閱到期日", "min"),
            最遠到期日=("訂閱到期日", "max"),
            經銷商=("經銷商", _safe_first),
            展碁業務=("展碁業務", _safe_first),
            客戶微軟網域=("客戶微軟網域", _safe_first),
            代表商品=("商品名稱", _safe_first),
            商品組合=("商品名稱", _product_mix),
            產品數=("商品名稱", pd.Series.nunique),
            訂閱筆數=(ANALYSIS_VALUE_COL, "size"),
        )
        .reset_index()
    )
    mart["今年金額"] = pd.to_numeric(mart["今年金額"], errors="coerce").fillna(0)
    mart["明年金額"] = 0.0

    next_lookup = build_group_renewal_lookup(df_next)
    next_amounts, days_lefts, risk_levels, loss_probs = [], [], [], []
    security_scores, health_scores, suggestions, risk_reasons, status_labels = [], [], [], [], []
    for _, row in mart.iterrows():
        cust = str(row.get("最終客戶", "") or "")
        yr = row.get(EXPIRY_YEAR_COL, pd.NA)
        try:
            next_key = int(float(yr)) + 1
        except Exception:
            next_key = str(yr)
        nxt = float((next_lookup.get((cust, next_key)) or {}).get("amount", 0) or 0)
        next_amounts.append(nxt)

        min_exp = pd.to_datetime(row.get("最近到期日", pd.NaT), errors="coerce")
        dl = (min_exp.date() - date.today()).days if pd.notna(min_exp) else pd.NA
        days_lefts.append(dl)
        rl = _risk_level(dl, nxt, row.get("今年金額", 0))
        risk_levels.append(rl)
        lp = _loss_probability(dl, nxt, row.get("今年金額", 0))
        loss_probs.append(lp)

        security_score = _security_presence(str(row.get("商品組合", "")) + " " + str(row.get("代表商品", "")))
        security_scores.append(security_score)
        health_scores.append(_health_score(nxt, row.get("今年金額", 0), row.get("產品數", 0), security_score))

        if nxt > 0:
            status_labels.append("已續約")
        elif pd.notna(dl) and int(dl) < 0:
            status_labels.append("已到期未續")
        elif pd.notna(dl) and int(dl) <= 30:
            status_labels.append("30天內待追")
        elif pd.notna(dl) and int(dl) <= 60:
            status_labels.append("60天內待追")
        elif pd.notna(dl) and int(dl) <= 90:
            status_labels.append("90天內暖身")
        else:
            status_labels.append("長天期 Pipeline")

        prod = str(row.get("代表商品", "") or "")
        mix = str(row.get("商品組合", "") or "")
        cat_row = df_this[df_this["最終客戶"].astype(str) == cust]
        cat = _safe_first(cat_row["產品分類"]) if not cat_row.empty and "產品分類" in cat_row.columns else ""
        recs = _pick_recommendations(prod or mix, cat)
        rec_names = [x[0] for x in recs.get("upsell", [])[:1]] + [x[0] for x in recs.get("cross_sell", [])[:1]]
        suggestions.append(" / ".join([x for x in rec_names if x]) or "維持續約追蹤")

        reason = []
        if nxt > 0:
            reason.append("已有明年度對應續約")
        else:
            if pd.notna(dl):
                if int(dl) < 0:
                    reason.append("已超過到期日")
                elif int(dl) <= 30:
                    reason.append("30天內到期但尚無明年度金額")
                elif int(dl) <= 60:
                    reason.append("60天內到期且尚未續約")
                elif int(dl) <= 90:
                    reason.append("90天內到期，應提前暖身")
            if row.get("今年金額", 0) > 0:
                reason.append(f"今年金額 {format_money(row.get('今年金額', 0))}")
        risk_reasons.append("；".join(reason))

    mart["明年金額"] = next_amounts
    mart["差異金額"] = mart["明年金額"] - mart["今年金額"]
    mart["續約剩餘天數"] = days_lefts
    mart["續約風險等級"] = risk_levels
    mart["流失機率"] = loss_probs
    mart["客戶健康度"] = health_scores
    mart["建議動作"] = suggestions
    mart["風險原因"] = risk_reasons
    mart["狀態"] = status_labels
    mart["商機類型"] = mart["建議動作"].apply(lambda x: "Upsell/Cross-sell" if x and x != "維持續約追蹤" else "Renewal")

    high_medium = mart[mart["續約風險等級"].isin(["🔴 High", "🟠 Medium"])]
    risk_amount = float(high_medium["今年金額"].sum()) if not high_medium.empty else 0.0
    lost_amount = float(mart[(mart["流失機率"] >= 80) & (mart["明年金額"] <= 0)]["今年金額"].sum()) if not mart.empty else 0.0
    pipeline_amount = float(mart[mart["續約風險等級"].isin(["🟡 Low", "🔵 Pipeline"])]["今年金額"].sum()) if not mart.empty else 0.0
    upsell_amount = float(mart[mart["商機類型"] == "Upsell/Cross-sell"]["今年金額"].sum()) if not mart.empty else 0.0
    high_customer_cnt = int(mart[mart["續約風險等級"] == "🔴 High"]["最終客戶"].astype(str).nunique()) if not mart.empty else 0
    high_reseller_cnt = int(mart[mart["續約風險等級"] == "🔴 High"]["經銷商"].astype(str).nunique()) if not mart.empty else 0
    high_sales_cnt = int(mart[mart["續約風險等級"] == "🔴 High"]["展碁業務"].astype(str).nunique()) if not mart.empty else 0

    st.markdown("## 1️⃣ Executive Summary")
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("今年訂閱金額", format_money(this_total))
    s2.metric("明年度已對應金額", format_money(next_total))
    s3.metric("Renewal Rate", _fmt_pct(renewal_rate))
    s4.metric("Revenue at Risk", format_money(risk_amount))

    s5, s6, s7, s8 = st.columns(4)
    s5.metric("Upsell Potential", format_money(upsell_amount))
    s6.metric("Lost Revenue", format_money(lost_amount))
    s7.metric("Pipeline", format_money(pipeline_amount))
    s8.metric("高風險客戶 / 經銷商 / 業務", f"{high_customer_cnt} / {high_reseller_cnt} / {high_sales_cnt}")

    top_risk_customer = "-"
    if not high_medium.empty:
        tr = high_medium.sort_values(["今年金額", "流失機率"], ascending=[False, False]).iloc[0]
        top_risk_customer = f"{tr['最終客戶']}（{format_money(tr['今年金額'])}）"
    top_growth_customer = "-"
    grow_df = mart[mart["差異金額"] > 0].sort_values("差異金額", ascending=False)
    if not grow_df.empty:
        gr = grow_df.iloc[0]
        top_growth_customer = f"{gr['最終客戶']}（+{format_money(gr['差異金額'])}）"
    top_risk_reseller = "-"
    if not high_medium.empty:
        rr = high_medium.groupby("經銷商", dropna=False)["今年金額"].sum().reset_index().sort_values("今年金額", ascending=False)
        if not rr.empty:
            top_risk_reseller = f"{rr.iloc[0]['經銷商']}（{format_money(rr.iloc[0]['今年金額'])}）"

    st.markdown(
        f"""
        <div style="background:white; border:1px solid rgba(0,0,0,0.06); border-radius:14px; padding:0.95rem 1.1rem; box-shadow:0 2px 10px rgba(0,0,0,0.05); margin-top:0.25rem; line-height:1.8;">
            <div style="font-size:0.82rem; color:#546e7a; font-weight:700; text-transform:uppercase; letter-spacing:0.8px; margin-bottom:0.35rem;">管理摘要</div>
            <div style="font-size:0.95rem; color:#263238;">
                目前本年度應續約金額為 <b>{format_money(this_total)}</b>，明年度已對應 <b>{format_money(next_total)}</b>，續約率 <b>{_fmt_pct(renewal_rate)}</b>。<br>
                其中高 / 中風險金額合計 <b>{format_money(risk_amount)}</b>，已流失或極高流失風險金額約 <b>{format_money(lost_amount)}</b>。<br>
                建議主管優先關注：<b>高風險客戶 {top_risk_customer}</b>、<b>高風險經銷商 {top_risk_reseller}</b>；若要推成長，優先查看 <b>{top_growth_customer}</b>。
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander("查看分頁 2 數據統計方式 / KPI 計算口徑", expanded=False):
        stats_df = pd.DataFrame([
            {"指標/欄位": "今年訂閱金額", "統計方式": "分頁 2 目前篩選後 df_this 的『成交價未稅小計』加總", "補充": "以目前篩選範圍視為今年度母體"},
            {"指標/欄位": "明年度已對應金額", "統計方式": "df_next 的『成交價未稅小計』加總", "補充": "df_next = 原篩選條件 +1 年"},
            {"指標/欄位": "Renewal Rate", "統計方式": "min(今年訂閱金額, 明年度已對應金額) / 今年訂閱金額", "補充": "避免明年度金額高於今年時續約率超過 100%"},
            {"指標/欄位": "Revenue at Risk", "統計方式": "客戶年度 mart 中屬於 🔴 High 或 🟠 Medium 的今年金額加總", "補充": "屬管理優先追蹤的風險池"},
            {"指標/欄位": "Lost Revenue", "統計方式": "流失機率 >= 80% 且 明年金額 <= 0 的今年金額加總", "補充": "視為已流失或極高流失風險"},
            {"指標/欄位": "Pipeline", "統計方式": "續約風險等級屬 🟡 Low 或 🔵 Pipeline 的今年金額加總", "補充": "屬可提前暖身/追蹤案件"},
            {"指標/欄位": "Upsell Potential", "統計方式": "建議動作不為『維持續約追蹤』且商機類型 = Upsell/Cross-sell 的今年金額加總", "補充": "為可延伸提案的商機池"},
            {"指標/欄位": "續約剩餘天數", "統計方式": "最近到期日 - 今日", "補充": "以客戶 + 到期年度的最早到期日為主"},
            {"指標/欄位": "續約風險等級", "統計方式": "已續約=🟢 Safe；<=30天=🔴；<=60天=🟠；<=90天=🟡；>90天=🔵", "補充": "若明年金額 > 0 優先視為已續約"},
            {"指標/欄位": "流失機率", "統計方式": "依剩餘天數與明年金額規則評分：已到期且未續=90%、30天內未續=80%、60天內未續=60%、90天內未續=40%", "補充": "規則引擎版，可再改成模型版"},
            {"指標/欄位": "客戶健康度", "統計方式": "是否續約 40 分 + 金額成長 20 分 + 產品數最多 20 分 + Security 關鍵字最多 20 分", "補充": "滿分 100 分"},
            {"指標/欄位": "四季 YOY", "統計方式": "依訂閱到期日季別將今年/明年金額分到 Q1~Q4，差異 = 明年 - 今年，YOY% = 差異 / 今年", "補充": "今年為 0 時，YOY% 顯示 0% 避免除零"},
            {"指標/欄位": "月趨勢", "統計方式": "每月到期金額 = df_this 依到期月份加總；每月續約金額 = df_next 依到期月份加總；每月流失金額 = max(到期-續約, 0)", "補充": "均以成交價未稅小計統計"},
        ])
        st.dataframe(stats_df, use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("## 2️⃣ 續約風險預警")
    c_left, c_right = st.columns([1, 1.35])

    risk_bucket = mart.groupby("續約風險等級", dropna=False).agg(客戶數=("最終客戶", "nunique"), 金額=("今年金額", "sum")).reset_index()
    order = ["🔴 High", "🟠 Medium", "🟡 Low", "🔵 Pipeline", "🟢 Safe", "⚪ 未知"]
    risk_bucket["續約風險等級"] = pd.Categorical(risk_bucket["續約風險等級"], categories=order, ordered=True)
    risk_bucket = risk_bucket.sort_values("續約風險等級").reset_index(drop=True)

    with c_left:
        st.markdown("#### Risk Score 分級")
        st.dataframe(risk_bucket.style.format({"金額": _safe_thousands_formatter, "客戶數": _safe_thousands_formatter}), use_container_width=True, hide_index=True)
        loss_top = mart.sort_values(["流失機率", "今年金額"], ascending=[False, False]).head(10)
        st.markdown("#### Top 10 流失風險")
        st.dataframe(loss_top[["最終客戶", "經銷商", "展碁業務", "代表商品", "今年金額", "流失機率", "續約風險等級"]].style.format({"今年金額": _safe_thousands_formatter}), use_container_width=True, hide_index=True)

    with c_right:
        chart = alt.Chart(risk_bucket).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
            x=alt.X("續約風險等級:N", sort=order, title="風險等級"),
            y=alt.Y("金額:Q", title="金額"),
            color=alt.Color("續約風險等級:N", scale=alt.Scale(domain=order, range=["#d32f2f", "#f57c00", "#c9a227", "#1976d2", "#2e7d32", "#90a4ae"]), legend=None),
            tooltip=["續約風險等級", alt.Tooltip("客戶數:Q", title="客戶數", format=",.0f"), alt.Tooltip("金額:Q", title="金額", format=",.0f")],
        ).properties(height=260)
        st.altair_chart(chart, use_container_width=True)

        mart_plot = mart.copy()
        bubble = alt.Chart(mart_plot).mark_circle(opacity=0.7).encode(
            x=alt.X("續約剩餘天數:Q", title="續約剩餘天數"),
            y=alt.Y("今年金額:Q", title="今年金額"),
            size=alt.Size("流失機率:Q", title="流失機率"),
            color=alt.Color("續約風險等級:N", scale=alt.Scale(domain=order, range=["#d32f2f", "#f57c00", "#c9a227", "#1976d2", "#2e7d32", "#90a4ae"])),
            tooltip=["最終客戶", "經銷商", "展碁業務", alt.Tooltip("今年金額:Q", format=",.0f"), "續約風險等級", alt.Tooltip("流失機率:Q", format=",.0f")],
        ).properties(height=290)
        st.altair_chart(bubble, use_container_width=True)

    r1, r2 = st.columns(2)
    with r1:
        reseller_risk = mart.groupby("經銷商", dropna=False).agg(風險金額=("今年金額", lambda s: float(s.sum())), 高風險客戶數=("續約風險等級", lambda s: int(pd.Series(s).isin(["🔴 High", "🟠 Medium"]).sum()))).reset_index()
        reseller_risk = reseller_risk.sort_values(["高風險客戶數", "風險金額"], ascending=[False, False]).head(10)
        st.markdown("#### 高風險經銷商")
        st.dataframe(reseller_risk.style.format({"風險金額": _safe_thousands_formatter}), use_container_width=True, hide_index=True)
    with r2:
        sales_risk = mart.groupby("展碁業務", dropna=False).agg(風險金額=("今年金額", lambda s: float(s.sum())), 高風險案件數=("續約風險等級", lambda s: int(pd.Series(s).isin(["🔴 High", "🟠 Medium"]).sum()))).reset_index()
        sales_risk = sales_risk.sort_values(["高風險案件數", "風險金額"], ascending=[False, False]).head(10)
        st.markdown("#### 高風險業務")
        st.dataframe(sales_risk.style.format({"風險金額": _safe_thousands_formatter}), use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("## 3️⃣ KPI 與年度趨勢")

    month_order = list(range(1, 13))
    expiry_trend = _month_sum(df_this).rename(columns={"金額": "到期金額"})
    renewal_trend = _month_sum(df_next).rename(columns={"金額": "續約金額"})
    lost_trend = expiry_trend.merge(renewal_trend, on="月份", how="left").fillna(0)
    lost_trend["流失金額"] = (lost_trend["到期金額"] - lost_trend["續約金額"]).clip(lower=0)

    lt_df = pd.concat([
        lost_trend[["月份", "到期金額"]].rename(columns={"到期金額": "金額"}).assign(指標="每月到期金額"),
        lost_trend[["月份", "續約金額"]].rename(columns={"續約金額": "金額"}).assign(指標="每月續約金額"),
        lost_trend[["月份", "流失金額"]].rename(columns={"流失金額": "金額"}).assign(指標="每月流失金額"),
    ], ignore_index=True)

    k_left, k_right = st.columns([1.25, 1])
    with k_left:
        line = alt.Chart(lt_df).mark_line(point=True).encode(
            x=alt.X("月份:O", sort=month_order, title="月份"),
            y=alt.Y("金額:Q", title="金額"),
            color=alt.Color("指標:N", scale=alt.Scale(domain=["每月到期金額", "每月續約金額", "每月流失金額"], range=["#1565c0", "#2e7d32", "#d32f2f"])),
            tooltip=["指標", "月份", alt.Tooltip("金額:Q", format=",.0f")],
        ).properties(height=310)
        st.markdown("#### Renewal vs Expiry")
        st.altair_chart(line, use_container_width=True)
    with k_right:
        funnel_df = pd.DataFrame([
            {"階段": "Total", "金額": this_total},
            {"階段": "Renewed", "金額": renewed_amount},
            {"階段": "Pending", "金額": max(this_total - renewed_amount - lost_amount, 0)},
            {"階段": "Lost", "金額": lost_amount},
        ])
        funnel = alt.Chart(funnel_df).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
            x=alt.X("階段:N", title="續約漏斗"),
            y=alt.Y("金額:Q", title="金額"),
            color=alt.Color("階段:N", scale=alt.Scale(domain=["Total", "Renewed", "Pending", "Lost"], range=["#1565c0", "#2e7d32", "#f57c00", "#d32f2f"]), legend=None),
            tooltip=["階段", alt.Tooltip("金額:Q", format=",.0f")],
        ).properties(height=150)
        st.markdown("#### Funnel")
        st.altair_chart(funnel, use_container_width=True)

        pipe_df = pd.DataFrame([
            {"區段": "30天", "金額": float(mart[(mart["續約剩餘天數"].fillna(999) <= 30) & (mart["明年金額"] <= 0)]["今年金額"].sum())},
            {"區段": "60天", "金額": float(mart[(mart["續約剩餘天數"].fillna(999) > 30) & (mart["續約剩餘天數"].fillna(999) <= 60) & (mart["明年金額"] <= 0)]["今年金額"].sum())},
            {"區段": "90天", "金額": float(mart[(mart["續約剩餘天數"].fillna(999) > 60) & (mart["續約剩餘天數"].fillna(999) <= 90) & (mart["明年金額"] <= 0)]["今年金額"].sum())},
        ])
        pipe = alt.Chart(pipe_df).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
            x=alt.X("區段:N", title="Pipeline"),
            y=alt.Y("金額:Q", title="待處理金額"),
            color=alt.Color("區段:N", scale=alt.Scale(domain=["30天", "60天", "90天"], range=["#d32f2f", "#f57c00", "#c9a227"]), legend=None),
            tooltip=["區段", alt.Tooltip("金額:Q", format=",.0f")],
        ).properties(height=150)
        st.altair_chart(pipe, use_container_width=True)

    st.markdown("#### 四季 YOY 分析")
    quarterly_kpi = build_quarterly_kpi_df(df_this, df_next)
    quarterly_kpi["YOY%"] = quarterly_kpi.apply(lambda r: ((float(r["差異"]) / float(r["本年度"]) * 100) if float(r["本年度"]) != 0 else 0), axis=1)
    q_line_df = pd.concat([
        quarterly_kpi[["季度", "本年度"]].rename(columns={"本年度": ANALYSIS_VALUE_COL}).assign(年度="本年度"),
        quarterly_kpi[["季度", "明年度"]].rename(columns={"明年度": ANALYSIS_VALUE_COL}).assign(年度="明年度"),
    ], ignore_index=True)

    q1, q2 = st.columns([1.2, 1])
    with q1:
        q_combo = alt.layer(
            alt.Chart(q_line_df).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
                x=alt.X("季度:O", sort=["Q1", "Q2", "Q3", "Q4"], title="季度"),
                y=alt.Y(f"{ANALYSIS_VALUE_COL}:Q", title="金額"),
                xOffset=alt.XOffset("年度:N"),
                color=alt.Color("年度:N", scale=alt.Scale(domain=["本年度", "明年度"], range=["#ef6c00", "#1565c0"])),
                tooltip=["年度", "季度", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
            ),
            alt.Chart(quarterly_kpi).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X("季度:O", sort=["Q1", "Q2", "Q3", "Q4"]),
                y=alt.Y("差異:Q", title="差異（明-今）"),
                color=alt.value("#2e7d32"),
                tooltip=["季度", alt.Tooltip("本年度:Q", format=",.0f"), alt.Tooltip("明年度:Q", format=",.0f"), alt.Tooltip("差異:Q", format=",.0f")],
            )
        ).resolve_scale(y='independent').properties(height=320)
        st.altair_chart(q_combo, use_container_width=True)
    with q2:
        q_yoy = alt.Chart(quarterly_kpi).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
            x=alt.X("季度:O", sort=["Q1", "Q2", "Q3", "Q4"], title="季度"),
            y=alt.Y("YOY%:Q", title="YOY%"),
            color=alt.condition(alt.datum['YOY%'] >= 0, alt.value('#2e7d32'), alt.value('#d32f2f')),
            tooltip=["季度", alt.Tooltip("本年度:Q", title="本年度", format=",.0f"), alt.Tooltip("明年度:Q", title="明年度", format=",.0f"), alt.Tooltip("差異:Q", title="差異", format=",.0f"), alt.Tooltip("YOY%:Q", title="YOY%", format=",.1f")],
        ).properties(height=220)
        st.altair_chart(q_yoy, use_container_width=True)
        st.dataframe(quarterly_kpi[["季度", "本年度", "明年度", "差異", "YOY%"]].style.format({"本年度": _safe_thousands_formatter, "明年度": _safe_thousands_formatter, "差異": _safe_thousands_formatter, "YOY%": lambda v: f"{float(v):,.1f}%"}), use_container_width=True, hide_index=True)

    st.caption("四季 YOY 統計方式：依『訂閱到期日』落點分到 Q1~Q4，統計金額以『成交價未稅小計』為主；差異 = 明年度 - 今年度；YOY% = 差異 / 今年度。")

    st.divider()
    st.markdown("## 4️⃣ 客戶與產品分析")
    cp1, cp2 = st.columns(2)

    cust_summary = mart.groupby("最終客戶", dropna=False).agg(
        今年金額=("今年金額", "sum"),
        明年金額=("明年金額", "sum"),
        風險最高等級=("續約風險等級", _safe_first),
        客戶健康度=("客戶健康度", "max"),
        經銷商=("經銷商", _safe_first),
        展碁業務=("展碁業務", _safe_first),
    ).reset_index()
    cust_summary["差異金額"] = cust_summary["明年金額"] - cust_summary["今年金額"]
    top_cust = cust_summary.sort_values("今年金額", ascending=False).head(10)
    top_cust_risk = cust_summary[cust_summary["最終客戶"].isin(high_medium["最終客戶"])].sort_values("今年金額", ascending=False).head(10)

    with cp1:
        st.markdown("#### Top 10 Revenue")
        bar = alt.Chart(top_cust).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
            x=alt.X("今年金額:Q", title="金額"),
            y=alt.Y("最終客戶:N", sort="-x", title="客戶"),
            color=alt.value("#1565c0"),
            tooltip=["最終客戶", alt.Tooltip("今年金額:Q", format=",.0f"), alt.Tooltip("明年金額:Q", format=",.0f"), alt.Tooltip("差異金額:Q", format=",.0f")],
        ).properties(height=320)
        st.altair_chart(bar, use_container_width=True)
    with cp2:
        st.markdown("#### Top 10 Risk")
        if top_cust_risk.empty:
            st.info("目前沒有高 / 中風險客戶")
        else:
            st.dataframe(top_cust_risk[["最終客戶", "經銷商", "展碁業務", "今年金額", "明年金額", "差異金額", "客戶健康度"]].style.format({"今年金額": _safe_thousands_formatter, "明年金額": _safe_thousands_formatter, "差異金額": _safe_thousands_formatter}), use_container_width=True, hide_index=True)

    prod_this = df_this.groupby("商品名稱", dropna=False).agg(今年金額=(ANALYSIS_VALUE_COL, "sum"), 今年數量=("數量", "sum")).reset_index()
    prod_next = df_next.groupby("商品名稱", dropna=False).agg(明年金額=(ANALYSIS_VALUE_COL, "sum"), 明年數量=("數量", "sum")).reset_index() if not df_next.empty else pd.DataFrame(columns=["商品名稱", "明年金額", "明年數量"])
    prod = prod_this.merge(prod_next, on="商品名稱", how="left").fillna(0)
    prod["金額差異"] = prod["明年金額"] - prod["今年金額"]
    prod["數量差異"] = prod["明年數量"] - prod["今年數量"]
    prod_top = prod.sort_values("今年金額", ascending=False).head(15)

    d_this = df_this.copy()
    d_this["訂閱到期日"] = pd.to_datetime(d_this["訂閱到期日"], errors="coerce")
    d_this = d_this[pd.notna(d_this["訂閱到期日"])].copy()
    d_this["月份"] = d_this["訂閱到期日"].dt.month

    p1, p2 = st.columns([1, 1.2])
    with p1:
        st.markdown("#### Top 15 商品金額")
        pbar = alt.Chart(prod_top).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
            x=alt.X("今年金額:Q", title="金額"),
            y=alt.Y("商品名稱:N", sort="-x", title="商品"),
            color=alt.value("#00897b"),
            tooltip=["商品名稱", alt.Tooltip("今年金額:Q", format=",.0f"), alt.Tooltip("明年金額:Q", format=",.0f"), alt.Tooltip("金額差異:Q", format=",.0f")],
        ).properties(height=420)
        st.altair_chart(pbar, use_container_width=True)
    with p2:
        st.markdown("#### 月份 × 產品 Heatmap")
        heat_base = d_this.copy()
        heat_base["商品名稱"] = heat_base["商品名稱"].astype(str)
        top_heat_products = prod_top["商品名稱"].astype(str).tolist()[:12]
        heat_base = heat_base[heat_base["商品名稱"].isin(top_heat_products)].copy()
        if heat_base.empty:
            st.info("無資料可顯示")
        else:
            heat = heat_base.groupby(["月份", "商品名稱"], dropna=False)[ANALYSIS_VALUE_COL].sum().reset_index()
            heat_chart = alt.Chart(heat).mark_rect().encode(
                x=alt.X("月份:O", sort=month_order, title="月份"),
                y=alt.Y("商品名稱:N", sort=top_heat_products, title="商品"),
                color=alt.Color(f"{ANALYSIS_VALUE_COL}:Q", title="金額", scale=alt.Scale(scheme="blues")),
                tooltip=["月份", "商品名稱", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
            ).properties(height=420)
            st.altair_chart(heat_chart, use_container_width=True)

    health_top = cust_summary.sort_values(["客戶健康度", "今年金額"], ascending=[False, False]).head(10)
    st.markdown("#### 客戶健康度 Top 10")
    st.dataframe(health_top[["最終客戶", "經銷商", "展碁業務", "客戶健康度", "今年金額", "明年金額"]].style.format({"今年金額": _safe_thousands_formatter, "明年金額": _safe_thousands_formatter}), use_container_width=True, hide_index=True)
    st.caption("客戶健康度統計方式：是否續約 40 分、金額成長 20 分、產品數最多 20 分、Security 關鍵字最多 20 分。")

    st.divider()
    st.markdown("## 5️⃣ 經銷商與業務分析")
    reseller = mart.groupby("經銷商", dropna=False).agg(
        今年金額=("今年金額", "sum"),
        明年金額=("明年金額", "sum"),
        風險金額=("今年金額", lambda s: float(s.sum())),
        高風險客戶數=("續約風險等級", lambda s: int(pd.Series(s).isin(["🔴 High", "🟠 Medium"]).sum())),
        Upsell潛力=("商機類型", lambda s: int(pd.Series(s).eq("Upsell/Cross-sell").sum())),
    ).reset_index()
    reseller["續約率"] = reseller.apply(lambda r: (min(r["今年金額"], r["明年金額"]) / r["今年金額"] * 100) if r["今年金額"] > 0 else 0, axis=1)
    reseller["差異金額"] = reseller["明年金額"] - reseller["今年金額"]
    reseller = reseller.sort_values("今年金額", ascending=False)

    sales = mart.groupby("展碁業務", dropna=False).agg(
        今年金額=("今年金額", "sum"),
        明年金額=("明年金額", "sum"),
        高風險案件數=("續約風險等級", lambda s: int(pd.Series(s).isin(["🔴 High", "🟠 Medium"]).sum())),
        失守金額=("今年金額", lambda s: float(s.sum())),
        Upsell潛力=("商機類型", lambda s: int(pd.Series(s).eq("Upsell/Cross-sell").sum())),
    ).reset_index()
    sales["續約率"] = sales.apply(lambda r: (min(r["今年金額"], r["明年金額"]) / r["今年金額"] * 100) if r["今年金額"] > 0 else 0, axis=1)
    sales["差異金額"] = sales["明年金額"] - sales["今年金額"]
    sales = sales.sort_values("今年金額", ascending=False)

    rs1, rs2 = st.columns(2)
    with rs1:
        st.markdown("#### 經銷商分析")
        st.dataframe(reseller[["經銷商", "今年金額", "明年金額", "差異金額", "續約率", "高風險客戶數", "Upsell潛力"]].head(15).style.format({"今年金額": _safe_thousands_formatter, "明年金額": _safe_thousands_formatter, "差異金額": _safe_thousands_formatter, "續約率": lambda v: f"{float(v):,.1f}%"}), use_container_width=True, hide_index=True)
    with rs2:
        st.markdown("#### 業務分析")
        st.dataframe(sales[["展碁業務", "今年金額", "明年金額", "差異金額", "續約率", "高風險案件數", "Upsell潛力"]].head(15).style.format({"今年金額": _safe_thousands_formatter, "明年金額": _safe_thousands_formatter, "差異金額": _safe_thousands_formatter, "續約率": lambda v: f"{float(v):,.1f}%"}), use_container_width=True, hide_index=True)

    rg1, rg2 = st.columns(2)
    with rg1:
        rchart = alt.Chart(reseller.head(10)).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
            x=alt.X("今年金額:Q", title="金額"),
            y=alt.Y("經銷商:N", sort="-x", title="經銷商"),
            color=alt.value("#3949ab"),
            tooltip=["經銷商", alt.Tooltip("今年金額:Q", format=",.0f"), alt.Tooltip("續約率:Q", format=",.1f"), "高風險客戶數"],
        ).properties(height=320, title="Top 10 經銷商 Revenue")
        st.altair_chart(rchart, use_container_width=True)
    with rg2:
        schart = alt.Chart(sales.head(10)).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
            x=alt.X("今年金額:Q", title="金額"),
            y=alt.Y("展碁業務:N", sort="-x", title="業務"),
            color=alt.value("#00897b"),
            tooltip=["展碁業務", alt.Tooltip("今年金額:Q", format=",.0f"), alt.Tooltip("續約率:Q", format=",.1f"), "高風險案件數"],
        ).properties(height=320, title="Top 10 業務 Revenue")
        st.altair_chart(schart, use_container_width=True)
    st.caption("經銷商 / 業務統計方式：以客戶 + 到期年度 mart 為基礎，先算每個客戶年度金額與風險，再依經銷商 / 展碁業務聚合。續約率 = min(今年, 明年) / 今年。")

    st.divider()
    st.markdown("## 6️⃣ 業務機會清單")
    opp_df = mart[mart["商機類型"] == "Upsell/Cross-sell"].copy().sort_values(["今年金額", "流失機率"], ascending=[False, False])
    if opp_df.empty:
        st.info("目前沒有明顯的 Upsell / Cross-sell 機會。")
    else:
        st.dataframe(
            opp_df[["最終客戶", "經銷商", "展碁業務", "代表商品", "今年金額", "續約風險等級", "建議動作", "流失機率", "客戶健康度"]].style.format({"今年金額": _safe_thousands_formatter}),
            use_container_width=True,
            hide_index=True,
        )
    st.caption("業務機會清單統計方式：以 _pick_recommendations() 對代表商品與產品分類做規則判斷，若存在升級或加購建議則標記為 Upsell/Cross-sell。")

    st.divider()
    st.markdown("## 7️⃣ 詳細明細")
    detail_tab1, detail_tab2, detail_tab3 = st.tabs(["風險明細", "商機明細", "分組明細"])
    with detail_tab1:
        risk_detail = mart.sort_values(["流失機率", "今年金額"], ascending=[False, False])
        st.dataframe(
            risk_detail[["最終客戶", "經銷商", "展碁業務", "客戶微軟網域", "代表商品", "商品組合", "今年金額", "明年金額", "差異金額", "續約剩餘天數", "續約風險等級", "流失機率", "風險原因", "建議動作"]].style.format({"今年金額": _safe_thousands_formatter, "明年金額": _safe_thousands_formatter, "差異金額": _safe_thousands_formatter}),
            use_container_width=True,
            hide_index=True,
        )
    with detail_tab2:
        st.dataframe(
            opp_df[["最終客戶", "經銷商", "展碁業務", "代表商品", "今年金額", "商機類型", "建議動作", "客戶健康度", "流失機率"]].style.format({"今年金額": _safe_thousands_formatter}),
            use_container_width=True,
            hide_index=True,
        )
    with detail_tab3:
        render_grouped_table(df_this, df_next=df_next)

    st.divider()
    st.markdown("## 8️⃣ AI 建議 / Copilot")
    prompt_col, result_col = st.columns([0.8, 1.2])
    with prompt_col:
        st.markdown("#### 建議提問")
        prompt_samples = pd.DataFrame([
            {"Prompt": "請分析續約風險"},
            {"Prompt": "哪些經銷商最需要優先輔導"},
            {"Prompt": "哪些業務名下風險最高"},
            {"Prompt": "哪些客戶適合推 Copilot / Security"},
            {"Prompt": "請幫我生成本週主管摘要"},
        ])
        st.dataframe(prompt_samples, use_container_width=True, hide_index=True)
    with result_col:
        top3_risk = mart.sort_values(["流失機率", "今年金額"], ascending=[False, False]).head(3)
        risk_lines = [f"- {r['最終客戶']}｜{r['經銷商']}｜{r['展碁業務']}｜{r['續約風險等級']}｜{format_money(r['今年金額'])}" for _, r in top3_risk.iterrows()]
        top3_opp = opp_df.sort_values("今年金額", ascending=False).head(3) if not opp_df.empty else pd.DataFrame()
        opp_lines = [f"- {r['最終客戶']}：建議 {r['建議動作']}" for _, r in top3_opp.iterrows()]
        top_sales_text = "-"
        if not sales.empty:
            srow = sales.sort_values(["高風險案件數", "今年金額"], ascending=[False, False]).iloc[0]
            top_sales_text = f"{srow['展碁業務']}（高風險 {int(srow['高風險案件數'])} 件）"
        ai_summary = f"""
【AI 建議摘要（規則引擎版，可直接接 AOAI / Copilot）】

1. 本週應優先聯絡客戶：
{chr(10).join(risk_lines) if risk_lines else '- 目前無高風險客戶'}

2. 本週建議優先輔導通路：
- {top_risk_reseller}

3. 本週建議主管關注業務：
- {top_sales_text}

4. 可優先推動的商機：
{chr(10).join(opp_lines) if opp_lines else '- 目前以續約追蹤為主'}

5. 建議下一步：
- 先處理 30 天內未續約且明年度金額為 0 的案件
- 對高金額客戶同步進行續約 + Security / Copilot 組合提案
- 對高風險經銷商安排專案檢視與週追蹤
"""
        st.text_area("AI 建議 / Copilot-ready 摘要", value=ai_summary, height=320)

    st.success("v2 已新增四季 YOY、月趨勢、統計口徑說明，並把目前做得到的內容完整整合為可執行版。")


tab_current, tab_v2 = st.tabs(["目前頁面", "CSP 續約儀表板 v2 架構"])

with tab_current:
    # ---------------------------
    # A) KPI（本年度 + 明年度 + 差異）
    # ---------------------------
    st.subheader(f"A) KPI 指標（以「{ANALYSIS_VALUE_COL}」分析）")
    show_filter_ranges_if_enabled(ui_state)

    # 本年度（原篩選）
    kpis_this = build_kpis(df_filtered)

    # 明年度（原篩選範圍 +1 年；若無資料視為 0）
    kpis_next = build_kpis(df_filtered_next)

    # 差異（明年度 - 今年度）
    kpis_diff = {
        "筆數": kpis_next["筆數"] - kpis_this["筆數"],
        "最終客戶數": kpis_next["最終客戶數"] - kpis_this["最終客戶數"],
        "經銷商數": kpis_next["經銷商數"] - kpis_this["經銷商數"],
        f"{ANALYSIS_VALUE_COL}合計": kpis_next[f"{ANALYSIS_VALUE_COL}合計"] - kpis_this[f"{ANALYSIS_VALUE_COL}合計"],
    }

    st.markdown('<div style="font-size:0.72rem; font-weight:700; color:#546e7a; text-transform:uppercase; letter-spacing:1px; margin-bottom:0.4rem; padding-left:2px;">📅 今年度（原篩選）</div>', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("筆數", f"{kpis_this['筆數']:,}")
    k2.metric("最終客戶數", f"{kpis_this['最終客戶數']:,}")
    k3.metric("經銷商數", f"{kpis_this['經銷商數']:,}")
    k4.metric(f"{ANALYSIS_VALUE_COL}合計", format_money(kpis_this[f"{ANALYSIS_VALUE_COL}合計"]))

    st.markdown('<div style="font-size:0.72rem; font-weight:700; color:#546e7a; text-transform:uppercase; letter-spacing:1px; margin-top:0.8rem; margin-bottom:0.4rem; padding-left:2px;">📅 明年度（原篩選範圍 +1 年；若無資料視為 0）</div>', unsafe_allow_html=True)
    n1, n2, n3, n4 = st.columns(4)
    n1.metric("筆數", f"{kpis_next['筆數']:,}")
    n2.metric("最終客戶數", f"{kpis_next['最終客戶數']:,}")
    n3.metric("經銷商數", f"{kpis_next['經銷商數']:,}")
    n4.metric(f"{ANALYSIS_VALUE_COL}合計", format_money(kpis_next[f"{ANALYSIS_VALUE_COL}合計"]))

    # ✅ 差異（明年度 - 今年度）數值以棕色呈現
    st.markdown('<div style="font-size:0.72rem; font-weight:700; color:#546e7a; text-transform:uppercase; letter-spacing:1px; margin-top:0.8rem; margin-bottom:0.4rem; padding-left:2px;">📈 差異（明年度 - 今年度）</div>', unsafe_allow_html=True)


    def _kpi_diff_card(label: str, value_str: str):
        is_positive = value_str.startswith("+")
        is_negative = value_str.startswith("-")
        arrow = "▲" if is_positive else ("▼" if is_negative else "—")
        bg_color = "rgba(46,125,50,0.07)" if is_positive else ("rgba(211,47,47,0.07)" if is_negative else "rgba(0,0,0,0.04)")
        border_color = "#2E7D32" if is_positive else ("#D32F2F" if is_negative else "#9e9e9e")
        text_color = "#2E7D32" if is_positive else ("#D32F2F" if is_negative else BROWN_COLOR)
        st.markdown(
            f"""
            <div style="
                background: {bg_color};
                border: 1px solid {border_color}22;
                border-left: 3px solid {border_color};
                border-radius: 10px;
                padding: 0.7rem 1rem;
                min-height: 72px;
                display: flex;
                flex-direction: column;
                justify-content: center;
            ">
                <div style="font-size: 0.7rem; color: #78909c; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 0.2rem;">{label} <span style="opacity:0.6;">YoY</span></div>
                <div style="font-size: 1.45rem; font-weight: 800; color: {text_color}; font-family: 'Inter', sans-serif; line-height: 1.1; letter-spacing: -0.3px;">{arrow} {value_str.lstrip('+-')}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


    d1, d2, d3, d4 = st.columns(4)
    with d1:
        _kpi_diff_card("筆數", format_signed_int(kpis_diff["筆數"]))
    with d2:
        _kpi_diff_card("最終客戶數", format_signed_int(kpis_diff["最終客戶數"]))
    with d3:
        _kpi_diff_card("經銷商數", format_signed_int(kpis_diff["經銷商數"]))
    with d4:
        _kpi_diff_card(f"{ANALYSIS_VALUE_COL}合計", format_signed_money(kpis_diff[f"{ANALYSIS_VALUE_COL}合計"]))

    st.markdown('<div style="font-size:0.72rem; font-weight:700; color:#546e7a; text-transform:uppercase; letter-spacing:1px; margin-top:0.8rem; margin-bottom:0.4rem; padding-left:2px;">📊 四季趨勢（本年度 / 明年度 / 差異）</div>', unsafe_allow_html=True)
    quarterly_kpi = build_quarterly_kpi_df(df_filtered, df_filtered_next)

    if quarterly_kpi.empty:
        st.info("無資料可顯示")
    else:
        quarterly_kpi = quarterly_kpi.copy()
        quarterly_kpi["YOY%"] = quarterly_kpi.apply(
            lambda r: ((float(r["差異"]) / float(r["本年度"]) * 100) if float(r["本年度"]) != 0 else 0),
            axis=1,
        )
        quarter_order = ["Q1", "Q2", "Q3", "Q4"]
        quarter_bar_df = pd.concat(
            [
                quarterly_kpi[["季度", "本年度"]].rename(columns={"本年度": "金額"}).assign(年度="本年度"),
                quarterly_kpi[["季度", "明年度"]].rename(columns={"明年度": "金額"}).assign(年度="明年度"),
            ],
            ignore_index=True,
        )

        left_col, right_col = st.columns([1.8, 1.0], gap="large")

        with left_col:
            bar_chart = alt.Chart(quarter_bar_df).mark_bar(size=40, cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
                x=alt.X("季度:O", sort=quarter_order, title="季度", axis=alt.Axis(labelAngle=0, labelFontSize=12, titleFontSize=12)),
                y=alt.Y("金額:Q", title="金額", axis=alt.Axis(format=",.0f", labelFontSize=11, titleFontSize=12)),
                xOffset=alt.XOffset("年度:N"),
                color=alt.Color(
                    "年度:N",
                    title="年度",
                    scale=alt.Scale(domain=["本年度", "明年度"], range=["#ef6c00", "#1565c0"]),
                    legend=alt.Legend(orient="top-right"),
                ),
                tooltip=["年度", "季度", alt.Tooltip("金額:Q", title="金額", format=",.0f")],
            )

            diff_line = alt.Chart(quarterly_kpi).mark_line(point=True, strokeWidth=3, color="#2e7d32").encode(
                x=alt.X("季度:O", sort=quarter_order),
                y=alt.Y("差異:Q", title="差異（明-今）", axis=alt.Axis(format=",.0f", orient="right", labelFontSize=11, titleFontSize=12)),
                tooltip=[
                    "季度",
                    alt.Tooltip("本年度:Q", title="本年度", format=",.0f"),
                    alt.Tooltip("明年度:Q", title="明年度", format=",.0f"),
                    alt.Tooltip("差異:Q", title="差異（明-今）", format=",.0f"),
                    alt.Tooltip("YOY%:Q", title="YOY%", format=",.1f"),
                ],
            )

            quarter_dual_chart = alt.layer(bar_chart, diff_line).resolve_scale(y="independent").properties(height=380)
            st.altair_chart(quarter_dual_chart, use_container_width=True)

        with right_col:
            table_show = quarterly_kpi[["季度", "本年度", "明年度", "差異", "YOY%"]].copy()
            st.dataframe(
                table_show.style.format(
                    {
                        "本年度": _safe_thousands_formatter,
                        "明年度": _safe_thousands_formatter,
                        "差異": _safe_thousands_formatter,
                        "YOY%": lambda v: f"{float(v):,.1f}%",
                    }
                ),
                use_container_width=True,
                hide_index=True,
                height=380,
            )

    st.divider()

    st.divider()

    # ---------------------------
    # 10) Top 10（續約客戶 / 經銷商）
    # ---------------------------
    st.subheader(f"B) Top 10 續約客戶（以 {ANALYSIS_VALUE_COL} 合計排序）")
    show_filter_ranges_if_enabled(ui_state)

    top10_customer_this = top10_by(df_filtered, "最終客戶", ANALYSIS_VALUE_COL, top_n=10, extra_col="商品名稱").rename(
        columns={ANALYSIS_VALUE_COL: "需約金額"}
    )
    top10_customer_this["需約金額"] = pd.to_numeric(top10_customer_this["需約金額"], errors="coerce").fillna(0)

    # 明年度續約金額（同客戶，範圍為原篩選 +1 年；若無視為 0）
    top10_customer_next_all = (
        df_filtered_next.groupby("最終客戶", dropna=False)[ANALYSIS_VALUE_COL]
        .sum()
        .reset_index()
        .rename(columns={ANALYSIS_VALUE_COL: "明年度續約金額"})
    )
    top10_customer_next_all["明年度續約金額"] = pd.to_numeric(top10_customer_next_all["明年度續約金額"], errors="coerce").fillna(0)

    top10_customer = top10_customer_this.merge(top10_customer_next_all, on="最終客戶", how="left")
    top10_customer["明年度續約金額"] = pd.to_numeric(top10_customer["明年度續約金額"], errors="coerce").fillna(0)

    # 差異金額 = 明年度續約金額 - 今年度需約金額
    top10_customer["差異金額"] = top10_customer["明年度續約金額"] - top10_customer["需約金額"]

    # 欄位順序：需約金額後 -> 明年度續約金額 -> 差異金額
    cols = top10_customer.columns.tolist()
    desired = ["最終客戶", "需約金額", "明年度續約金額", "差異金額"]
    ordered_cols = [c for c in desired if c in cols] + [c for c in cols if c not in set(desired)]
    top10_customer = top10_customer[ordered_cols]

    # ✅ 右側主區：數字排序欄
    top10_customer.insert(0, "排序", [str(i+1) for i in range(len(top10_customer))])

    # ✅ 差異金額欄位名稱棕色（表頭）
    cust_styler = top10_customer.style.format(
        {
            "需約金額": _safe_thousands_formatter,
            "明年度續約金額": _safe_thousands_formatter,
            "差異金額": _safe_thousands_formatter,
        }
    ).set_table_styles(_style_header_color_for_column(top10_customer, "差異金額", BROWN_COLOR), overwrite=False)

    st.dataframe(cust_styler, use_container_width=True, hide_index=True)

    st.subheader(f"C) Top 10 經銷商（以 {ANALYSIS_VALUE_COL} 合計排序）")
    show_filter_ranges_if_enabled(ui_state)

    # 本年度 Top10 經銷商（金額）
    top10_dealer_this = top10_by(df_filtered, "經銷商", ANALYSIS_VALUE_COL, top_n=10).rename(columns={ANALYSIS_VALUE_COL: "金額"})
    top10_dealer_this["金額"] = pd.to_numeric(top10_dealer_this["金額"], errors="coerce").fillna(0)

    # ✅ 明年度續約金額（原篩選範圍 +1 年；若無視為 0）
    top10_dealer_next_all = (
        df_filtered_next.groupby("經銷商", dropna=False)[ANALYSIS_VALUE_COL]
        .sum()
        .reset_index()
        .rename(columns={ANALYSIS_VALUE_COL: "明年度續約金額"})
    )
    top10_dealer_next_all["明年度續約金額"] = pd.to_numeric(top10_dealer_next_all["明年度續約金額"], errors="coerce").fillna(0)

    top10_dealer = top10_dealer_this.merge(top10_dealer_next_all, on="經銷商", how="left")
    top10_dealer["明年度續約金額"] = pd.to_numeric(top10_dealer["明年度續約金額"], errors="coerce").fillna(0)

    # ✅ 差異金額 = 明年度續約金額 - 今年度金額
    top10_dealer["差異金額"] = top10_dealer["明年度續約金額"] - top10_dealer["金額"]

    # 欄位順序：金額 -> 明年度續約金額 -> 差異金額
    dealer_cols = top10_dealer.columns.tolist()
    dealer_desired = ["經銷商", "金額", "明年度續約金額", "差異金額"]
    dealer_ordered = [c for c in dealer_desired if c in dealer_cols] + [c for c in dealer_cols if c not in set(dealer_desired)]
    top10_dealer = top10_dealer[dealer_ordered]

    # ✅ 右側主區：數字排序欄
    top10_dealer.insert(0, "排序", [str(i+1) for i in range(len(top10_dealer))])

    # ✅ 差異金額欄位名稱棕色（表頭）
    dealer_styler = top10_dealer.style.format(
        {
            "金額": _safe_thousands_formatter,
            "明年度續約金額": _safe_thousands_formatter,
            "差異金額": _safe_thousands_formatter,
        }
    ).set_table_styles(_style_header_color_for_column(top10_dealer, "差異金額", BROWN_COLOR), overwrite=False)

    st.dataframe(dealer_styler, use_container_width=True, hide_index=True)

    st.divider()

    # ---------------------------
    # 圖表
    # ---------------------------
    st.subheader(f"D) 資訊圖表（以「{ANALYSIS_VALUE_COL}」分析）")
    show_filter_ranges_if_enabled(ui_state)

    c1, c3 = st.columns(2)

    with c1:
        st.markdown('<div style="font-size:0.78rem; font-weight:700; color:#1565c0; margin-bottom:0.4rem; padding-left:2px;">📊 Top 10 續約客戶（需約金額）</div>', unsafe_allow_html=True)
        if top10_customer.empty:
            st.info("無資料可顯示")
        else:
            tmp = top10_customer[["最終客戶", "需約金額"]].rename(columns={"需約金額": ANALYSIS_VALUE_COL}).copy()
            tmp[ANALYSIS_VALUE_COL] = pd.to_numeric(tmp[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)
            st.altair_chart(
                alt.Chart(tmp)
                .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    x=alt.X(ANALYSIS_VALUE_COL, title="金額", axis=alt.Axis(format=",.0s", labelFontSize=10)),
                    y=alt.Y("最終客戶:N", sort="-x", title="最終客戶", axis=alt.Axis(labelFontSize=11)),
                    color=alt.Color("最終客戶:N", legend=None, scale=alt.Scale(scheme="blues")),
                    tooltip=["最終客戶", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
                )
                .properties(height=380),
                use_container_width=True,
            )


    with c3:
        st.markdown('<div style="font-size:0.78rem; font-weight:700; color:#1565c0; margin-bottom:0.4rem; padding-left:2px;">📈 到期月份金額趨勢（今年度 vs 明年度）</div>', unsafe_allow_html=True)

        def _months_in_range(start_d: date, end_d: date) -> list[int]:
            """
            回傳在 start_d~end_d 之間會涵蓋到的月份（1~12），只看月份區間，不強制補齊 1~12。
            若跨年，會回傳跨年的月份序列（例如 11,12,1,2）。
            """
            if start_d is None or end_d is None:
                return list(range(1, 13))

            if start_d > end_d:
                start_d, end_d = end_d, start_d

            m1, m2 = start_d.month, end_d.month
            y1, y2 = start_d.year, end_d.year

            if y1 == y2:
                return list(range(m1, m2 + 1))

            # 跨年：例如 2026/11~2027/02 -> [11,12,1,2]
            return list(range(m1, 13)) + list(range(1, m2 + 1))

        def _month_trend(df_in: pd.DataFrame, months: list[int]) -> pd.DataFrame:
            """回傳欄位：到期月份, ANALYSIS_VALUE_COL（已補齊 months_scope）"""
            if df_in is None or df_in.empty:
                return pd.DataFrame({"到期月份": months, ANALYSIS_VALUE_COL: [0] * len(months)})

            d = df_in[pd.notna(df_in["訂閱到期日"])].copy()
            if d.empty:
                return pd.DataFrame({"到期月份": months, ANALYSIS_VALUE_COL: [0] * len(months)})

            d["到期月份"] = d["訂閱到期日"].dt.month
            trend = d.groupby("到期月份")[ANALYSIS_VALUE_COL].sum()
            trend = pd.to_numeric(trend, errors="coerce").fillna(0)

            # ✅ 只補齊「篩選範圍涵蓋的月份」
            trend = trend.reindex(months, fill_value=0).reset_index()
            return trend

        # ✅ 依左側篩選決定月份範圍（若勾選「未來N月」，則以 today~end_date 的月份範圍）
        if ui_state["future_expiry_enabled"]:
            base_today = date.today()
            end_date = base_today + relativedelta(months=ui_state["future_expiry_months"])
            months_scope = _months_in_range(base_today, end_date)
        else:
            exp_from, exp_to = ui_state["expiry_range"]
            months_scope = _months_in_range(exp_from, exp_to)

        # 本年度 / 隔年度（月趨勢）
        t_this = _month_trend(df_filtered, months_scope).rename(columns={ANALYSIS_VALUE_COL: f"{ANALYSIS_VALUE_COL}_本年度"})
        t_next = _month_trend(df_filtered_next, months_scope).rename(columns={ANALYSIS_VALUE_COL: f"{ANALYSIS_VALUE_COL}_隔年度"})

        # 合併後計算差異（隔年度 - 本年度）
        merged = t_this.merge(t_next, on="到期月份", how="outer").fillna(0)

        merged[f"{ANALYSIS_VALUE_COL}_本年度"] = pd.to_numeric(merged[f"{ANALYSIS_VALUE_COL}_本年度"], errors="coerce").fillna(0)
        merged[f"{ANALYSIS_VALUE_COL}_隔年度"] = pd.to_numeric(merged[f"{ANALYSIS_VALUE_COL}_隔年度"], errors="coerce").fillna(0)

        merged["差異"] = merged[f"{ANALYSIS_VALUE_COL}_隔年度"] - merged[f"{ANALYSIS_VALUE_COL}_本年度"]

        if merged.empty:
            st.info("無資料可顯示")
        else:
            # 上半部：雙線圖（本年度 vs 隔年度）
            line_df = pd.concat(
                [
                    merged[["到期月份", f"{ANALYSIS_VALUE_COL}_本年度"]]
                    .rename(columns={f"{ANALYSIS_VALUE_COL}_本年度": ANALYSIS_VALUE_COL})
                    .assign(年度="本年度"),
                    merged[["到期月份", f"{ANALYSIS_VALUE_COL}_隔年度"]]
                    .rename(columns={f"{ANALYSIS_VALUE_COL}_隔年度": ANALYSIS_VALUE_COL})
                    .assign(年度="明年度"),
                ],
                ignore_index=True,
            )

            line_chart = (
                alt.Chart(line_df)
                .mark_line(point=True)
                .encode(
                    x=alt.X("到期月份:O", sort=months_scope, title=""),
                    y=alt.Y(f"{ANALYSIS_VALUE_COL}:Q", title="金額"),
                    color=alt.Color("年度:N", title="年度"),
                    tooltip=["年度", "到期月份", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
                )
                .properties(height=180)
            )

            # 下半部：差異柱狀圖（隔年度 - 本年度）
            bar_chart = (
                alt.Chart(merged)
                .mark_bar()
                .encode(
                    x=alt.X("到期月份:O", sort=months_scope, title="到期月份"),
                    y=alt.Y("差異:Q", title="差異（明年度 - 今年度）"),
                    # 正值（隔 > 本）用棕色；負值用藍色
                    color=alt.condition(
                        alt.datum.差異 >= 0,
                        alt.value("#8B4513"),
                        alt.value("#4682B4"),
                    ),
                    tooltip=[
                        "到期月份",
                        alt.Tooltip(f"{ANALYSIS_VALUE_COL}_本年度:Q", title="本年度", format=",.0f"),
                        alt.Tooltip(f"{ANALYSIS_VALUE_COL}_隔年度:Q", title="明年度", format=",.0f"),
                        alt.Tooltip("差異:Q", title="差異（明-今）", format=",.0f"),
                    ],
                )
                .properties(height=210)
            )

            final_chart = alt.vconcat(line_chart, bar_chart).resolve_scale(color="independent")
            st.altair_chart(final_chart, use_container_width=True)




    c4, c5 = st.columns([1, 2])
    # -------------------------------------------------
    # 4 🆕 新增：Top 15 商品名稱（金額）
    # -------------------------------------------------
    with c4:
        st.markdown('<div style="font-size:0.78rem; font-weight:700; color:#1565c0; margin-bottom:0.4rem; padding-left:2px;">📦 Top 15 商品名稱（金額）</div>', unsafe_allow_html=True)

        prod_amt = (
            df_filtered.groupby("商品名稱", dropna=False)[ANALYSIS_VALUE_COL]
            .sum()
            .sort_values(ascending=False)
            .head(15)
            .reset_index()
        )

        prod_amt[ANALYSIS_VALUE_COL] = pd.to_numeric(prod_amt[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)

        if prod_amt.empty:
            st.info("無資料可顯示")
        else:
            st.altair_chart(
                alt.Chart(prod_amt)
                .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    x=alt.X(ANALYSIS_VALUE_COL, title="金額", axis=alt.Axis(format=",.0s", labelFontSize=10)),
                    y=alt.Y("商品名稱:N", sort="-x",
                            axis=alt.Axis(labelFontSize=10, labelLimit=400, labelOverlap=False)),
                    color=alt.Color("商品名稱:N", legend=None, scale=alt.Scale(scheme="teals")),
                    tooltip=["商品名稱", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
                )
                .properties(height=480),
                use_container_width=True,
            )
    with c5:
        st.markdown('<div style="font-size:0.78rem; font-weight:700; color:#1565c0; margin-bottom:0.4rem; padding-left:2px;">📊 商品名稱 Top 15（金額及數量之增減｜今年度 vs 明年度）</div>', unsafe_allow_html=True)

        # 依本年度金額排序取 Top15 商品
        prod_amt_this = df_filtered.groupby("商品名稱", dropna=False)[ANALYSIS_VALUE_COL].sum()
        prod_amt_this = pd.to_numeric(prod_amt_this, errors="coerce").fillna(0).sort_values(ascending=False)
        top_products = prod_amt_this.head(15).index.tolist()

        if not top_products:
            st.info("無資料可顯示")
        else:

            # ===== 金額 =====
            amt_this = prod_amt_this.reindex(top_products, fill_value=0).reset_index()
            amt_this.columns = ["商品名稱", "金額_本年度"]

            prod_amt_next = df_filtered_next.groupby("商品名稱", dropna=False)[ANALYSIS_VALUE_COL].sum()
            prod_amt_next = pd.to_numeric(prod_amt_next, errors="coerce").fillna(0)
            amt_next = prod_amt_next.reindex(top_products, fill_value=0).reset_index()
            amt_next.columns = ["商品名稱", "金額_明年度"]

            merged = amt_this.merge(amt_next, on="商品名稱", how="outer").fillna(0)
            merged["金額_差異"] = merged["金額_明年度"] - merged["金額_本年度"]

            # ===== 數量 =====
            qty_this = df_filtered.groupby("商品名稱", dropna=False)["數量"].sum()
            qty_this = pd.to_numeric(qty_this, errors="coerce").fillna(0)
            qty_this = qty_this.reindex(top_products, fill_value=0).reset_index()
            qty_this.columns = ["商品名稱", "數量_本年度"]

            qty_next = df_filtered_next.groupby("商品名稱", dropna=False)["數量"].sum()
            qty_next = pd.to_numeric(qty_next, errors="coerce").fillna(0)
            qty_next = qty_next.reindex(top_products, fill_value=0).reset_index()
            qty_next.columns = ["商品名稱", "數量_明年度"]

            merged = merged.merge(qty_this, on="商品名稱", how="left")
            merged = merged.merge(qty_next, on="商品名稱", how="left")
            merged["數量_差異"] = merged["數量_明年度"] - merged["數量_本年度"]

            merged = merged.sort_values("金額_本年度", ascending=False)
            # ✅ 讓「差異直條圖」X 軸改顯示差異數值（取代商品名稱）
            merged["金額差異_label"] = merged["金額_差異"].apply(lambda v: format_signed_money(v))
            merged["數量差異_label"] = merged["數量_差異"].apply(lambda v: format_signed_int(v))
            sort_list = merged["商品名稱"].tolist()
            sort_list_amt_diff = merged["金額差異_label"].tolist()
            sort_list_qty_diff = merged["數量差異_label"].tolist()

            # ------------------------
            # 金額雙線圖
            # ------------------------
            line_amt_df = pd.concat([
                merged[["商品名稱", "金額_本年度"]].rename(columns={"金額_本年度": "值"}).assign(年度="本年度"),
                merged[["商品名稱", "金額_明年度"]].rename(columns={"金額_明年度": "值"}).assign(年度="明年度")
            ])

            line_amt = (
                alt.Chart(line_amt_df)
                .mark_line(point=True)
                .encode(
                    x=alt.X("商品名稱:O", sort=sort_list),
                    y=alt.Y("值:Q", title="金額"),
                    color=alt.Color("年度:N"),
                    tooltip=["年度", "商品名稱", alt.Tooltip("值:Q", format=",.0f")]
                )
                .properties(height=175)
            )

            bar_amt = (
                alt.Chart(merged)
                .mark_bar()
                .encode(
                    # ✅ X 軸改用「金額差異_label」(顯示差異數值)，不再顯示商品名稱
                    x=alt.X("金額差異_label:O", sort=sort_list_amt_diff, title="金額差異（明-今）"),
                    y=alt.Y("金額_差異:Q", title="金額差異（明-今）"),
                    color=alt.condition(
                        alt.datum.金額_差異 >= 0,
                        alt.value("#6A0DAD"),
                        alt.value("#DC143C")
                    ),
                    # ✅ Tooltip：商品名稱、金額_本年度、金額_明年度、金額_差異（依你指定的順序）
                    tooltip=[
                        alt.Tooltip("商品名稱:N", title="商品名稱"),
                        alt.Tooltip("金額_本年度:Q", title="金額_本年度", format=",.0f"),
                        alt.Tooltip("金額_明年度:Q", title="金額_明年度", format=",.0f"),
                        alt.Tooltip("金額_差異:Q", title="金額_差異（明-今）", format=",.0f"),
                    ],
                )
                .properties(height=175)
            )

            # ------------------------
            # 數量雙線圖
            # ------------------------
            line_qty_df = pd.concat([
                merged[["商品名稱", "數量_本年度"]].rename(columns={"數量_本年度": "值"}).assign(年度="本年度"),
                merged[["商品名稱", "數量_明年度"]].rename(columns={"數量_明年度": "值"}).assign(年度="明年度")
            ])

            line_qty = (
                alt.Chart(line_qty_df)
                .mark_line(point=True)
                .encode(
                    x=alt.X("商品名稱:O", sort=sort_list),
                    y=alt.Y("值:Q", title="數量"),
                    color=alt.Color("年度:N", scale=alt.Scale(range=["#2E8B57", "#FF8C00"])),
                    tooltip=["年度", "商品名稱", "值"]
                )
                .properties(height=175)
            )

            bar_qty = (
                alt.Chart(merged)
                .mark_bar()
                .encode(
                    x=alt.X("商品名稱:N", sort=sort_list, title="商品名稱"),
                    y=alt.Y("數量_差異:Q", title="數量差異（明-今）"),
                    color=alt.condition(
                        alt.datum.數量_差異 >= 0,
                        alt.value("#2E8B57"),
                        alt.value("#B22222")
                    ),
                    tooltip=[
                        alt.Tooltip("商品名稱:N", title="商品名稱"),
                        alt.Tooltip("數量_本年度:Q", title="數量_本年度", format=",.0f"),
                        alt.Tooltip("數量_明年度:Q", title="數量_明年度", format=",.0f"),
                        alt.Tooltip("數量_差異:Q", title="數量_差異（明-今）", format=",.0f"),
                    ],
                )
                .properties(height=175)
            )

            left_block = alt.vconcat(
                line_amt,
                bar_amt
            )

            right_block = alt.vconcat(
                line_qty,
                bar_qty
            )

            final_chart = alt.hconcat(
                left_block,
                right_block
            ).resolve_scale(color="independent")

            st.altair_chart(final_chart, use_container_width=True)



    c7, c8 = st.columns(2)
    with c7:
        st.markdown('<div style="font-size:0.78rem; font-weight:700; color:#1565c0; margin-bottom:0.4rem; padding-left:2px;">👤 業務人員分析：各展碁業務金額</div>', unsafe_allow_html=True)
        sales_sum = df_filtered.groupby("展碁業務", dropna=False)[ANALYSIS_VALUE_COL].sum().reset_index()
        sales_sum["展碁業務"] = sales_sum["展碁業務"].astype(str).replace({"nan": "未填"})
        sales_sum[ANALYSIS_VALUE_COL] = pd.to_numeric(sales_sum[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)
        sales_sum = sales_sum.sort_values(ANALYSIS_VALUE_COL, ascending=False).head(20)
        if sales_sum.empty:
            st.info("無資料可顯示")
        else:
            st.altair_chart(
                alt.Chart(sales_sum)
                .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    x=alt.X(ANALYSIS_VALUE_COL, title="金額", axis=alt.Axis(format=",.0s", labelFontSize=10)),
                    y=alt.Y("展碁業務:N", sort="-x", title="展碁業務", axis=alt.Axis(labelFontSize=11)),
                    color=alt.Color("展碁業務:N", legend=None, scale=alt.Scale(scheme="tealblues")),
                    tooltip=["展碁業務", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
                )
                .properties(height=480),
                use_container_width=True,
            )

    with c8:
        st.markdown('<div style="font-size:0.78rem; font-weight:700; color:#1565c0; margin-bottom:0.4rem; padding-left:2px;">🔄 訂閱動作分析：各訂閱動作金額</div>', unsafe_allow_html=True)
        action_sum = df_filtered.groupby("訂閱動作", dropna=False)[ANALYSIS_VALUE_COL].sum().reset_index()
        action_sum["訂閱動作"] = action_sum["訂閱動作"].astype(str).replace({"nan": "未填"})
        action_sum[ANALYSIS_VALUE_COL] = pd.to_numeric(action_sum[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)
        action_sum = action_sum.sort_values(ANALYSIS_VALUE_COL, ascending=False)
        if action_sum.empty:
            st.info("無資料可顯示")
        else:
            st.altair_chart(
                alt.Chart(action_sum)
                .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    x=alt.X(ANALYSIS_VALUE_COL, title="金額", axis=alt.Axis(format=",.0s", labelFontSize=10)),
                    y=alt.Y("訂閱動作:N", sort="-x", title="訂閱動作", axis=alt.Axis(labelFontSize=11)),
                    color=alt.Color("訂閱動作:N", legend=None, scale=alt.Scale(scheme="purples")),
                    tooltip=["訂閱動作", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
                )
                .properties(height=480),
                use_container_width=True,
            )




    st.divider()

    # ---------------------------
    # 12) 明細表（點選後顯示浮動視窗）
    # ---------------------------
    st.subheader("E) 今年度明細表格（點選資料列後顯示 Email）")
    show_filter_ranges_if_enabled(ui_state)
    st.caption("提示：勾選一筆明細列會顯示 Email 浮動視窗；取消勾選即可關閉。")

    if "detail_editor_version" not in st.session_state:
        st.session_state["detail_editor_version"] = 0

    # ✅ 今年度明細（維持既有勾選/浮動視窗行為）
    editor_key_this = f"detail_editor_{st.session_state['detail_editor_version']}"
    edited_this, selected_row_key = render_detail_table(df_filtered, editor_key_this, selectable=True)

    if st.session_state.get("skip_selection_once"):
        selected_row_key = None
        st.session_state["skip_selection_once"] = False

    # ✅ 明年度明細（複製一份相同表格：欄位與格示一樣；範圍為原篩選+1年；不影響既有功能）
    st.divider()
    st.subheader("E) 明年度明細表格（原篩選範圍 +1 年；若尚無明年度資料則視為 0）")
    show_filter_ranges_if_enabled(ui_state)
    st.caption("此表為今年度明細表格的複製版（欄位與格示一致），資料範圍為原篩選條件 +1 年。")

    # 使用獨立 key，且不提供勾選互動（避免干擾既有 overlay 行為）
    editor_key_next = f"detail_editor_next_{st.session_state['detail_editor_version']}"
    _edited_next, _ = render_detail_table(df_filtered_next, editor_key_next, selectable=False)

    # ==========================================================
    # ✅ 12) 明細表格下的兩個區域
    # ==========================================================
    st.divider()

    # ---------------------------
    # F) 年度總額（改為以「成交價未稅小計」分析 + 明年度(範圍+1年) + 差異）
    # ---------------------------
    st.subheader(f"F) 年度總額（以「{ANALYSIS_VALUE_COL}」分析）")
    show_filter_ranges_if_enabled(ui_state)

    this_total = float(pd.to_numeric(df_filtered[ANALYSIS_VALUE_COL], errors="coerce").fillna(0).sum()) if (df_filtered is not None and not df_filtered.empty) else 0.0
    next_total = float(pd.to_numeric(df_filtered_next[ANALYSIS_VALUE_COL], errors="coerce").fillna(0).sum()) if (df_filtered_next is not None and not df_filtered_next.empty) else 0.0
    diff_total = next_total - this_total

    a1, a2, a3 = st.columns(3)
    a1.metric("今年度總額", format_money(this_total))
    a2.metric("明年度總額（原篩選範圍 +1 年）", format_money(next_total))
    a3.metric("年度總額差異（明年 - 今年）", format_signed_money(diff_total))

    st.divider()

    # ---------------------------
    # 12-B) 分組明細表（大綱模式）
    # Group By：最終客戶 > 訂閱到期日年度 > 訂閱動作 > 訂閱到期日
    # ✅ 12-B 今年度標題紫色字體
    # ✅ 兩份分組明細表均依「訂閱總金額」由高至低排序（已於 build_grouped_detail_report_v2 實作）
    # ---------------------------
    st.markdown(
        f'''<div style="
            background: linear-gradient(135deg, {PURPLE_COLOR}15 0%, {PURPLE_COLOR}08 100%);
            border-left: 4px solid {PURPLE_COLOR};
            border-radius: 0 8px 8px 0;
            padding: 0.6rem 1rem;
            margin: 0.5rem 0;
            font-size: 1.1rem;
            font-weight: 800;
            color: {PURPLE_COLOR};
            font-family: 'Noto Sans TC', sans-serif;
            letter-spacing: -0.2px;
        ">F) 今年度分組明細表（Grouped Detail Report）</div>''',
        unsafe_allow_html=True,
    )
    show_filter_ranges_if_enabled(ui_state)
    st.caption("大綱模式：每個【最終客戶 + 年度】先顯示一列 Header（客戶/年度/訂閱總金額），下方列出訂閱動作/到期日/單筆金額等明細。")

    # ----------------------------------------------------------
    # ✅ 依「Header（最終客戶 + 到期年度）」產生 Email 範例（仿照 12) 明細表格）
    # ✅ 新增：電話訪談勾選欄位 → 右下方浮動視窗顯示口語化訪談腳本
    # ----------------------------------------------------------
    st.caption("操作：勾選下表任一 Header（最終客戶 + 到期年度）即可產生 Email 範例或電話訪談腳本；「未續約示警」除原本 15/30/45/60/90 天提醒外，若明年度已有續約資訊則顯示「已續約」，若明年度尚無續約且已超過到期日但未超過 30 天則顯示「寬限期」，超過 30 天則顯示「已到期」；取消勾選即可關閉視窗。")


    def _first_non_empty(series: pd.Series) -> str:
        if series is None:
            return ""
        s = series.dropna().astype(str)
        s = s[s.str.strip() != ""]
        return s.iloc[0] if not s.empty else ""


    headers_this = (
        df_filtered.groupby(["最終客戶", EXPIRY_YEAR_COL], dropna=False)
        .agg(
            訂閱總金額=(ANALYSIS_VALUE_COL, "sum"),
            最近到期日=("訂閱到期日", "min"),
            客戶微軟網域=("客戶微軟網域", _first_non_empty),
            經銷商=("經銷商", _first_non_empty),
            展碁業務=("展碁業務", _first_non_empty),
        )
        .reset_index()
    )

    if not headers_this.empty:
        headers_view = headers_this.copy()
        headers_view["訂閱總金額"] = pd.to_numeric(headers_view["訂閱總金額"], errors="coerce").fillna(0)
        _header_renewal_lookup = build_group_renewal_lookup(df_filtered_next)

        def _compute_header_warning_meta(row: pd.Series) -> pd.Series:
            warning_text, warning_color, warning_threshold = get_group_warning_meta_with_renewal(
                min_expiry_dt=row.get("最近到期日", pd.NaT),
                current_total=row.get("訂閱總金額", 0),
                customer=row.get("最終客戶", ""),
                expiry_year=row.get(EXPIRY_YEAR_COL, pd.NA),
                renewal_lookup=_header_renewal_lookup,
            )
            return pd.Series(
                {
                    "未續約示警": warning_text,
                    "_warning_color": warning_color,
                    "_warning_threshold": warning_threshold,
                    "未續約示警顯示": format_warning_display_text(warning_text, warning_color, warning_threshold),
                }
            )

        headers_view = pd.concat([headers_view, headers_view.apply(_compute_header_warning_meta, axis=1)], axis=1)
        headers_view = headers_view.drop(columns=["最近到期日"], errors="ignore")

        ordered_cols = [
            "未續約示警顯示",
            "最終客戶",
            EXPIRY_YEAR_COL,
            "訂閱總金額",
            "客戶微軟網域",
            "經銷商",
            "展碁業務",
            "未續約示警",
            "_warning_color",
            "_warning_threshold",
        ]
        headers_view = headers_view[[c for c in ordered_cols if c in headers_view.columns]]

        # ✅ Header 清單也依訂閱總金額排序（與分組表一致）
        headers_view = headers_view.sort_values("訂閱總金額", ascending=False).reset_index(drop=True)

        # ✅ 套用「未續約示警篩選」（sidebar 多選）
        _warning_filter = st.session_state.get("warning_filter_pick", [])
        if _warning_filter and isinstance(_warning_filter, list) and len(_warning_filter) > 0 and "未續約示警顯示" in headers_view.columns:
            _wf_stripped = [w.strip() for w in _warning_filter]
            headers_view = headers_view[
                headers_view["未續約示警顯示"].astype(str).str.strip().isin(_wf_stripped)
            ].reset_index(drop=True)

        # ✅ 兩個勾選欄位：Email（原本）+ 電話訪談（新增）
        headers_view.insert(0, "選取", False)
        headers_view.insert(1, "電話訪談", False)

        header_editor_key = f"group_header_editor_this_{st.session_state.get('detail_editor_version', 0)}"
        edited_headers = st.data_editor(
            headers_view,
            use_container_width=True,
            hide_index=True,
            # 僅開放「選取」與「電話訪談」兩欄可勾選，其他欄位維持不可編輯（避免影響既有 UI/行為）
            disabled=[c for c in headers_view.columns if c not in {"選取", "電話訪談"}],
            column_config={
                "選取": st.column_config.CheckboxColumn(help="勾選一筆 Header 以產生 Email", default=False),
                "電話訪談": st.column_config.CheckboxColumn(help="勾選一筆 Header 以產生電話訪談口語腳本", default=False),
                "未續約示警顯示": st.column_config.TextColumn("未續約示警", help="🟢 已續約：明年度金額 <= 今年度；🔴 已續約：明年度金額 > 今年度；🟤 寬限期：已過到期日但未超過 30 天；其餘為天數提醒"),
                "未續約示警": None,
                "_warning_color": None,
                "_warning_threshold": None,
            },
            key=header_editor_key,
        )

        # ---------------------------
        # 勾選判斷：優先顯示「電話訪談」視窗（避免同一時間多視窗重疊）
        # ---------------------------
        selected_customer = None
        selected_year = None
        selected_mode = None  # "CALL" | "EMAIL"

        if edited_headers is not None:
            sel_call = edited_headers.loc[edited_headers["電話訪談"] == True] if "電話訪談" in edited_headers.columns else pd.DataFrame()
            if sel_call is not None and not sel_call.empty:
                selected_customer = sel_call.iloc[0].get("最終客戶")
                selected_year = sel_call.iloc[0].get(EXPIRY_YEAR_COL)
                selected_mode = "CALL"
            else:
                sel_email = edited_headers.loc[edited_headers["選取"] == True] if "選取" in edited_headers.columns else pd.DataFrame()
                if sel_email is not None and not sel_email.empty:
                    selected_customer = sel_email.iloc[0].get("最終客戶")
                    selected_year = sel_email.iloc[0].get(EXPIRY_YEAR_COL)
                    selected_mode = "EMAIL"

        if selected_customer is not None and selected_year is not None and selected_mode == "EMAIL":
            sig = (str(selected_customer), str(selected_year))
            if st.session_state.get("group_last_selected_sig") != sig:
                txt, subj, mlink = build_group_email_from_header(df_filtered, selected_customer, selected_year)
                st.session_state["group_email_text"] = txt
                st.session_state["group_email_subject"] = subj
                st.session_state["group_mailto_link"] = mlink
                st.session_state["group_last_selected_sig"] = sig
            st.session_state["group_selected_sig"] = sig

            # ✅ 清掉電話訪談狀態（避免兩種視窗互相干擾）
            st.session_state["group_call_selected_sig"] = None
            st.session_state["group_call_text"] = ""
            st.session_state["group_call_last_selected_sig"] = None

        elif selected_customer is not None and selected_year is not None and selected_mode == "CALL":
            sig = (str(selected_customer), str(selected_year))
            if st.session_state.get("group_call_last_selected_sig") != sig:
                st.session_state["group_call_text"] = build_group_call_script_from_header(df_filtered, selected_customer, selected_year)
                st.session_state["group_call_last_selected_sig"] = sig
            st.session_state["group_call_selected_sig"] = sig

            # ✅ 清掉 Email 狀態（避免兩種視窗互相干擾）
            st.session_state["group_selected_sig"] = None
            st.session_state["group_email_text"] = ""
            st.session_state["group_email_subject"] = ""
            st.session_state["group_mailto_link"] = ""
            st.session_state["group_last_selected_sig"] = None

        else:
            # 都沒勾選：兩種視窗都關閉
            st.session_state["group_selected_sig"] = None
            st.session_state["group_email_text"] = ""
            st.session_state["group_email_subject"] = ""
            st.session_state["group_mailto_link"] = ""
            st.session_state["group_last_selected_sig"] = None

            st.session_state["group_call_selected_sig"] = None
            st.session_state["group_call_text"] = ""
            st.session_state["group_call_last_selected_sig"] = None
    else:
        st.info("無 Header 可選（此區塊視為 0）")

    # 仍保留原本的大綱分組明細表（顯示用）
    render_grouped_table(df_filtered, df_next=df_filtered_next)

    st.divider()

    st.subheader("F) 明年度分組明細表（原篩選範圍 +1 年；若尚無明年度資料則視為 0）")
    show_filter_ranges_if_enabled(ui_state)
    st.caption("此表為今年度分組明細表的複製版（欄位與格示一致），資料範圍為原篩選條件 +1 年。")
    render_grouped_table(df_filtered_next)

    # ---------------------------
    # 浮動視窗（純 HTML overlay，不占位，不推擠）
    # ---------------------------
    st.markdown(
        """
        <style>
        html, body { overflow-x: hidden !important; max-width: 100% !important; }
        .stApp, .main, section.main, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
            overflow-x: hidden !important;
            max-width: 100% !important;
        }
        pre, code, .stMarkdown pre, .stMarkdown code {
            white-space: pre-wrap !important;
            overflow-wrap: anywhere !important;
            word-break: break-word !important;
            overflow-x: hidden !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
        }

        /* 12) 明細 Email overlay（原本既有：右下） */
        #emailOverlay {
            position: fixed;
            right: 24px;
            bottom: 24px;
            width: min(900px, 48vw, calc(100vw - 48px));
            max-width: calc(100vw - 48px);
            min-width: 0;
            height: 72vh;
            z-index: 2147483000;
            background: rgba(255,255,255,0.97);
            border: none;
            box-shadow: 0 24px 64px rgba(0,0,0,0.18), 0 0 0 1px rgba(21,101,192,0.1);
            border-radius: 20px;
            padding: 16px;
            backdrop-filter: blur(16px) saturate(1.6);
            box-sizing: border-box;
            display: none;
            flex-direction: column;
            gap: 10px;
            overflow-y: hidden !important;
            overflow-x: hidden !important;
        }
        #emailOverlay .hdr {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            flex: 0 0 auto;
            max-width: 100% !important;
            overflow-x: hidden !important;
        }
        #emailOverlay .title { font-size: 16px; font-weight: 800; line-height: 1.1; color: #1565c0; }
        #emailOverlay .sub { opacity: .65; font-size: 12px; margin-top: 2px; }

        #emailOverlay .body {
            flex: 1 1 auto;
            min-height: 0;
            overflow-y: auto !important;
            overflow-x: hidden !important;
            background: rgba(240,245,255,0.9);
            border-radius: 12px;
            border: 1px solid rgba(21,101,192,0.1);
            padding: 10px 12px 18px 12px;
            box-sizing: border-box;
            max-width: 100% !important;
        }
        #emailOverlay .emailText {
            white-space: pre-wrap !important;
            overflow-wrap: anywhere !important;
            word-break: break-word !important;
            overflow-x: hidden !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace !important;
            font-size: 13px !important;
            line-height: 1.55 !important;
        }
        #emailOverlay .emailText * {
            overflow-wrap: anywhere !important;
            word-break: break-word !important;
            max-width: 100% !important;
        }

        /* ✅ 12-B Header 彙整 Email overlay（右下、同 12) 格式、有 Email 按鈕） */
        #groupEmailOverlay {
            position: fixed;
            right: 24px;
            bottom: 24px;
            width: min(900px, 48vw, calc(100vw - 48px));
            max-width: calc(100vw - 48px);
            min-width: 0;
            height: 72vh;
            z-index: 2147483001;
            background: rgba(255,255,255,0.97);
            border: none;
            box-shadow: 0 24px 64px rgba(0,0,0,0.18), 0 0 0 1px rgba(106,13,173,0.1);
            border-radius: 20px;
            padding: 16px;
            backdrop-filter: blur(16px) saturate(1.6);
            box-sizing: border-box;
            display: none;
            flex-direction: column;
            gap: 10px;
            overflow-y: hidden !important;
            overflow-x: hidden !important;
        }
        #groupEmailOverlay .hdr {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            flex: 0 0 auto;
            max-width: 100% !important;
            overflow-x: hidden !important;
        }
        #groupEmailOverlay .title { font-size: 16px; font-weight: 800; line-height: 1.1; color: #6A0DAD; }
        #groupEmailOverlay .sub { opacity: .65; font-size: 12px; margin-top: 2px; }

        #groupEmailOverlay .body {
            flex: 1 1 auto;
            min-height: 0;
            overflow-y: auto !important;
            overflow-x: hidden !important;
            background: rgba(240,245,255,0.9);
            border-radius: 12px;
            border: 1px solid rgba(106,13,173,0.1);
            padding: 10px 12px 18px 12px;
            box-sizing: border-box;
            max-width: 100% !important;
        }
        #groupEmailOverlay .emailText {
            white-space: pre-wrap !important;
            overflow-wrap: anywhere !important;
            word-break: break-word !important;
            overflow-x: hidden !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace !important;
            font-size: 13px !important;
            line-height: 1.55 !important;
        }

        /* ✅ 12-B Header 電話訪談 overlay（新：右下、同 12) 格式、有 Y 軸捲軸） */
        #groupCallOverlay {
            position: fixed;
            right: 24px;
            bottom: 24px;
            width: min(900px, 48vw, calc(100vw - 48px));
            max-width: calc(100vw - 48px);
            min-width: 0;
            height: 72vh;
            z-index: 2147483002;
            background: rgba(255,255,255,0.98);
            border: 1px solid rgba(0,0,0,0.12);
            box-shadow: 0 14px 38px rgba(0,0,0,0.22);
            border-radius: 16px;
            padding: 14px;
            backdrop-filter: blur(6px);
            box-sizing: border-box;
            display: none;
            flex-direction: column;
            gap: 10px;
            overflow-y: hidden !important;
            overflow-x: hidden !important;
        }
        #groupCallOverlay .hdr {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            flex: 0 0 auto;
            max-width: 100% !important;
            overflow-x: hidden !important;
        }
        #groupCallOverlay .title { font-size: 16px; font-weight: 800; line-height: 1.1; }
        #groupCallOverlay .sub { opacity: .7; font-size: 12px; margin-top: 2px; }

        #groupCallOverlay .body {
            flex: 1 1 auto;
            min-height: 0;
            overflow-y: auto !important;
            overflow-x: hidden !important;
            background: rgba(245,246,248,0.9);
            border-radius: 12px;
            border: 1px solid rgba(0,0,0,0.08);
            padding: 10px 12px 18px 12px;
            box-sizing: border-box;
            max-width: 100% !important;
        }
        #groupCallOverlay .emailText {
            white-space: pre-wrap !important;
            overflow-wrap: anywhere !important;
            word-break: break-word !important;
            overflow-x: hidden !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace !important;
            font-size: 13px !important;
            line-height: 1.55 !important;
        }

        .overlayBtn {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 7px 14px;
            border-radius: 8px;
            border: 1px solid rgba(21,101,192,0.25);
            background: linear-gradient(135deg, #1565c0, #1976d2);
            font-weight: 700;
            font-size: 13px;
            text-decoration: none;
            color: white !important;
            white-space: nowrap;
            box-shadow: 0 2px 8px rgba(21,101,192,0.3);
            transition: all 0.2s ease;
            cursor: pointer;
        }
        .overlayBtn:hover {
            background: linear-gradient(135deg, #1976d2, #1e88e5);
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(21,101,192,0.4);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ✅ overlay 維持既有：僅綁定「今年度明細表格」的勾選結果
    if selected_row_key is not None and selected_row_key in df_filtered.index:
        row = df_filtered.loc[selected_row_key]
        sig = (selected_row_key, str(row.get("最終客戶", "")), str(row.get("商品名稱", "")), str(row.get("訂閱到期日", "")))
        if st.session_state.get("last_selected_sig") != sig:
            st.session_state["email_text"] = build_email_from_row(row)
            st.session_state["last_selected_sig"] = sig
        st.session_state["selected_row_key"] = selected_row_key
    else:
        st.session_state["selected_row_key"] = None
        st.session_state["email_text"] = ""
        st.session_state["last_selected_sig"] = None

    email_text = st.session_state.get("email_text", "")
    overlay_visible = st.session_state.get("selected_row_key") is not None and bool(email_text)

    group_email_text = st.session_state.get("group_email_text", "")
    group_mailto_link = st.session_state.get("group_mailto_link", "")
    group_overlay_visible = st.session_state.get("group_selected_sig") is not None and bool(group_email_text)

    group_call_text = st.session_state.get("group_call_text", "")
    group_call_overlay_visible = st.session_state.get("group_call_selected_sig") is not None and bool(group_call_text)

    # ✅ 若 12-B 任一 Header 視窗顯示中，為避免多個右下視窗重疊：暫時隱藏 12) 明細 overlay（不改變既有勾選邏輯）
    any_group_overlay_visible = group_overlay_visible or group_call_overlay_visible
    effective_detail_overlay_visible = overlay_visible and (not any_group_overlay_visible)

    st.markdown(
        f"""
        <div id="emailOverlay" style="display:{'flex' if effective_detail_overlay_visible else 'none'};">
            <div class="hdr">
                <div>
                    <div class="title">✉️ 續約提醒與建議加購產品 Email 範例</div>
                    <div class="sub">（可直接複製貼上；取消勾選明細列即可關閉此視窗）</div>
                </div>
            </div>
            <div class="body"><div class="emailText">{_html_escape(email_text)}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.caption("✅ 勾選明細列＝顯示 Email；取消勾選＝關閉。")

    # ---------------------------
    # ✅ 12-B) 分組 Header Email 視窗（右下、同 12) 格式、有 Y 軸捲軸、有 Email 按鈕）
    # ---------------------------
    st.markdown(
        f"""
        <div id="groupEmailOverlay" style="display:{'flex' if group_overlay_visible else 'none'};">
            <div class="hdr">
                <div>
                    <div class="title">✉️ 續約提醒與建議加購產品 Email 範例</div>
                    <div class="sub">（依 Header 彙整；可直接複製貼上；取消勾選 Header 即可關閉。收件人請於 Outlook 中填寫）</div>
                </div>
                <div style="display:flex; gap:8px; align-items:center;">
                    <a class="overlayBtn" href="{group_mailto_link}">Email</a>
                </div>
            </div>
            <div class="body"><div class="emailText">{_html_escape(group_email_text)}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---------------------------
    # ✅ 12-B) 分組 Header 電話訪談 視窗（新：右下、同 12) 格式、有 Y 軸捲軸）
    # ---------------------------
    st.markdown(
        f"""
        <div id="groupCallOverlay" style="display:{'flex' if group_call_overlay_visible else 'none'};">
            <div class="hdr">
                <div>
                    <div class="title">📞 電話訪談內容範例</div>
                    <div class="sub">（依 Header 彙整；可直接照念或複製改寫；取消勾選「電話訪談」即可關閉）</div>
                </div>
            </div>
            <div class="body"><div class="emailText">{_html_escape(group_call_text)}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with tab_v2:
    render_csp_dashboard_v2_architecture(df_filtered, df_filtered_next, ui_state)
