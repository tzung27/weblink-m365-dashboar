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
    if days_left < 0:
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


def show_filter_ranges_if_enabled(ui_state: dict):
    """每個區塊標題下方顯示（若 sidebar checkbox 勾選）。"""
    if not st.session_state.get("show_filter_ranges", False):
        return
    this_rng, next_rng = _get_effective_ranges(ui_state)
    st.caption(
        f"篩選時間範圍｜本年度：{_fmt_d(this_rng[0])} ~ {_fmt_d(this_rng[1])}　｜　明年度：{_fmt_d(next_rng[0])} ~ {_fmt_d(next_rng[1])}"
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
st.sidebar.title("⚙️ 篩選與匯入")

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
st.title("📊 Weblink M365 續約精準行銷｜企業儀表板")

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

    dealer_opts = uniq_options(df, "經銷商")
    qual_opts = uniq_options(df, "資格")
    customer_opts = uniq_options(df, "最終客戶")
    action_opts = uniq_options(df, "訂閱動作")
    product_opts = uniq_options(df, "商品名稱")
    sales_opts = uniq_options(df, "展碁業務")

    st.sidebar.divider()
    st.sidebar.subheader("8) 多條件篩選（可多選）")
    dealer_sel = st.sidebar.multiselect("經銷商", dealer_opts, default=st.session_state.get("dealer", []), key="dealer")
    qual_sel = st.sidebar.multiselect("資格", qual_opts, default=st.session_state.get("qual", []), key="qual")



    # ---------------------------
    # ✅ Top 30 最終客戶下拉選單（依「篩選範圍」金額由高至低）
    # 規則：
    # - 會套用：未來N月 / 訂閱到期日 / 訂單下單日 + 其他多選（經銷商/資格/訂閱動作/商品名稱/展碁業務）
    # - 但「不套用最終客戶」本身，避免排名被自己限制住
    # - 選到某客戶時：自動把 customer multiselect 設為該客戶（避免 session_state 錯誤用 on_change）
    # ---------------------------
    def _compute_top30_customers_by_filtered_scope(df_all: pd.DataFrame) -> list[str]:
        if df_all is None or df_all.empty:
            return []

        # 以目前 sidebar 的 session_state 組出「用於 Top30 計算」的篩選條件
        # ✅ 不套用 customer（最終客戶）本身
        ui_state_for_top30 = {
            "future_expiry_enabled": bool(st.session_state.get("future_expiry_enabled", False)),
            "future_expiry_months": int(st.session_state.get("future_expiry_months", 3)),
            "expiry_range": (
                st.session_state.get("expiry_from", None),
                st.session_state.get("expiry_to", None),
            ),
            "order_range": (
                st.session_state.get("order_from", None),
                st.session_state.get("order_to", None),
            ),
            "dealer": st.session_state.get("dealer", []) or [],
            "qual": st.session_state.get("qual", []) or [],
            "customer": [],  # ✅ 關鍵：Top30 排名不受 customer 已選影響
            "action": st.session_state.get("action", []) or [],
            "product": st.session_state.get("product", []) or [],
            "sales": st.session_state.get("sales", []) or [],
        }

        # 套用篩選（依你既有 apply_filters 邏輯）
        df_scope = apply_filters(df_all, ui_state_for_top30)

        if df_scope is None or df_scope.empty or "最終客戶" not in df_scope.columns or ANALYSIS_VALUE_COL not in df_scope.columns:
            return []

        g = (
            df_scope.groupby("最終客戶", dropna=False)[ANALYSIS_VALUE_COL]
            .sum()
            .reset_index()
        )
        g["最終客戶"] = g["最終客戶"].astype(str).replace({"nan": "未填"})
        g[ANALYSIS_VALUE_COL] = pd.to_numeric(g[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)
        g = g.sort_values(ANALYSIS_VALUE_COL, ascending=False, kind="mergesort").head(30)

        return g["最終客戶"].tolist()

    top30_customers = _compute_top30_customers_by_filtered_scope(df)

    # 下拉顯示文字：01｜xxx
    top30_cust_labels = ["（不套用 Top 30）"]
    top30_cust_label_to_value = {"（不套用 Top 30）": None}
    for i, name in enumerate(top30_customers, 1):
        label = f"{i:02d}｜{name}"
        top30_cust_labels.append(label)
        top30_cust_label_to_value[label] = name

    st.session_state.setdefault("top30_customer_pick", "（不套用 Top 30）")

    def _on_pick_top30_customer():
        picked_label = st.session_state.get("top30_customer_pick", "（不套用 Top 30）")
        picked_customer = top30_cust_label_to_value.get(picked_label)

        if not picked_customer:
            return

        # ✅ 套用：覆蓋「最終客戶」多選為單一客戶（最穩定、不干擾既有邏輯）
        st.session_state["customer"] = [picked_customer]

    st.sidebar.selectbox(
        "Top 30 最終客戶（依篩選範圍金額排序）",
        options=top30_cust_labels,
        key="top30_customer_pick",
        on_change=_on_pick_top30_customer,
        help=f"依 {ANALYSIS_VALUE_COL} 加總，並套用目前篩選範圍與其他條件（不含最終客戶本身）計算。",
    )







    customer_sel = st.sidebar.multiselect("最終客戶", customer_opts, default=st.session_state.get("customer", []), key="customer")
    action_sel = st.sidebar.multiselect("訂閱動作", action_opts, default=st.session_state.get("action", []), key="action")
    # ✅ 商品名稱篩選（可多選）
    product_sel = st.sidebar.multiselect("商品名稱", product_opts, default=st.session_state.get("product", []), key="product")
    sales_sel = st.sidebar.multiselect("展碁業務", sales_opts, default=st.session_state.get("sales", []), key="sales")

    progress.progress(90)
    status.update(label="套用篩選…", state="running", expanded=False)

    ui_state = {
        "future_expiry_enabled": bool(future_expiry_enabled),
        "future_expiry_months": int(future_expiry_months),
        "expiry_range": (expiry_from, expiry_to),
        "order_range": (order_from, order_to),
        "dealer": dealer_sel,
        "qual": qual_sel,
        "customer": customer_sel,
        "action": action_sel,
        "product": product_sel,
        "sales": sales_sel,
    }

    # 今年度（原篩選）
    df_filtered = apply_filters(df, ui_state)

    # 明年度（原篩選範圍 +1 年；若無資料視為 0）
    ui_next = shift_ui_state_one_year(ui_state)
    base_today_next = date.today() + relativedelta(years=1)
    df_filtered_next = apply_filters(df, ui_next, base_today=base_today_next)

    progress.progress(100)
    status.update(label=f"完成 ✅ 目前顯示 {len(df_filtered):,} 筆資料", state="complete", expanded=False)

except FileNotFoundError:
    status.update(label=f"找不到 Excel：請上傳檔案或確認路徑存在：{DEFAULT_XLSX_PATH}", state="error", expanded=True)
    st.stop()
except Exception as e:
    status.update(label=f"處理失敗：{e}", state="error", expanded=True)
    st.stop()

# ---------------------------
# 9) KPI（本年度 + 明年度 + 差異）
# ---------------------------
st.subheader(f"9) KPI 指標（以「{ANALYSIS_VALUE_COL}」分析）")
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
    f"{ANALYSIS_VALUE_COL}平均每筆": kpis_next[f"{ANALYSIS_VALUE_COL}平均每筆"] - kpis_this[f"{ANALYSIS_VALUE_COL}平均每筆"],
    f"{ANALYSIS_VALUE_COL}平均每客戶": kpis_next[f"{ANALYSIS_VALUE_COL}平均每客戶"] - kpis_this[f"{ANALYSIS_VALUE_COL}平均每客戶"],
}

st.caption("今年度（原篩選）")
k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("筆數", f"{kpis_this['筆數']:,}")
k2.metric("最終客戶數", f"{kpis_this['最終客戶數']:,}")
k3.metric("經銷商數", f"{kpis_this['經銷商數']:,}")
k4.metric(f"{ANALYSIS_VALUE_COL}合計", format_money(kpis_this[f"{ANALYSIS_VALUE_COL}合計"]))
k5.metric(f"{ANALYSIS_VALUE_COL}平均每筆", format_money(kpis_this[f"{ANALYSIS_VALUE_COL}平均每筆"]))
k6.metric(f"{ANALYSIS_VALUE_COL}平均每客戶", format_money(kpis_this[f"{ANALYSIS_VALUE_COL}平均每客戶"]))

st.caption("明年度（原篩選範圍 +1 年；若無資料視為 0）")
n1, n2, n3, n4, n5, n6 = st.columns(6)
n1.metric("筆數", f"{kpis_next['筆數']:,}")
n2.metric("最終客戶數", f"{kpis_next['最終客戶數']:,}")
n3.metric("經銷商數", f"{kpis_next['經銷商數']:,}")
n4.metric(f"{ANALYSIS_VALUE_COL}合計", format_money(kpis_next[f"{ANALYSIS_VALUE_COL}合計"]))
n5.metric(f"{ANALYSIS_VALUE_COL}平均每筆", format_money(kpis_next[f"{ANALYSIS_VALUE_COL}平均每筆"]))
n6.metric(f"{ANALYSIS_VALUE_COL}平均每客戶", format_money(kpis_next[f"{ANALYSIS_VALUE_COL}平均每客戶"]))

# ✅ 差異（明年度 - 今年度）數值以棕色呈現（你的需求）
st.caption("差異（明年度 - 今年度）")


def _kpi_diff_card(label: str, value_str: str):
    st.markdown(
        f"""
        <div style="padding: 6px 2px;">
            <div style="font-size: 0.85rem; opacity: 0.75; line-height: 1.1;">{label}</div>
            <div style="font-size: 1.55rem; font-weight: 800; color: {BROWN_COLOR}; line-height: 1.2;">{value_str}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


d1, d2, d3, d4, d5, d6 = st.columns(6)
with d1:
    _kpi_diff_card("筆數", format_signed_int(kpis_diff["筆數"]))
with d2:
    _kpi_diff_card("最終客戶數", format_signed_int(kpis_diff["最終客戶數"]))
with d3:
    _kpi_diff_card("經銷商數", format_signed_int(kpis_diff["經銷商數"]))
with d4:
    _kpi_diff_card(f"{ANALYSIS_VALUE_COL}合計", format_signed_money(kpis_diff[f"{ANALYSIS_VALUE_COL}合計"]))
with d5:
    _kpi_diff_card(f"{ANALYSIS_VALUE_COL}平均每筆", format_signed_money(kpis_diff[f"{ANALYSIS_VALUE_COL}平均每筆"]))
with d6:
    _kpi_diff_card(f"{ANALYSIS_VALUE_COL}平均每客戶", format_signed_money(kpis_diff[f"{ANALYSIS_VALUE_COL}平均每客戶"]))

st.divider()

# ---------------------------
# 10) Top 10（續約客戶 / 經銷商）
# ---------------------------
st.subheader(f"10) Top 10 續約客戶（以 {ANALYSIS_VALUE_COL} 合計排序）")
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

# ✅ 差異金額欄位名稱棕色（表頭）
cust_styler = top10_customer.style.format(
    {
        "需約金額": _safe_thousands_formatter,
        "明年度續約金額": _safe_thousands_formatter,
        "差異金額": _safe_thousands_formatter,
    }
).set_table_styles(_style_header_color_for_column(top10_customer, "差異金額", BROWN_COLOR), overwrite=False)

st.dataframe(cust_styler, use_container_width=True, hide_index=True)

st.subheader(f"10) Top 10 經銷商（以 {ANALYSIS_VALUE_COL} 合計排序）")
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
st.subheader(f"11) 圖表（以「{ANALYSIS_VALUE_COL}」分析）")
show_filter_ranges_if_enabled(ui_state)

c1, c3 = st.columns(2)

with c1:
    st.caption("Bar：Top 10 續約客戶（需約金額）")
    if top10_customer.empty:
        st.info("無資料可顯示")
    else:
        tmp = top10_customer[["最終客戶", "需約金額"]].rename(columns={"需約金額": ANALYSIS_VALUE_COL}).copy()
        tmp[ANALYSIS_VALUE_COL] = pd.to_numeric(tmp[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)
        st.altair_chart(
            alt.Chart(tmp)
            .mark_bar()
            .encode(
                x=alt.X(ANALYSIS_VALUE_COL, title="金額"),
                y=alt.Y("最終客戶:N", sort="-x", title="最終客戶"),
                color=alt.Color("最終客戶:N", legend=None),
                tooltip=["最終客戶", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
            )
            .properties(height=320),
            use_container_width=True,
        )


with c3:
    st.caption("Line：到期月份金額趨勢（今年度 vs 明年度）")

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
            .properties(height=160)
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
            .properties(height=160)
        )

        final_chart = alt.vconcat(line_chart, bar_chart).resolve_scale(color="independent")
        st.altair_chart(final_chart, use_container_width=True)




c4, c5 = st.columns([1, 2])
# -------------------------------------------------
# 4 🆕 新增：Top 15 商品名稱（金額）
# -------------------------------------------------
with c4:
    st.caption("Bar：Top 15 商品名稱（金額）")

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
            .mark_bar()
            .encode(
                x=alt.X(ANALYSIS_VALUE_COL, title="金額"),
                y=alt.Y("商品名稱:N", sort="-x"),
                tooltip=["商品名稱", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
            )
            .properties(height=320),
            use_container_width=True,
        )
with c5:
    st.caption("商品名稱：Top 15（金額以及數量之增減｜今年度 vs 明年度）")

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
            .properties(height=150)
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
            .properties(height=150)
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
            .properties(height=150)
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
            .properties(height=150)
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
    st.caption("業務人員分析：各展碁業務金額")
    sales_sum = df_filtered.groupby("展碁業務", dropna=False)[ANALYSIS_VALUE_COL].sum().reset_index()
    sales_sum["展碁業務"] = sales_sum["展碁業務"].astype(str).replace({"nan": "未填"})
    sales_sum[ANALYSIS_VALUE_COL] = pd.to_numeric(sales_sum[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)
    sales_sum = sales_sum.sort_values(ANALYSIS_VALUE_COL, ascending=False).head(20)
    if sales_sum.empty:
        st.info("無資料可顯示")
    else:
        st.altair_chart(
            alt.Chart(sales_sum)
            .mark_bar()
            .encode(
                x=alt.X(ANALYSIS_VALUE_COL, title="金額"),
                y=alt.Y("展碁業務:N", sort="-x", title="展碁業務"),
                color=alt.Color("展碁業務:N", legend=None),
                tooltip=["展碁業務", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
            )
            .properties(height=420),
            use_container_width=True,
        )

with c8:
    st.caption("訂閱動作分析：各訂閱動作金額")
    action_sum = df_filtered.groupby("訂閱動作", dropna=False)[ANALYSIS_VALUE_COL].sum().reset_index()
    action_sum["訂閱動作"] = action_sum["訂閱動作"].astype(str).replace({"nan": "未填"})
    action_sum[ANALYSIS_VALUE_COL] = pd.to_numeric(action_sum[ANALYSIS_VALUE_COL], errors="coerce").fillna(0)
    action_sum = action_sum.sort_values(ANALYSIS_VALUE_COL, ascending=False)
    if action_sum.empty:
        st.info("無資料可顯示")
    else:
        st.altair_chart(
            alt.Chart(action_sum)
            .mark_bar()
            .encode(
                x=alt.X(ANALYSIS_VALUE_COL, title="金額"),
                y=alt.Y("訂閱動作:N", sort="-x", title="訂閱動作"),
                color=alt.Color("訂閱動作:N", legend=None),
                tooltip=["訂閱動作", alt.Tooltip(ANALYSIS_VALUE_COL, format=",.0f")],
            )
            .properties(height=420),
            use_container_width=True,
        )




st.divider()

# ---------------------------
# 12) 明細表（點選後顯示浮動視窗）
# ---------------------------
st.subheader("12) 今年度明細表格（點選資料列後顯示 Email）")
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
st.subheader("12) 明年度明細表格（原篩選範圍 +1 年；若尚無明年度資料則視為 0）")
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
# 12-A) 年度總額（改為以「成交價未稅小計」分析 + 明年度(範圍+1年) + 差異）
# ---------------------------
st.subheader(f"12-A) 年度總額（以「{ANALYSIS_VALUE_COL}」分析）")
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
    f'<div style="font-size: 1.35rem; font-weight: 800; color:{PURPLE_COLOR};">12-B) 今年度分組明細表（Grouped Detail Report）</div>',
    unsafe_allow_html=True,
)
show_filter_ranges_if_enabled(ui_state)
st.caption("大綱模式：每個【最終客戶 + 年度】先顯示一列 Header（客戶/年度/訂閱總金額），下方列出訂閱動作/到期日/單筆金額等明細。")

# ----------------------------------------------------------
# ✅ 依「Header（最終客戶 + 到期年度）」產生 Email 範例（仿照 12) 明細表格）
# ✅ 新增：電話訪談勾選欄位 → 右下方浮動視窗顯示口語化訪談腳本
# ----------------------------------------------------------
st.caption("操作：勾選下表任一 Header（最終客戶 + 到期年度）即可產生 Email 範例或電話訪談腳本；「未續約示警」除原本 15/30/45/60/90 天提醒外，若明年度已有續約資訊則顯示「已續約」，且明年度金額 <= 今年度為綠色、明年度金額 > 今年度為紅色；取消勾選即可關閉視窗。")


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
            "未續約示警顯示": st.column_config.TextColumn("未續約示警", help="🟢 已續約：明年度金額 <= 今年度；🔴 已續約：明年度金額 > 今年度；其餘為天數提醒"),
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

st.subheader("12-B) 明年度分組明細表（原篩選範圍 +1 年；若尚無明年度資料則視為 0）")
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
    #emailOverlay .hdr {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        flex: 0 0 auto;
        max-width: 100% !important;
        overflow-x: hidden !important;
    }
    #emailOverlay .title { font-size: 16px; font-weight: 800; line-height: 1.1; }
    #emailOverlay .sub { opacity: .7; font-size: 12px; margin-top: 2px; }

    #emailOverlay .body {
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
    #groupEmailOverlay .hdr {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        flex: 0 0 auto;
        max-width: 100% !important;
        overflow-x: hidden !important;
    }
    #groupEmailOverlay .title { font-size: 16px; font-weight: 800; line-height: 1.1; }
    #groupEmailOverlay .sub { opacity: .7; font-size: 12px; margin-top: 2px; }

    #groupEmailOverlay .body {
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
        display: inline-block;
        padding: 8px 12px;
        border-radius: 10px;
        border: 1px solid rgba(0,0,0,0.14);
        background: rgba(255,255,255,0.92);
        font-weight: 700;
        text-decoration: none;
        color: inherit;
        white-space: nowrap;
    }
    .overlayBtn:hover { background: rgba(245,246,248,0.95); }
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