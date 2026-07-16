
from __future__ import annotations

import io
from pathlib import Path
import re
from datetime import datetime
from typing import Iterable

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import requests
import urllib3
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt


st.set_page_config(
    page_title="MMS AI Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
:root {
    --bg: #f5f7fb;
    --surface: #ffffff;
    --surface-soft: #f8fafc;
    --border: #e4e8ef;
    --text: #1f2937;
    --muted: #6b7280;
    --primary: #2f6fec;
    --primary-soft: #eef4ff;
    --success: #2e8b57;
    --warning: #d69e00;
    --danger: #d64545;
    --shadow: 0 6px 20px rgba(25, 42, 70, 0.06);
}

html, body, [class*="css"] {
    font-family: "Pretendard", "Noto Sans KR", "Apple SD Gothic Neo", sans-serif;
    color: var(--text);
}

.stApp {
    background: var(--bg);
}

.block-container {
    max-width: 100%;
    padding: 2.4rem 1.4rem 3rem;
}

[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid var(--border);
}

[data-testid="stSidebar"] .block-container {
    padding-top: 1.5rem;
}

.app-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 20px;
    background: linear-gradient(135deg, #ffffff 0%, #f7faff 100%);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 22px 24px;
    box-shadow: var(--shadow);
    margin-bottom: 16px;
}

.app-title {
    font-size: 30px;
    font-weight: 800;
    letter-spacing: -0.7px;
    margin: 0;
    line-height: 1.35;
}

.app-subtitle {
    color: var(--muted);
    font-size: 14px;
    margin-top: 6px;
}

.status-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: var(--primary-soft);
    color: #2859bb;
    border: 1px solid #d8e5ff;
    border-radius: 999px;
    padding: 9px 14px;
    font-size: 13px;
    white-space: nowrap;
}

.section-title {
    font-size: 21px;
    font-weight: 800;
    letter-spacing: -0.4px;
    margin: 24px 0 12px;
    padding-left: 12px;
    border-left: 4px solid var(--primary);
}

.subsection-title {
    font-size: 17px;
    font-weight: 750;
    margin: 18px 0 10px;
}

.card-shell {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 18px;
    box-shadow: var(--shadow);
    margin-bottom: 14px;
}

.metric-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 16px 17px;
    min-height: 98px;
    box-shadow: 0 3px 12px rgba(25, 42, 70, 0.04);
    transition: transform .15s ease, box-shadow .15s ease;
}

.metric-card:hover {
    transform: translateY(-1px);
    box-shadow: 0 7px 18px rgba(25, 42, 70, 0.08);
}

.metric-label {
    font-size: 12px;
    color: var(--muted);
    font-weight: 650;
    margin-bottom: 8px;
}

.metric-value {
    font-size: 23px;
    font-weight: 800;
    letter-spacing: -0.5px;
    line-height: 1.2;
}

.metric-delta {
    font-size: 12px;
    margin-top: 8px;
    color: var(--muted);
}

.insight-box {
    background: linear-gradient(180deg, #f8fbff 0%, #f3f7fd 100%);
    border: 1px solid #dfe8f5;
    border-radius: 14px;
    padding: 18px 20px;
    line-height: 1.6;
    white-space: pre-wrap;
    box-shadow: 0 3px 12px rgba(25, 42, 70, 0.035);
}

.asset-card {
    height: 430px;
    border: 1px solid var(--border);
    border-radius: 14px;
    background: var(--surface);
    padding: 14px;
    box-sizing: border-box;
    overflow: hidden;
    box-shadow: 0 3px 12px rgba(25, 42, 70, 0.045);
}

.asset-image-card {
    display: flex;
    align-items: center;
    justify-content: center;
}

.asset-image-card img {
    width: 100%;
    height: 398px;
    object-fit: contain;
    border-radius: 10px;
}

.asset-message-card {
    height: 430px;
    overflow-y: auto;
    white-space: pre-wrap;
    font-size: 15px;
    line-height: 1.75;
    background: #fbfcfe;
}

.asset-empty {
    height: 398px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: var(--muted);
    text-align: center;
}

[data-testid="stDataFrame"] {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 3px 12px rgba(25, 42, 70, 0.04);
}

[data-testid="stDataFrame"] [role="columnheader"] {
    background: #f5f8fc !important;
    font-weight: 750 !important;
}

[data-testid="stSelectbox"] > div,
[data-testid="stDateInput"] > div,
[data-testid="stMultiSelect"] > div {
    background: var(--surface);
    border-radius: 10px;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background: #eef2f7;
    padding: 5px;
    border-radius: 12px;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 9px;
    padding: 8px 14px;
    font-weight: 700;
}

.stTabs [aria-selected="true"] {
    background: #ffffff !important;
    color: var(--primary) !important;
    box-shadow: 0 2px 8px rgba(25, 42, 70, 0.08);
}

div[data-testid="stPlotlyChart"] {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 8px;
    box-shadow: 0 3px 12px rgba(25, 42, 70, 0.04);
}

.stButton > button,
.stDownloadButton > button {
    border-radius: 10px;
    font-weight: 750;
    border: 1px solid #d5deed;
}

.stDownloadButton > button {
    background: var(--primary);
    color: #ffffff;
    border-color: var(--primary);
}

hr {
    border: 0;
    border-top: 1px solid var(--border);
}

@media (max-width: 900px) {
    .block-container {
        padding: 1.6rem 0.75rem 2rem;
    }

    .app-header {
        flex-direction: column;
        align-items: flex-start;
    }

    .app-title {
        font-size: 25px;
    }

    .status-badge {
        white-space: normal;
    }

    .metric-value {
        font-size: 20px;
    }

    .asset-card,
    .asset-message-card {
        height: auto;
        min-height: 320px;
    }

    .asset-image-card img {
        height: auto;
        max-height: 380px;
    }
}
</style>
    """,
    unsafe_allow_html=True,
)

GRADE_ORDER = ["핵심 상품", "우수 상품", "유망 상품", "관찰 상품", "부진 상품"]
CASE_ORDER = [
    "가격 경쟁력 부족 사례",
    "기네스 갱신 사례",
    "타겟 확대 운영 사례",
    "시즌 상품 사례",
    "운영 피로도 사례",
    "보답프로그램 영향 사례",
]


def first_col(df: pd.DataFrame, names: Iterable[str]) -> str | None:
    for name in names:
        if name in df.columns:
            return name
    return None


def num(series: pd.Series) -> pd.Series:
    if series.dtype == object:
        series = (
            series.astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("%", "", regex=False)
            .replace({"-": "0", "nan": "0", "None": "0", "": "0"})
        )
    return pd.to_numeric(series, errors="coerce").fillna(0)


def safe_div(a, b):
    return a / b.replace(0, pd.NA)


def fmt_num(v) -> str:
    return f"{int(round(float(v))):,}"


def fmt_pct(v) -> str:
    return f"{float(v) * 100:.1f}%"


def compact_money(v: float) -> str:
    v = float(v)
    if abs(v) >= 10_000_000:
        return f"{v/10_000_000:.1f}천만원"
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.1f}백만원"
    if abs(v) >= 10_000:
        return f"{v/10_000:.1f}만원"
    return f"{int(v):,}원"


def product_grade(amount: float) -> str:
    if amount < 1_000_000:
        return "부진 상품"
    if amount < 2_000_000:
        return "관찰 상품"
    if amount < 3_000_000:
        return "유망 상품"
    if amount < 5_000_000:
        return "우수 상품"
    return "핵심 상품"


def normalize_product(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    date_col = first_col(d, ["발송일", "날짜", "일자"])
    if date_col is None:
        raise ValueError("상품 시트에서 발송일/날짜/일자 열을 찾을 수 없습니다.")
    d["_date"] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d["_date"].notna()].copy()
    for c in ["정상가", "멤버십혜택가", "할인율", "전시순서", "주문건수", "주문수량", "주문금액", "발송일 최저가"]:
        if c in d.columns:
            d[c] = num(d[c])
    d["_year"] = d["_date"].dt.year.astype(int)
    d["_month"] = d["_date"].dt.month.astype(int)
    return d


def normalize_send(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    date_col = first_col(d, ["발송일시2", "발송일", "날짜", "일자"])
    if date_col is None:
        raise ValueError("소재 시트에서 발송일시2/발송일/날짜/일자 열을 찾을 수 없습니다.")
    d["_date"] = pd.to_datetime(d[date_col], errors="coerce")
    d = d[d["_date"].notna()].copy()
    for c in [
        "상품수", "URL", "총 발송 건수", "발송 성공 건수",
        "클릭 수", "클릭 수(uniq)", "반응율", "반응율(uniq)",
        "주문건수", "주문수량", "주문금액", "객단가",
        "클릭>구매 전환율", "발송>구매 전환율", "SPM",
    ]:
        if c in d.columns:
            d[c] = num(d[c])
    d["_year"] = d["_date"].dt.year.astype(int)
    d["_month"] = d["_date"].dt.month.astype(int)
    return d


def normalize_lowest(df: pd.DataFrame | None) -> pd.DataFrame:
    """선택적인 '최저가' 시트를 정규화합니다."""
    if df is None or df.empty:
        return pd.DataFrame()

    d = df.copy()
    for c in ["쇼라코드", "알파코드"]:
        if c in d.columns:
            d[c] = d[c].astype(str).str.replace(".0", "", regex=False).str.strip()

    date_col = first_col(d, ["발송일", "날짜", "일자"])
    if date_col:
        d["_date"] = pd.to_datetime(d[date_col], errors="coerce")

    for c in [
        "네이버 최저가", "최저가", "비교가", "네이버가",
        "멤버십혜택가", "가격차이"
    ]:
        if c in d.columns:
            d[c] = num(d[c])

    price_col = first_col(d, ["네이버 최저가", "최저가", "네이버가", "비교가"])
    if price_col and price_col != "최저가":
        d = d.rename(columns={price_col: "최저가"})

    return d


@st.cache_data(show_spinner=False)
def load_excel_bytes(file_bytes: bytes):
    workbook = pd.ExcelFile(io.BytesIO(file_bytes))
    product = pd.read_excel(workbook, sheet_name="상품")
    send = pd.read_excel(workbook, sheet_name="소재")
    lowest = (
        pd.read_excel(workbook, sheet_name="최저가")
        if "최저가" in workbook.sheet_names
        else pd.DataFrame()
    )
    return normalize_product(product), normalize_send(send), normalize_lowest(lowest)


def extract_google_sheet_id(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError("구글시트 주소에서 문서 ID를 찾을 수 없습니다.")
    return match.group(1)



def google_requests_get(url: str, **kwargs):
    """
    Google 요청은 먼저 정상 SSL 검증으로 시도합니다.
    회사망 자체 서명 인증서 오류가 발생할 때만 검증 없이 한 번 재시도합니다.
    """
    try:
        return requests.get(url, **kwargs)
    except requests.exceptions.SSLError:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        retry_kwargs = dict(kwargs)
        retry_kwargs["verify"] = False
        return requests.get(url, **retry_kwargs)

def read_google_csv(sheet_id: str, sheet_name: str) -> pd.DataFrame:
    csv_url = (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
        f"?tqx=out:csv&sheet={requests.utils.quote(sheet_name)}"
    )
    response = google_requests_get(csv_url, timeout=30)
    response.raise_for_status()
    text = response.text
    if text.lstrip().lower().startswith("<!doctype html") or "<html" in text[:300].lower():
        raise PermissionError(
            f"'{sheet_name}' 탭을 불러오지 못했습니다. "
            "구글시트 공유 권한을 '링크가 있는 모든 사용자/뷰어'로 설정해주세요."
        )
    return pd.read_csv(io.StringIO(text))


@st.cache_data(ttl=300, show_spinner=False)
def load_google_sheet(url: str):
    """구글시트를 불러오고 5분간 캐시합니다."""
    sheet_id = extract_google_sheet_id(url)
    errors = []

    # 1차: 전체 문서를 XLSX로 내보내기
    try:
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
        response = google_requests_get(
            export_url,
            timeout=45,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        response.raise_for_status()
        content_type = response.headers.get("content-type", "").lower()
        if "text/html" not in content_type and response.content[:2] == b"PK":
            return load_excel_bytes(response.content)
        errors.append("전체 XLSX 내보내기 응답이 엑셀 파일이 아니었습니다.")
    except Exception as exc:
        errors.append(
            "전체 XLSX 내보내기 실패: "
            f"{exc}"
        )

    # 2차: 상품/소재 탭을 각각 CSV로 불러오기
    try:
        product = read_google_csv(sheet_id, "상품")
        send = read_google_csv(sheet_id, "소재")
        try:
            lowest = read_google_csv(sheet_id, "최저가")
        except Exception:
            lowest = pd.DataFrame()
        return normalize_product(product), normalize_send(send), normalize_lowest(lowest)
    except Exception as exc:
        errors.append(f"상품·소재 CSV 불러오기 실패: {exc}")

    raise RuntimeError(" / ".join(errors))


def sync_google_sheet(url: str, force: bool = False):
    """자동 또는 수동으로 구글시트를 세션 데이터에 반영합니다."""
    if force:
        load_google_sheet.clear()

    products, sends, lowest = load_google_sheet(url)
    st.session_state.products = products
    st.session_state.sends = sends
    st.session_state.lowest = lowest
    st.session_state.source_name = "구글시트 자동연동"
    st.session_state.synced_at = datetime.now()
    st.session_state.google_sync_error = None


def aggregate_send(data: pd.DataFrame, mode: str) -> pd.DataFrame:
    d = data.copy()
    if d.empty:
        return pd.DataFrame()

    if mode == "Monthly":
        d["_label"] = d["_date"].dt.strftime("%Y-%m")
        d["_sort1"] = d["_date"].dt.year
        d["_sort2"] = d["_date"].dt.month
    elif mode == "Weekly":
        d["_label"] = d["주차"].astype(str)
        d["_sort1"] = d["_date"].dt.year
        d["_sort2"] = pd.to_numeric(
            d["_label"].str.extract(r"(\d+)")[0], errors="coerce"
        ).fillna(0)
    else:
        d["_label"] = d["_date"].dt.strftime("%m%d")
        d["_sort1"] = d["_date"].dt.year
        d["_sort2"] = d["_date"].dt.dayofyear

    send_col = first_col(d, ["발송 성공 건수", "총 발송 건수"])
    click_col = first_col(d, ["클릭 수(uniq)", "클릭 수"])

    g = d.groupby("_label", as_index=False).agg(
        연도=("_sort1", "max"),
        _sort1=("_sort1", "max"),
        _sort2=("_sort2", "max"),
        발송횟수=("_label", "size"),
        상품수=("상품수", "sum"),
        URL=("URL", "sum"),
        발송건수=(send_col, "sum"),
        클릭수=(click_col, "sum"),
        주문건수=("주문건수", "sum"),
        주문수량=("주문수량", "sum"),
        주문금액=("주문금액", "sum"),
    ).sort_values(["_sort1", "_sort2"])

    if mode == "Monthly":
        g["월"] = g["_label"].str[-2:].astype(int)
    elif mode == "Daily":
        month_map = d.groupby("_label")["_month"].max()
        g["월"] = g["_label"].map(month_map).astype(int)

    g["반응율(Uniq CTR)"] = safe_div(g["클릭수"], g["발송건수"])
    g["객단가"] = safe_div(g["주문금액"], g["주문건수"])
    g["클릭 CVR"] = safe_div(g["주문건수"], g["클릭수"])
    g["발송 CVR"] = safe_div(g["주문건수"], g["발송건수"])
    g["클릭당매출(RPC)"] = safe_div(g["주문금액"], g["클릭수"])
    g["발송대비매출(SPM)"] = safe_div(g["주문금액"], g["발송건수"])
    g["발송당매출(발송횟수)"] = safe_div(g["주문금액"], g["발송횟수"])

    return g.fillna(0).reset_index(drop=True)


def add_changes(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["주문금액 증감"] = out["주문금액"].pct_change()
    out["CTR 증감"] = out["반응율(Uniq CTR)"].diff()
    out["SPM 증감"] = out["발송대비매출(SPM)"].pct_change()
    out["발송당매출 증감"] = out["발송당매출(발송횟수)"].pct_change()
    return out.replace([float("inf"), float("-inf")], pd.NA)


def change_label(value, pp: bool = False) -> str:
    if pd.isna(value):
        return "-"
    arrow = "▲" if value > 0 else ("▼" if value < 0 else "-")
    if pp:
        return f"{arrow}{abs(value)*100:.1f}%p"
    return f"{arrow}{abs(value)*100:.1f}%"


def format_home_table(df: pd.DataFrame, mode: str) -> pd.DataFrame:
    out = add_changes(df).copy()
    if mode == "Monthly":
        out = out.rename(columns={"_label": "기간"})
        order = [
            "연도", "월", "발송횟수", "상품수", "URL", "발송건수", "클릭수",
            "반응율(Uniq CTR)", "주문건수", "주문수량", "주문금액", "객단가",
            "클릭 CVR", "발송 CVR", "클릭당매출(RPC)",
            "발송대비매출(SPM)", "발송당매출(발송횟수)",
            "주문금액 증감", "CTR 증감", "SPM 증감", "발송당매출 증감",
        ]
    elif mode == "Weekly":
        out = out.rename(columns={"_label": "주차"})
        order = [
            "연도", "주차", "발송횟수", "상품수", "URL", "발송건수", "클릭수",
            "반응율(Uniq CTR)", "주문건수", "주문수량", "주문금액", "객단가",
            "클릭 CVR", "발송 CVR", "클릭당매출(RPC)",
            "발송대비매출(SPM)", "발송당매출(발송횟수)",
            "주문금액 증감", "CTR 증감", "SPM 증감", "발송당매출 증감",
        ]
    else:
        out = out.rename(columns={"_label": "일자"})
        order = [
            "연도", "월", "일자", "발송횟수", "상품수", "URL", "발송건수", "클릭수",
            "반응율(Uniq CTR)", "주문건수", "주문수량", "주문금액", "객단가",
            "클릭 CVR", "발송 CVR", "클릭당매출(RPC)",
            "발송대비매출(SPM)", "발송당매출(발송횟수)",
            "주문금액 증감", "CTR 증감", "SPM 증감", "발송당매출 증감",
        ]

    out = out[[c for c in order if c in out.columns]].copy()

    for c in ["반응율(Uniq CTR)", "클릭 CVR", "발송 CVR"]:
        if c in out.columns:
            out[c] = out[c].map(fmt_pct)

    for c in ["주문금액 증감", "SPM 증감", "발송당매출 증감"]:
        if c in out.columns:
            out[c] = out[c].map(change_label)

    if "CTR 증감" in out.columns:
        out["CTR 증감"] = out["CTR 증감"].map(lambda x: change_label(x, pp=True))

    for c in [
        "발송횟수", "상품수", "URL", "발송건수", "클릭수",
        "주문건수", "주문수량", "주문금액", "객단가",
        "클릭당매출(RPC)", "발송당매출(발송횟수)",
    ]:
        if c in out.columns:
            out[c] = out[c].map(fmt_num)

    if "발송대비매출(SPM)" in out.columns:
        out["발송대비매출(SPM)"] = out["발송대비매출(SPM)"].map(lambda x: f"{x:.1f}")

    return out


def filter_monthly_period(df: pd.DataFrame, option: str, start_key: str | None, end_key: str | None):
    if df.empty or option == "전체":
        return df
    if option.startswith("최근"):
        count = int(re.search(r"\d+", option).group())
        return df.tail(count)
    if option == "직접 선택" and start_key and end_key:
        labels = df["_label"].astype(str)
        return df[(labels >= start_key) & (labels <= end_key)]
    return df


def filter_weekly_period(df: pd.DataFrame, start_week: str, end_week: str):
    if df.empty:
        return df
    labels = df["_label"].astype(str).tolist()
    if start_week not in labels or end_week not in labels:
        return df
    start_idx, end_idx = labels.index(start_week), labels.index(end_week)
    lo, hi = sorted([start_idx, end_idx])
    return df.iloc[lo : hi + 1]


def trend_chart(df: pd.DataFrame, title: str, color: str) -> go.Figure:
    labels = df["_label"].astype(str).tolist()
    vals = df["발송대비매출(SPM)"].astype(float).tolist()
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=labels,
            y=vals,
            marker_color=color,
            name="SPM",
            text=[f"{v:.1f}" for v in vals],
            textposition="outside",
            cliponaxis=False,
        )
    )
    if len(vals) >= 2:
        trend = pd.Series(vals).rolling(3, min_periods=1).mean()
        fig.add_trace(
            go.Scatter(
                x=labels,
                y=trend,
                mode="lines",
                name="추세",
                line=dict(color="#333", width=2, dash="dot"),
            )
        )
    ymax = max(vals) if vals else 0
    fig.update_layout(
        title=dict(text=title, x=.5, font=dict(size=23)),
        height=480,
        margin=dict(l=60, r=30, t=72, b=100),
        plot_bgcolor="#ffffff",
        barmode="overlay",
        yaxis=dict(
            tickformat=",.1f",
            gridcolor="#ddd",
            range=[0, max(ymax * 1.22, 10)],
        ),
        xaxis=dict(
            tickangle=-35 if len(labels) > 12 else 0,
            automargin=True,
        ),
        legend=dict(orientation="h", y=-.23),
    )
    return fig


def delta_for_latest(df: pd.DataFrame, metric: str, pp: bool = False) -> str:
    if len(df) < 2:
        return "-"
    cur, prev = df.iloc[-1][metric], df.iloc[-2][metric]
    if pp:
        return change_label(cur - prev, pp=True)
    if prev == 0:
        return "-"
    return change_label((cur - prev) / abs(prev))


def classify_cases(row: pd.Series, history: pd.DataFrame) -> list[str]:
    cases = []
    name = str(row["상품명"])
    amount = float(row["주문금액"])
    prior = history[(history["상품명"] == name) & (history["_date"] < row["_date"])].sort_values("_date")

    if len(prior) and amount > prior["주문금액"].max() and amount >= 3_000_000:
        cases.append("기네스 갱신 사례")

    if len(prior):
        last = prior.iloc[-1]
        if (
            str(last.get("성별", "")) != str(row.get("성별", ""))
            or str(last.get("연령", "")) != str(row.get("연령", ""))
        ):
            cases.append("타겟 확대 운영 사례")
        gap = (row["_date"] - last["_date"]).days
        if gap <= 21 and amount < last["주문금액"] * 0.75:
            cases.append("운영 피로도 사례")
        if row.get("멤버십혜택가", 0) < last.get("멤버십혜택가", 0) and amount < last["주문금액"]:
            cases.append("가격 경쟁력 부족 사례")

    season_words = ["선풍기", "에어컨", "우양산", "래쉬가드", "삼계탕", "장어", "아이스크림", "드라이기"]
    if any(word in name for word in season_words):
        cases.append("시즌 상품 사례")

    if amount < 1_000_000 and float(row.get("할인율", 0)) >= 0.5:
        cases.append("가격 경쟁력 부족 사례")

    return list(dict.fromkeys(cases))


def make_insight(row: pd.Series, history: pd.DataFrame) -> str:
    name = str(row["상품명"])
    amount = float(row["주문금액"])
    grade = product_grade(amount)
    prior = history[(history["상품명"] == name) & (history["_date"] < row["_date"])].sort_values("_date")
    parts = []

    if len(prior) == 0:
        parts.append("MMS 첫 운영 상품으로 신규 TEST 진행")
    elif len(prior) == 1:
        parts.append(f"기존 운영 시 {compact_money(prior.iloc[-1]['주문금액'])} 기록")
    else:
        parts.append(
            f"기존 운영 시 최소 {compact_money(prior['주문금액'].min())}에서 "
            f"최대 {compact_money(prior['주문금액'].max())} 기록"
        )

    exposure = row.get("추가노출")
    if pd.notna(exposure) and str(exposure) not in ["", "-", "nan"]:
        parts.append(f"{exposure} 추가 노출 운영")

    dt = row.get("_date")
    current_date = f"{dt.month}/{dt.day}" if pd.notna(dt) and hasattr(dt, "month") else ""
    target_text = " ".join(
        str(row.get(c, "")) for c in ["성별", "연령"] if str(row.get(c, "")) not in ["", "nan"]
    ).strip()
    context = f"{current_date} {target_text} 운영에서".strip()
    parts.append(f"금번 {context} {compact_money(amount)} 기록")
    cases = classify_cases(row, history)

    if "기네스 갱신 사례" in cases:
        parts.append("기존 최고 주문금액을 상회하며 MMS 발송 기준 기네스 갱신")
    elif grade == "핵심 상품":
        parts.append("매우 우수한 판매 성과 확인")
    elif grade == "우수 상품":
        parts.append("우수한 판매 성과 확인")
    elif grade == "유망 상품":
        parts.append("추가 재편성 검토가 가능한 실적")
    elif grade == "관찰 상품":
        parts.append("기대 대비 다소 아쉬운 실적")
    else:
        parts.append("매우 저조한 실적으로 MMS 메인 상품 적합도가 낮은 것으로 확인")

    if "타겟 확대 운영 사례" in cases:
        parts.append("추가 운영 결과 비교를 통해 우수 타겟 확인 필요")
    elif "운영 피로도 사례" in cases:
        parts.append("단기간 동일 상품 재노출에 따른 운영 피로도 영향 가능성 존재")
    elif "시즌 상품 사례" in cases and grade in ["핵심 상품", "우수 상품"]:
        parts.append("시즌 종료 전까지 지속 운영 검토 필요")
    elif grade in ["핵심 상품", "우수 상품"]:
        parts.append("추가 운영 검토 필요")
    elif grade == "유망 상품":
        parts.append("타겟 및 전시순서 변경 후 추가 테스트 검토")
    elif grade == "관찰 상품":
        parts.append("가격 경쟁력·타겟 적합성 추가 확인 필요")

    return f"[{name}] " + " > ".join(parts)


def build_ppt(title: str, lines: list[str], table_df: pd.DataFrame | None = None) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(.6), Inches(.35), Inches(12), Inches(.65))
    p = title_box.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(26)
    p.font.bold = True

    body = slide.shapes.add_textbox(Inches(.7), Inches(1.2), Inches(12), Inches(5.7))
    tf = body.text_frame
    tf.word_wrap = True
    for idx, line in enumerate(lines):
        para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        para.text = line
        para.font.size = Pt(14)
        para.space_after = Pt(7)

    if table_df is not None and not table_df.empty:
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        head = slide2.shapes.add_textbox(Inches(.6), Inches(.3), Inches(12), Inches(.6))
        hp = head.text_frame.paragraphs[0]
        hp.text = "상세 실적"
        hp.font.size = Pt(24)
        hp.font.bold = True

        view = table_df.head(14).copy()
        rows, cols = len(view) + 1, len(view.columns)
        table = slide2.shapes.add_table(
            rows, cols, Inches(.35), Inches(1), Inches(12.6), Inches(5.9)
        ).table
        for j, c in enumerate(view.columns):
            table.cell(0, j).text = str(c)
        for i, (_, r) in enumerate(view.iterrows(), start=1):
            for j, c in enumerate(view.columns):
                table.cell(i, j).text = str(r[c])

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()


def append_total_and_change_rows(raw: pd.DataFrame, mode: str) -> pd.DataFrame:
    """상세 기간 행 아래에 총합계와 최신 기간 증감을 추가합니다."""
    if raw.empty:
        return raw.copy()

    d = raw.copy()
    total = {
        "_label": "총합계",
        "연도": "",
        "월": "",
        "발송횟수": d["발송횟수"].sum(),
        "상품수": d["상품수"].sum(),
        "URL": d["URL"].sum(),
        "발송건수": d["발송건수"].sum(),
        "클릭수": d["클릭수"].sum(),
        "주문건수": d["주문건수"].sum(),
        "주문수량": d["주문수량"].sum(),
        "주문금액": d["주문금액"].sum(),
    }
    total["반응율(Uniq CTR)"] = total["클릭수"] / total["발송건수"] if total["발송건수"] else 0
    total["객단가"] = total["주문금액"] / total["주문건수"] if total["주문건수"] else 0
    total["클릭 CVR"] = total["주문건수"] / total["클릭수"] if total["클릭수"] else 0
    total["발송 CVR"] = total["주문건수"] / total["발송건수"] if total["발송건수"] else 0
    total["클릭당매출(RPC)"] = total["주문금액"] / total["클릭수"] if total["클릭수"] else 0
    total["발송대비매출(SPM)"] = total["주문금액"] / total["발송건수"] if total["발송건수"] else 0
    total["발송당매출(발송횟수)"] = total["주문금액"] / total["발송횟수"] if total["발송횟수"] else 0

    change = {"_label": "증감", "연도": "", "월": ""}
    numeric_cols = [
        "발송횟수", "상품수", "URL", "발송건수", "클릭수",
        "주문건수", "주문수량", "주문금액", "객단가",
        "클릭당매출(RPC)", "발송대비매출(SPM)", "발송당매출(발송횟수)"
    ]
    rate_cols = ["반응율(Uniq CTR)", "클릭 CVR", "발송 CVR"]

    if len(d) >= 2:
        cur, prev = d.iloc[-1], d.iloc[-2]
        for col in numeric_cols:
            if col in d.columns:
                change[col] = (cur[col] - prev[col]) / abs(prev[col]) if prev[col] else pd.NA
        for col in rate_cols:
            if col in d.columns:
                change[col] = cur[col] - prev[col]
    else:
        for col in numeric_cols + rate_cols:
            change[col] = pd.NA

    return pd.concat([d, pd.DataFrame([total, change])], ignore_index=True)


def format_home_table_with_summary(df: pd.DataFrame, mode: str) -> pd.DataFrame:
    """홈 표 전용: 일반행 + 총합계 + 증감 행."""
    raw = append_total_and_change_rows(df, mode)
    if raw.empty:
        return raw

    label_name = {"Monthly": "구분", "Weekly": "구분", "Daily": "구분"}[mode]
    raw = raw.rename(columns={"_label": label_name})

    if mode == "Monthly":
        order = [
            label_name, "연도", "월", "발송횟수", "상품수", "URL", "발송건수", "클릭수",
            "반응율(Uniq CTR)", "주문건수", "주문수량", "주문금액", "객단가",
            "클릭 CVR", "발송 CVR", "클릭당매출(RPC)",
            "발송대비매출(SPM)", "발송당매출(발송횟수)"
        ]
    elif mode == "Weekly":
        order = [
            label_name, "연도", "발송횟수", "상품수", "URL", "발송건수", "클릭수",
            "반응율(Uniq CTR)", "주문건수", "주문수량", "주문금액", "객단가",
            "클릭 CVR", "발송 CVR", "클릭당매출(RPC)",
            "발송대비매출(SPM)", "발송당매출(발송횟수)"
        ]
    else:
        order = [
            label_name, "연도", "월", "발송횟수", "상품수", "URL", "발송건수", "클릭수",
            "반응율(Uniq CTR)", "주문건수", "주문수량", "주문금액", "객단가",
            "클릭 CVR", "발송 CVR", "클릭당매출(RPC)", "발송대비매출(SPM)"
        ]

    view = raw[[c for c in order if c in raw.columns]].copy()

    # pandas 최신 버전에서는 숫자형 열에 "5.7%" 같은 문자열을 다시 대입할 수 없으므로
    # 화면 표시용 데이터프레임 전체를 object 형식으로 변환합니다.
    view = view.astype("object")

    change_mask = view[label_name].astype(str).eq("증감")

    # 일반행/총합계 포맷
    for col in ["반응율(Uniq CTR)", "클릭 CVR", "발송 CVR"]:
        if col in view.columns:
            view.loc[~change_mask, col] = view.loc[~change_mask, col].map(fmt_pct)
            view.loc[change_mask, col] = view.loc[change_mask, col].map(
                lambda x: change_label(x, pp=True)
            )

    for col in [
        "발송횟수", "상품수", "URL", "발송건수", "클릭수",
        "주문건수", "주문수량", "주문금액", "객단가",
        "클릭당매출(RPC)", "발송당매출(발송횟수)"
    ]:
        if col in view.columns:
            view.loc[~change_mask, col] = view.loc[~change_mask, col].map(fmt_num)
            view.loc[change_mask, col] = view.loc[change_mask, col].map(change_label)

    if "발송대비매출(SPM)" in view.columns:
        view.loc[~change_mask, "발송대비매출(SPM)"] = (
            view.loc[~change_mask, "발송대비매출(SPM)"].map(lambda x: f"{float(x):.1f}")
        )
        view.loc[change_mask, "발송대비매출(SPM)"] = (
            view.loc[change_mask, "발송대비매출(SPM)"].map(change_label)
        )

    return view


def merge_lowest_price(product_df: pd.DataFrame, lowest_df: pd.DataFrame | None = None) -> pd.DataFrame:
    """상품 RAW의 '발송일 최저가' 컬럼을 직접 사용합니다."""
    d = product_df.copy()

    if "발송일 최저가" in d.columns:
        d["최저가"] = num(d["발송일 최저가"])
        d.loc[d["최저가"] <= 0, "최저가"] = pd.NA
        d["가격차이"] = d["최저가"] - d["멤버십혜택가"]
        d["최저가 확보"] = d.apply(
            lambda r: (
                "확보" if pd.notna(r.get("최저가")) and r["멤버십혜택가"] < r["최저가"]
                else (
                    "동일가" if pd.notna(r.get("최저가")) and r["멤버십혜택가"] == r["최저가"]
                    else ("미확보" if pd.notna(r.get("최저가")) else "")
                )
            ),
            axis=1,
        )
        return d

    # 과거 파일 호환: 발송일 최저가 컬럼이 없으면 공란 처리
    d["최저가"] = pd.NA
    d["가격차이"] = pd.NA
    d["최저가 확보"] = ""
    return d


def weekly_product_chart(sw: pd.DataFrame) -> go.Figure:
    f = sw.sort_values("_date").copy()
    labels = [
        f"{r.get('요일','')}<br>{r.get('시간대','')}<br>{r.get('소재','')}"
        for _, r in f.iterrows()
    ]
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(
            x=labels, y=f["주문금액"], name="주문금액",
            marker_color="#70ad47",
            text=[fmt_num(v) for v in f["주문금액"]],
            textposition="inside",
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=labels, y=f["주문수량"], name="주문수량",
            mode="lines+markers+text",
            line=dict(color="#f4b000", width=3),
            text=[fmt_num(v) for v in f["주문수량"]],
            textposition="top center",
        ),
        secondary_y=True,
    )
    fig.add_trace(
        go.Scatter(
            x=labels, y=f["상품수"], name="상품수",
            mode="lines+markers+text",
            line=dict(color="#ed7d31", width=2),
            text=[fmt_num(v) for v in f["상품수"]],
            textposition="bottom center",
        ),
        secondary_y=True,
    )
    fig.update_yaxes(tickformat=",", gridcolor="#ddd", secondary_y=False)
    fig.update_yaxes(tickformat=",", showgrid=False, secondary_y=True)
    fig.update_layout(
        title=dict(text="MMS 상품 실적", x=.5, font=dict(size=23)),
        height=560,
        margin=dict(l=60, r=70, t=70, b=150),
        plot_bgcolor="#ffffff",
        barmode="overlay",
        xaxis=dict(automargin=True, tickfont=dict(size=10)),
        legend=dict(orientation="h", y=-.24),
    )
    return fig


def weekly_send_chart(sw: pd.DataFrame) -> go.Figure:
    f = sw.sort_values("_date").copy()
    labels = [
        f"{r.get('요일','')}<br>{r.get('시간대','')}<br>{r.get('소재','')}"
        for _, r in f.iterrows()
    ]
    click_all = first_col(f, ["클릭 수", "클릭 수(uniq)"])
    click_uniq = first_col(f, ["클릭 수(uniq)", "클릭 수"])
    ctr_all = first_col(f, ["반응율", "반응율(uniq)"])
    ctr_uniq = first_col(f, ["반응율(uniq)", "반응율"])

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(
            x=labels, y=f[click_all], name="클릭 수",
            marker_color="#5b9bd5",
            text=[fmt_num(v) for v in f[click_all]],
            textposition="inside",
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Bar(
            x=labels, y=f[click_uniq], name="클릭 수(uniq)",
            marker_color="#a5a5a5",
            text=[fmt_num(v) for v in f[click_uniq]],
            textposition="inside",
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=labels, y=f[ctr_all] * 100, name="반응율",
            mode="lines+markers+text",
            line=dict(color="#ed7d31", width=3),
            text=[f"{v*100:.0f}%" for v in f[ctr_all]],
            textposition="top center",
        ),
        secondary_y=True,
    )
    fig.add_trace(
        go.Scatter(
            x=labels, y=f[ctr_uniq] * 100, name="반응율(uniq)",
            mode="lines+markers+text",
            line=dict(color="#f4b000", width=3),
            text=[f"{v*100:.0f}%" for v in f[ctr_uniq]],
            textposition="bottom center",
        ),
        secondary_y=True,
    )
    fig.update_yaxes(tickformat=",", gridcolor="#ddd", secondary_y=False)
    fig.update_yaxes(ticksuffix="%", showgrid=False, secondary_y=True)
    fig.update_layout(
        title=dict(text="MMS 발송 통계", x=.5, font=dict(size=23)),
        height=560,
        margin=dict(l=60, r=70, t=70, b=150),
        plot_bgcolor="#ffffff",
        barmode="group",
        xaxis=dict(automargin=True, tickfont=dict(size=10)),
        legend=dict(orientation="h", y=-.24),
    )
    return fig


def grouped_send_table(sw: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
    send_col = first_col(sw, ["발송 성공 건수", "총 발송 건수"])
    click_col = first_col(sw, ["클릭 수(uniq)", "클릭 수"])
    g = sw.groupby(keys, as_index=False).agg(
        발송횟수=("소재", "size"),
        발송건수=(send_col, "sum"),
        클릭수=(click_col, "sum"),
        주문건수=("주문건수", "sum"),
        주문금액=("주문금액", "sum"),
    )
    g["CTR(uniq)"] = safe_div(g["클릭수"], g["발송건수"])
    g["CVR(클릭>구매)"] = safe_div(g["주문건수"], g["클릭수"])
    g["객단가"] = safe_div(g["주문금액"], g["주문건수"])
    g["SPM"] = safe_div(g["주문금액"], g["발송건수"])
    g["발송당매출(발송횟수)"] = safe_div(g["주문금액"], g["발송횟수"])
    return g.fillna(0)


def add_total_row(df: pd.DataFrame, label_col: str, label="총합계") -> pd.DataFrame:
    if df.empty:
        return df
    total = {}
    for c in df.columns:
        if c == label_col:
            total[c] = label
        elif pd.api.types.is_numeric_dtype(df[c]):
            total[c] = df[c].sum()
        else:
            total[c] = ""
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def weekly_display_format(df: pd.DataFrame) -> pd.DataFrame:
    """주간 표 표시 형식을 안전하게 통일합니다."""
    out = df.copy()

    def format_percent_value(x):
        if pd.isna(x) or str(x).strip() in ["", "nan", "None"]:
            return ""
        if isinstance(x, str):
            value = x.strip()
            if value.endswith("%"):
                return value
            value = value.replace(",", "")
        else:
            value = x
        try:
            numeric = float(value)
            return fmt_pct(numeric)
        except (TypeError, ValueError):
            return str(x)

    def format_number_value(x):
        if pd.isna(x) or str(x).strip() in ["", "nan", "None"]:
            return ""
        if isinstance(x, str):
            value = x.strip()
            if value.endswith("%"):
                return value
            value = value.replace(",", "")
        else:
            value = x
        try:
            return fmt_num(float(value))
        except (TypeError, ValueError):
            return str(x)

    def format_spm_value(x):
        if pd.isna(x) or str(x).strip() in ["", "nan", "None"]:
            return ""
        if isinstance(x, str):
            value = x.strip().replace(",", "")
        else:
            value = x
        try:
            return f"{float(value):.1f}"
        except (TypeError, ValueError):
            return str(x)

    percent_cols = [
        "편성비중", "주문비중", "CTR(uniq)", "CVR(클릭>구매)",
        "CTR", "CVR"
    ]
    for c in percent_cols:
        if c in out.columns:
            out[c] = out[c].map(format_percent_value)

    if "SPM" in out.columns:
        out["SPM"] = out["SPM"].map(format_spm_value)

    number_cols = [
        "발송횟수", "상품수", "URL", "발송성공건수", "발송건수",
        "클릭수(uniq)", "클릭수", "객단가", "주문건수", "주문수량",
        "주문금액", "발송당매출(발송횟수)", "멤버십혜택가",
        "최저가", "가격차이", "전시순서"
    ]
    for c in number_cols:
        if c in out.columns:
            out[c] = out[c].map(format_number_value)

    return out


def category_summary_table(
    pw: pd.DataFrame,
    category_col: str,
    week: str,
    year: int,
) -> pd.DataFrame:
    """편성비중·주문비중·주문금액과 총합계를 생성합니다."""
    cat = (
        pw.groupby(category_col, dropna=False, as_index=False)
        .agg(
            편성수=("상품명", "size"),
            주문금액=("주문금액", "sum"),
        )
    )
    cat[category_col] = cat[category_col].fillna("미분류").astype(str)
    total_count = cat["편성수"].sum()
    total_amount = cat["주문금액"].sum()

    cat["편성비중"] = cat["편성수"] / total_count if total_count else 0
    cat["주문비중"] = cat["주문금액"] / total_amount if total_amount else 0

    # 주문비중 큰 순서
    cat = cat.sort_values(
        ["주문비중", "주문금액", category_col],
        ascending=[False, False, True],
    ).reset_index(drop=True)

    view = cat[[category_col, "편성비중", "주문비중", "주문금액"]].copy()
    view = view.rename(columns={category_col: "행 레이블"})
    view.insert(0, "주차", week)
    view.insert(0, "연도", year)

    total_row = pd.DataFrame([{
        "연도": "",
        "주차": "",
        "행 레이블": "총합계",
        "편성비중": 1.0 if total_count else 0,
        "주문비중": 1.0 if total_amount else 0,
        "주문금액": total_amount,
    }])
    return pd.concat([view, total_row], ignore_index=True)


def category_pie_chart(
    table: pd.DataFrame,
    title: str,
) -> go.Figure:
    """주문비중 큰 순서부터 시계방향으로 표시합니다."""
    data = table[table["행 레이블"] != "총합계"].copy()
    data = data.sort_values(
        ["주문비중", "주문금액"],
        ascending=[False, False],
    ).reset_index(drop=True)

    # 음수 주문금액은 파이에서 표현할 수 없어 0으로 처리하되 표에는 원값 유지
    values = pd.to_numeric(data["주문금액"], errors="coerce").fillna(0).clip(lower=0)

    fig = go.Figure(
        go.Pie(
            labels=data["행 레이블"],
            values=values,
            sort=False,
            direction="clockwise",
            rotation=0,
            textinfo="label+percent",
            textposition="inside",
            hovertemplate="%{label}<br>주문금액 %{value:,.0f}원<br>비중 %{percent}<extra></extra>",
        )
    )
    fig.update_layout(
        title=dict(text=title, x=.5),
        height=560,
        margin=dict(l=40, r=40, t=70, b=40),
        uniformtext_minsize=8,
        uniformtext_mode="show",
    )
    return fig


def clean_identifier_value(x):
    if pd.isna(x) or str(x).strip() in ["", "nan", "None"]:
        return ""
    value = str(x).strip().replace(",", "")
    if value.endswith(".0"):
        value = value[:-2]
    return value


def clean_identifier_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in ["연도", "월", "일", "일자", "연령", "SEG", "전시순서", "쇼라코드", "알파코드"]:
        if c in out.columns:
            out[c] = out[c].map(clean_identifier_value)
    return out


def style_weekly_product_rows(formatted_df: pd.DataFrame, raw_amounts: list):
    styles = pd.DataFrame("", index=formatted_df.index, columns=formatted_df.columns)
    for idx, amount in enumerate(raw_amounts):
        if idx >= len(formatted_df):
            break
        try:
            value = float(amount)
        except (TypeError, ValueError):
            continue
        if value >= 3_000_000:
            styles.iloc[idx, :] = "background-color: #fff2cc"
        elif value < 1_000_000:
            styles.iloc[idx, :] = "background-color: #e7e6e6"
    return styles


def weekly_delta(cur: float, prev: float, pp: bool = False) -> str:
    if pd.isna(cur) or pd.isna(prev):
        return "-"
    if pp:
        return change_label(cur - prev, pp=True)
    if prev == 0:
        return "-"
    return change_label((cur - prev) / abs(prev))


def build_weekly_analysis(
    week: str,
    year: int,
    pw: pd.DataFrame,
    sw: pd.DataFrame,
    products_all: pd.DataFrame,
    sends_all: pd.DataFrame,
) -> str:
    """선택 주차의 실제 수치만 사용해 주간 분석 문구를 생성합니다."""
    send_col = first_col(sw, ["발송 성공 건수", "총 발송 건수"])
    click_col = first_col(sw, ["클릭 수(uniq)", "클릭 수"])

    send_count = float(sw[send_col].sum())
    click_count = float(sw[click_col].sum())
    order_count = float(sw["주문건수"].sum())
    qty = float(sw["주문수량"].sum())
    amount = float(sw["주문금액"].sum())
    ctr = click_count / send_count if send_count else 0
    cvr = order_count / click_count if click_count else 0
    aov = amount / order_count if order_count else 0
    spm = amount / send_count if send_count else 0

    all_weeks = (
        sends_all[sends_all["_year"] == year]
        .groupby("주차")["_date"].min()
        .sort_values()
    )
    week_names = [str(x) for x in all_weeks.index]
    prev_sw = pd.DataFrame()
    if week in week_names and week_names.index(week) > 0:
        prev_week = week_names[week_names.index(week) - 1]
        prev_sw = sends_all[
            (sends_all["_year"] == year)
            & (sends_all["주차"].astype(str) == prev_week)
        ]

    if not prev_sw.empty:
        psend = float(prev_sw[send_col].sum())
        pclick = float(prev_sw[click_col].sum())
        porders = float(prev_sw["주문건수"].sum())
        pqty = float(prev_sw["주문수량"].sum())
        pamount = float(prev_sw["주문금액"].sum())
        pctr = pclick / psend if psend else 0
        pcvr = porders / pclick if pclick else 0
        paov = pamount / porders if porders else 0
        pspm = pamount / psend if psend else 0
        prev_compare = (
            f"전주 대비 발송횟수 {weekly_delta(len(sw), len(prev_sw))} / "
            f"상품수 {len(pw):,}건 {weekly_delta(len(pw), float(prev_sw['상품수'].sum()))} / "
            f"발송건수 {weekly_delta(send_count, psend)} 운영\n"
            f"주문건수 {int(order_count):,}건 {weekly_delta(order_count, porders)} / "
            f"주문수량 {int(qty):,}건 {weekly_delta(qty, pqty)} / "
            f"주문금액 {compact_money(amount)} {weekly_delta(amount, pamount)} 기록\n"
            f"CTR {ctr*100:.1f}% {weekly_delta(ctr, pctr, pp=True)} / "
            f"CVR {cvr*100:.1f}% {weekly_delta(cvr, pcvr, pp=True)} / "
            f"객단가 {int(aov):,}원 {weekly_delta(aov, paov)} / "
            f"SPM {spm:.1f} {weekly_delta(spm, pspm)} 기록"
        )
    else:
        prev_compare = (
            f"발송횟수 {len(sw)}회 / 상품수 {len(pw)}건 / 발송건수 {int(send_count):,}건 운영\n"
            f"주문건수 {int(order_count):,}건 / 주문수량 {int(qty):,}건 / "
            f"주문금액 {compact_money(amount)} 기록\n"
            f"CTR {ctr*100:.1f}% / CVR {cvr*100:.1f}% / "
            f"객단가 {int(aov):,}원 / SPM {spm:.1f} 기록"
        )

    product_rank = (
        pw.groupby("상품명", as_index=False)
        .agg(주문금액=("주문금액", "sum"))
        .sort_values("주문금액", ascending=False)
    )
    top = product_rank.head(1)
    over5 = product_rank[product_rank["주문금액"] >= 5_000_000]["상품명"].tolist()
    over3 = product_rank[
        (product_rank["주문금액"] >= 3_000_000)
        & (product_rank["주문금액"] < 5_000_000)
    ]["상품명"].tolist()
    under1 = product_rank[product_rank["주문금액"] < 1_000_000]["상품명"].tolist()

    top_line = ""
    if not top.empty:
        top_line = (
            f"[{top.iloc[0]['상품명']}] {compact_money(top.iloc[0]['주문금액'])}으로 "
            "금주 최고 매출 기록"
        )

    send_stats = sw.copy()
    send_stats["_SPM"] = safe_div(send_stats["주문금액"], send_stats[send_col])
    send_stats["_CTR"] = safe_div(send_stats[click_col], send_stats[send_col])
    max_spm = send_stats.loc[send_stats["_SPM"].idxmax()]
    max_ctr = send_stats.loc[send_stats["_CTR"].idxmax()]
    min_spm = send_stats.loc[send_stats["_SPM"].idxmin()]

    big_cat = pw.groupby("대카", as_index=False)["주문금액"].sum().sort_values("주문금액", ascending=False)
    mid_cat = pw.groupby("중카", as_index=False)["주문금액"].sum().sort_values("주문금액", ascending=False)
    big_total = big_cat["주문금액"].sum()
    big_lines = ", ".join(
        f"{r['대카']} {r['주문금액']/big_total*100:.1f}%"
        for _, r in big_cat.head(3).iterrows()
    )
    mid_lines = " > ".join(mid_cat.head(5)["중카"].astype(str).tolist())

    seg = grouped_send_table(sw, ["성별", "연령"])
    best_seg_spm = seg.loc[seg["SPM"].idxmax()]
    best_seg_amt = seg.loc[seg["주문금액"].idxmax()]

    weekday = grouped_send_table(sw, ["요일"])
    time_df = grouped_send_table(sw, ["시간대"])
    best_day = weekday.loc[weekday["SPM"].idxmax()]
    best_time = time_df.loc[time_df["SPM"].idxmax()]

    product_insights = "\n".join(
        make_insight(r, products_all) for _, r in pw.sort_values("주문금액", ascending=False).head(6).iterrows()
    )

    price_df = merge_lowest_price(pw)
    unavailable = price_df[price_df["최저가 확보"] == "미확보"]["상품명"].dropna().astype(str).unique().tolist()
    price_line = (
        "[발송일 최저가 미확보 상품] " + " / ".join(unavailable)
        if unavailable else "발송일 최저가가 입력된 상품 중 미확보 사례 없음"
    )

    lines = [
        "■ 주간 요약",
        prev_compare,
        "",
        "■ MMS 상품 실적",
        top_line,
        ("[5백만원 이상] " + " / ".join(over5)) if over5 else "5백만원 이상 상품 없음",
        ("[3백만원 이상] " + " / ".join(over3)) if over3 else "3~5백만원 상품 없음",
        ("[1백만원 미만] " + " / ".join(under1)) if under1 else "1백만원 미만 상품 없음",
        "",
        "■ MMS 발송 통계",
        f"{max_spm.get('요일','')}({max_spm.get('시간대','')}) {max_spm.get('소재','')} "
        f"SPM {max_spm['_SPM']:.1f} 금주 최고 효율 기록",
        f"{max_ctr.get('요일','')}({max_ctr.get('시간대','')}) {max_ctr.get('소재','')} "
        f"CTR {max_ctr['_CTR']*100:.1f}% 금주 최고 반응 기록",
        f"{min_spm.get('요일','')}({min_spm.get('시간대','')}) {min_spm.get('소재','')} "
        f"SPM {min_spm['_SPM']:.1f} 금주 최저 효율 기록",
        "",
        "■ 카테고리 분석",
        f"대카테고리 기준 {big_lines} 매출 비중 차지",
        f"중카테고리 기준 {mid_lines} 중심 매출 구성",
        "",
        "■ SEG 분석",
        f"{best_seg_spm['성별']}{best_seg_spm['연령']} SPM {best_seg_spm['SPM']:.1f}로 최고 효율 기록",
        f"{best_seg_amt['성별']}{best_seg_amt['연령']} 주문금액 {compact_money(best_seg_amt['주문금액'])}으로 최고 매출 기여",
        "",
        "■ 요일·시간대 분석",
        f"{best_day['요일']}요일 SPM {best_day['SPM']:.1f}로 최고 효율 기록",
        f"{best_time['시간대']} 시간대 SPM {best_time['SPM']:.1f}로 시간대 최고 효율 기록",
        "",
        "■ 상품 인사이트",
        product_insights,
        "",
        "■ 가격·카테고리 인사이트",
        price_line,
    ]
    return "\n".join(str(x) for x in lines if x is not None)


APP_DIR = Path(__file__).resolve().parent
IMAGE_DIR = APP_DIR / "images"
MESSAGE_DIR = APP_DIR / "messages"


def daily_asset_key(date_value, time_value) -> str:
    dt = pd.to_datetime(date_value, errors="coerce")
    if pd.isna(dt):
        return ""

    time_text = str(time_value).strip()
    if time_text.startswith("10") or "오전" in time_text or time_text in ["1", "01"]:
        slot = "01"
    elif time_text.startswith("16") or time_text.startswith("17") or "오후" in time_text or time_text in ["2", "02"]:
        slot = "02"
    else:
        parsed = pd.to_datetime(time_text, errors="coerce")
        if pd.notna(parsed):
            slot = "01" if parsed.hour < 12 else "02"
        else:
            return ""
    return f"{dt:%Y%m%d}_{slot}"


def find_daily_image(asset_key: str):
    if not asset_key or not IMAGE_DIR.exists():
        return None
    matches = [
        p for p in IMAGE_DIR.iterdir()
        if p.is_file()
        and p.name.startswith(asset_key)
        and p.suffix.lower() in [".jpg", ".jpeg", ".png", ".webp"]
    ]
    return sorted(matches, key=lambda p: p.name)[0] if matches else None


def clean_mms_message(value) -> str:
    """앞뒤 큰따옴표만 제거하고 내부 줄바꿈은 그대로 유지합니다."""
    if value is None or pd.isna(value):
        return ""

    text_value = str(value)
    stripped = text_value.strip()

    if len(stripped) >= 2 and stripped.startswith('"') and stripped.endswith('"'):
        stripped = stripped[1:-1]

    return stripped.strip("\r\n")


def extract_mms_message(matched: pd.DataFrame, send_row: pd.Series) -> str:
    """상품 RAW 또는 소재 RAW의 MMS문구 컬럼에서 발송 문구를 읽습니다."""
    candidate_cols = ["MMS문구", "MMS 문구", "발송문구", "문구"]

    for col in candidate_cols:
        if col in matched.columns:
            for value in matched[col].tolist():
                cleaned = clean_mms_message(value)
                if cleaned:
                    return cleaned

    for col in candidate_cols:
        if col in send_row.index:
            cleaned = clean_mms_message(send_row.get(col))
            if cleaned:
                return cleaned

    return ""


def format_discount_percent(x):
    if pd.isna(x) or str(x).strip() in ["", "nan", "None"]:
        return ""
    try:
        value = float(str(x).replace("%", "").replace(",", "").strip())
        if abs(value) <= 1:
            value *= 100
        return f"{value:.0f}%"
    except (TypeError, ValueError):
        return str(x)


def format_integer_price(x):
    if pd.isna(x) or str(x).strip() in ["", "nan", "None"]:
        return ""
    try:
        return f"{float(str(x).replace(',', '')):,.0f}"
    except (TypeError, ValueError):
        return str(x)


# ─────────────────────────────────────────────────────────────────────────────
# 데이터 연결
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.markdown("## 📊 MMS AI Dashboard")

DEFAULT_GOOGLE_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1I8sAfs8kfMAFThHa_o-aeb2GLWLbFtxf3FxBhA8q-tQ/edit?gid=0#gid=0"
)

if "products" not in st.session_state:
    st.session_state.products = None
    st.session_state.sends = None
    st.session_state.source_name = None
    st.session_state.synced_at = None
    st.session_state.google_sync_error = None
    st.session_state.lowest = pd.DataFrame()
    st.session_state.auto_sync_attempted = False

source = st.sidebar.radio(
    "데이터 연결",
    ["구글시트 자동연동", "엑셀 업로드"],
    index=0,
)

if source == "구글시트 자동연동":
    sheet_url = st.sidebar.text_input(
        "🔗 구글시트 링크",
        value=DEFAULT_GOOGLE_SHEET_URL,
    )

    # 앱 최초 실행 시 자동 동기화
    if not st.session_state.auto_sync_attempted:
        st.session_state.auto_sync_attempted = True
        try:
            with st.spinner("구글시트 최신 데이터를 자동으로 불러오는 중입니다."):
                sync_google_sheet(sheet_url)
        except Exception as exc:
            st.session_state.google_sync_error = str(exc)

    if st.sidebar.button("🔄 지금 새로고침", use_container_width=True):
        try:
            with st.spinner("구글시트 최신 데이터를 다시 불러오는 중입니다."):
                sync_google_sheet(sheet_url, force=True)
            st.sidebar.success("최신 데이터로 갱신했습니다.")
            st.rerun()
        except Exception as exc:
            st.session_state.google_sync_error = str(exc)
            st.sidebar.error(f"새로고침 실패: {exc}")

    st.sidebar.caption("자동 갱신 주기: 5분")
    if st.session_state.google_sync_error:
        st.sidebar.warning(
            "구글시트 자동연동에 실패했습니다. "
            "공유 권한과 탭 이름(상품, 소재)을 확인하거나 엑셀 업로드를 이용해주세요."
        )
        with st.sidebar.expander("오류 상세"):
            st.code(st.session_state.google_sync_error)

else:
    uploaded = st.sidebar.file_uploader("📁 MMS 파일 업로드", type=["xlsx", "xlsm"])
    if uploaded is not None:
        try:
            products, sends, lowest = load_excel_bytes(uploaded.getvalue())
            st.session_state.products = products
            st.session_state.sends = sends
            st.session_state.lowest = lowest
            st.session_state.source_name = uploaded.name
            st.session_state.synced_at = datetime.now()
            st.session_state.google_sync_error = None
        except Exception as exc:
            st.sidebar.error(f"파일을 불러오지 못했습니다: {exc}")

if st.session_state.products is None or st.session_state.sends is None:
    st.info(
        "구글시트 자동연동을 확인 중입니다. 연결되지 않으면 "
        "왼쪽에서 엑셀 업로드로 전환해주세요."
    )
    st.stop()

products = st.session_state.products
sends = st.session_state.sends
lowest = st.session_state.get("lowest", pd.DataFrame())

st.markdown(
    f"""
    <div class="app-header">
        <div>
            <div class="app-title">MMS AI 대시보드</div>
            <div class="app-subtitle">MMS 일일·주간·월간 실적 및 상품 인사이트 통합 분석</div>
        </div>
        <div class="status-badge">
            현재 데이터: {st.session_state.get("source_name", "-")}
            · 마지막 동기화: {st.session_state.get("synced_at", "-")}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

menu = st.sidebar.radio(
    "메뉴",
    ["홈", "일일실적", "주간실적", "상품구분", "상품분석", "타겟분석", "설정"],
)

st.markdown('<div class="app-title">MMS AI Dashboard</div>', unsafe_allow_html=True)
sync_text = (
    st.session_state.synced_at.strftime("%Y-%m-%d %H:%M:%S")
    if st.session_state.synced_at
    else "-"
)
st.markdown(
    f'<div class="data-source">현재 데이터: {st.session_state.source_name} · '
    f'마지막 동기화: {sync_text}</div>',
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────────────────────────────────────
# 홈
# ─────────────────────────────────────────────────────────────────────────────
if menu == "홈":
    st.markdown('<div class="section-title">홈 · 기간별 실적</div>', unsafe_allow_html=True)

    monthly_all = aggregate_send(sends, "Monthly")
    weekly_all = aggregate_send(sends, "Weekly")
    daily_all = aggregate_send(sends, "Daily")

    # 월간 기간 필터
    c1, c2, c3 = st.columns([1.2, 1.2, 1.2])
    with c1:
        month_option = st.selectbox("월간 조회 기간", ["전체", "최근 3개월", "최근 6개월", "최근 12개월", "직접 선택"])
    month_labels = monthly_all["_label"].astype(str).tolist()
    start_month = end_month = None
    if month_option == "직접 선택" and month_labels:
        with c2:
            start_month = st.selectbox("시작월", month_labels, index=0)
        with c3:
            end_month = st.selectbox("종료월", month_labels, index=len(month_labels)-1)

    monthly = filter_monthly_period(monthly_all, month_option, start_month, end_month)

    weekly = weekly_all.copy()

    # KPI
    latest_df = monthly if not monthly.empty else monthly_all
    if not latest_df.empty:
        latest = latest_df.iloc[-1]
        cards = [
            ("주문금액", f"{fmt_num(latest['주문금액'])}원", delta_for_latest(latest_df, "주문금액")),
            ("발송건수", fmt_num(latest["발송건수"]), delta_for_latest(latest_df, "발송건수")),
            ("CTR", fmt_pct(latest["반응율(Uniq CTR)"]), delta_for_latest(latest_df, "반응율(Uniq CTR)", pp=True)),
            ("클릭 CVR", fmt_pct(latest["클릭 CVR"]), delta_for_latest(latest_df, "클릭 CVR", pp=True)),
            ("SPM", f"{latest['발송대비매출(SPM)']:.1f}", delta_for_latest(latest_df, "발송대비매출(SPM)")),
            ("발송당매출", f"{fmt_num(latest['발송당매출(발송횟수)'])}원", delta_for_latest(latest_df, "발송당매출(발송횟수)")),
        ]
        cols = st.columns(6)
        for col, (label, value, delta) in zip(cols, cards):
            col.markdown(
                f'<div class="metric-card"><div class="metric-label">{label}</div>'
                f'<div class="metric-value">{value}</div>'
                f'<div class="metric-delta">직전 대비 {delta}</div></div>',
                unsafe_allow_html=True,
            )

    # 월간 그래프와 표
    st.markdown('<div class="section-title">월별 SPM / 발송대비매출</div>', unsafe_allow_html=True)
    st.plotly_chart(
        trend_chart(monthly, "월별 SPM / 발송대비매출", "#fdbb00"),
        use_container_width=True,
        config={"displayModeBar": False},
    )
    st.dataframe(
        clean_identifier_columns(format_home_table_with_summary(monthly, "Monthly")),
        use_container_width=True,
        hide_index=True,
        height=400,
    )

    # 주간 그래프와 표 - 독립 조회 기간
    st.markdown('<div class="section-title">주간 SPM / 발송대비매출</div>', unsafe_allow_html=True)
    st.caption("주간 조회 기간")
    week_labels = weekly_all["_label"].astype(str).tolist()
    if week_labels:
        week_start_col, week_end_col = st.columns(2)
        with week_start_col:
            start_week = st.selectbox("시작 주차", week_labels, index=0, key="home_week_start")
        with week_end_col:
            end_week = st.selectbox("종료 주차", week_labels, index=len(week_labels)-1, key="home_week_end")
        weekly = filter_weekly_period(weekly_all, start_week, end_week)
    else:
        weekly = weekly_all.copy()
    if weekly.empty:
        st.info("선택한 기간의 주간 데이터가 없습니다.")
    else:
        st.plotly_chart(
            trend_chart(weekly, "주간 SPM / 발송대비매출", "#70ad47"),
            use_container_width=True,
            config={"displayModeBar": False},
        )
        st.dataframe(
            clean_identifier_columns(format_home_table_with_summary(weekly, "Weekly")),
            use_container_width=True,
            hide_index=True,
            height=520,
        )

    # 일간 표: 실제 날짜 필터
    st.markdown('<div class="section-title">Daily</div>', unsafe_allow_html=True)
    st.caption("일간 조회 기간")
    daily_start_col, daily_end_col = st.columns(2)
    with daily_start_col:
        daily_start_date = st.date_input(
            "시작일",
            value=sends["_date"].min().date(),
            min_value=sends["_date"].min().date(),
            max_value=sends["_date"].max().date(),
            format="YYYY/MM/DD",
            key="home_daily_start",
        )
    with daily_end_col:
        daily_end_date = st.date_input(
            "종료일",
            value=sends["_date"].max().date(),
            min_value=sends["_date"].min().date(),
            max_value=sends["_date"].max().date(),
            format="YYYY/MM/DD",
            key="home_daily_end",
        )
    if daily_start_date > daily_end_date:
        daily_start_date, daily_end_date = daily_end_date, daily_start_date
    daily_source = sends[
        (sends["_date"].dt.date >= daily_start_date)
        & (sends["_date"].dt.date <= daily_end_date)
    ].copy()
    daily = aggregate_send(daily_source, "Daily")
    if daily.empty:
        st.info("선택한 기간의 일간 데이터가 없습니다.")
    else:
        st.dataframe(
            clean_identifier_columns(format_home_table_with_summary(daily, "Daily")),
            use_container_width=True,
            hide_index=True,
            height=480,
        )


# ─────────────────────────────────────────────────────────────────────────────
# 일일실적
# ─────────────────────────────────────────────────────────────────────────────
elif menu == "일일실적":
    st.markdown('<div class="section-title">📊 일일실적 분석</div>', unsafe_allow_html=True)
    dates = sorted(products["_date"].dt.date.unique(), reverse=True)
    selected_date = st.selectbox("📅 날짜 선택", dates)

    pday = products[products["_date"].dt.date == selected_date].copy()
    sday = sends[sends["_date"].dt.date == selected_date].copy()
    pday = merge_lowest_price(pday, lowest)

    if pday.empty and sday.empty:
        st.info("선택한 날짜의 데이터가 없습니다.")
        st.stop()

    total_amount = pday["주문금액"].sum()
    total_orders = pday["주문건수"].sum()
    total_qty = pday["주문수량"].sum()
    send_col = first_col(sday, ["발송 성공 건수", "총 발송 건수"])
    click_col = first_col(sday, ["클릭 수(uniq)", "클릭 수"])
    send_success = sday[send_col].sum() if send_col else 0
    clicks = sday[click_col].sum() if click_col else 0

    cards = [
        ("주문건수", f"{int(total_orders):,}건"),
        ("주문수량", f"{int(total_qty):,}개"),
        ("주문금액", f"{int(total_amount):,}원"),
        ("CTR", f"{(clicks/send_success*100 if send_success else 0):.1f}%"),
        ("CVR", f"{(total_orders/clicks*100 if clicks else 0):.1f}%"),
        ("SPM", f"{(total_amount/send_success if send_success else 0):.1f}"),
    ]
    metric_cols = st.columns(len(cards))
    for c, (label, value) in zip(metric_cols, cards):
        c.markdown(
            f'<div class="metric-card"><div class="metric-label">{label}</div>'
            f'<div class="metric-value">{value}</div></div>',
            unsafe_allow_html=True,
        )

    # 오전/오후 또는 소재 단위로 분리
    if "시간대" in sday.columns:
        sday["_sort_time"] = pd.to_datetime(sday["시간대"].astype(str), errors="coerce")
        sday = sday.sort_values(["_sort_time", "소재"] if "소재" in sday.columns else ["_sort_time"])

    for idx, send_row in sday.iterrows():
        time_value = str(send_row.get("시간대", ""))
        material = str(send_row.get("소재", ""))
        part_title = f"{time_value} · {material}" if material else time_value
        st.markdown(f'<div class="section-title">🕒 {part_title}</div>', unsafe_allow_html=True)

        matched = pday.copy()
        if "시간대" in matched.columns and time_value:
            matched = matched[matched["시간대"].astype(str) == time_value]
        if "소재" in matched.columns and material:
            material_match = matched[matched["소재"].astype(str) == material]
            if not material_match.empty:
                matched = material_match

        if matched.empty:
            st.info("해당 발송 건과 연결된 상품 실적이 없습니다.")
            continue

        asset_key = daily_asset_key(selected_date, time_value)
        image_path = find_daily_image(asset_key)
        message_text = extract_mms_message(matched, send_row)

        st.markdown('<div class="subsection-title">발송 소재</div>', unsafe_allow_html=True)
        asset_image_col, asset_text_col = st.columns([1, 1.35], gap="medium")

        with asset_image_col:
            if image_path is not None:
                import base64
                mime_map = {
                    ".jpg": "image/jpeg",
                    ".jpeg": "image/jpeg",
                    ".png": "image/png",
                    ".webp": "image/webp",
                }
                mime = mime_map.get(image_path.suffix.lower(), "image/jpeg")
                encoded = base64.b64encode(image_path.read_bytes()).decode("utf-8")
                st.markdown(
                    f"""
                    <div class="asset-card asset-image-card">
                        <img src="data:{mime};base64,{encoded}" alt="{image_path.name}">
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"""
                    <div class="asset-card asset-image-card">
                        <div class="asset-empty">
                            images 폴더에<br>
                            {asset_key}_ 로 시작하는 이미지가 없습니다.
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        with asset_text_col:
            import html
            if message_text:
                message_body = html.escape(message_text).replace("\n", "<br>")
            else:
                message_body = (
                    "로우데이터의 MMS문구 컬럼에 문구를 입력하면 "
                    "이곳에 자동으로 표시됩니다."
                )

            st.markdown(
                f"""
                <div class="asset-card asset-message-card" style="padding-top:18px;">
                    {message_body}
                </div>
                """,
                unsafe_allow_html=True,
            )

        # 발송 통계: 요청 컬럼만 표시
        send_count = float(send_row.get(send_col, 0)) if send_col else 0
        click_count = float(send_row.get(click_col, 0)) if click_col else 0
        orders = float(send_row.get("주문건수", 0))
        amount = float(send_row.get("주문금액", 0))
        send_view = pd.DataFrame([{
            "성별": send_row.get("성별", ""),
            "연령": send_row.get("연령", ""),
            "SEG": send_row.get("SEG", ""),
            "소재": material,
            "URL": send_row.get("URL", ""),
            "발송건수": fmt_num(send_count),
            "클릭수": fmt_num(click_count),
            "CTR": fmt_pct(click_count / send_count if send_count else 0),
            "CVR": fmt_pct(orders / click_count if click_count else 0),
            "객단가": fmt_num(amount / orders if orders else 0),
            "SPM": f"{(amount/send_count if send_count else 0):.1f}",
        }])
        st.markdown('<div class="subsection-title">발송 통계</div>', unsafe_allow_html=True)
        st.dataframe(
            clean_identifier_columns(send_view),
            use_container_width=True,
            hide_index=True,
        )

        matched["상품등급"] = matched["주문금액"].apply(product_grade)
        sort_cols = [c for c in ["전시순서", "상품명"] if c in matched.columns]
        if sort_cols:
            matched = matched.sort_values(sort_cols)

        display_cols = [
            "전시순서", "MD", "알파코드", "쇼라코드", "상품명",
            "정상가", "멤버십혜택가", "할인율", "추가노출",
            "주문건수", "주문수량", "주문금액", "최저가",
            "가격차이", "최저가 확보"
        ]
        product_view = matched[[c for c in display_cols if c in matched.columns]].copy()

        if "할인율" in product_view.columns:
            product_view["할인율"] = product_view["할인율"].map(format_discount_percent)
        for price_col in ["정상가", "멤버십혜택가", "최저가", "가격차이"]:
            if price_col in product_view.columns:
                product_view[price_col] = product_view[price_col].map(format_integer_price)

        # 합계행
        total_row = {c: "" for c in product_view.columns}
        first_display = product_view.columns[0]
        total_row[first_display] = "합계"
        for c in ["주문건수", "주문수량", "주문금액"]:
            if c in product_view.columns:
                total_row[c] = matched[c].sum()
        product_view = pd.concat([product_view, pd.DataFrame([total_row])], ignore_index=True)

        st.markdown('<div class="subsection-title">상품 실적</div>', unsafe_allow_html=True)
        st.dataframe(
            clean_identifier_columns(product_view),
            use_container_width=True,
            hide_index=True,
            height=280,
        )

        st.markdown('<div class="subsection-title">주요 인사이트</div>', unsafe_allow_html=True)
        insight_lines = [make_insight(r, products) for _, r in matched.iterrows()]
        st.markdown(
            f'<div class="insight-box">{"<br>".join(insight_lines)}</div>',
            unsafe_allow_html=True,
        )

    # 최저가 원본 탭 요약
    with st.expander("최저가 분석"):
        if lowest is None or lowest.empty:
            st.info("상품 RAW의 '발송일 최저가' 컬럼을 입력하면 가격 비교가 자동 반영됩니다.")
        else:
            lowest_day = pday[[c for c in [
                "일자", "쇼라코드", "알파코드", "상품명", "멤버십혜택가",
                "최저가", "가격차이", "최저가 확보"
            ] if c in pday.columns]]
            st.dataframe(lowest_day, use_container_width=True, hide_index=True)

    all_insights = [make_insight(r, products) for _, r in pday.iterrows()]
    ppt = build_ppt(
        f"{selected_date} MMS 일일실적",
        all_insights,
        pday[[c for c in ["상품명", "주문건수", "주문수량", "주문금액", "최저가"] if c in pday.columns]],
    )
    st.download_button(
        "📥 PPT 다운로드",
        ppt,
        file_name=f"{selected_date}_MMS_일일실적.pptx",
    )


# ─────────────────────────────────────────────────────────────────────────────
# 주간실적
# ─────────────────────────────────────────────────────────────────────────────
elif menu == "주간실적":
    st.markdown('<div class="section-title">📈 주간실적 분석</div>', unsafe_allow_html=True)

    # 연도 → 주차 순서로 선택하여 같은 주차명이 연도별로 섞이지 않도록 함
    weekly_years = sorted(products["_year"].dropna().astype(int).unique(), reverse=True)
    selected_year = st.selectbox("연도 선택", weekly_years)

    year_products = products[products["_year"] == selected_year].copy()
    year_sends = sends[sends["_year"] == selected_year].copy()
    week_order = (
        year_products.groupby("주차")["_date"].min().sort_values()
        if not year_products.empty else pd.Series(dtype="datetime64[ns]")
    )
    weeks = [str(x) for x in week_order.index]
    if not weeks:
        st.info("선택한 연도의 주차 데이터가 없습니다.")
        st.stop()

    week = st.selectbox("주차 선택", weeks, index=len(weeks)-1)
    pw = year_products[year_products["주차"].astype(str) == week].copy()
    sw = year_sends[year_sends["주차"].astype(str) == week].copy()

    if pw.empty or sw.empty:
        st.info("선택한 연도·주차의 상품 또는 소재 데이터가 없습니다.")
        st.stop()

    send_col = first_col(sw, ["발송 성공 건수", "총 발송 건수"])
    click_col = first_col(sw, ["클릭 수(uniq)", "클릭 수"])
    send_count = sw[send_col].sum()
    click_count = sw[click_col].sum()
    order_count = sw["주문건수"].sum()
    amount = sw["주문금액"].sum()
    ctr = click_count / send_count if send_count else 0
    cvr = order_count / click_count if click_count else 0
    aov = amount / order_count if order_count else 0
    spm = amount / send_count if send_count else 0

    # 전주 대비 카드 증감
    year_week_names = [str(x) for x in (
        year_sends.groupby("주차")["_date"].min().sort_values().index
    )]
    prev_sw = pd.DataFrame()
    if week in year_week_names and year_week_names.index(week) > 0:
        prev_week = year_week_names[year_week_names.index(week) - 1]
        prev_sw = year_sends[year_sends["주차"].astype(str) == prev_week]

    def prev_metric(column, default=0):
        return float(prev_sw[column].sum()) if not prev_sw.empty and column in prev_sw.columns else default

    prev_send = prev_metric(send_col)
    prev_click = prev_metric(click_col)
    prev_orders = prev_metric("주문건수")
    prev_amount = prev_metric("주문금액")
    prev_ctr = prev_click / prev_send if prev_send else 0
    prev_cvr = prev_orders / prev_click if prev_click else 0
    prev_aov = prev_amount / prev_orders if prev_orders else 0
    prev_spm = prev_amount / prev_send if prev_send else 0

    cards = [
        ("발송횟수", f"{len(sw):,}회", weekly_delta(len(sw), len(prev_sw)) if not prev_sw.empty else "-"),
        ("상품수", f"{len(pw):,}건", "-"),
        ("발송건수", f"{int(send_count):,}", weekly_delta(send_count, prev_send) if prev_sw is not None and not prev_sw.empty else "-"),
        ("주문건수", f"{int(order_count):,}건", weekly_delta(order_count, prev_orders) if not prev_sw.empty else "-"),
        ("주문금액", f"{int(amount):,}원", weekly_delta(amount, prev_amount) if not prev_sw.empty else "-"),
        ("CTR", f"{ctr*100:.1f}%", weekly_delta(ctr, prev_ctr, pp=True) if not prev_sw.empty else "-"),
        ("CVR", f"{cvr*100:.1f}%", weekly_delta(cvr, prev_cvr, pp=True) if not prev_sw.empty else "-"),
        ("객단가", f"{int(aov):,}원", weekly_delta(aov, prev_aov) if not prev_sw.empty else "-"),
        ("SPM", f"{spm:.1f}", weekly_delta(spm, prev_spm) if not prev_sw.empty else "-"),
    ]
    card_cols = st.columns(5)
    for i, (label, value, delta) in enumerate(cards):
        with card_cols[i % 5]:
            st.markdown(
                f'<div class="metric-card"><div class="metric-label">{label}</div>'
                f'<div class="metric-value">{value}</div>'
                f'<div class="metric-delta">전주 대비 {delta}</div></div>',
                unsafe_allow_html=True,
            )
        if i == 4:
            card_cols = st.columns(4)

    chart_left, chart_right = st.columns(2)
    with chart_left:
        st.plotly_chart(
            weekly_product_chart(sw),
            use_container_width=True,
            config={"displayModeBar": False},
        )
    with chart_right:
        st.plotly_chart(
            weekly_send_chart(sw),
            use_container_width=True,
            config={"displayModeBar": False},
        )

    # 대·중카테고리 편성 및 주문 비중
    cat_left, cat_right = st.columns(2)

    with cat_left:
        big_table = category_summary_table(pw, "대카", week, selected_year)
        st.plotly_chart(
            category_pie_chart(big_table, "대카테고리 주문비중"),
            use_container_width=True,
            config={"displayModeBar": False},
        )
        st.markdown("**대카테고리 편성 및 주문 비중**")
        st.dataframe(
            clean_identifier_columns(weekly_display_format(big_table)),
            use_container_width=True,
            hide_index=True,
            height=430,
        )

    with cat_right:
        mid_table = category_summary_table(pw, "중카", week, selected_year)
        st.plotly_chart(
            category_pie_chart(mid_table, "중카테고리 주문비중"),
            use_container_width=True,
            config={"displayModeBar": False},
        )
        st.markdown("**중카테고리 편성 및 주문 비중**")
        st.dataframe(
            clean_identifier_columns(weekly_display_format(mid_table)),
            use_container_width=True,
            hide_index=True,
            height=560,
        )

    tabs = st.tabs([
        "주간실적 분석", "상품 실적", "소재 실적",
        "SEG 실적", "요일 실적", "시간대 실적", "MMS 상품"
    ])

    with tabs[0]:
        report = build_weekly_analysis(
            week, selected_year, pw, sw, products, sends
        )
        report = re.sub(r"\n{2,}", "\n", report).strip()
        st.markdown(
            f'<div class="insight-box">{report}</div>',
            unsafe_allow_html=True,
        )

    with tabs[1]:
        total_amount = pw["주문금액"].sum()
        pw["주문비중"] = pw["주문금액"] / total_amount if total_amount else 0
        sort_cols = [c for c in ["_date", "시간대", "전시순서"] if c in pw.columns]
        product_sorted = pw.sort_values(sort_cols) if sort_cols else pw
        cols = [
            "일자", "요일", "시간대", "성별", "연령", "소재",
            "전시순서", "추가노출", "상품명", "멤버십혜택가",
            "주문건수", "주문수량", "주문금액", "주문비중"
        ]
        view = product_sorted[[c for c in cols if c in product_sorted.columns]].copy()
        total = {c: "" for c in view.columns}
        total[view.columns[0]] = "총합계"
        for c in ["주문건수", "주문수량", "주문금액"]:
            if c in view.columns:
                total[c] = product_sorted[c].sum()
        if "주문비중" in view.columns:
            total["주문비중"] = 1.0
        view = pd.concat([view, pd.DataFrame([total])], ignore_index=True)
        raw_amounts = (
            list(pd.to_numeric(view["주문금액"], errors="coerce"))
            if "주문금액" in view.columns else []
        )
        formatted_view = clean_identifier_columns(weekly_display_format(view))
        styled_view = formatted_view.style.apply(
            lambda _: style_weekly_product_rows(formatted_view, raw_amounts),
            axis=None,
        )
        st.dataframe(
            styled_view,
            use_container_width=True,
            hide_index=True,
            height=680,
        )

    with tabs[2]:
        material = pd.DataFrame({
            "연도": selected_year,
            "주차": sw["주차"],
            "일자": sw.get("일자", sw["_date"].dt.strftime("%m%d")),
            "요일": sw.get("요일", ""),
            "시간대": sw.get("시간대", ""),
            "소재": sw.get("소재", ""),
            "성별": sw.get("성별", ""),
            "연령": sw.get("연령", ""),
            "상품수": sw.get("상품수", 0),
            "발송성공건수": sw[send_col],
            "클릭수(uniq)": sw[click_col],
            "CTR(uniq)": safe_div(sw[click_col], sw[send_col]),
            "CVR(클릭>구매)": safe_div(sw["주문건수"], sw[click_col]),
            "객단가": safe_div(sw["주문금액"], sw["주문건수"]),
            "SPM": safe_div(sw["주문금액"], sw[send_col]),
            "주문건수": sw["주문건수"],
            "주문수량": sw["주문수량"],
            "주문금액": sw["주문금액"],
        }).fillna(0)
        st.dataframe(
            clean_identifier_columns(weekly_display_format(material)),
            use_container_width=True, hide_index=True, height=570
        )

    with tabs[3]:
        seg = grouped_send_table(sw, ["성별", "연령"])
        seg.insert(0, "주차", week)
        seg.insert(0, "연도", selected_year)
        cols = ["연도", "주차", "성별", "연령", "발송횟수", "CTR(uniq)", "CVR(클릭>구매)", "객단가", "SPM", "주문금액", "발송당매출(발송횟수)"]
        st.dataframe(
            clean_identifier_columns(weekly_display_format(seg[cols])),
            use_container_width=True, hide_index=True
        )

    with tabs[4]:
        weekday = grouped_send_table(sw, ["요일"])
        weekday.insert(0, "주차", week)
        weekday.insert(0, "연도", selected_year)
        cols = ["연도", "주차", "요일", "발송횟수", "CTR(uniq)", "CVR(클릭>구매)", "객단가", "SPM", "주문금액", "발송당매출(발송횟수)"]
        st.dataframe(
            clean_identifier_columns(weekly_display_format(weekday[cols])),
            use_container_width=True, hide_index=True
        )

    with tabs[5]:
        time_df = grouped_send_table(sw, ["시간대"])
        time_df.insert(0, "주차", week)
        time_df.insert(0, "연도", selected_year)
        cols = ["연도", "주차", "시간대", "발송횟수", "CTR(uniq)", "CVR(클릭>구매)", "객단가", "SPM", "주문금액", "발송당매출(발송횟수)"]
        st.dataframe(
            clean_identifier_columns(weekly_display_format(time_df[cols])),
            use_container_width=True, hide_index=True
        )

    with tabs[6]:
        rank = pw.groupby(["쇼라코드", "상품명"], as_index=False).agg(
            발송횟수=("상품명", "size"),
            주문건수=("주문건수", "sum"),
            주문수량=("주문수량", "sum"),
            주문금액=("주문금액", "sum"),
        ).sort_values("주문금액", ascending=False)
        rank.insert(0, "주차", week)
        rank.insert(0, "연도", selected_year)
        st.dataframe(
            clean_identifier_columns(weekly_display_format(rank)),
            use_container_width=True, hide_index=True, height=680
        )

    weekly_summary = [
        f"{selected_year}년 {week} 주문금액 {fmt_num(pw['주문금액'].sum())}원",
        f"발송횟수 {len(sw)}회 / 상품수 {len(pw)}건",
        f"핵심 상품 {sum(pw['주문금액'].apply(product_grade) == '핵심 상품')}개",
    ]
    ppt = build_ppt(
        f"{selected_year}년 {week} MMS 주간실적",
        weekly_summary,
        rank[[c for c in ["상품명", "발송횟수", "주문건수", "주문수량", "주문금액"] if c in rank.columns]],
    )
    st.download_button(
        "📥 주간실적 PPT 다운로드",
        ppt,
        file_name=f"{selected_year}_{week}_MMS_주간실적.pptx",
    )


# ─────────────────────────────────────────────────────────────────────────────
# 상품구분
# ─────────────────────────────────────────────────────────────────────────────
elif menu == "상품구분":
    st.markdown('<div class="section-title">상품구분</div>', unsafe_allow_html=True)
    date_range = st.date_input(
        "기간 선택",
        [products["_date"].min().date(), products["_date"].max().date()],
    )

    if len(date_range) == 2:
        start, end = date_range
        filt = products[
            (products["_date"].dt.date >= start)
            & (products["_date"].dt.date <= end)
        ].copy()
    else:
        filt = products.copy()

    grouped = filt.groupby(["쇼라코드", "상품명"], as_index=False).agg(
        운영횟수=("상품명", "size"),
        주문건수=("주문건수", "sum"),
        주문수량=("주문수량", "sum"),
        주문금액=("주문금액", "sum"),
        최고실적=("주문금액", "max"),
        평균실적=("주문금액", "mean"),
    )
    grouped["등급"] = grouped["주문금액"].apply(product_grade)

    grade_filter = st.multiselect("상품 등급", GRADE_ORDER, default=GRADE_ORDER)
    case_filter = st.multiselect("특수 사례", CASE_ORDER)

    rows = []
    for _, group_row in grouped.iterrows():
        name = group_row["상품명"]
        hist = filt[filt["상품명"] == name].sort_values("_date")
        cases = []
        insights = []
        for _, row in hist.iterrows():
            cases += classify_cases(row, products)
            insights.append(make_insight(row, products))
        cases = list(dict.fromkeys(cases))
        rows.append({
            **group_row.to_dict(),
            "사례": ", ".join(cases),
            "인사이트": "\n".join(insights),
        })

    result = pd.DataFrame(rows)
    result = result[result["등급"].isin(grade_filter)]
    if case_filter:
        result = result[result["사례"].apply(lambda x: any(case in x for case in case_filter))]

    search = st.text_input("상품명 검색")
    if search:
        result = result[result["상품명"].astype(str).str.contains(search, case=False, na=False)]

    st.dataframe(
        result.drop(columns=["인사이트"]),
        use_container_width=True,
        hide_index=True,
        height=500,
    )

    if not result.empty:
        selected_product = st.selectbox("인사이트 확인 상품", result["상품명"].tolist())
        insight = result.loc[result["상품명"] == selected_product, "인사이트"].iloc[0]
        st.markdown(f'<div class="insight-box">{insight}</div>', unsafe_allow_html=True)


elif menu == "상품분석":
    names = sorted(products["상품명"].dropna().astype(str).unique())
    name = st.selectbox("상품 선택", names)
    hist = products[products["상품명"].astype(str) == name].sort_values("_date")
    st.dataframe(hist, use_container_width=True, hide_index=True)
    st.markdown('<div class="section-title">상품 이력 인사이트</div>', unsafe_allow_html=True)
    insights = "\n".join(make_insight(row, products) for _, row in hist.iterrows())
    st.markdown(f'<div class="insight-box">{insights}</div>', unsafe_allow_html=True)


elif menu == "타겟분석":
    send_col = first_col(sends, ["발송 성공 건수", "총 발송 건수"])
    click_col = first_col(sends, ["클릭 수(uniq)", "클릭 수"])
    group = sends.groupby(["성별", "연령", "SEG"], as_index=False).agg(
        발송횟수=("소재", "size"),
        발송건수=(send_col, "sum"),
        클릭수=(click_col, "sum"),
        주문건수=("주문건수", "sum"),
        주문금액=("주문금액", "sum"),
    )
    group["CTR"] = safe_div(group["클릭수"], group["발송건수"])
    group["CVR"] = safe_div(group["주문건수"], group["클릭수"])
    group["SPM"] = safe_div(group["주문금액"], group["발송건수"])
    group["발송당매출"] = safe_div(group["주문금액"], group["발송횟수"])
    st.dataframe(
        group.sort_values("발송당매출", ascending=False),
        use_container_width=True,
        hide_index=True,
    )


else:
    st.markdown('<div class="section-title">설정</div>', unsafe_allow_html=True)
    st.table(pd.DataFrame({
        "주문금액": ["100만원 미만", "100~200만원", "200~300만원", "300~500만원", "500만원 이상"],
        "등급": ["🔴 부진 상품", "🟠 관찰 상품", "🟡 유망 상품", "🟢 우수 상품", "🔵 핵심 상품"],
    }))
    st.caption("구글시트는 '상품', '소재' 탭이 필수이며 상품 탭의 '발송일 최저가' 컬럼을 지원합니다. 링크 공유 권한이 필요합니다.")
