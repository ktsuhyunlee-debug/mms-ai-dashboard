# -*- coding: utf-8 -*-
"""쇼핑라운지 GA4 인사이트 · compact, read-only dashboard.

Run: streamlit run ga4_app.py

Uses the *existing* Streamlit secrets:
[google_oauth] client_id/client_secret/redirect_uri/state_signing_secret
[ga4] property_id = "544313553"

Analytical boundaries:
- Property contains all KT membership; title-filtered views ≠ guaranteed full
  lounge user/session segment. All reports identify page-context scope.
- Purchase events only show when they occur on a lounge-labelled screen.
  Their association with a promotion/product is NOT automatically proved.
- User counts cannot be summed across dimension rows without deduplication.
- GA4 UI sequence funnels are not faked from unrelated event totals.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import html
import io
import json
import re
import secrets as pysecrets
from collections import OrderedDict

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
from requests_oauthlib import OAuth2Session

st.set_page_config(page_title="쇼핑라운지 | GA4 인사이트", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
GA4_API_URL = "https://analyticsdata.googleapis.com/v1beta"
READONLY_SCOPE = "https://www.googleapis.com/auth/analytics.readonly"
DEFAULT_PROPERTY = "544313553"
NAV = [
    "01  종합 현황", "02  이용 고객", "03  유입 경로", "04  페이지 분석",
    "05  기획전 성과", "06  상품 성과", "07  고객 행동", "08  캠페인 분석",
]
BLUE = "#4255C7"
TEAL = "#0F9D99"
MUTED = "#7E8AA3"
COLORS = [BLUE, TEAL, "#F49B55", "#8567C5", "#D35B82", "#84A4BF", "#4BBA91"]

CSS = """
<style>
:root {--primary:#4255c7;}
.block-container {padding-top:.75rem!important; padding-bottom:1.1rem!important;
    max-width:1540px!important;}
section[data-testid="stSidebar"] {background:#f5f7fc;}
section[data-testid="stSidebar"] .block-container {padding-top:1rem!important;}
div[data-testid="stVerticalBlock"] {gap:.55rem;}
div[data-testid="stHorizontalBlock"] {gap:.6rem;}
div[data-testid="stPlotlyChart"] {margin-top:-8px;}
div[data-testid="stDataFrame"] {margin-top:-6px;}
.stTabs [data-baseweb="tab-list"] {gap:6px;}
.stTabs [data-baseweb="tab"] {height:37px;padding:6px 14px;}
[data-testid="stWidgetLabel"] p {font-size:.84rem!important;}
#MainMenu, footer {visibility:hidden;}
.app-title {font-size:1.5rem;font-weight:800;letter-spacing:-.045em;line-height:1.3;
 color:#233047;margin:0;}
.app-sub {color:#7c8ba3;font-size:.8rem;margin:2px 0 12px;}
.section-heading {font-weight:760;font-size:1rem; color:#26344d;margin:2px 0 8px;}
.section-helper {font-size:.76rem;color:#8491a6;margin-top:-8px;}
.kpi {border:1px solid #e8edf6;background:#fff;border-radius:13px;
 box-shadow:0 1px 5px rgba(28,49,91,.045);padding:11px 12px 10px;min-height:79px;}
.kpi-label {font-size:.75rem;color:#78869d;font-weight:600;margin-bottom:5px;}
.kpi-value {font-size:clamp(1rem,1.6vw,1.52rem);color:#22324b;letter-spacing:-.04em;
 font-weight:800;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.kpi-delta {font-size:.68rem;color:#6f819c;margin-top:4px;}
.kpi-up {color:#0b9b75;} .kpi-down{color:#df6262;}
.pill {display:inline-block;background:#eef2ff;color:#445bc1;padding:3px 8px;
 border-radius:100px;font-size:.68rem;font-weight:700;}
.microcopy {font-size:.74rem;color:#7d8ba3;line-height:1.55;}
@media(max-width:900px){.block-container {padding-top:.5rem!important;}
 .kpi {padding:9px;min-height:72px;} .kpi-value {font-size:.98rem;}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ───────────────────── Google sign-in (existing OAuth flow) ────────────────────
def load_config():
    try:
        oauth = st.secrets["google_oauth"]
        ga4 = st.secrets.get("ga4", {})
        vals = {
            "client_id": str(oauth["client_id"]),
            "client_secret": str(oauth["client_secret"]),
            "redirect_uri": str(oauth["redirect_uri"]),
            "state_signing_secret": str(oauth["state_signing_secret"]),
            "property_id": str(ga4.get("property_id", DEFAULT_PROPERTY)),
            "keyword": str(ga4.get("page_keyword", "쇼핑라운지")),
        }
        return vals if all(vals.values()) else None
    except (KeyError, FileNotFoundError):
        return None


def signer(secret):
    return URLSafeTimedSerializer(secret, salt="ga4-oauth-state-v1")


def auth_url(cfg):
    state = signer(cfg["state_signing_secret"]).dumps({"nonce": pysecrets.token_urlsafe(24)})
    client = OAuth2Session(
        client_id=cfg["client_id"], scope=[READONLY_SCOPE],
        redirect_uri=cfg["redirect_uri"], state=state,
    )
    url, _ = client.authorization_url(AUTH_URL, access_type="online", prompt="select_account")
    return url


def receive_callback(cfg):
    params = st.query_params
    if "error" in params:
        st.query_params.clear()
        st.error("Google 계정 인증이 취소되었어요. 다시 연결해 주세요.")
        return
    code, state = params.get("code"), params.get("state")
    if not code:
        return
    try:
        if not state:
            raise BadSignature("missing state")
        signer(cfg["state_signing_secret"]).loads(state, max_age=600)
        client = OAuth2Session(
            client_id=cfg["client_id"], scope=[READONLY_SCOPE],
            redirect_uri=cfg["redirect_uri"], state=state,
        )
        token = client.fetch_token(
            TOKEN_URL, code=code, client_secret=cfg["client_secret"],
            include_client_id=True, timeout=20,
        )
        access = token.get("access_token")
        if not access:
            raise ValueError("no access token")
        st.session_state["ga4_access_token"] = access
        st.session_state["ga4_token_expires_at"] = dt.datetime.now(dt.timezone.utc) + dt.timedelta(
            seconds=max(0, int(token.get("expires_in", 3600)) - 60)
        )
        st.session_state.pop("ga4_report_cache", None)
        st.query_params.clear()
        st.rerun()
    except (BadSignature, SignatureExpired):
        st.query_params.clear()
        st.error("로그인 요청이 만료되었어요. 다시 로그인해 주세요.")
    except Exception as exc:
        st.query_params.clear()
        st.error(f"로그인 실패: {type(exc).__name__}. Google OAuth 설정을 확인해 주세요.")


def current_token():
    token = st.session_state.get("ga4_access_token")
    expiry = st.session_state.get("ga4_token_expires_at") or st.session_state.get("ga4_expiry")
    if not token or not expiry:
        return None
    if dt.datetime.now(dt.timezone.utc) >= expiry:
        disconnect()
        return None
    return token


def disconnect():
    for k in ["ga4_access_token", "ga4_token_expires_at", "ga4_expiry", "ga4_report_cache"]:
        st.session_state.pop(k, None)


# ─────────────────────── GA4 reporting and safeguards ────────────────────────
def screen_filter(keyword):
    return {"filter": {"fieldName": "unifiedScreenName", "stringFilter": {
        "matchType": "CONTAINS", "value": keyword, "caseSensitive": False}}}


def eq_filter(field, value):
    return {"filter": {"fieldName": field, "stringFilter": {
        "matchType": "EXACT", "value": value}}}


def and_filter(*filters):
    valid = [f for f in filters if f]
    if not valid:
        return None
    if len(valid) == 1:
        return valid[0]
    return {"andGroup": {"expressions": valid}}


def report(cfg, token, start, end, dims, mets, filt=None, limit=200, sort=None):
    """Small per-session cache; never caches across users or persists tokens to disk."""
    payload = {
        "dateRanges": [{"startDate": str(start), "endDate": str(end)}],
        "dimensions": [{"name": x} for x in dims],
        "metrics": [{"name": x} for x in mets],
        "limit": str(limit),
    }
    if filt:
        payload["dimensionFilter"] = filt
    if sort:
        payload["orderBys"] = [{"metric": {"metricName": sort}, "desc": True}]
    ck = hashlib.sha256((cfg["property_id"] + json.dumps(payload, sort_keys=True, ensure_ascii=False)).encode()).hexdigest()
    cache = st.session_state.setdefault("ga4_report_cache", OrderedDict())
    now = dt.datetime.now(dt.timezone.utc)
    if ck in cache:
        ts, data = cache[ck]
        if (now - ts).total_seconds() <= 300:
            return data
        del cache[ck]
    try:
        res = requests.post(
            f"{GA4_API_URL}/properties/{cfg['property_id']}:runReport",
            json=payload,
            headers={"Authorization": f"Bearer {token}"},
            timeout=30,
        )
        if not res.ok:
            try:
                msg = str(res.json().get("error", {}).get("message", "요청을 확인해주세요"))
            except ValueError:
                msg = "알 수 없는 서버 오류"
            raise RuntimeError(f"GA4 API {res.status_code}: {msg[:280]}")
        data = res.json()
    except requests.RequestException as exc:
        raise RuntimeError(f"GA4 통신 오류: {type(exc).__name__}") from None
    cache[ck] = (now, data)
    while len(cache) > 50:
        cache.popitem(last=False)
    return data


def to_frame(result):
    dim = [x["name"] for x in result.get("dimensionHeaders", [])]
    met = [x["name"] for x in result.get("metricHeaders", [])]
    rows = []
    for r in result.get("rows", []):
        a = [v.get("value", "") for v in r.get("dimensionValues", [])]
        b = [v.get("value", "") for v in r.get("metricValues", [])]
        rows.append(dict(zip(dim + met, a + b)))
    df = pd.DataFrame(rows, columns=dim + met)
    for k in met:
        df[k] = pd.to_numeric(df[k], errors="coerce").fillna(0)
    return df


def fetch(cfg, token, start, end, dims, mets, filt=None, limit=200, sort=None, quiet=False):
    try:
        raw = report(cfg, token, start, end, dims, mets, filt, limit, sort)
        df = to_frame(raw)
        if raw.get("metadata", {}).get("subjectToThresholding"):
            st.caption("⚠️ GA4 개인정보 보호 임계값이 적용된 보고서예요.")
        if raw.get("metadata", {}).get("dataLossFromOtherRow"):
            st.caption("⚠️ 일부 항목은 GA4의 (other) 행으로 합쳐졌을 수 있어요.")
        if raw.get("rowCount", 0) > limit:
            st.caption(f"⚠️ 총 {raw['rowCount']:,}개 항목 중 상위 {limit:,}개만 표시해요.")
        return df
    except RuntimeError as exc:
        if not quiet:
            st.warning(str(exc))
        return pd.DataFrame(columns=dims + mets)


def lounge(cfg):
    return screen_filter(cfg["keyword"])


def shop_fetch(cfg, token, start, end, dims, mets, other=None, limit=200, sort=None, quiet=False):
    return fetch(cfg, token, start, end, dims, mets,
                 and_filter(lounge(cfg), other), limit=limit, sort=sort, quiet=quiet)


def val(df, key):
    if df.empty or key not in df.columns:
        return None
    return float(df.iloc[0][key])


def get_summary(cfg, token, start, end):
    df = shop_fetch(cfg, token, start, end, [],
        ["activeUsers", "screenPageViews", "sessions", "newUsers", "userEngagementDuration"],
        quiet=True)
    return {k: val(df, k) for k in ["activeUsers", "screenPageViews", "sessions", "newUsers", "userEngagementDuration"]}


def get_purchase(cfg, token, start, end):
    """Purchases *logged on a lounge-labelled screen*; not sales attributed to campaigns/items."""
    df = shop_fetch(cfg, token, start, end, [],
        ["ecommercePurchases", "purchaseRevenue", "totalPurchasers"],
        eq_filter("eventName", "purchase"), quiet=True)
    if df.empty or (val(df, "ecommercePurchases") or 0) <= 0:
        return None
    return {k: val(df, k) for k in ["ecommercePurchases", "purchaseRevenue", "totalPurchasers"]}


def get_dates(cfg, token, start, end, metric, title=None):
    extra = eq_filter("unifiedScreenName", title) if title else None
    if metric in ("ecommercePurchases", "purchaseRevenue"):
        extra = and_filter(extra, eq_filter("eventName", "purchase"))
    df = shop_fetch(cfg, token, start, end, ["date"], [metric], extra, limit=600)
    if df.empty:
        return df
    df["날짜"] = pd.to_datetime(df["date"], format="%Y%m%d", errors="coerce")
    return df.dropna(subset=["날짜"]).sort_values("날짜")


def get_hours(cfg, token, day, metric, title=None, extra=None):
    extras = [eq_filter("unifiedScreenName", title)] if title else []
    if metric in ("ecommercePurchases", "purchaseRevenue"):
        extras.append(eq_filter("eventName", "purchase"))
    if extra:
        extras.append(extra)
    df = shop_fetch(cfg, token, day, day, ["hour"], [metric], and_filter(*extras), limit=24, quiet=True)
    base = pd.DataFrame({"hour": [f"{i:02d}" for i in range(24)]})
    if df.empty:
        base[metric] = 0
    else:
        df["hour"] = df["hour"].astype(str).str.zfill(2)
        base = base.merge(df[["hour", metric]], on="hour", how="left").fillna(0)
    base["시간"] = base["hour"] + "시"
    return base


def get_pages(cfg, token, start, end):
    return shop_fetch(cfg, token, start, end, ["unifiedScreenName"],
        ["screenPageViews", "activeUsers"], limit=5000, sort="screenPageViews")


# ───────────────────────────── Display widgets ───────────────────────────────
def num(n, suffix=""):
    if n is None or pd.isna(n):
        return "—"
    return f"{n:,.0f}{suffix}"


def money(n):
    if n is None:
        return "—"
    n = float(n)
    if abs(n) >= 1e8:
        return f"{n/1e8:,.2f}억원"
    if abs(n) >= 1e4:
        return f"{n/1e4:,.1f}만원"
    return f"{n:,.0f}원"


def duration(seconds):
    if seconds is None:
        return "—"
    s = max(0, int(seconds))
    return f"{s//60}분 {s%60:02d}초" if s >= 60 else f"{s}초"


def kpi(label, value, previous=None, fmt="number", note=""):
    formatter = money if fmt == "money" else duration if fmt == "time" else num
    value_s = formatter(value)
    delta = ""
    if value is not None and previous not in (None, 0):
        pct = (value - previous) / abs(previous) * 100
        cls = "kpi-up" if pct >= 0 else "kpi-down"
        delta = f'<span class="{cls}">{"▲" if pct >= 0 else "▼"} {abs(pct):.1f}%</span> 전 기간 대비'
    if note:
        delta = html.escape(note) if not delta else delta + " · " + html.escape(note)
    st.markdown(f'<div class="kpi"><div class="kpi-label">{html.escape(label)}</div>'
                f'<div class="kpi-value" title="{html.escape(value_s)}">{html.escape(value_s)}</div>'
                f'<div class="kpi-delta">{delta or "&nbsp;"}</div></div>', unsafe_allow_html=True)


def heading(title, helper=""):
    st.markdown(f'<div class="section-heading">{html.escape(title)}</div>', unsafe_allow_html=True)
    if helper:
        st.markdown(f'<div class="section-helper">{html.escape(helper)}</div>', unsafe_allow_html=True)


def empty(message="선택한 기간에 해당하는 데이터가 없어요."):
    st.info(message)


def linechart(df, x, y, label=None, height=240, clickable=False, key=None):
    if df.empty:
        empty()
        return None
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df[x], y=df[y], mode="lines+markers", name=label or y,
        line=dict(color=BLUE, width=2.6), marker=dict(size=6),
        fill="tozeroy", fillcolor="rgba(66,85,199,.07)",
        hovertemplate="%{x}<br>%{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(height=height, margin=dict(l=4, r=8, t=5, b=8),
        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)", showlegend=False,
        hovermode="x unified", font=dict(size=11, color="#62718a"),
        dragmode=False)
    fig.update_xaxes(showgrid=False, tickfont=dict(size=10))
    fig.update_yaxes(gridcolor="#edf1f8", zeroline=False, tickformat=",d")
    if clickable:
        state = st.plotly_chart(fig, key=key, on_select="rerun", selection_mode="points",
                                use_container_width=True, config={"displayModeBar": False})
        points = state.get("selection", {}).get("points", []) if state else []
        if points:
            xval = points[0].get("x")
            try:
                return pd.to_datetime(xval).date()
            except (ValueError, TypeError):
                pass
        return None
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    return None


def bars(df, label, value, height=230, horizontal=False, limit=12):
    if df.empty:
        empty()
        return
    view = df.head(limit).copy()
    fig = go.Figure()
    if horizontal:
        view = view.iloc[::-1]
        names = view[label].astype(str).str.slice(0, 32)
        fig.add_trace(go.Bar(x=view[value], y=names, orientation="h", marker_color=BLUE,
             hovertext=view[label], hovertemplate="%{hovertext}<br>%{x:,.0f}<extra></extra>"))
    else:
        fig.add_trace(go.Bar(x=view[label], y=view[value], marker_color=BLUE,
             hovertemplate="%{x}<br>%{y:,.0f}<extra></extra>"))
    fig.update_layout(height=height, margin=dict(l=4,r=5,t=8,b=3),
        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)", showlegend=False,
        font=dict(size=11,color="#64728a"))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#edf1f8", zeroline=False)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})


def donut(df, label, metric, height=220):
    if df.empty:
        empty()
        return
    fig = go.Figure(go.Pie(labels=df[label], values=df[metric], hole=.67,
        sort=True, marker=dict(colors=COLORS),
        textinfo="percent", textposition="outside", textfont=dict(size=10),
        hovertemplate="%{label}: %{value:,.0f} (%{percent})<extra></extra>"))
    fig.update_layout(height=height, margin=dict(l=5,r=5,t=3,b=3),
        paper_bgcolor="rgba(0,0,0,0)", showlegend=True,
        legend=dict(orientation="h",y=-.13,x=.5,xanchor="center",font=dict(size=10)))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})


def export_table(df, filename="ga4_data"):
    if df.empty:
        return
    # Save in popover: one line instead of two buttons extending the page.
    with st.popover("↓ 데이터 다운로드"):
        st.download_button("CSV 다운로드", df.to_csv(index=False).encode("utf-8-sig"),
            file_name=filename+".csv", mime="text/csv", use_container_width=True,
            key="csv_"+filename)
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="GA4")
        st.download_button("엑셀 다운로드", bio.getvalue(),
            file_name=filename+".xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="xlsx_"+filename)


def compact_table(df, filename, max_rows=100, height=290):
    if df.empty:
        empty()
        return
    st.dataframe(df.head(max_rows), hide_index=True, use_container_width=True, height=height)
    export_table(df, filename)


def selected_day(chart_df, metric, cfg, token, key, title=None, other=None):
    """Date drill-down only calls hourly API when chosen; minimizes initial render."""
    if chart_df.empty:
        return
    clicked = linechart(chart_df, "날짜", metric, clickable=True, key=f"click_{key}")
    state_key = f"picked_day_{key}"
    if clicked is not None:
        st.session_state[state_key] = str(clicked)
    valid = set(chart_df["날짜"].dt.strftime("%Y-%m-%d"))
    stored = st.session_state.get(state_key)
    if stored not in valid:
        stored = chart_df["날짜"].max().strftime("%Y-%m-%d")
        st.session_state[state_key] = stored
    labels = list(chart_df["날짜"].dt.strftime("%Y-%m-%d"))
    with st.expander(f"🕒  시간대별 상세 · {stored}", expanded=clicked is not None):
        choose_key = f"choose_{key}"
        if clicked is not None or st.session_state.get(choose_key) not in labels:
            st.session_state[choose_key] = stored
        chosen = st.selectbox("상세 날짜", labels, index=labels.index(stored),
                              key=choose_key)
        if chosen != stored:
            st.session_state[state_key] = chosen
            stored = chosen
        hourly = get_hours(cfg, token, stored, metric, title=title, extra=other)
        bars(hourly, "시간", metric, height=235, limit=24)
        top = hourly.nlargest(3, metric)[["시간", metric]]
        if not top.empty:
            st.caption("이용 집중 시간 TOP 3: " + " · ".join(
                f"{r['시간']} ({num(r[metric])})" for _,r in top.iterrows()))
        st.caption("시간별 활성 사용자는 중복될 수 있어 합계가 하루의 고유 사용자 수와 다를 수 있어요. 시간은 GA4 속성 시간대 기준입니다.")


def scope_notice(purchases=False):
    st.caption("조회 범위: GA4 전체 KT멤버십 속성 중 **화면 이름에 '쇼핑라운지'가 포함된 이벤트**. "
               "쇼핑라운지 방문 세션·고객 전체를 완전하게 구분하는 사용자 세그먼트는 아닙니다.")
    if purchases:
        st.caption("구매 지표는 쇼핑라운지 이름의 화면에서 `purchase` 이벤트가 기록될 때만 참고값으로 표시합니다. "
                   "페이지·기획전·상품 기여 매출과는 다를 수 있어요.")


def date_menu():
    today = dt.date.today()
    # GA4 finalized reports are best compared using yesterday as the latest day.
    yesterday = today - dt.timedelta(days=1)
    period = st.sidebar.radio("조회 기간", ["7일", "30일", "90일", "직접 선택"],
                              horizontal=True, index=1, key="period_option")
    if period == "직접 선택":
        chosen = st.sidebar.date_input("시작일 ~ 종료일",
            value=(yesterday - dt.timedelta(days=29), yesterday), max_value=yesterday,
            key="custom_period")
        if not isinstance(chosen, (tuple, list)) or len(chosen)!=2:
            st.info("시작 날짜와 종료 날짜를 선택해주세요.")
            st.stop()
        start, end = chosen
    else:
        days = int(period.replace("일", ""))
        start, end = yesterday - dt.timedelta(days=days-1), yesterday
    if start > end:
        st.warning("시작 날짜가 종료 날짜보다 늦어요.")
        st.stop()
    return start, end


def prev_range(start, end):
    delta = (end-start).days + 1
    return start - dt.timedelta(days=delta), start - dt.timedelta(days=1)


def classify(title):
    t = str(title).lower()
    if "기획전" in t or "프로모션" in t:
        return "기획전"
    if any(x in t for x in ["상품상세", "상품 상세", ">상품>", ">상품상세", "상품정보"]):
        return "상품 상세"
    if any(x in t for x in ["장바구니", "카트"]):
        return "장바구니"
    if any(x in t for x in ["쿠폰", "혜택함"]):
        return "쿠폰"
    if any(x in t for x in [">검색", "검색하기", "검색결과"]):
        return "검색"
    if any(x in t for x in ["주문", "결제"]):
        return "주문·결제"
    if any(x in t for x in [">메인", "홈", "쇼핑라운지>main"]):
        return "메인"
    if "베스트" in t or "카테고리" in t or "멤버스특가" in t:
        return "카테고리·베스트"
    return "기타"


def add_category(pages):
    if pages.empty:
        return pages
    x = pages.copy()
    x["유형"] = x["unifiedScreenName"].map(classify)
    return x


def source_type(text):
    s = str(text).lower()
    if re.search(r"(^|[^a-z])(mms|rcs|k-?ads|sms)([^a-z]|$)", s) or "문자" in s:
        return "MMS/RCS 통합"
    if any(k in s for k in ["ktmembership", "kt_membership", "kt-membership", "kt멤버십", "ktmember", "멤버십앱"]):
        return "KT멤버십 앱 (태그 후보)"
    if re.search(r"(google|naver|daum|bing).*(organic|search|/ organic)", s):
        return "검색"
    return "기타·미분류"


# ────────────────────────────── Page implementations ─────────────────────────
def page_overview(cfg, token, start, end):
    scope_notice(purchases=True)
    current = get_summary(cfg, token, start, end)
    ps, pe = prev_range(start, end)
    prior = get_summary(cfg, token, ps, pe)
    if current["activeUsers"] is None:
        empty("선택한 기간의 쇼핑라운지 화면 데이터를 확인하지 못했어요. GA4 화면 이름을 확인해주세요.")
        return
    cust = shop_fetch(cfg, token, start, end, ["newVsReturning"], ["activeUsers"], quiet=True)
    pcust = shop_fetch(cfg, token, ps, pe, ["newVsReturning"], ["activeUsers"], quiet=True)
    types = dict(zip(cust.get("newVsReturning", []), cust.get("activeUsers", [])))
    ptypes = dict(zip(pcust.get("newVsReturning", []), pcust.get("activeUsers", [])))
    purchase = get_purchase(cfg, token, start, end)
    old_purchase = get_purchase(cfg, token, ps, pe) if purchase else None
    act, old_act = current["activeUsers"], prior.get("activeUsers")
    # Two rows instead of a tall card stack.
    first = st.columns(5)
    widgets = [
        ("활성 사용자", act, old_act, "number", ""),
        ("신규 유형 사용자", types.get("new"), ptypes.get("new"), "number", "유형 기준"),
        ("재방문 유형 사용자", types.get("returning"), ptypes.get("returning"), "number", "유형 기준"),
        ("세션 수", current.get("sessions") if (current.get("sessions") or 0)>0 else None,
         prior.get("sessions") if (prior.get("sessions") or 0)>0 else None, "number", "화면 조건"),
        ("페이지 조회수", current["screenPageViews"], prior.get("screenPageViews"), "number", ""),
        ("평균 참여 시간", current["userEngagementDuration"]/act if act and (current.get("userEngagementDuration") or 0)>0 else None,
         prior["userEngagementDuration"]/old_act if old_act and (prior.get("userEngagementDuration") or 0)>0 else None,
         "time", "사용자당"),
        ("구매 금액*", purchase.get("purchaseRevenue") if purchase else None,
         old_purchase.get("purchaseRevenue") if old_purchase else None, "money", "화면 기록"),
        ("구매 건수*", purchase.get("ecommercePurchases") if purchase else None,
         old_purchase.get("ecommercePurchases") if old_purchase else None, "number", "화면 기록"),
        ("구매 전환율*", ((purchase["totalPurchasers"]/act)*100 if purchase and act else None),
         ((old_purchase["totalPurchasers"]/old_act)*100 if old_purchase and old_act else None),
         "percent", "참고값"),
    ]
    # Use the same 9 KPIs across a 5+4 grid.
    for col, (label,v,p,format_type,note) in zip(first,widgets[:5]):
        with col: kpi(label,v,p,fmt=format_type,note=note)
    second = st.columns(4)
    for col, (label,v,p,format_type,note) in zip(second,widgets[5:]):
        with col:
            if format_type == "percent":
                # Converted percentage points, not growth %, for a rate.
                rate_value = f"{v:.2f}%" if v is not None else "—"
                diff = f"전 기간 대비 {v-p:+.2f}%p" if v is not None and p is not None else note
                st.markdown(f'<div class="kpi"><div class="kpi-label">{html.escape(label)}</div>'
                    f'<div class="kpi-value">{html.escape(rate_value)}</div>'
                    f'<div class="kpi-delta">{html.escape(diff)}</div></div>', unsafe_allow_html=True)
            else:
                kpi(label,v,p,fmt=format_type,note=note)
    st.caption("* 구매 이벤트가 화면 이름에 연결된 경우만 표시합니다. 신규/재방문은 GA4 사용자 유형(newVsReturning)이며 두 유형의 합이 전체 고유 사용자와 일치하지 않을 수 있습니다.")
    st.divider()
    ctrl, note_col = st.columns([1.6, 2.4], vertical_alignment="center")
    with ctrl:
        choice = st.selectbox("일별 추이 지표", ["활성 사용자", "페이지 조회수", "세션 수", "구매 건수*", "구매 금액*"], label_visibility="collapsed")
    with note_col:
        st.caption("그래프의 날짜를 선택하면 아래 시간대별 상세 분석이 열려요.")
    chosen_metric = {"활성 사용자":"activeUsers", "페이지 조회수":"screenPageViews",
                     "세션 수":"sessions", "구매 건수*":"ecommercePurchases",
                     "구매 금액*":"purchaseRevenue"}[choice]
    if chosen_metric in ("ecommercePurchases","purchaseRevenue") and not purchase:
        st.info("현재 기간에 쇼핑라운지 화면으로 기록된 구매 이벤트가 확인되지 않아 그래프를 표시하지 않아요.")
        return
    df = get_dates(cfg, token, start, end, chosen_metric)
    selected_day(df, chosen_metric, cfg, token, f"overview_{choice}_{start}_{end}")


def page_customers(cfg, token, start, end):
    scope_notice()
    section = st.radio("고객 분석 기준", ["성별", "연령대", "운영체제", "성별 × 기기", "일별 변화"],
                       horizontal=True, label_visibility="collapsed", key="customer_view")
    # Only the selected sub-view runs its GA4 query, minimizing latency.
    if section == "성별":
        df = shop_fetch(cfg, token, start, end, ["userGender"], ["activeUsers"], limit=30)
        left,right=st.columns([1,1.25])
        with left: donut(df,"userGender","activeUsers")
        with right: compact_table(df.rename(columns={"userGender":"성별","activeUsers":"활성 사용자"}),"ga4_gender",height=240)
        st.caption("인구통계 데이터는 개인정보 보호 임계값·미수집으로 일부 누락될 수 있어요.")
    elif section == "연령대":
        df = shop_fetch(cfg, token, start, end, ["userAgeBracket"], ["activeUsers"], limit=30)
        left,right=st.columns([1,1.15])
        with left: bars(df,"userAgeBracket","activeUsers")
        with right: compact_table(df.rename(columns={"userAgeBracket":"연령대","activeUsers":"활성 사용자"}),"ga4_age",height=245)
    elif section == "운영체제":
        df = shop_fetch(cfg, token, start, end, ["operatingSystem"], ["activeUsers"], limit=40)
        left,right=st.columns([1,1.15])
        with left: donut(df,"operatingSystem","activeUsers")
        with right: compact_table(df.rename(columns={"operatingSystem":"운영체제","activeUsers":"활성 사용자"}),"ga4_os",height=245)
        st.caption("iOS/Android 외에 데스크톱 OS가 포함될 수 있어요.")
    elif section == "성별 × 기기":
        df = shop_fetch(cfg, token, start, end, ["userGender","operatingSystem"], ["activeUsers"], limit=60)
        if not df.empty:
            compact_table(df.rename(columns={"userGender":"성별","operatingSystem":"운영체제", "activeUsers":"활성 사용자"}),"ga4_gender_os")
        st.caption("성별 × 기기 데이터는 수집·공개된 범위에서만 표시합니다. 성별 × 연령 분석은 요청에 따라 제외했어요.")
    elif section == "일별 변화":
        metric_type = st.selectbox("비교 기준", ["운영체제", "성별", "연령대"])
        dim = {"운영체제":"operatingSystem","성별":"userGender","연령대":"userAgeBracket"}[metric_type]
        df = shop_fetch(cfg, token, start, end, ["date",dim], ["activeUsers"], limit=8000)
        if not df.empty:
            df["날짜"]=pd.to_datetime(df["date"],format="%Y%m%d")
            pivot=df.pivot_table(index="날짜",columns=dim,values="activeUsers",aggfunc="sum").fillna(0)
            fig=go.Figure()
            for i,col in enumerate(pivot.columns[:7]):
                fig.add_trace(go.Scatter(x=pivot.index,y=pivot[col],name=str(col),mode="lines",
                                         line=dict(width=2,color=COLORS[i%len(COLORS)])))
            fig.update_layout(height=260,margin=dict(l=5,r=8,t=12,b=8),paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="white",legend=dict(orientation="h",y=-.2),
                font=dict(size=11,color="#62718a"))
            fig.update_yaxes(gridcolor="#eef0f7")
            st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
            compact_table(df.rename(columns={dim:metric_type,"activeUsers":"활성 사용자"})[["날짜",metric_type,"활성 사용자"]],
                          "ga4_customers_trend",height=215)

def page_acquisition(cfg, token, start, end):
    scope_notice(purchases=True)
    st.caption("MMS/RCS는 통합합니다. KT멤버십 앱 경로는 태그가 명확할 때만 후보로 표시하며, (direct)를 앱 유입으로 간주하지 않아요.")
    df = shop_fetch(cfg, token, start, end,
        ["sessionSourceMedium", "sessionCampaignName"],
        ["screenPageViews", "activeUsers", "sessions"], limit=5000, sort="screenPageViews")
    if df.empty:
        empty()
        return
    df["유입 경로"]=(df["sessionSourceMedium"].fillna("")+" "+df["sessionCampaignName"].fillna("")).map(source_type)
    agg=df.groupby("유입 경로",as_index=False)["screenPageViews"].sum().sort_values("screenPageViews",ascending=False)
    left,right=st.columns([1,1.2])
    with left:
        heading("유입 경로별 조회수 비중")
        donut(agg,"유입 경로","screenPageViews",height=240)
    with right:
        heading("유입 채널별 상세")
        compact_table(agg.rename(columns={"screenPageViews":"조회수"}),"ga4_acquisition_channels",height=225)
    st.caption("같은 사람이 여러 유입 경로에서 활동할 수 있으므로 경로별 활성 사용자를 단순 합산하지 않습니다. 세션 수는 화면 조건이 붙은 참고값이에요.")
    with st.expander("유입 소스·캠페인 원본 상세", expanded=True):
        show=df.rename(columns={"sessionSourceMedium":"소스 / 매체", "sessionCampaignName":"캠페인",
            "screenPageViews":"조회수", "activeUsers":"활성 사용자", "sessions":"세션"})
        compact_table(show[["유입 경로","소스 / 매체","캠페인","조회수","활성 사용자","세션"]],"ga4_acquisition_raw",height=245)
    st.info("정확한 채널별 구매 귀속 및 랜딩 페이지별 전환을 보려면 UTM/딥링크와 쇼핑라운지 구매 이벤트 연결 검증이 필요합니다. 현재는 출처를 추정해 매출로 계산하지 않습니다.")


def page_pages(cfg,token,start,end):
    scope_notice()
    pages = add_category(get_pages(cfg,token,start,end))
    if pages.empty:
        empty()
        return
    stats=pages.groupby("유형",as_index=False)["screenPageViews"].sum().sort_values("screenPageViews",ascending=False)
    left,right=st.columns([1,1.2])
    with left:
        heading("페이지 유형별 조회수")
        bars(stats,"유형","screenPageViews",height=250)
    with right:
        heading("페이지 조회수 TOP 10")
        top=pages[["unifiedScreenName","screenPageViews"]].head(10)
        bars(top,"unifiedScreenName","screenPageViews",height=250,horizontal=True,limit=10)
    types=["전체"]+sorted(pages["유형"].unique())
    fcol,scol=st.columns([1,2])
    with fcol: typ=st.selectbox("페이지 유형",types)
    with scol: q=st.text_input("페이지명 검색",placeholder="검색어 또는 페이지 ID")
    shown=pages.copy()
    if typ!="전체":shown=shown[shown["유형"]==typ]
    if q:shown=shown[shown["unifiedScreenName"].str.contains(re.escape(q),case=False,na=False)]
    compact_table(shown.rename(columns={"unifiedScreenName":"페이지 이름", "screenPageViews":"조회수",
        "activeUsers":"활성 사용자"})[["유형","페이지 이름","조회수","활성 사용자"]],"ga4_pages",height=275)
    st.caption("개별 기획전과 상품은 각각 ⑤·⑥ 메뉴에서 더 자세히 분석합니다. 유형 분류는 페이지 제목 규칙에 따른 자동 분류예요.")


def detail_trend(cfg,token,start,end,title,identity):
    heading("선택 항목 · 일별 조회수")
    df=get_dates(cfg,token,start,end,"screenPageViews",title=title)
    selected_day(df,"screenPageViews",cfg,token,f"detail_{identity}_{start}_{end}",title=title)


def page_promotions(cfg,token,start,end):
    scope_notice(purchases=True)
    pages=add_category(get_pages(cfg,token,start,end))
    pro=pages[pages["유형"]=="기획전"].copy() if not pages.empty else pd.DataFrame()
    if pro.empty:
        empty("기획전으로 분류되는 화면을 찾지 못했어요. 화면 이름에 '기획전' 등이 있는지 확인해주세요.")
        return
    pro["기획전 ID"] = pro["unifiedScreenName"].astype(str).str.extract(r"(?<!\d)(\d{4,})(?!\d)",expand=False).fillna("—")
    st.caption("조회수 TOP은 실제 화면 조회 기준입니다. '매출 TOP'은 기획전 식별정보가 구매 이벤트와 연결된 뒤에만 표시해요.")
    left,right=st.columns([1,1.3])
    with left:
        heading("기획전 조회수 TOP 10")
        bars(pro.head(10),"unifiedScreenName","screenPageViews",height=260,horizontal=True,limit=10)
    with right:
        heading("기획전별 실적")
        compact_table(pro.rename(columns={"unifiedScreenName":"기획전명", "screenPageViews":"조회수",
            "activeUsers":"활성 사용자"})[["기획전 ID","기획전명","조회수","활성 사용자"]],"ga4_promotions",height=255)
    st.divider()
    options=pro["unifiedScreenName"].tolist()
    chosen=st.selectbox("기획전 선택 · 클릭 시 상세 보기",options,index=0)
    chosen_row=pro[pro["unifiedScreenName"]==chosen].iloc[0]
    a,b,c=st.columns(3)
    with a:kpi("기획전 조회수",chosen_row["screenPageViews"])
    with b:kpi("기획전 활성 사용자",chosen_row["activeUsers"])
    with c:kpi("기획전 연결 매출",None,note="구매 귀속 검증 필요")
    detail_trend(cfg,token,start,end,chosen,"promotion_"+hashlib.md5(chosen.encode()).hexdigest()[:7])
    with st.expander("기획전 내부 상품 성과",expanded=False):
        st.info("이 기획전에 포함된 상품과 매출을 연결하려면 기획전 ID와 상품 ID가 함께 수집되어야 해요. "
                "현재 조회수만으로 개별 상품의 기획전 기여 매출을 만들지 않습니다.")
        df=shop_fetch(cfg,token,start,end,["itemPromotionName","itemName"],["itemsViewed","itemsPurchased","itemRevenue"],
                      limit=250,quiet=True)
        if not df.empty:
            # Only display rows explicitly linked by a nonempty promotion name matching the selected title.
            # Otherwise no justified promotion attribution can be made.
            term = str(chosen_row["기획전 ID"])
            candidates = df[df["itemPromotionName"].astype(str).apply(
                lambda x: x not in ("", "(not set)") and (x.lower() in chosen.lower() or
                      (term!="—" and term in x)))].copy()
            if not candidates.empty:
                compact_table(candidates.rename(columns={"itemPromotionName":"기획전 태그", "itemName":"상품명",
                    "itemsViewed":"상품 조회 수량", "itemsPurchased":"구매 수량", "itemRevenue":"상품 매출"}),
                    "ga4_promotion_items",height=240)


def page_products(cfg,token,start,end):
    scope_notice(purchases=True)
    pages=add_category(get_pages(cfg,token,start,end))
    pro=pages[pages["유형"]=="상품 상세"].copy() if not pages.empty else pd.DataFrame()
    if pro.empty:
        empty("'상품 상세'로 분류되는 화면이 아직 없어요. 실제 상품 페이지 제목 규칙을 확인해주세요.")
        return
    pro["상품코드 후보"] = pro["unifiedScreenName"].astype(str).str.extract(r"(?<!\d)(\d{5,})(?!\d)",expand=False).fillna("—")
    left,right=st.columns([1,1.25])
    with left:
        heading("상품 상세 조회수 TOP 10")
        bars(pro,"unifiedScreenName","screenPageViews",height=260,horizontal=True,limit=10)
    with right:
        heading("상품 상세페이지 실적")
        compact_table(pro.rename(columns={"unifiedScreenName":"상품 페이지", "screenPageViews":"조회수",
            "activeUsers":"활성 사용자"})[["상품코드 후보","상품 페이지","조회수","활성 사용자"]],"ga4_product_pages",height=258)
    chosen=st.selectbox("상품 상세 페이지 선택",pro["unifiedScreenName"].tolist(),index=0)
    detail_trend(cfg,token,start,end,chosen,"product_"+hashlib.md5(chosen.encode()).hexdigest()[:7])
    with st.expander("상품 구매·장바구니 상세 / 데이터 수집 확인",expanded=False):
        st.info("상품별 구매 건수·매출·장바구니 수치는 `item_id`/`item_name` 전자상거래 이벤트와 "
                "쇼핑라운지 소속이 함께 검증된 후 연결해요. 지금은 KT멤버십 전체 상품 매출을 쇼핑라운지 매출로 표시하지 않습니다.")
        ev=shop_fetch(cfg,token,start,end,["eventName"],["eventCount"],limit=100,quiet=True)
        if not ev.empty:
            ev=ev[ev["eventName"].isin(["view_item","add_to_cart","begin_checkout","purchase"])]
            if not ev.empty:compact_table(ev.rename(columns={"eventName":"이벤트", "eventCount":"기록 건수"}),"ga4_product_event_check",height=215)


def page_behavior(cfg,token,start,end):
    scope_notice()
    pages=add_category(get_pages(cfg,token,start,end))
    if pages.empty:
        empty()
        return
    stats=pages.groupby("유형",as_index=False)["screenPageViews"].sum().sort_values("screenPageViews",ascending=False)
    left,right=st.columns([1,1.2])
    with left:
        heading("행동 관련 페이지 조회수")
        chosen_types=["메인","기획전","상품 상세","장바구니","주문·결제"]
        funnel=stats.set_index("유형").reindex(chosen_types).fillna(0).reset_index()
        bars(funnel,"유형","screenPageViews",height=250)
    with right:
        heading("행동 이벤트 실제 수집")
        events=shop_fetch(cfg,token,start,end,["eventName"],["eventCount"],limit=300,sort="eventCount")
        if not events.empty:
            monitored=["page_view","screen_view","view_item","add_to_cart","begin_checkout","purchase"]
            ev=events[events["eventName"].isin(monitored)].copy()
            compact_table(ev.rename(columns={"eventName":"이벤트명", "eventCount":"기록 건수"}),"ga4_behavior",height=246)
        else:empty()
    st.caption("위 그래프는 단계별 **조회/이벤트 빈도**이지, 동일 고객의 순차 이동으로 검증된 구매 퍼널이나 이탈률은 아닙니다.")
    with st.expander("페이지 이동·구매 퍼널 검증 조건",expanded=False):
        st.write("GA4 Data API의 화면별 조회수를 단순히 이어 붙이면 동일 사용자의 이동 경로를 재현할 수 없어요. "
                 "실제 사용자 단위 경로·이탈 분석은 GA4 탐색의 경로 탐색/유입경로 탐색 결과 또는 추가 이벤트 설계가 필요합니다.")
    st.caption("요청에 따라 미구매 고객 개별 분석 기능은 포함하지 않았어요.")


def page_campaigns(cfg,token,start,end):
    scope_notice(purchases=True)
    st.caption("③ 유입 분석은 소스/매체를, ⑧ 캠페인 분석은 `sessionCampaignName`을 기준으로 상세하게 봅니다. MMS/RCS는 합쳐서 표시해요.")
    df=shop_fetch(cfg,token,start,end,["sessionCampaignName","sessionSourceMedium"],
        ["screenPageViews","sessions","activeUsers"],limit=5000,sort="screenPageViews")
    if df.empty:
        empty()
        return
    valid=df[~df["sessionCampaignName"].isin(["(not set)","(direct)","", "(organic)"])].copy()
    if valid.empty:
        empty("쇼핑라운지 화면과 연결된 캠페인 태그를 확인하지 못했어요. UTM 태깅을 점검해주세요.")
        return
    camp=valid.groupby("sessionCampaignName",as_index=False)["screenPageViews"].sum().sort_values("screenPageViews",ascending=False)
    left,right=st.columns([1,1.3])
    with left:
        heading("캠페인별 조회수 TOP 10")
        bars(camp,"sessionCampaignName","screenPageViews",height=260,horizontal=True,limit=10)
    with right:
        heading("캠페인별 상세")
        compact_table(valid.rename(columns={"sessionCampaignName":"캠페인명", "sessionSourceMedium":"소스/매체",
                 "screenPageViews":"조회수", "sessions":"세션(화면조건)", "activeUsers":"활성 사용자"}),
                 "ga4_campaigns",height=262)
    st.caption("캠페인별 사용자 수는 중복 가능하며, 구매 귀속은 확인되지 않았으므로 매출 순위를 만들지 않습니다.")
    chosen=st.selectbox("캠페인 선택",camp["sessionCampaignName"].tolist(),index=0)
    trend=shop_fetch(cfg,token,start,end,["date"],["screenPageViews"],
                     eq_filter("sessionCampaignName",chosen),limit=500)
    if not trend.empty:
        trend["날짜"]=pd.to_datetime(trend["date"],format="%Y%m%d")
        # For a campaign drill-down, the hour query must retain the campaign filter too.
        selected_day(trend.sort_values("날짜"),"screenPageViews",cfg,token,
                     "campaign_"+hashlib.md5(chosen.encode()).hexdigest()[:7]+f"_{start}_{end}",
                     other=eq_filter("sessionCampaignName",chosen))
    with st.expander("랜딩 페이지 / 구매 연결 상태",expanded=False):
        st.info("캠페인별 실제 첫 방문 페이지와 구매를 검증하려면 세션 유입정보·쇼핑라운지 식별정보·구매 이벤트를 "
                "연결하는 추적 설계가 필요합니다. `(not set)`은 임의로 분류하지 않습니다.")


def page_guide(cfg,token,start,end):
    st.caption("정확한 실적을 위한 데이터 정의를 이 메뉴에서 확인해 주세요.")
    choices=[
        ("화면 이름", "쇼핑라운지 포함 여부", "화면 이벤트 필터 · 적용"),
        ("활성 사용자", "화면 조건 하 고유 사용자 수", "사용자 규모 참고"),
        ("세션", "screenName 이벤트 필터로는 세션 전체를 정확히 복원할 수 없음", "화면 조건 참고"),
        ("구매", "purchase 이벤트가 쇼핑라운지 화면과 연결되어야 함", "화면 기록 참고"),
        ("기획전 기여 매출", "promotion_id와 구매 항목 연결", "추가 태깅 필요"),
        ("상품별 매출", "item_id / item_name + 쇼핑라운지 식별정보", "추가 검증 필요"),
        ("MMS/RCS vs 앱", "유입 캠페인/소스 식별값", "UTM 검증 필요"),
    ]
    compact_table(pd.DataFrame(choices,columns=["항목","기준","상태"]),"ga4_guide",height=290)


def main():
    cfg=load_config()
    st.markdown('<div class="app-title">쇼핑라운지 <span style="color:#4255c7">GA4 인사이트</span></div>',
                unsafe_allow_html=True)
    st.markdown('<div class="app-sub">방문·고객·유입·기획전·상품·캠페인 성과를 한 화면에서</div>',unsafe_allow_html=True)
    if not cfg:
        st.warning("Google OAuth 설정이 없어요. 기존 Streamlit Cloud 앱의 Secrets 값을 그대로 사용해 주세요.")
        st.code('[google_oauth]\nclient_id = "..."\nclient_secret = "..."\nredirect_uri = "https://shoppinglounge-ga4.streamlit.app/"\nstate_signing_secret = "..."\n\n[ga4]\nproperty_id = "544313553"',language="toml")
        st.stop()
    if not cfg["redirect_uri"].startswith(("https://","http://localhost","http://127.0.0.1")):
        st.error("OAuth 리디렉션 URI는 HTTPS 또는 localhost여야 해요.")
        st.stop()
    receive_callback(cfg)
    token=current_token()
    if not token:
        st.info("GA4 조회 권한이 있는 Google 계정으로 연결해 주세요. 데이터는 읽기 전용입니다.")
        st.link_button("Google 계정 연결",auth_url(cfg),type="primary")
        st.stop()
    st.sidebar.markdown("### 쇼핑라운지 분석")
    menu=st.sidebar.radio("메뉴",NAV,label_visibility="collapsed",key="nav_menu")
    st.sidebar.divider()
    start,end=date_menu()
    st.sidebar.caption(f"📅 {start:%Y.%m.%d} ~ {end:%Y.%m.%d}")
    if st.sidebar.button("↻ 데이터 새로고침",use_container_width=True):
        st.session_state.pop("ga4_report_cache",None)
        st.rerun()
    if st.sidebar.button("로그아웃",use_container_width=True):
        disconnect()
        st.rerun()
    st.sidebar.divider()
    st.sidebar.caption(f"GA4 속성: {cfg['property_id']}  ·  읽기 전용")
    st.sidebar.caption("화면 이름: '"+cfg['keyword']+"' 포함")
    st.markdown(f'<span class="pill">{html.escape(menu)}</span> &nbsp; '
                f'<span class="microcopy">조회 기간: {start:%Y.%m.%d} – {end:%Y.%m.%d}</span>',
                unsafe_allow_html=True)
    pages={
        NAV[0]:page_overview, NAV[1]:page_customers, NAV[2]:page_acquisition,
        NAV[3]:page_pages, NAV[4]:page_promotions, NAV[5]:page_products,
        NAV[6]:page_behavior, NAV[7]:page_campaigns,
    }
    with st.spinner("GA4 데이터를 불러오는 중..."):
        pages[menu](cfg,token,start,end)
    with st.expander("ℹ️ 데이터 집계 기준 · 제약사항",expanded=False):
        page_guide(cfg,token,start,end)
    st.caption("© 쇼핑라운지 내부 분석 | GA4 실데이터 연결 · 테스트용 공개 배포 시 접근권한 및 OAuth 보안 검토 필요")


if __name__ == "__main__":
    main()
