# -*- coding: utf-8 -*-
"""쇼핑라운지 GA4 운영 인사이트 v6 — 실제 수집값 검증 기반, 7메뉴 compact dashboard.

기존 Streamlit Secrets 유지:
[google_oauth] client_id, client_secret, redirect_uri, state_signing_secret
[ga4] property_id = '544313553'

중요: KT멤버십 속성의 쇼핑라운지 이벤트/화면 실적과 사용자 세그먼트
(사용자가 쇼핑라운지에 진입한 후 타 화면에서 발생시킨 이벤트 포함 가능)은 다릅니다.
구매금액은 주문완료 화면에 기록된 purchase 이벤트 기준입니다.
상품 수익은 상품 ID가 매칭돼도 채널 귀속 매출이 아닙니다.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import html
import json
import re
import secrets
from collections import OrderedDict

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
from requests_oauthlib import OAuth2Session

st.set_page_config(page_title="쇼핑라운지 | GA4 인사이트", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

API = "https://analyticsdata.googleapis.com/v1beta"
ALPHA = "https://analyticsdata.googleapis.com/v1alpha"
SCOPE = "https://www.googleapis.com/auth/analytics.readonly"
AUTH = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN = "https://oauth2.googleapis.com/token"
DEFAULT_PROPERTY = "544313553"
LOUNGE_TITLE = "KT멤버십>쇼핑라운지"
# GA4 탐색의 「총 구매 수익」과 같은 grossPurchaseRevenue를 우선 사용합니다.
GROSS_REVENUE = "grossPurchaseRevenue"
PURCHASE_TITLE = "KT멤버십>쇼핑라운지>주문>주문완료"
MENUS = ["01 · 종합 현황", "02 · 이용 고객", "03 · 유입·앱 클릭", "04 · 페이지", "05 · 기획전", "06 · 상품·브랜드", "07 · 쿠폰·구매 행동"]
NAV_HELP = ["9개 KPI · 주요 변화 · 인기 TOP", "성별·연령·기기·등급·로그인", "유입 채널 · KT앱 영역별·배너별 클릭 순위", "자주 보는 화면 · 유형 · 상세", "인기 기획전 · 조회/이용자 순위", "상품명·브랜드 · 구매 효율 참고", "쿠폰 사용 거래 · 수익 · 퍼널"]
BLUE, TEAL, ORANGE, LIGHT = "#4765D2", "#169C99", "#F6A34F", "#ECF1FB"

st.markdown("""<style>
/* No negative margins or global stVerticalBlock gap overrides. */
.block-container{padding:4.6rem 1.4rem 2.5rem!important;max-width:1540px!important}
section[data-testid="stSidebar"]{background:#F7F9FD}
section[data-testid="stSidebar"] .block-container{padding-top:1.8rem!important}
#MainMenu,footer{visibility:hidden}
.header-title{font-size:clamp(1.4rem,2.1vw,1.9rem);letter-spacing:-.035em;font-weight:820;color:#253451;line-height:1.4;margin:.05rem 0 .2rem}
.header-note{font-size:.81rem;color:#718199;margin:.25rem 0 1rem;line-height:1.6}
.section-title{font-size:1.08rem;font-weight:760;color:#263C58;margin:1rem 0 .35rem}
.scope-label{display:inline-block;border-radius:20px;background:#EAF0FC;color:#4055B6;font-size:.72rem;padding:3px 10px;font-weight:750}
.kpi-overview{display:flex;flex-direction:column;gap:12px;width:100%;margin:12px 0 18px;clear:both}
.kpi-overview-row{display:grid;gap:12px;align-items:stretch;min-width:0}
.kpi-overview-row.row-five{grid-template-columns:repeat(5,minmax(0,1fr))}
.kpi-overview-row.row-four{grid-template-columns:repeat(4,minmax(0,1fr))}
.kpi{box-sizing:border-box;border:1px solid #E5EBF5;border-radius:14px;background:#FFF;min-width:0;min-height:134px;padding:14px 14px 15px;box-shadow:0 2px 8px rgba(30,49,80,.04);display:flex;flex-direction:column;gap:4px;overflow-wrap:anywhere}
.kpi-name{color:#72829A;font-size:.78rem;font-weight:650;line-height:1.35}
.kpi-value{color:#1E3150;font-size:clamp(1.13rem,1.55vw,1.57rem);font-weight:790;line-height:1.35;overflow-wrap:normal;word-break:keep-all;white-space:normal}
.kpi-foot{font-size:.69rem;color:#7F8EA6;line-height:1.5;margin-top:auto}
.kpi-delta{font-weight:700;font-size:.73rem;margin-top:2px}
.kpi-delta.up{color:#187F6E}.kpi-delta.down{color:#C05D55}.kpi-delta.neutral{color:#7F8EA6}
.insight-row{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin:6px 0 14px}
.insight-card{box-sizing:border-box;background:#F7F9FD;border:1px solid #E8EEF6;border-radius:12px;min-width:0;padding:15px;min-height:98px}
.insight-head{font-size:.72rem;color:#6C7E98;font-weight:650;margin-bottom:8px}
.insight-body{color:#273956;font-size:.87rem;font-weight:660;line-height:1.5;overflow-wrap:anywhere}
.insight-sub{color:#8191A6;font-size:.7rem;margin-top:5px;line-height:1.4}
@media(max-width:1160px){.kpi-overview-row.row-five,.kpi-overview-row.row-four{grid-template-columns:repeat(3,minmax(0,1fr))}}
@media(max-width:780px){.kpi-overview-row.row-five,.kpi-overview-row.row-four,.insight-row{grid-template-columns:repeat(2,minmax(0,1fr))}.block-container{padding:4.3rem .85rem 2rem!important}}
@media(max-width:470px){.kpi-overview-row.row-five,.kpi-overview-row.row-four,.insight-row{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}.kpi{min-height:142px;padding:11px}.kpi-value{font-size:1.05rem}.kpi-foot{font-size:.64rem}}
</style>""", unsafe_allow_html=True)

# ------------------------ OAuth: compatible with existing Secrets ------------------------
def config():
    try:
        o = st.secrets["google_oauth"]
        g = st.secrets.get("ga4", {})
        return {"client_id":str(o["client_id"]), "client_secret":str(o["client_secret"]),
                "redirect_uri":str(o["redirect_uri"]), "state_signing_secret":str(o["state_signing_secret"]),
                "property":str(g.get("property_id",DEFAULT_PROPERTY)),
                "title":str(g.get("page_keyword",LOUNGE_TITLE))}
    except (KeyError, FileNotFoundError):
        return None


def state_signer(cfg):
    return URLSafeTimedSerializer(cfg["state_signing_secret"], salt="ga4-oauth-state-v1")


def login_url(cfg):
    state = state_signer(cfg).dumps({"nonce":secrets.token_urlsafe(24)})
    oauth = OAuth2Session(cfg["client_id"], scope=[SCOPE], redirect_uri=cfg["redirect_uri"], state=state)
    url, _ = oauth.authorization_url(AUTH, access_type="online", prompt="select_account")
    return url


def logout():
    for k in ["ga4_access_token","ga4_token_expires_at","ga4_expiry","ga4_v6_cache","ga4_report_cache"]:
        st.session_state.pop(k, None)


def oauth_callback(cfg):
    params=st.query_params
    if "error" in params:
        st.query_params.clear(); st.error("Google 인증이 취소됐어요."); return
    code,state=params.get("code"),params.get("state")
    if not code:
        return
    try:
        if not state: raise BadSignature("state missing")
        state_signer(cfg).loads(state,max_age=600)
        oauth=OAuth2Session(cfg["client_id"],scope=[SCOPE],redirect_uri=cfg["redirect_uri"],state=state)
        result=oauth.fetch_token(TOKEN,code=code,client_secret=cfg["client_secret"],include_client_id=True,timeout=25)
        if not result.get("access_token"): raise ValueError("no token")
        st.session_state["ga4_access_token"]=result["access_token"]
        st.session_state["ga4_token_expires_at"]=dt.datetime.now(dt.timezone.utc)+dt.timedelta(seconds=max(30,int(result.get("expires_in",3600))-60))
        st.session_state.pop("ga4_v6_cache",None)
        st.query_params.clear(); st.rerun()
    except (BadSignature,SignatureExpired):
        st.query_params.clear();st.error("인증 요청이 만료됐어요. 다시 로그인해 주세요.")
    except Exception as exc:
        st.query_params.clear();st.error(f"Google 인증 오류: {type(exc).__name__}")


def active_token():
    val=st.session_state.get("ga4_access_token")
    until=st.session_state.get("ga4_token_expires_at") or st.session_state.get("ga4_expiry")
    if not val or not until: return None
    if dt.datetime.now(dt.timezone.utc)>=until:
        logout();return None
    return val

# ----------------------------- REST & metadata -----------------------------
class GAError(Exception):
    pass


def api(cfg, token, action, payload=None, method="POST", alpha=False):
    url=f"{ALPHA if alpha else API}/properties/{cfg['property']}"+action
    cache=st.session_state.setdefault("ga4_v6_cache",OrderedDict())
    key=hashlib.sha256(json.dumps([url,method,payload],sort_keys=True,ensure_ascii=False).encode()).hexdigest()
    now=dt.datetime.now(dt.timezone.utc)
    if key in cache:
        stamp, saved=cache[key]
        if (now-stamp).total_seconds()<300: return saved
        del cache[key]
    try:
        kw={"headers":{"Authorization":f"Bearer {token}"},"timeout":35}
        if method=="POST": kw["json"]=payload or {}
        result=requests.request(method,url,**kw)
        if not result.ok:
            try: msg=result.json().get("error",{}).get("message","요청 오류")
            except ValueError: msg=result.text[:120]
            raise GAError(f"GA4 {result.status_code}: {str(msg)[:230]}")
        data=result.json()
    except requests.RequestException as exc:
        raise GAError(f"네트워크 오류: {type(exc).__name__}") from None
    cache[key]=(now,data)
    while len(cache)>90: cache.popitem(last=False)
    return data


def metadata(cfg,token):
    return api(cfg,token,"/metadata",method="GET")


def canonical(x):
    return re.sub(r"[\s_\-]","",str(x).casefold())


def custom_field(meta, names):
    # Never guess customEvent:... names. Only use values advertised by property metadata.
    dims=[x for x in meta.get("dimensions",[]) if x.get("customDefinition")]
    for candidate in names:
        for d in dims:
            if canonical(d.get("uiName",""))==canonical(candidate): return d["apiName"]
    for candidate in names:
        for d in dims:
            if canonical(candidate) in canonical(d.get("uiName","")): return d["apiName"]
    return None


def available_meta(meta, field):
    return any(x.get("apiName")==field for x in meta.get("dimensions",[])+meta.get("metrics",[]))


def field_health(meta):
    # Availability in metadata only; presence of populated rows requires testing.
    groups={
       "기본 이용": ["activeUsers","screenPageViews","newUsers","userEngagementDuration"],
       "시간대별": ["date","hour","screenPageViews"],
       "구매": [GROSS_REVENUE,"transactions","totalPurchasers"],
       "상품": ["itemId","itemName","itemBrand","itemsViewed","itemsCheckedOut","itemsPurchased","itemRevenue"],
       "유입": ["sessionDefaultChannelGroup","sessionSourceMedium"],
    }
    return {name:all(available_meta(meta,f) for f in fields) for name,fields in groups.items()}


def info_fields(meta):
    d={"page1": ["페이지_1Depth","페이지 1Depth","페이지 1뎁스"],
       "gender":["고객_성별","고객 성별"],"age":["고객_연령대","고객 연령대"],
       "grade":["고객_등급","고객 등급"],"login":["고객_고객 로그인 유형","고객_로그인 유형","로그인 유형"],
       "coupon":["주문 쿠폰","주문_쿠폰","order_coupon"],
       "evt_page":["이벤트_페이지","이벤트 페이지"],
       "evt_area":["이벤트_영역","이벤트 영역"],
       "evt_label":["이벤트_라벨","이벤트 라벨"]}
    return {key:custom_field(meta, names) for key,names in d.items()}


def f_eq(field,value):
    return {"filter":{"fieldName":field,"stringFilter":{"matchType":"EXACT","value":value}}}

def f_contains(field,value):
    return {"filter":{"fieldName":field,"stringFilter":{"matchType":"CONTAINS","value":value,"caseSensitive":False}}}

def f_regex(field,value):
    return {"filter":{"fieldName":field,"stringFilter":{"matchType":"FULL_REGEXP","value":value,"caseSensitive":False}}}

def f_and(*args):
    parts=[x for x in args if x]
    if not parts:return None
    return parts[0] if len(parts)==1 else {"andGroup":{"expressions":parts}}

def f_not(x):return {"notExpression":x}


def report(cfg,token,start,end,dimensions,metrics,filters=None,limit=250,sort=None):
    req={"dateRanges":[{"startDate":str(start),"endDate":str(end)}],
         "dimensions":[{"name":d} for d in dimensions], "metrics":[{"name":m} for m in metrics],
         "limit":str(limit),"keepEmptyRows":False}
    if filters:req["dimensionFilter"]=filters
    if sort: req["orderBys"]=[{"metric":{"metricName":sort},"desc":True}]
    return api(cfg,token,":runReport",req)


def frame(raw):
    """Convert a GA4 report without failing on repeated header names.

    A column label can occur more than once in an API subreport. Selecting
    df[label] in that case produces a DataFrame rather than the 1-D Series
    expected by pd.to_numeric. Preserve the first metric's public name and
    rename later occurrences instead of crashing the whole dashboard.
    """
    ds=[x.get("name", "") for x in raw.get("dimensionHeaders", [])]
    ms=[x.get("name", "") for x in raw.get("metricHeaders", [])]
    original=ds+ms
    used=set()
    columns=[]
    for position,name in enumerate(original):
        label=name or f"unnamed_{position}"
        if label in used:
            index=2
            while f"{label}__{index}" in used:
                index+=1
            label=f"{label}__{index}"
        used.add(label)
        columns.append(label)
    rows=[]
    for entry in raw.get("rows", []):
        vals=([v.get("value", "") for v in entry.get("dimensionValues", [])]
              +[v.get("value", "") for v in entry.get("metricValues", [])])
        rows.append((vals+[""]*len(columns))[:len(columns)])
    df=pd.DataFrame(rows, columns=columns)
    for position in range(len(ds),len(columns)):
        field=columns[position]
        df[field]=pd.to_numeric(df[field],errors="coerce").fillna(0)
    return df


def query(cfg,token,start,end,dimensions,metrics,filters=None,limit=250,sort=None):
    raw=report(cfg,token,start,end,dimensions,metrics,filters,limit,sort)
    return frame(raw),raw


def safe_query(*args,show_error=True,**kwargs):
    try: return query(*args,**kwargs)
    except GAError as exc:
        if show_error: st.warning(f"이 보고서는 조회할 수 없어요. {exc}")
        return None,None


def one(df,metric):
    return None if df is None or df.empty or metric not in df else float(df.iloc[0][metric])


def quality(raw):
    if not raw:return
    meta=raw.get("metadata",{})
    messages=[]
    if meta.get("subjectToThresholding"): messages.append("개인정보 보호 임계값 적용")
    if meta.get("dataLossFromOtherRow"): messages.append("(other) 행 병합 발생")
    if raw.get("rowCount",0)>len(raw.get("rows",[])):messages.append(f"일부 행만 표시 ({len(raw.get('rows',[])):,}/{raw.get('rowCount',0):,})")
    if messages: st.caption("⚠️ "+" · ".join(messages))


def shop_filter():
    return f_contains("unifiedScreenName",LOUNGE_TITLE)


def order_filter():
    return f_and(f_contains("unifiedScreenName",PURCHASE_TITLE),f_eq("eventName","purchase"))


def human_number(x):
    return "—" if x is None or pd.isna(x) else f"{x:,.0f}"

def krw(x):
    if x is None or pd.isna(x): return "조회 불가"
    x=float(x)
    if abs(x)>=100000000:return f"{x/100000000:,.2f}억원"
    if abs(x)>=10000:return f"{x/10000:,.1f}만원"
    return f"{x:,.0f}원"


def engagement(x):
    if x is None or pd.isna(x):return "—"
    s=round(float(x))
    return f"{s//60}분 {s%60:02d}초" if s>=60 else f"{s}초"


def kpi(name,value,note="",type_="num"):
    fmt=(krw(value) if type_=="krw" else engagement(value) if type_=="time" else
         f"{value*100:,.1f}%" if type_=="rate" and value is not None else human_number(value))
    note=html.escape(note)
    st.markdown(f'<div class="kpi" title="{note}"><div class="kpi-name">{html.escape(name)}</div><div class="kpi-value">{fmt}</div>'
                f'<div class="kpi-foot">{note or " "}</div></div>',unsafe_allow_html=True)


def kpi_grid(items, changes=None):
    """A single HTML flow with no columns prevents Streamlit element-height overlap."""
    changes=changes or {}
    def card(item):
        name,value,*other=item
        note=other[0] if other else ""
        unit=other[1] if len(other)>1 else "num"
        val=(krw(value) if unit=="krw" else engagement(value) if unit=="time" else
             f"{value*100:,.1f}%" if unit=="rate" and value is not None else human_number(value))
        if value is None and unit!="krw": val="조회 불가"
        delta=changes.get(name)
        if delta is None:
            change_html='<div class="kpi-delta neutral">&nbsp;</div>'
        else:
            tone='up' if delta>0 else 'down' if delta<0 else 'neutral'
            label=f"직전 기간 {delta:+,.1f}%" if delta else '직전 기간 0.0%'
            change_html=f'<div class="kpi-delta {tone}">{html.escape(label)}</div>'
        note_safe=html.escape(str(note))
        return (f'<div class="kpi" title="{note_safe}">'
                f'<div class="kpi-name">{html.escape(str(name))}</div>'
                f'<div class="kpi-value">{html.escape(str(val))}</div>'
                +change_html+f'<div class="kpi-foot">{note_safe or "&nbsp;"}</div></div>')
    first=''.join(card(x) for x in items[:5])
    second=''.join(card(x) for x in items[5:])
    st.markdown('<div class="kpi-overview"><div class="kpi-overview-row row-five">'+first+
                '</div><div class="kpi-overview-row row-four">'+second+'</div></div>',unsafe_allow_html=True)


def heading(title,desc=""):
    st.markdown(f'<div class="section-title">{html.escape(title)}</div>',unsafe_allow_html=True)
    if desc:st.caption(desc)


def table(df,key,height=330):
    if df is None or df.empty:
        st.info("선택 기간에 표시할 데이터가 없습니다.")
        return
    st.dataframe(df,hide_index=True,use_container_width=True,height=height)
    a,b=st.columns(2)
    with a:
        st.download_button("엑셀 다운로드",excel_bytes(df),file_name=f"{key}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"x_{key}",use_container_width=True)
    with b:
        st.download_button("CSV 다운로드",df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{key}.csv",mime="text/csv",key=f"c_{key}",use_container_width=True)


def excel_bytes(df):
    import io
    buffer=io.BytesIO()
    with pd.ExcelWriter(buffer,engine="xlsxwriter") as writer:
        df.to_excel(writer,sheet_name="분석 결과",index=False)
        ws=writer.sheets["분석 결과"]
        for idx,col in enumerate(df.columns):
            try: max_len=int(df[col].astype(str).str.len().clip(upper=44).max())
            except (TypeError,ValueError):max_len=14
            ws.set_column(idx,idx,min(46,max(13,max_len+3)))
        ws.freeze_panes(1,0)
        ws.autofilter(0,0,len(df),len(df.columns)-1)
    return buffer.getvalue()


def bar(df,label,value,h=295,n=12):
    if df is None or df.empty:st.info("조회된 데이터가 없어요.");return
    values=df.head(n).iloc[::-1].copy()
    original=values[label].fillna("(미설정)").astype(str)
    short=original.map(lambda x:x if len(x)<=24 else x[:22]+"…")
    fig=go.Figure(go.Bar(x=values[value],y=short,customdata=original,orientation="h",marker_color=BLUE,
                         text=values[value].map(lambda x:f"{x:,.0f}"),textposition="auto",
                         hovertemplate="%{customdata}<br>%{x:,.0f}<extra></extra>"))
    fig.update_layout(height=h,margin=dict(l=3,r=8,t=7,b=5),showlegend=False,
                      xaxis=dict(showgrid=True,gridcolor="#EDF1F7"),yaxis=dict(tickfont=dict(size=10)),
                      plot_bgcolor="white",paper_bgcolor="white")
    st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})


def trend_plot(df,metric,period,widget_key):
    if df is None or df.empty:
        st.info("선택 기간의 추이 데이터가 없어요.")
        return None
    field={"일별":"date","주별":"yearWeek","월별":"yearMonth"}[period]
    if field not in df or metric not in df:
        st.warning("차트에 필요한 항목을 찾지 못했어요.")
        return None
    if period=="일별": labels=pd.to_datetime(df[field],format="%Y%m%d",errors="coerce").dt.strftime("%m/%d")
    elif period=="주별":labels=df[field].astype(str).str.slice(0,4)+"-"+df[field].astype(str).str.slice(4)+"주"
    else:labels=df[field].astype(str).str.slice(0,4)+"."+df[field].astype(str).str.slice(4)
    show_numbers=len(df)<=16
    fig=go.Figure(go.Scatter(x=labels,y=df[metric],mode="lines+markers+text" if show_numbers else "lines+markers",
        line=dict(color=BLUE,width=2.8),marker=dict(size=7),
        text=[f"{x:,.0f}" for x in df[metric]] if show_numbers else None,
        textposition="top center",textfont=dict(size=9,color="#687c9c"),
        customdata=df[field],hovertemplate="%{x}<br>%{y:,.0f}<extra></extra>"))
    fig.update_layout(height=300,margin=dict(l=4,r=12,t=28,b=10),clickmode="event+select",
        xaxis=dict(showgrid=False,tickfont=dict(size=10)),
        yaxis=dict(showgrid=True,gridcolor="#EDF1F7"),plot_bgcolor="white",paper_bgcolor="white")
    try:
        event=st.plotly_chart(fig,key=widget_key,use_container_width=True,on_select="rerun",
                              selection_mode="points",config={"displayModeBar":False})
        points=event.selection.points if event and hasattr(event,"selection") else []
        if points and period=="일별":
            x=points[0].get("customdata")
            if isinstance(x,(list,tuple)):x=x[0]
            return str(x) if x else None
    except TypeError: # old Streamlit
        st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
    return None


def hour_table(cfg,token,day,metric,extra=None):
    filter_=f_and(shop_filter(),extra)
    df,raw=safe_query(cfg,token,day,day,["hour"],[metric],filter_,limit=24)
    if df is None:return None
    base=pd.DataFrame({"hour":[f"{k:02d}" for k in range(24)]})
    if df.empty:base[metric]=0
    else:
        df["hour"]=df["hour"].astype(str).str.zfill(2)
        base=base.merge(df[["hour",metric]],on="hour",how="left").fillna(0)
    base["시간"]=base["hour"]+"시"
    return base


def day_drill(cfg,token,start,end,name,kind,chosen,extra=None):
    """Graph → chosen date → hourly report. No API error is replaced with zero."""
    dimension={"일별":"date","주별":"yearWeek","월별":"yearMonth"}[kind]
    df,raw=safe_query(cfg,token,start,end,[dimension],[chosen],f_and(shop_filter(),extra),limit=1100)
    if df is None:return
    if df.empty:st.info("이 지표에 해당하는 날짜별 데이터가 없어요.");return
    df=df.sort_values(dimension).reset_index(drop=True)
    picked=trend_plot(df,chosen,kind,f"chart_{name}_{kind}_{chosen}_{start}_{end}")
    quality(raw)
    if kind!="일별":st.caption("시간대별 상세는 '일별' 모드에서 날짜를 선택하면 나타납니다.");return
    date_options=pd.to_datetime(df["date"],format="%Y%m%d",errors="coerce").dt.strftime("%Y-%m-%d").tolist()
    if not date_options:return
    key=f"day_{name}_{start}_{end}"
    if key not in st.session_state:st.session_state[key]=date_options[-1]
    if picked:
        parsed=pd.to_datetime(picked,format="%Y%m%d",errors="coerce")
        if pd.notna(parsed):
            candidate=parsed.strftime("%Y-%m-%d")
            if candidate in date_options:st.session_state[key]=candidate
    if st.session_state[key] not in date_options:st.session_state[key]=date_options[-1]
    a,b=st.columns([1,4],vertical_alignment="center")
    with a: chosen_day=st.selectbox("시간대별 상세 날짜",date_options,key=key)
    with b: st.caption("일별 그래프의 점을 클릭하거나 날짜를 선택하면 해당 일자의 00~23시를 표시해요.")
    h=hour_table(cfg,token,chosen_day,chosen,extra)
    if h is None:return
    fig=go.Figure(go.Bar(x=h["시간"],y=h[chosen],marker_color=TEAL,hovertemplate="%{x} · %{y:,.0f}<extra></extra>"))
    fig.update_layout(height=200,margin=dict(l=4,r=8,t=6,b=5),xaxis=dict(tickfont=dict(size=10)),
                      yaxis=dict(showgrid=True,gridcolor="#EDF1F7"),paper_bgcolor="white",plot_bgcolor="white")
    st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})


def previous_period(start,end):
    length=(end-start).days+1
    last=start-dt.timedelta(days=1)
    return last-dt.timedelta(days=length-1),last


def change_ratio(current,previous):
    if current is None or previous is None or previous<=0:return None
    return (current-previous)/previous*100


def insights(cards):
    chunks=[]
    for head,body,note in cards[:3]:
        chunks.append('<div class="insight-card">'+
            f'<div class="insight-head">{html.escape(str(head))}</div>'+
            f'<div class="insight-body">{html.escape(str(body))}</div>'+
            f'<div class="insight-sub">{html.escape(str(note))}</div></div>')
    st.markdown('<div class="insight-row">'+''.join(chunks)+'</div>',unsafe_allow_html=True)


def empty_label(v):
    return str(v) in ('', '(not set)', '(other)', 'None', 'nan')


def date_hour_detail(cfg,token,start,end,metric,filter_,key_prefix,subtitle=''):
    """Reuse the same dimension filter for the daily and hourly requests."""
    daily,raw=safe_query(cfg,token,start,end,['date'],[metric],filter_,limit=550)
    if daily is None or daily.empty:
        st.info('선택한 항목의 일별 기록이 없어요.');return
    daily=daily.sort_values('date').reset_index(drop=True)
    picked=trend_plot(daily,metric,'일별',f'{key_prefix}_trend_{start}_{end}')
    dates=pd.to_datetime(daily['date'],format='%Y%m%d',errors='coerce').dropna().dt.strftime('%Y-%m-%d').tolist()
    if not dates:return
    key=f'{key_prefix}_day_{start}_{end}'
    if key not in st.session_state or st.session_state[key] not in dates:st.session_state[key]=dates[-1]
    if picked:
        parsed=pd.to_datetime(picked,format='%Y%m%d',errors='coerce')
        if pd.notna(parsed) and parsed.strftime('%Y-%m-%d') in dates:
            st.session_state[key]=parsed.strftime('%Y-%m-%d')
    selected=st.selectbox('시간대별 상세 날짜',dates,key=key)
    h,_=safe_query(cfg,token,selected,selected,['hour'],[metric],filter_,limit=24)
    if h is None:return
    full=pd.DataFrame({'hour':[f'{i:02d}' for i in range(24)]})
    if h.empty:full[metric]=0
    else:
        h['hour']=h['hour'].astype(str).str.zfill(2)
        full=full.merge(h[['hour',metric]],on='hour',how='left').fillna(0)
    bars=['#7B93CE' if hh in ('11','16') else TEAL for hh in full['hour']]
    fig=go.Figure(go.Bar(x=full['hour']+'시',y=full[metric],marker_color=bars,
                         hovertemplate='%{x}: %{y:,.0f}<extra></extra>'))
    st.caption('11시·16시 막대는 MMS/RCS 발송 시간대(11:30·16:00) 참고용으로 색을 구분했어요. 발송 효과를 확정하는 지표는 아닙니다.')
    fig.update_layout(height=210,margin=dict(l=5,r=8,t=7,b=10),
        xaxis=dict(tickfont=dict(size=10)),yaxis=dict(gridcolor='#EDF1F7'),
        paper_bgcolor='white',plot_bgcolor='white')
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
    if subtitle:st.caption(subtitle)
    quality(raw)


def weekday_hour_heatmap(cfg,token,start,end,mode):
    metric={'페이지 조회수':'screenPageViews','구매 건수':'transactions',
            '구매 금액':GROSS_REVENUE}[mode]
    filter_=order_filter() if mode in ('구매 건수','구매 금액') else shop_filter()
    df,raw=safe_query(cfg,token,start,end,['dayOfWeek','hour'],[metric],filter_,limit=300)
    if df is None or df.empty:return
    # dayOfWeek: 0=Sunday .. 6=Saturday (GA4 property timezone)
    df['d']=pd.to_numeric(df['dayOfWeek'],errors='coerce')
    df['h']=pd.to_numeric(df['hour'],errors='coerce')
    week=[1,2,3,4,5,6,0]
    matrix=[[df.loc[(df['d']==day)&(df['h']==hour),metric].sum() for hour in range(24)] for day in week]
    fig=go.Figure(go.Heatmap(z=matrix,x=[f'{x:02d}시' for x in range(24)],
        y=['월','화','수','목','금','토','일'],colorscale='Blues',
        hovertemplate='%{y}요일 %{x}<br>'+mode+': %{z:,.0f}<extra></extra>'))
    fig.update_layout(height=310,margin=dict(l=5,r=6,t=8,b=8),
                      yaxis=dict(autorange='reversed'),paper_bgcolor='white')
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
    st.caption('GA4 속성 시간대 기준. 11시·16시대 이용과 MMS/RCS의 직접적인 인과관계를 의미하지 않습니다.')
    quality(raw)


def purchase_summary(cfg,token,start,end):
    # Google Analytics UI `총 구매 수익` is grossPurchaseRevenue, not purchaseRevenue.
    df,_=safe_query(cfg,token,start,end,[],[GROSS_REVENUE,"transactions","totalPurchasers"],order_filter(),show_error=False)
    if df is None:return {}
    return {k:one(df,k) for k in [GROSS_REVENUE,"transactions","totalPurchasers"]}


def lounge_summary(cfg,token,start,end):
    df,_=safe_query(cfg,token,start,end,[],["activeUsers","screenPageViews","sessions","userEngagementDuration","newUsers"],shop_filter(),show_error=False)
    if df is None:return {}
    return {k:one(df,k) for k in ["activeUsers","screenPageViews","sessions","userEngagementDuration","newUsers"]}


def returning_summary(cfg,token,start,end):
    df,_=safe_query(cfg,token,start,end,["newVsReturning"],["activeUsers"],shop_filter(),show_error=False)
    if df is None or df.empty:return None
    z=df[df["newVsReturning"].astype(str).str.lower().isin(["returning","established"])]
    return float(z["activeUsers"].sum()) if not z.empty else None


def funnel(cfg,token,start,end,field):
    """Closed funnel: lounge screen_view -> purchase recorded on order-complete screen."""
    if field:
        first={'andGroup':{'expressions':[
            {'funnelEventFilter':{'eventName':'screen_view'}},
            {'funnelFieldFilter':{'fieldName':field,'stringFilter':{'matchType':'EXACT','value':LOUNGE_TITLE}}}
        ]}}
    else:
        first={'andGroup':{'expressions':[
            {'funnelEventFilter':{'eventName':'screen_view'}},
            {'funnelFieldFilter':{'fieldName':'unifiedScreenName',
                'stringFilter':{'matchType':'CONTAINS','value':LOUNGE_TITLE}}}
        ]}}
    second={'andGroup':{'expressions':[
        {'funnelEventFilter':{'eventName':'purchase'}},
        {'funnelFieldFilter':{'fieldName':'unifiedScreenName',
            'stringFilter':{'matchType':'CONTAINS','value':PURCHASE_TITLE}}}
    ]}}
    body={'dateRanges':[{'startDate':str(start),'endDate':str(end)}],
          'funnel':{'isOpenFunnel':False,'steps':[
              {'name':'쇼핑라운지 진입','filterExpression':first},
              {'name':'주문완료 구매','filterExpression':second}]},'limit':'10'}
    try:
        data=api(cfg,token,':runFunnelReport',body,alpha=True)
        for section in ('funnelVisualization','funnelTable'):
            part=data.get(section) or {}
            metrics=[x.get('name') for x in part.get('metricHeaders',[])]
            dims=[x.get('name') for x in part.get('dimensionHeaders',[])]
            if 'activeUsers' not in metrics or 'funnelStepName' not in dims:continue
            mi=metrics.index('activeUsers');di=dims.index('funnelStepName')
            found={}
            for row in part.get('rows',[]):
                dvals=row.get('dimensionValues',[]);mvals=row.get('metricValues',[])
                if di>=len(dvals) or mi>=len(mvals):continue
                label=dvals[di].get('value','')
                try:number=float(mvals[mi].get('value',''))
                except (ValueError,TypeError):continue
                if label.startswith('1.') or '쇼핑라운지 진입' in label:found[1]=number
                elif label.startswith('2.') or '주문완료 구매' in label:found[2]=number
            if 1 in found and 2 in found:return found[1],found[2]
    except (GAError,TypeError,ValueError,KeyError,IndexError):pass
    return None


def coupon_data(cfg,token,start,end,fields,sort_metric='transactions'):
    field=fields.get('coupon')
    if not field:return None,None
    filt=f_and(order_filter(),f_not(f_eq(field,'(not set)')),f_not(f_eq(field,'')))
    return safe_query(cfg,token,start,end,[field],['transactions',GROSS_REVENUE],
                      filt,limit=1000,sort=sort_metric)


def top_cards(cfg,token,start,end,fields):
    """Three independently sourced highlights, without inventing an attribution link."""
    p,_=page_data(cfg,token,start,end)
    best_promo=None
    codes=set()
    if p is not None and not p.empty:
        subset=p[p['페이지 유형']=='기획전'].sort_values('screenPageViews',ascending=False)
        if not subset.empty:
            row=subset.iloc[0]
            best_promo=(f"기획전 {promotion_id(row['unifiedScreenName']) or '화면'}",
                        f"조회 {human_number(row['screenPageViews'])}회")
        items=p[p['페이지 유형']=='상품 상세']['unifiedScreenName'].astype(str)
        codes=set(items.str.extract(r'>상품상세>(\d{6,})',expand=False).dropna())
    c,_=coupon_data(cfg,token,start,end,fields)
    best_coupon=None
    if c is not None and not c.empty:
        c=c[~c[fields['coupon']].astype(str).map(empty_label)]
        if not c.empty:
            row=c.sort_values('transactions',ascending=False).iloc[0]
            best_coupon=(str(row[fields['coupon']]),f"쿠폰 포함 거래 {human_number(row['transactions'])}건")
    best_item=None
    if codes:
        d,_=safe_query(cfg,token,start,end,['itemId','itemName'],['itemsPurchased'],
                       limit=6000,sort='itemsPurchased',show_error=False)
        if d is not None and not d.empty:
            d=d[d['itemId'].astype(str).isin(codes)]
            if not d.empty:
                row=d.sort_values('itemsPurchased',ascending=False).iloc[0]
                best_item=(str(row['itemName']),f"구매 수량 {human_number(row['itemsPurchased'])}개 · 전체 속성")
    cards=[('인기 기획전',best_promo,'쇼핑라운지 기획전 조회수 기준'),
           ('많이 사용한 쿠폰',best_coupon,'주문완료 화면의 쿠폰 포함 거래 수'),
           ('인기 상품(참고)',best_item,'상품 ID 매칭 · 전체 KT멤버십 속성 구매 수량')]
    return [(h,x[0] if x else '표시할 데이터 없음',x[1] if x else scope) for h,x,scope in cards]


def overview(cfg,token,start,end,fields):
    heading('핵심 운영 지표','활성 고객 · 구매 · 페이지 이용을 한눈에 확인합니다. 전체 KPI 9개는 그대로 유지해요.')
    current=lounge_summary(cfg,token,start,end)
    sales=purchase_summary(cfg,token,start,end)
    returning=returning_summary(cfg,token,start,end)
    active=current.get('activeUsers')
    first=current.get('newUsers')
    if first==0 and active and active>0:first=None
    total_time=current.get('userEngagementDuration')
    avg=(total_time/active) if active and total_time is not None else None
    closed=funnel(cfg,token,start,end,fields.get('page1'))
    conversion=closed[1]/closed[0] if closed and closed[0] else None
    changes={}
    notes=[]
    prev_label=''
    if st.session_state.get('ga4_compare_v6',True):
        p_start,p_end=st.session_state.get('ga4_compare_dates_v6',previous_period(start,end))
        prev_label=f'{p_start:%m/%d}~{p_end:%m/%d}'
        prior=lounge_summary(cfg,token,p_start,p_end)
        prev_sales=purchase_summary(cfg,token,p_start,p_end)
        for name,now,old in [
            ('활성 사용자',active,prior.get('activeUsers')),
            ('세션 수',current.get('sessions'),prior.get('sessions')),
            ('페이지 조회수',current.get('screenPageViews'),prior.get('screenPageViews')),
            ('구매 금액',sales.get(GROSS_REVENUE),prev_sales.get(GROSS_REVENUE)),
            ('구매 건수',sales.get('transactions'),prev_sales.get('transactions'))]:
            changes[name]=change_ratio(now,old)
        for name in ('활성 사용자','구매 금액','구매 건수'):
            delta=changes.get(name)
            if delta is not None:
                notes.append((name,f"직전 동일 기간보다 {abs(delta):.1f}% {'증가' if delta>0 else '감소' if delta<0 else '변화 없음'}",
                              '원인 확정이 아닌 관찰된 기간별 변화'))
        if not notes:notes.append(('기간 비교','비교 가능한 지표가 부족해요.','집계 범위와 기간을 확인해 주세요.'))
        st.caption(f'비교 기간: {p_start:%Y.%m.%d} ~ {p_end:%Y.%m.%d}' +
                   (' · 일수가 달라 직접 비교에 유의하세요' if (p_end-p_start).days != (end-start).days else ' · 동일 일수'))
    if not notes:
        notes=[('활성 사용자',f'{human_number(active)}명' if active is not None else '조회할 수 없음','쇼핑라운지 화면 기록 기준'),
               ('구매 수익',krw(sales.get(GROSS_REVENUE)),'주문완료 화면 purchase 이벤트 기준'),
               ('구매 거래',f"{human_number(sales.get('transactions'))}건",'동일 화면의 거래 수')]
    items=[
        ('활성 사용자',active,'쇼핑라운지 화면을 이용한 활성 고객'),
        ('세션 수',current.get('sessions'),'쇼핑라운지 화면과 연결된 방문 세션'),
        ('페이지 조회수',current.get('screenPageViews'),'쇼핑라운지 화면이 열린 횟수'),
        ('구매 금액',sales.get(GROSS_REVENUE),'주문완료 화면 총 구매 수익','krw'),
        ('신규 사용자',first,'앱 first_open 성격 · 쇼핑라운지 최초 방문과 다름'),
        ('평균 참여 시간',avg,'활성 사용자당 참여 시간','time'),
        ('구매 건수',sales.get('transactions'),'쇼핑라운지 주문완료 화면 거래 수'),
        ('구매 전환율',conversion,'닫힌 진입→주문완료 퍼널 기준','rate'),
        ('재방문 사용자',returning,'GA4 기존 이용자 기준 · 쇼핑라운지 반복 방문과 다름'),
    ]
    kpi_grid(items,changes)
    with st.expander('ⓘ 숫자와 집계 기준을 쉽게 설명해 드려요',expanded=False):
        st.markdown("""**활성 사용자**: 선택 기간에 쇼핑라운지 관련 화면을 이용한 고객 수.  
**페이지 조회수**: 화면이 열린 횟수로, 같은 사람이 여러 번 열면 여러 번 집계돼요.  
**세션 수**: 방문 세션 수로, 사람 수와 다릅니다.  
**구매 금액·건수**: 쇼핑라운지 *주문완료 화면의 purchase* 기준이며 결제 시스템 정산 확정액과는 다를 수 있어요.  
**신규 사용자**: GA4 앱 첫 이용 기준입니다. 0이 나오는 등 수집 방식이 불분명하면 '조회 불가'로 표시합니다.  
**재방문 사용자**: GA4의 returning 상태 기준이며 쇼핑라운지 자체 재방문을 뜻하지는 않아요.  
**구매 전환율**: 동일한 퍼널에서 진입 후 주문완료 단계에 도달한 사용자 비율로, API에서 확인되면 표시합니다.""")
    heading('이번 기간의 변화')
    insights(notes)
    if any(abs(changes.get(k) or 0)>=30 for k in ('활성 사용자','구매 금액','구매 건수')):
        st.caption('참고: 직전 기간 대비 30% 이상 움직인 지표가 있습니다. 원인을 확정하는 경고가 아니라 변화 폭을 알려주는 안내예요.')
    heading('가장 많이 본 기획전 · 사용한 쿠폰 · 구매한 상품')
    insights(top_cards(cfg,token,start,end,fields))
    st.caption('상품은 쇼핑라운지 상품 상세 화면 ID와 일치하는 상품의 **KT멤버십 전체 속성** 구매 수량입니다. 쇼핑라운지 단독 매출로 해석하지 마세요.')
    heading('일별 추이와 시간대 패턴')
    mode=st.radio('추이 종류',['날짜별 추이','요일 × 시간대'],horizontal=True,label_visibility='collapsed',key='overview_graph_mode')
    if mode=='요일 × 시간대':
        heat_metric=st.selectbox('시간대 분석 지표',['페이지 조회수','구매 건수','구매 금액'],key='heat_metric')
        weekday_hour_heatmap(cfg,token,start,end,heat_metric)
    else:
        a,b=st.columns([1,2],vertical_alignment='bottom')
        with a:gran=st.segmented_control('묶음 단위',['일별','주별','월별'],default='일별',key='overview_gran')
        with b:label=st.selectbox('차트에 표시할 항목',['활성 사용자','세션 수','페이지 조회수','구매 건수','구매 금액'],key='overview_plot_metric')
        metric={'활성 사용자':'activeUsers','세션 수':'sessions','페이지 조회수':'screenPageViews',
                '구매 건수':'transactions','구매 금액':GROSS_REVENUE}[label]
        if label in ('구매 건수','구매 금액'):purchase_drill(cfg,token,start,end,gran,metric)
        else:day_drill(cfg,token,start,end,'overview',gran,metric)
        st.caption('일별 그래프의 날짜를 클릭하면 00~23시 상세를 확인할 수 있어요.')


def purchase_drill(cfg,token,start,end,kind,metric):
    dim={"일별":"date","주별":"yearWeek","월별":"yearMonth"}[kind]
    df,raw=safe_query(cfg,token,start,end,[dim],[metric],order_filter(),limit=1100)
    if df is None or df.empty:return
    df=df.sort_values(dim).reset_index(drop=True)
    picked=trend_plot(df,metric,kind,f"purch_plot_{kind}_{metric}_{start}_{end}")
    if kind!="일별":return
    choices=pd.to_datetime(df["date"],format="%Y%m%d",errors="coerce").dt.strftime("%Y-%m-%d").tolist()
    if not choices:return
    key=f"purchase_day_{start}_{end}"
    if key not in st.session_state or st.session_state[key] not in choices:st.session_state[key]=choices[-1]
    if picked:
        day=pd.to_datetime(picked,format="%Y%m%d",errors="coerce")
        if pd.notna(day) and day.strftime("%Y-%m-%d") in choices:st.session_state[key]=day.strftime("%Y-%m-%d")
    chosen=st.selectbox("시간대별 상세 날짜",choices,key=key)
    h,r=safe_query(cfg,token,chosen,chosen,["hour"],[metric],order_filter(),limit=24)
    if h is None:return
    base=pd.DataFrame({"hour":[f"{i:02d}" for i in range(24)]})
    if h.empty:base[metric]=0
    else:
        h["hour"]=h["hour"].astype(str).str.zfill(2)
        base=base.merge(h[["hour",metric]],on="hour",how="left").fillna(0)
    fig=go.Figure(go.Bar(x=base["hour"].map(lambda x:x+"시"),y=base[metric],marker_color=TEAL))
    fig.update_layout(height=205,margin=dict(l=5,r=8,t=8,b=5),xaxis=dict(tickfont=dict(size=10)),
                      yaxis=dict(showgrid=True,gridcolor="#EDF1F7"),plot_bgcolor="white",paper_bgcolor="white")
    st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})


def customers(cfg,token,start,end,fields):
    heading('어떤 고객이 많이 이용할까요?',
            '성별과 연령을 각각 확인합니다. 동일 고객이 여러 값에 집계되는 경우가 있어 전체를 단순 합산하지 않아요.')
    selected=st.radio('고객 구분',['기기','성별','연령','멤버십 등급','로그인 유형','기기별 구매'],
                      horizontal=True,key='v6_customer_select',label_visibility='collapsed')
    dims={'기기':('operatingSystem','운영체제'),
          '성별':(fields.get('gender') or 'userGender','성별'),
          '연령':(fields.get('age') or 'userAgeBracket','연령대'),
          '멤버십 등급':(fields.get('grade'),'등급'),
          '로그인 유형':(fields.get('login'),'로그인 방식')}
    if selected=='기기별 구매':
        d,r=safe_query(cfg,token,start,end,['operatingSystem'],['transactions',GROSS_REVENUE],
                       order_filter(),limit=30,sort=GROSS_REVENUE)
        if d is None or d.empty:return
        d['operatingSystem']=d['operatingSystem'].replace({'iOS':'아이폰(iOS)','Android':'안드로이드(Android)'})
        a,b=st.columns([.85,1.35],gap='large')
        with a:bar(d,'operatingSystem','transactions',h=290,n=10)
        with b:table(d.rename(columns={'operatingSystem':'운영체제','transactions':'구매 거래 수',
                                        GROSS_REVENUE:'구매 수익'}),f'ga4_기기별구매_{start}_{end}',height=285)
        st.caption('쇼핑라운지 주문완료 화면의 purchase 이벤트에 기록된 운영체제 기준입니다.')
        quality(r)
        return
    dim,title=dims[selected]
    if not dim:
        st.info(f'{selected} 분석에 필요한 GA4 맞춤 항목을 찾지 못했어요. 아래 데이터 진단에서 확인할 수 있어요.')
        return
    df,raw=safe_query(cfg,token,start,end,[dim],['activeUsers','screenPageViews'],shop_filter(),
                      limit=150,sort='activeUsers')
    if df is None or df.empty:return
    if selected=='기기':
        df[dim]=df[dim].replace({'iOS':'아이폰(iOS)','Android':'안드로이드(Android)'})
    if selected=='연령':df[dim]=df[dim].replace({'(not set)':'미설정'})
    df=df.rename(columns={dim:title,'activeUsers':'이용 고객 수','screenPageViews':'페이지 조회수'})
    left,right=st.columns([1,1.2],gap='large')
    with left:bar(df,title,'이용 고객 수',h=340,n=15)
    with right:table(df,f'ga4_고객_{selected}_{start}_{end}',height=320)
    if selected=='성별':st.caption('성별 × 연령 교차표는 요청에 따라 제외했습니다.')
    st.caption('이용 고객 수는 분류별 중복 집계가 가능하며, 앱 첫 이용 기준인 GA4 신규 사용자와는 다릅니다.')
    quality(raw)


def click_category(event_page,event_area,event_label):
    """Classify only when event tags contain identifiable placement words."""
    s=' '.join([str(x) for x in (event_area,event_label,event_page)]).casefold()
    if any(k in s for k in ('퀵버튼','퀵 버튼','퀵메뉴','quick','바로가기')):return '퀵버튼·바로가기'
    if any(k in s for k in ('배너','banner','bnr')):return '배너'
    if any(k in s for k in ('하단','바텀','bottom','footer','하단메뉴','gnb')):return '하단 메뉴'
    if any(k in s for k in ('팝업','popup','pop-up')):return '팝업'
    if any(k in s for k in ('상단메뉴','상단 메뉴','top menu')):return '상단 메뉴'
    if any(k in s for k in ('탭','tab','메뉴')):return '기타 메뉴'
    return '유형 미분류'


def click_reports(cfg,token,start,end,fields):
    if not fields.get('evt_label'):
        st.warning('이벤트_라벨 맞춤 측정기준을 찾지 못해 앱 내 쇼핑 관련 클릭 경로를 구분할 수 없어요.')
        return None,None,None
    dims=[(k,fields.get(k)) for k in ('evt_page','evt_area','evt_label') if fields.get(k)]
    columns=[val for key,val in dims]
    filt=f_and(f_eq('eventName','click_event'),f_eq('sessionDefaultChannelGroup','Direct'),
               f_regex(fields['evt_label'],r'.*(쇼핑|shopping).*'))
    df,raw=safe_query(cfg,token,start,end,columns,['eventCount'],filt,limit=10000,sort='eventCount')
    if df is None or df.empty:return df,raw,filt
    for key,api_name in dims:
        df[key]=df[api_name].fillna('').astype(str)
    for key in ('evt_page','evt_area','evt_label'):
        if key not in df:df[key]=''
    df['유형']=df.apply(lambda row:click_category(row['evt_page'],row['evt_area'],row['evt_label']),axis=1)
    df['경로명']=df.apply(lambda r:' · '.join(v for v in [r['evt_page'],r['evt_area'],r['evt_label']]
                                        if v and not empty_label(v)),axis=1)
    df['경로명']=df['경로명'].replace('', '라벨 확인 필요')
    return df,raw,filt


def channels(cfg,token,start,end,fields):
    heading('고객은 어디에서 들어오고 무엇을 클릭할까요?',
            'KT멤버십 앱 내 쇼핑 관련 클릭 경로와 GA4 세션 유입 채널을 구분해 보여드려요.')
    mode=st.radio('유입 보고서',['KT멤버십 앱 내 쇼핑 클릭 순위','세션 유입 채널'],
                  horizontal=True,label_visibility='collapsed',key='v6_acquisition')
    if mode=='세션 유입 채널':
        df,raw=safe_query(cfg,token,start,end,['sessionDefaultChannelGroup','sessionSourceMedium'],
                          ['screenPageViews','activeUsers'],shop_filter(),limit=1000,sort='screenPageViews')
        if df is None or df.empty:return
        grouped=df.groupby('sessionDefaultChannelGroup',as_index=False)['screenPageViews'].sum()
        grouped=grouped.sort_values('screenPageViews',ascending=False)
        left,right=st.columns([.85,1.3])
        with left:bar(grouped,'sessionDefaultChannelGroup','screenPageViews',h=330,n=12)
        with right:table(df.rename(columns={'sessionDefaultChannelGroup':'유입 채널','sessionSourceMedium':'소스 / 매체',
                                            'screenPageViews':'페이지 조회수','activeUsers':'이용 고객 수'}),
                          f'ga4_유입채널_{start}_{end}',height=320)
        st.info('**Direct는 곧 KT멤버십 앱 유입이 아니며**, Unassigned를 MMS라고 단정할 수도 없어요. 별도 추적값이 필요합니다.')
        quality(raw)
        return
    df,raw,base_filter=click_reports(cfg,token,start,end,fields)
    if df is None or df.empty:
        st.info('선택 기간에 쇼핑 관련 클릭 이벤트를 조회하지 못했어요.')
        return
    if raw and raw.get('rowCount',0)>len(df):
        st.warning('클릭 경로가 많아 일부 행만 조회됐어요. 이 순위는 전체 경로 순위가 아닐 수 있습니다.')
    total=int(df['eventCount'].sum())
    st.caption(f"조회된 쇼핑 관련 클릭 이벤트 합계 **{total:,}회** · 클릭 수 기준, 방문 완료 기준 아님")
    tab=st.radio('순위 유형',['영역별 순위','개별 메뉴·배너 순위'],horizontal=True,
                 label_visibility='collapsed',key='v6_click_group')
    if tab=='영역별 순위':
        grouped=(df.groupby('유형',as_index=False)['eventCount'].sum()
                 .sort_values('eventCount',ascending=False).reset_index(drop=True))
        grouped['클릭 비중 (%)']=(grouped['eventCount']/total*100).round(1) if total else 0
        col1,col2=st.columns([1,1.28],gap='large')
        with col1:bar(grouped,'유형','eventCount',h=325,n=10)
        with col2:table(grouped.rename(columns={'유형':'클릭 영역','eventCount':'클릭 수'}),
                        f'ga4_KT앱클릭_영역순위_{start}_{end}',height=320)
        st.caption('클릭 위치가 이벤트 태그에 기록된 경우에만 자동 분류해요. 「유형 미분류」는 임의로 메인배너·메뉴로 분류하지 않았습니다.')
    else:
        df=df.sort_values('eventCount',ascending=False).reset_index(drop=True)
        left,right=st.columns([1,1.35],gap='large')
        with left:bar(df,'경로명','eventCount',h=355,n=12)
        with right:
            show=df[['유형','evt_page','evt_area','evt_label','eventCount']].rename(
                columns={'유형':'클릭 영역','evt_page':'이벤트 페이지','evt_area':'이벤트 위치',
                         'evt_label':'이벤트 라벨','eventCount':'클릭 수'})
            table(show.head(300),f'ga4_KT앱클릭_개별순위_{start}_{end}',height=345)
        with st.expander('특정 메뉴·배너 선택 → 일별·시간대별 클릭',expanded=False):
            choices=df.head(250).index.to_list()
            if choices:
                ix=st.selectbox('분석할 클릭 경로',choices,
                    format_func=lambda n:f"{str(df.loc[n,'경로명'])[:90]} · {int(df.loc[n,'eventCount']):,}회",
                    key='v6_click_selected')
                r=df.loc[ix]
                detail=base_filter
                for key in ('evt_page','evt_area','evt_label'):
                    if fields.get(key):detail=f_and(detail,f_eq(fields[key],str(r[key])))
                date_hour_detail(cfg,token,start,end,'eventCount',detail,
                                 f'click_{ix}_{hashlib.md5(str(r["경로명"]).encode()).hexdigest()[:10]}',
                                 '이 데이터는 클릭 기록이며 쇼핑라운지 방문 완료나 구매 여부는 나타내지 않습니다.')
    st.caption('집계 기준: GA4 click_event · Direct · 이벤트 라벨에 쇼핑/shopping 포함. 다른 쇼핑 관련 링크가 일부 포함될 수 있어요.')
    quality(raw)


def page_class(title):
    s=str(title)
    if ">기획전>" in s:return "기획전"
    if ">상품>상품상세>" in s:return "상품 상세"
    for match, name in [("장바구니","장바구니"),("쿠폰","쿠폰"),("검색","검색"),("주문","주문"),("마이","마이"),("메인","메인"),("베스트","베스트")]:
        if match in s:return name
    return "기타"


def page_data(cfg,token,start,end):
    df,raw=safe_query(cfg,token,start,end,["unifiedScreenName"],["screenPageViews","activeUsers"],shop_filter(),limit=8000,sort="screenPageViews")
    if df is not None and not df.empty:df["페이지 유형"]=df["unifiedScreenName"].apply(page_class)
    return df,raw


def pages(cfg,token,start,end):
    heading('고객이 가장 많이 이용한 화면은?',
            '쇼핑라운지 화면을 유형별로 묶고, 실제 화면별 조회수와 활성 이용자를 확인합니다.')
    df,raw=page_data(cfg,token,start,end)
    if df is None or df.empty:return
    grouped=df.groupby('페이지 유형',as_index=False)['screenPageViews'].sum().sort_values('screenPageViews',ascending=False)
    a,b=st.columns([.85,1.4],gap='large')
    with a:bar(grouped,'페이지 유형','screenPageViews',h=345)
    with b:
        phrase=st.text_input('화면 검색',placeholder='예: 메인, 쿠폰, 장바구니',key='v6_page_search')
        d=df.copy()
        if phrase:d=d[d['unifiedScreenName'].astype(str).str.contains(re.escape(phrase),case=False,na=False)]
        show=d.rename(columns={'unifiedScreenName':'화면 이름','screenPageViews':'조회수',
                               'activeUsers':'이용 고객 수'})
        table(show[['페이지 유형','화면 이름','조회수','이용 고객 수']].head(250),
              f'ga4_페이지순위_{start}_{end}',height=335)
    with st.expander('화면 선택 → 일별·시간대별 조회',expanded=False):
        if not d.empty:
            name=st.selectbox('보고 싶은 화면',d['unifiedScreenName'].head(300).astype(str).tolist(),key='v6_page_selected')
            date_hour_detail(cfg,token,start,end,'screenPageViews',f_eq('unifiedScreenName',name),
                             'singlepage_'+hashlib.md5(name.encode()).hexdigest()[:9],
                             '화면 조회 이벤트 기준입니다. 페이지 사이의 실제 이동 순서는 의미하지 않습니다.')
    quality(raw)


def promotion_id(title):
    s=str(title)
    found=re.search(r">기획전>(\d{4,})",s)
    return found.group(1) if found else ""


def promotions(cfg,token,start,end):
    heading('어떤 기획전이 가장 많은 관심을 받았을까요?',
            '기획전별 조회수와 이용 고객 수를 각각 비교할 수 있어요. 실제 구매 귀속 매출은 별도 태깅이 필요합니다.')
    df,raw=page_data(cfg,token,start,end)
    if df is None or df.empty:return
    p=df[df['페이지 유형']=='기획전'].copy()
    if p.empty:st.info('해당 기간의 기획전 페이지 기록이 없습니다.');return
    p['기획전 ID']=p['unifiedScreenName'].map(promotion_id)
    rank=st.segmented_control('순위 기준',['조회수','이용 고객 수','조회수 급상승'],default='조회수',key='v6_promo_rank')
    key='screenPageViews' if rank!='이용 고객 수' else 'activeUsers'
    if rank=='조회수 급상승':
        ps,pe=st.session_state.get('ga4_compare_dates_v6',previous_period(start,end))
        before,_=page_data(cfg,token,ps,pe)
        if before is not None and not before.empty:
            old=before.set_index('unifiedScreenName')['screenPageViews'].to_dict()
            p['비교 기간 조회수']=p['unifiedScreenName'].map(old).fillna(0)
            p['조회수 증감(%)']=p.apply(lambda r:
                round((r['screenPageViews']-r['비교 기간 조회수'])/r['비교 기간 조회수']*100,1)
                if r['비교 기간 조회수']>=50 else None,axis=1)
            p=p[p['조회수 증감(%)'].notna()].copy()
            if p.empty:st.info('비교 기간 조회수 50회 이상인 기획전이 없어 급상승 순위를 만들 수 없어요.');return
            key='조회수 증감(%)'
        else:st.warning('비교 기간 기획전 데이터를 확인할 수 없습니다.');return
    p=p.sort_values(key,ascending=False).reset_index(drop=True)
    p['선택 기간 일평균 조회수']=(p['screenPageViews']/max(1,(end-start).days+1)).round(1)
    p['기획전']=p.apply(lambda row:'기획전 '+str(row['기획전 ID'] or 'ID 미확인'),axis=1)
    a,b=st.columns([.85,1.4],gap='large')
    with a:bar(p,'기획전',key,h=355,n=12)
    with b:
        show_fields=['기획전 ID','unifiedScreenName','screenPageViews','activeUsers','선택 기간 일평균 조회수']
        if rank=='조회수 급상승':show_fields+=['비교 기간 조회수','조회수 증감(%)']
        d=p[show_fields].rename(columns={'unifiedScreenName':'기획전 화면명',
                                      'screenPageViews':'조회수','activeUsers':'이용 고객 수'})
        table(d.head(200),f'ga4_기획전_{rank}_{start}_{end}',height=345)
    st.caption('일평균은 선택 기간의 달력 일수로 나눈 값이며 기획전 **실제 운영일수**는 아닙니다. 급상승 순위는 비교 기간 조회 50회 이상인 기획전만 계산해 극단적인 변화율을 줄였어요.')
    with st.expander('조회 발생일당 평균도 비교하기',expanded=False):
        if st.checkbox('기획전 조회 발생일별 평균 계산',key='v6_promo_active_days'):
            detail,detail_raw=safe_query(cfg,token,start,end,['unifiedScreenName','date'],
                ['screenPageViews'],f_and(shop_filter(),f_contains('unifiedScreenName','>기획전>')),
                limit=20000,show_error=False)
            if detail is not None and not detail.empty:
                days=detail.groupby('unifiedScreenName')['date'].nunique()
                rows=p[['기획전 ID','unifiedScreenName','screenPageViews']].copy()
                rows['조회 발생일 수']=rows['unifiedScreenName'].map(days)
                rows['조회 발생일당 평균']=rows.apply(lambda r:round(r['screenPageViews']/r['조회 발생일 수'],1)
                    if pd.notna(r['조회 발생일 수']) and r['조회 발생일 수']>0 else None,axis=1)
                table(rows.rename(columns={'unifiedScreenName':'기획전 화면명','screenPageViews':'조회수'}).head(100),
                    f'ga4_기획전_조회발생일_{start}_{end}',height=285)
                st.caption('조회 발생일은 실제 기획전 운영일이 아닙니다. 그날 조회 데이터가 있었다는 뜻이에요.')
                quality(detail_raw)
            else:st.info('날짜별 상세 조회가 가능하지 않습니다.')
    with st.expander('기획전 하나 선택 → 일별·시간대별 조회',expanded=False):
        names=p['unifiedScreenName'].astype(str).head(200).tolist()
        if names:
            choice=st.selectbox('기획전',names,format_func=lambda s:
                                '기획전 '+(promotion_id(s) or '화면')+' · '+s[-55:],key='v6_promo_selected')
            date_hour_detail(cfg,token,start,end,'screenPageViews',
                             f_eq('unifiedScreenName',choice),
                             'promo_'+hashlib.md5(choice.encode()).hexdigest()[:9])
    quality(raw)


def shop_item_ids(cfg,token,start,end):
    # IDs observed in lounge item detail screens. This is not a user/session revenue attribution.
    df,raw=page_data(cfg,token,start,end)
    if df is None or df.empty:return set(),raw
    p=df[df["페이지 유형"]=="상품 상세"]
    codes=p["unifiedScreenName"].str.extract(r">상품상세>(\d{6,})",expand=False).dropna()
    return set(codes.astype(str)),raw


def products(cfg,token,start,end):
    heading('어떤 상품과 브랜드의 반응이 좋을까요?',
            '쇼핑라운지 상품 상세 화면에서 발견된 상품 ID로 전체 속성 상품 이벤트를 찾아 비교합니다.')
    category=st.radio('상품 분석',['쇼핑라운지 상품 ID 매칭','브랜드 전체 참고'],horizontal=True,
                      label_visibility='collapsed',key='v6_product_category')
    is_item=category=='쇼핑라운지 상품 ID 매칭'
    codes,_=shop_item_ids(cfg,token,start,end)
    if is_item and not codes:
        st.warning('해당 기간 쇼핑라운지 상품 상세 화면에 기록된 상품 ID가 없어 연결할 수 없어요.')
        return
    dimensions=['itemId','itemName','itemBrand'] if is_item else ['itemBrand']
    measures=['itemsViewed','itemsCheckedOut','itemsPurchased','itemRevenue']
    df,raw=safe_query(cfg,token,start,end,dimensions,measures,
                      limit=15000,sort='itemRevenue')
    if df is None or df.empty:return
    if is_item:
        df=df[df['itemId'].astype(str).isin(codes)].copy()
        st.caption(f'쇼핑라운지 화면에서 관찰된 상품 ID {len(codes):,}개와 매칭한 결과입니다.')
        if df.empty:
            st.warning('조회된 상품 항목에서 ID가 일치하는 상품이 없어요. 실제 매출 0원을 뜻하지 않습니다.')
            return
        find=st.text_input('상품명·상품 ID 검색',placeholder='예: 비타500 / 132237997',key='v6_product_search')
        if find:
            mask=df['itemName'].astype(str).str.contains(re.escape(find),case=False,na=False)|\
                 df['itemId'].astype(str).str.contains(re.escape(find),case=False,na=False)
            df=df[mask].copy()
    label=st.selectbox('순위를 비교할 항목',['많이 조회한 상품','많이 구매한 상품','상품 수익 높은 순'],key='v6_product_sort')
    ranking={'많이 조회한 상품':'itemsViewed','많이 구매한 상품':'itemsPurchased',
             '상품 수익 높은 순':'itemRevenue'}[label]
    df=df.sort_values(ranking,ascending=False).reset_index(drop=True)
    if df.empty:st.info('조건에 맞는 상품이 없어요.');return
    show=df.rename(columns={'itemId':'상품 ID','itemName':'상품명','itemBrand':'브랜드',
        'itemsViewed':'조회 수량','itemsCheckedOut':'결제 진행 수량',
        'itemsPurchased':'구매 수량','itemRevenue':'상품 수익'})
    if is_item and 'itemsViewed' in df:
        # Important: quantity ratio is NOT GA4 purchase conversion rate.
        show['참고: 조회 수량 대비 구매 수량 (%)']=df.apply(lambda r:
            round(r['itemsPurchased']/r['itemsViewed']*100,1) if r['itemsViewed']>0 else None,axis=1)
    a,b=st.columns([1,1.35],gap='large')
    with a:bar(df,'itemName' if is_item else 'itemBrand',ranking,h=350,n=10)
    with b:table(show.head(300),f'ga4_상품_{label}_{start}_{end}',height=340)
    st.info('**주의:** 상품 수익·구매 수량은 동일 상품 ID의 KT멤버십 **전체 속성**에서 기록된 값이에요. 쇼핑라운지 단독 기여 매출로 볼 수 없어요. 조회 수량 대비 구매 수량은 단순 비율이지 고객 구매 전환율이 아닙니다.')
    if raw and raw.get('rowCount',0)>len(raw.get('rows',[])):
        st.caption('상품 결과 일부만 불러왔습니다. 상위 목록에 포함되지 않은 상품이 있을 수 있어요.')
    quality(raw)
    if is_item:product_detail(cfg,token,start,end,df)


def product_detail(cfg,token,start,end,df):
    if df.empty or 'itemId' not in df:return
    with st.expander('상품 선택 → 일별·시간대별 상세 (전체 속성)',expanded=False):
        names={str(r['itemId']):str(r['itemName']) for _,r in df.head(400).iterrows()}
        ids=list(names)
        if not ids:return
        chosen=st.selectbox('상품',ids,format_func=lambda x:names[x][:60]+' · '+x,key='v6_product_detail')
        title=st.segmented_control('상세 측정 항목',['조회 수량','결제 진행 수량','구매 수량','상품 수익'],default='구매 수량')
        metric={'조회 수량':'itemsViewed','결제 진행 수량':'itemsCheckedOut',
                '구매 수량':'itemsPurchased','상품 수익':'itemRevenue'}[title]
        date_hour_detail(cfg,token,start,end,metric,f_eq('itemId',chosen),
                         f'item_{chosen}_{metric}',
                         '상품 ID 일치 기준 KT멤버십 전체 속성 이벤트입니다. 쇼핑라운지 단독 구매 전환율이 아닙니다.')


def purchase_actions(cfg,token,start,end,fields):
    heading('어떤 쿠폰이 많이 사용되고, 구매까지 이어졌을까요?',
            '쿠폰 발급 수가 아니라 주문완료 화면에서 쿠폰이 기록된 거래 수와 구매 수익을 보여드려요.')
    mode=st.radio('구매 분석',['쿠폰 순위','구매 전환 흐름','고객 등급별 구매'],horizontal=True,
                  label_visibility='collapsed',key='v6_coupon_mode')
    if mode=='구매 전환 흐름':
        values=funnel(cfg,token,start,end,fields.get('page1'))
        if not values:
            st.warning('GA4 퍼널 API에서 동일 단계의 결과를 가져올 수 없어요. 임의의 비율로 전환율을 만들지 않습니다.')
            return
        a,b,c=st.columns(3)
        with a:st.metric('쇼핑라운지 진입 고객',f'{values[0]:,.0f}명')
        with b:st.metric('이후 주문완료 고객',f'{values[1]:,.0f}명')
        with c:st.metric('퍼널 완료 비율',f'{values[1]/values[0]*100:.1f}%' if values[0] else '조회 불가')
        st.caption('순서대로 진입→주문완료 조건을 만족한 고객 기준의 **닫힌 퍼널**입니다. 동일 기간의 단순 방문자 수 대비 구매 건수 비율과 다릅니다.')
        return
    if mode=='쿠폰 순위':
        key=fields.get('coupon')
        if not key:
            st.warning('「주문 쿠폰」 맞춤 측정기준의 실제 API 이름을 확인할 수 없습니다. 데이터 진단을 확인하세요.')
            return
        rank=st.segmented_control('정렬 기준',['사용 거래 수 TOP','구매 수익 TOP'],default='사용 거래 수 TOP',key='v6_coupon_rank')
        metric='transactions' if rank=='사용 거래 수 TOP' else GROSS_REVENUE
        df,raw=coupon_data(cfg,token,start,end,fields,sort_metric=metric)
        if df is None or df.empty:
            st.info('이 필터 조건의 쿠폰 거래 데이터가 없어요.');return
        df=df[~df[key].astype(str).map(empty_label)].sort_values(metric,ascending=False).reset_index(drop=True)
        if df.empty:st.info('쿠폰명으로 분류 가능한 구매 거래가 없어요.');return
        a,b=st.columns([1,1.3],gap='large')
        with a:bar(df,key,metric,h=350,n=10)
        with b:
            show=df.rename(columns={key:'쿠폰명','transactions':'쿠폰이 기록된 구매 거래 수',
                                    GROSS_REVENUE:'쿠폰 적용 구매 수익'})
            table(show.head(250),f'ga4_쿠폰_{rank}_{start}_{end}',height=340)
        sum_tx=float(df['transactions'].sum())
        if sum_tx>0:
            top3=float(df.head(3)['transactions'].sum())/sum_tx*100 if metric=='transactions' else float(df.nlargest(3,'transactions')['transactions'].sum())/sum_tx*100
            st.caption(f'쿠폰별 조회 거래 수 합계 중 상위 3개 쿠폰 비중: **{top3:.1f}%** (쿠폰별 중복 거래 가능)')
        st.caption('주문완료 화면의 purchase와 쿠폰명 기록이 있는 거래를 집계합니다. 할인 비용, 쿠폰 발급 수, 쿠폰 자체의 고유 사용 이벤트 수와는 다릅니다. 복수 쿠폰/반복 기록이 있으면 쿠폰별 합계가 중복될 수 있어요.')
        with st.expander('쿠폰 하나 선택 → 일별·시간대별 실적',expanded=False):
            names=df[key].astype(str).head(250).tolist()
            if names:
                selected=st.selectbox('쿠폰 선택',names,key='v6_coupon_selected')
                cond=f_and(order_filter(),f_eq(key,selected))
                detailed=st.segmented_control('쿠폰 상세 지표',['구매 거래 수','구매 수익'],default='구매 거래 수',key='v6_coupon_detail_metric')
                detail_metric='transactions' if detailed=='구매 거래 수' else GROSS_REVENUE
                date_hour_detail(cfg,token,start,end,detail_metric,cond,
                                 f'coupon_{hashlib.md5(selected.encode()).hexdigest()[:9]}_{detail_metric}')
        quality(raw)
        return
    # Membership-tier purchase: only when field present in current GA4 metadata.
    dim=fields.get('grade')
    if not dim:st.warning('등급 맞춤 측정기준을 확인할 수 없습니다.');return
    df,raw=safe_query(cfg,token,start,end,[dim],['transactions',GROSS_REVENUE],
                      order_filter(),limit=150,sort=GROSS_REVENUE)
    if df is None or df.empty:return
    df=df.rename(columns={dim:'멤버십 등급','transactions':'구매 거래 수',GROSS_REVENUE:'구매 수익'})
    a,b=st.columns([1,1.25],gap='large')
    with a:bar(df,'멤버십 등급','구매 수익',h=300,n=12)
    with b:table(df,f'ga4_등급별구매_{start}_{end}',height=290)
    st.caption('주문완료 화면과 연결된 고객 등급 기준이며 사용자 수의 단순 합계와 다른 값이에요.')
    quality(raw)


def diagnostic(cfg,token,meta,fields,start,end):
    with st.expander("🛠 데이터 진단 · 사용 가능한 GA4 필드 확인",expanded=False):
        st.write("이 표는 실제 Google 계정 인증 상태에서 해당 GA4 속성의 메타데이터를 읽어 구성합니다.")
        labels={"page1":"페이지_1Depth","gender":"고객 성별","age":"고객 연령대","grade":"고객 등급",
                "login":"로그인 유형","coupon":"주문 쿠폰","evt_page":"이벤트 페이지","evt_area":"이벤트 영역","evt_label":"이벤트 라벨"}
        stat=pd.DataFrame([{"항목":v,"실제 API 필드":fields.get(k) or "미발견","상태":"발견" if fields.get(k) else "추가 확인"}
                           for k,v in labels.items()])
        st.dataframe(stat,hide_index=True,use_container_width=True,height=320)
        st.caption("메타데이터에 필드가 있어도 데이터가 있는지는 별도 보고서 조회와 집계 조건 검증이 필요해요.")
        support=field_health(meta)
        st.dataframe(pd.DataFrame([{"분석 영역":k,"필드 상태":"메타데이터 지원" if v else "일부 필드 미지원"}
                                   for k,v in support.items()]),hide_index=True,use_container_width=True,height=215)
        st.caption("표준 필드는 API 스키마로, 맞춤 필드는 이 GA4 속성의 메타데이터로 확인합니다. "
                   "보고서 조회가 정상 완료됐는지는 각 메뉴의 결과·오류 안내로 확인해 주세요.")
        st.markdown("**공통 범위 주의**: 쇼핑라운지 화면 필터, 방문자 사용자 세그먼트, 주문완료 화면 필터는 각각 다른 집계입니다.")
        d,_=safe_query(cfg,token,start,end,["eventName"],["eventCount"],shop_filter(),limit=80,show_error=False)
        if d is not None:
            st.write("쇼핑라운지 화면에 기록된 이벤트")
            table(d.rename(columns={"eventName":"이벤트명","eventCount":"이벤트 수"}),"ga4_event_diagnostics",height=230)


def choose_dates():
    today=dt.datetime.now(dt.timezone(dt.timedelta(hours=9))).date()
    yesterday=today-dt.timedelta(days=1)
    with st.sidebar:
        st.markdown('**분석 기간**')
        preset=st.selectbox('조회 구간',['30일','7일','90일','어제','오늘','직접 선택'],
                           label_visibility='collapsed',key='v6_preset')
        if preset=='직접 선택':
            interval=st.date_input('시작일 ~ 종료일',(yesterday-dt.timedelta(days=29),yesterday),
                                   max_value=today,format='YYYY/MM/DD',key='v6_dates')
            if not isinstance(interval,(tuple,list)) or len(interval)!=2:
                st.info('시작일과 종료일을 모두 지정해 주세요.');st.stop()
            start,end=interval
        elif preset=='오늘':start=end=today
        elif preset=='어제':start=end=yesterday
        else:
            length={'7일':7,'30일':30,'90일':90}[preset]
            end=yesterday;start=end-dt.timedelta(days=length-1)
        if start>end:start,end=end,start
        st.caption(f'{start:%Y.%m.%d} ~ {end:%Y.%m.%d}')
        do_compare=st.checkbox('비교 기간 표시',value=True,key='ga4_compare_v6')
        if do_compare:
            basis=st.selectbox('비교 기준',['직전 동일 기간','전월 같은 날짜','비교 기간 직접 지정'],key='ga4_compare_basis_v6')
            if basis=='전월 같은 날짜':
                from dateutil.relativedelta import relativedelta
                prior_start=start-relativedelta(months=1)
                prior_end=end-relativedelta(months=1)
            elif basis=='비교 기간 직접 지정':
                previous=previous_period(start,end)
                choice=st.date_input('비교 기간 선택',previous,max_value=today,
                                     format='YYYY/MM/DD',key='ga4_compare_custom_v6')
                if isinstance(choice,(tuple,list)) and len(choice)==2:prior_start,prior_end=choice
                else:prior_start,prior_end=previous
            else:prior_start,prior_end=previous_period(start,end)
            if prior_start>prior_end:prior_start,prior_end=prior_end,prior_start
            st.session_state['ga4_compare_dates_v6']=(prior_start,prior_end)
            if (prior_end-prior_start).days != (end-start).days:
                st.caption('비교 기간 일수가 달라 증감률에 기간 길이 차이가 반영됩니다.')
        else:st.session_state.pop('ga4_compare_dates_v6',None)
        if end==today:st.caption('오늘 데이터는 GA4 처리 중이어서 일부 늦게 반영될 수 있어요.')
    return start,end


def main():
    cfg=config()
    if cfg is None:
        st.error("Streamlit Secrets의 [google_oauth] 와 [ga4] 설정을 확인해 주세요.")
        st.code('[google_oauth]\nclient_id="..."\nclient_secret="..."\nredirect_uri="https://shoppinglounge-ga4.streamlit.app/"\nstate_signing_secret="..."\n\n[ga4]\nproperty_id="544313553"',language="toml")
        st.stop()
    if not cfg["redirect_uri"].startswith(("https://","http://localhost","http://127.0.0.1")):
        st.error("Google OAuth redirect_uri는 HTTPS 또는 localhost여야 합니다.");st.stop()
    oauth_callback(cfg)
    token=active_token()
    if not token:
        st.markdown('<div class="header-title">📊 쇼핑라운지 GA4 인사이트</div>',unsafe_allow_html=True)
        st.caption("Google Analytics Data API · 읽기 전용")
        st.link_button("Google 계정 연결",login_url(cfg),type="primary")
        st.stop()
    with st.sidebar:
        st.markdown("### 📊 쇼핑라운지 운영 분석")
        selected=st.radio("메뉴",MENUS,label_visibility="collapsed",key="menu_v6")
        st.caption(NAV_HELP[MENUS.index(selected)])
        st.divider()
    start,end=choose_dates()
    with st.sidebar:
        if st.button("↻ 데이터 다시 조회",use_container_width=True):
            st.session_state.pop("ga4_v6_cache",None);st.rerun()
        if st.button("로그아웃",use_container_width=True):
            logout();st.rerun()
        st.caption(f"GA4 속성 {cfg['property']} · 읽기 전용")
    st.markdown('<div class="header-title">쇼핑라운지 운영 인사이트</div>',unsafe_allow_html=True)
    scope = ("상품 수익은 KT멤버십 전체 속성 기준" if selected == MENUS[5]
             else "쇼핑라운지 관련 이벤트 · 메뉴별 집계 기준 적용")
    st.markdown(f'<div class="header-note"><span class="scope-label">{html.escape(selected)}</span> &nbsp;'
                f'{start:%Y.%m.%d} — {end:%Y.%m.%d} &nbsp; | &nbsp; {html.escape(scope)}</div>',
                unsafe_allow_html=True)
    try:meta=metadata(cfg,token)
    except GAError as exc:st.error(f"GA4 메타데이터 연결 실패: {exc}");st.stop()
    fields=info_fields(meta)
    pages=[overview,customers,channels,pages_view, promotions, products,purchase_actions]
    handler=pages[MENUS.index(selected)]
    with st.spinner("GA4에서 필요한 데이터만 조회하는 중..."):
        handler(cfg,token,start,end,fields) if handler not in [pages_view,promotions,products] else handler(cfg,token,start,end)
    diagnostic(cfg,token,meta,fields,start,end)
    st.caption("Google Analytics Data API · 읽기 전용 · 지표별 집계 조건은 각 화면의 안내를 확인하세요. 데이터 처리 시점에 따라 GA4 탐색 결과와 차이가 날 수 있습니다.")

# alias avoids shadowing `pages` local while keeping readable handler map
pages_view=pages
if __name__=="__main__":main()
