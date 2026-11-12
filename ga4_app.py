# -*- coding: utf-8 -*-
"""쇼핑라운지 GA4 인사이트 v3 — 실제 수집값 검증 기반, 7메뉴 compact dashboard.

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
NAV_HELP = ["UV/PV·구매·일별/시간대별", "기기·성별·연령·등급·로그인", "유입 채널·앱 내 쇼핑 클릭", "인기 페이지·이용 추이", "기획전 순위·상세", "상품 ID 매칭·수익 참고", "쿠폰·구매 퍼널"]
BLUE, TEAL, ORANGE, LIGHT = "#4765D2", "#169C99", "#F6A34F", "#ECF1FB"

st.markdown("""<style>
.block-container{padding:0.75rem 1.4rem 1rem!important;max-width:1700px!important}
section[data-testid=stSidebar]{background:#F7F9FD}
section[data-testid=stSidebar] .block-container{padding-top:1rem!important}
[data-testid=stVerticalBlock]{gap:.47rem!important}
[data-testid=stHorizontalBlock]{gap:.55rem!important}
.stTabs [data-baseweb=tab-list]{gap:4px}
.stTabs [data-baseweb=tab]{height:37px;padding:6px 12px}
div[data-testid=stPlotlyChart]{margin-top:-9px}
div[data-testid=stDataFrame]{margin-top:-8px}
#MainMenu,footer{visibility:hidden}
.header-title{font-size:1.55rem;letter-spacing:-.04em;font-weight:800;color:#24324B;line-height:1.2;margin:0}
.header-note{font-size:.81rem;color:#718199;margin:.32rem 0 .6rem}
.section-title{font-size:1.02rem;font-weight:750;color:#263C58;letter-spacing:-.02em}
.muted-note{font-size:.77rem;color:#72829A}
.kpi{border:1px solid #E6EBF5;border-radius:12px;background:white;min-height:79px;padding:12px 13px;box-shadow:0 1px 3px #24324b0a}
.kpi-name{color:#76859B;font-size:.76rem;margin-bottom:7px;font-weight:600}
.kpi-value{color:#1E3150;font-size:clamp(1rem,1.55vw,1.62rem);font-weight:780;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.kpi-foot{font-size:.68rem;color:#8B98AB;margin-top:4px}
.scope-label{display:inline-block;border-radius:16px;background:#EDF2FF;color:#4055B6;font-size:.68rem;padding:3px 9px;font-weight:750}
@media(max-width:900px){.block-container{padding:0.6rem!important}.kpi{padding:9px;min-height:70px}}
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
    for k in ["ga4_access_token","ga4_token_expires_at","ga4_expiry","ga4_v3_cache","ga4_report_cache"]:
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
        st.session_state.pop("ga4_v3_cache",None)
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
    cache=st.session_state.setdefault("ga4_v3_cache",OrderedDict())
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
    ds=[x["name"] for x in raw.get("dimensionHeaders",[])]
    ms=[x["name"] for x in raw.get("metricHeaders",[])]
    rows=[]
    for x in raw.get("rows",[]):
        vals=[v.get("value","") for v in x.get("dimensionValues",[])]+[v.get("value","") for v in x.get("metricValues",[])]
        rows.append(dict(zip(ds+ms,vals)))
    df=pd.DataFrame(rows,columns=ds+ms)
    for m in ms: df[m]=pd.to_numeric(df[m],errors="coerce").fillna(0)
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
    if x is None or pd.isna(x):return "—"
    x=float(x)
    return f"{x/100000000:,.2f}억" if abs(x)>=100000000 else (f"{x/10000:,.0f}만" if abs(x)>=10000 else f"{x:,.0f}")

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


def kpi_grid(items):
    first=st.columns(5,gap="small")
    for col,item in zip(first,items[:5]):
        with col: kpi(*item)
    second=st.columns(4,gap="small")
    for col,item in zip(second,items[5:]):
        with col:kpi(*item)


def heading(title,desc=""):
    st.markdown(f'<div class="section-title">{html.escape(title)}</div>',unsafe_allow_html=True)
    if desc:st.caption(desc)


def table(df,key,height=320):
    if df is None or df.empty:
        st.info("선택한 기간에 조회된 데이터가 없어요.")
        return
    st.dataframe(df,hide_index=True,use_container_width=True,height=height)
    st.download_button("↓ CSV",df.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"{key}.csv",mime="text/csv",key=f"dl_{key}")


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
    if df is None or df.empty:st.info("날짜별 조회 결과가 없어요.");return None
    labels=df["date"]
    if period=="일별": labels=pd.to_datetime(df["date"],format="%Y%m%d",errors="coerce").dt.strftime("%m/%d")
    elif period=="주별": labels=df["yearWeek"].astype(str)+"주"
    else: labels=df["yearMonth"].astype(str).str.slice(0,4)+"."+df["yearMonth"].astype(str).str.slice(4)
    fig=go.Figure(go.Scatter(x=labels,y=df[metric],mode="lines+markers+text",line=dict(color=BLUE,width=2.6),
                              marker=dict(size=7),text=[f"{x:,.0f}" for x in df[metric]],textposition="top center",
                              textfont=dict(size=9,color="#7386a2"),customdata=df["date"] if "date" in df else None,
                              hovertemplate="%{x}<br>%{y:,.0f}<extra></extra>"))
    fig.update_layout(height=290,margin=dict(l=4,r=12,t=22,b=15),clickmode="event+select",
                      xaxis=dict(showgrid=False,tickfont=dict(size=10)),yaxis=dict(showgrid=True,gridcolor="#EDF1F7"),
                      plot_bgcolor="white",paper_bgcolor="white")
    try:
        event=st.plotly_chart(fig,key=widget_key,use_container_width=True,on_select="rerun",selection_mode="points",config={"displayModeBar":False})
        points=event.selection.points if event and hasattr(event,"selection") else []
        if points and period=="일별":
            x=points[0].get("customdata")
            if isinstance(x,(list,tuple)):x=x[0]
            if x:return str(x)
    except TypeError:
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
    # GA4 Data API v1alpha (preview). If unavailable, no fabricated conversion rate.
    if field:
        first={"andGroup":{"expressions":[
            {"funnelEventFilter":{"eventName":"screen_view"}},
            {"funnelFieldFilter":{"fieldName":field,"stringFilter":{"matchType":"EXACT","value":LOUNGE_TITLE}}}
        ]}}
    else:
        first={"andGroup":{"expressions":[
            {"funnelEventFilter":{"eventName":"screen_view"}},
            {"funnelFieldFilter":{"fieldName":"unifiedScreenName","stringFilter":{"matchType":"CONTAINS","value":LOUNGE_TITLE}}}
        ]}}
    body={"dateRanges":[{"startDate":str(start),"endDate":str(end)}],
          "funnel":{"isOpenFunnel":False,"steps":[
              {"name":"쇼핑라운지 진입","filterExpression":first},
              {"name":"구매 완료","filterExpression":{"funnelEventFilter":{"eventName":"purchase"}}}
          ]},"limit":"10"}
    try:
        data=api(cfg,token,":runFunnelReport",body,alpha=True)
        df=frame(data.get("funnelTable",{}))
        if df.empty:return None
        # v1alpha response contains funnelStepName + activeUsers, and often funnelStep/ completion rate columns
        active=[x for x in df.columns if x=="activeUsers"]
        if not active:return None
        values=[float(v) for v in df["activeUsers"].tolist()]
        if len(values)<2:return None
        return values[0],values[1]
    except GAError:return None


def overview(cfg,token,start,end,fields):
    st.markdown('<div class="section-title">핵심 지표</div>',unsafe_allow_html=True)
    use=lounge_summary(cfg,token,start,end)
    purch=purchase_summary(cfg,token,start,end)
    ret=returning_summary(cfg,token,start,end)
    active=use.get("activeUsers")
    f=funnel(cfg,token,start,end,fields.get("page1"))
    conversion=(f[1]/f[0]) if f and f[0] else None
    # Note: newUsers==0 in user's observed shop-segment exploration; don't claim no first time users.
    first=use.get("newUsers")
    if first==0 and active and active>0:first=None
    avg=(use.get("userEngagementDuration") or 0)/active if active and use.get("userEngagementDuration") is not None else None
    kpi_grid([
        ("활성 사용자",active,"쇼핑라운지 화면 기록 기준"),
        ("세션 수",use.get("sessions"),"화면 조건 연관 세션·확인 필요"),
        ("페이지 조회수",use.get("screenPageViews"),"쇼핑라운지 화면 기준"),
        ("구매 금액",purch.get(GROSS_REVENUE),"주문완료 화면·총 구매 수익","krw"),
        ("신규 사용자",first,"GA4 first_open·수집 상태 확인"),
        ("평균 참여 시간",avg,"활성 사용자당 참여 시간","time"),
        ("구매 건수",purch.get("transactions"),"주문완료 화면의 거래 수"),
        ("구매 전환율",conversion,"진입→구매·퍼널 API 사용 가능 시","rate"),
        ("재방문 사용자",ret,"KT멤버십 기준 재방문 성격"),
    ])
    st.caption("집계 범위: 이용=쇼핑라운지 화면 기록, 구매=주문완료 화면의 purchase(총 구매 수익). "
               "신규·재방문·세션 지표는 GA4 탐색의 사용자 세그먼트와 다를 수 있고, "
               "구매 전환율은 동일한 퍼널 API 결과가 확인될 때만 표시합니다.")
    c1,c2=st.columns([1,3],vertical_alignment="bottom")
    with c1: gran=st.segmented_control("추이 기준",["일별","주별","월별"],default="일별",key="overview_period")
    with c2: metric=st.selectbox("그래프 지표",["활성 사용자","세션 수","페이지 조회수","구매 건수","구매 금액"],key="overview_metric")
    map_m={"활성 사용자":"activeUsers","세션 수":"sessions","페이지 조회수":"screenPageViews",
           "구매 건수":"transactions","구매 금액":GROSS_REVENUE}
    extra=f_eq("eventName","purchase") if metric in ("구매 건수","구매 금액") else None
    if extra:
        # For purchase charts use purchase screen ONLY, not all lounge pages.
        purchase_drill(cfg,token,start,end,gran,map_m[metric])
    else:
        day_drill(cfg,token,start,end,"overview",gran,map_m[metric])
    if f:
        st.caption(f"구매 퍼널(참고): 진입 {f[0]:,.0f}명 → 구매 완료 {f[1]:,.0f}명. 탐색과의 정확한 일치는 단계 조건·세그먼트 비교가 필요해요.")
    else:
        st.caption("구매 전환율: GA4 퍼널 API에서 동일 조건을 조회할 수 없으면 표시하지 않습니다.")


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
    heading("고객 특성","쇼핑라운지 화면에서 확인된 사용자 특성입니다. 동일 고객이 여러 분류에 포함될 수 있어 합산 비중을 단정하지 않아요.")
    selected=st.radio("분석 항목",["기기","성별","연령","멤버십 등급","로그인 유형"],horizontal=True,label_visibility="collapsed")
    mapping={"기기":("operatingSystem","운영체제"),"성별":(fields.get("gender") or "userGender","성별"),
             "연령":(fields.get("age") or "userAgeBracket","연령대"),
             "멤버십 등급":(fields.get("grade"),"등급"),"로그인 유형":(fields.get("login"),"로그인 유형")}
    dim,name=mapping[selected]
    if not dim:
        st.warning(f"'{selected}'에 해당하는 GA4 맞춤 측정기준을 메타데이터에서 찾지 못했어요. 아래 '데이터 진단'에서 확인해 주세요.")
        return
    df,raw=safe_query(cfg,token,start,end,[dim],["activeUsers"],shop_filter(),limit=200,sort="activeUsers")
    if df is None:return
    quality(raw)
    if df.empty:st.info("선택한 기간에 집계된 사용자 데이터가 없어요.");return
    if selected=="기기":df[dim]=df[dim].replace({"iOS":"iOS (아이폰)","Android":"Android"})
    df=df.rename(columns={dim:name,"activeUsers":"활성 사용자"})
    l,r=st.columns([1,1.35])
    with l:bar(df,name,"활성 사용자",h=300,n=13)
    with r:table(df,f"ga4_customer_{selected}_{start}_{end}",height=300)
    if selected=="성별":st.caption("성별 × 연령 교차 분석은 요청에 따라 제외했습니다.")
    if selected=="연령":st.caption("고객_연령대 맞춤 정의가 없으면 Google 표준 연령 구간을 대신 표시합니다. (not set)은 별도 표시됩니다.")


def channels(cfg,token,start,end,fields):
    heading("유입 채널·KT멤버십 앱 내 쇼핑 클릭","Direct를 KT앱 내부 유입이라고 임의 분류하지 않아요. MMS/RCS는 확실한 식별값이 발견될 때까지 통합·미분류합니다.")
    mode=st.radio("유입 분석",["GA4 세션 채널","KT멤버십 앱 내 쇼핑 클릭"],horizontal=True,label_visibility="collapsed")
    if mode=="GA4 세션 채널":
        df,raw=safe_query(cfg,token,start,end,["sessionDefaultChannelGroup","sessionSourceMedium"],
                          ["screenPageViews","activeUsers"],shop_filter(),limit=1000,sort="screenPageViews")
        if df is None:return
        if df.empty:st.info("해당 조건에서 유입 데이터가 조회되지 않았어요.");return
        grouped=df.groupby("sessionDefaultChannelGroup",as_index=False)["screenPageViews"].sum().sort_values("screenPageViews",ascending=False)
        a,b=st.columns([1,1.25])
        with a:bar(grouped,"sessionDefaultChannelGroup","screenPageViews",h=300)
        with b:table(df.rename(columns={"sessionDefaultChannelGroup":"채널","sessionSourceMedium":"소스/매체",
                                       "screenPageViews":"화면 조회수","activeUsers":"활성 사용자"}),
                     f"ga4_acquisition_{start}_{end}",height=300)
        st.caption("유입 채널은 쇼핑라운지 화면 기록과 연결된 세션의 출처입니다. 사용자 수를 채널별로 합산하지 마세요.")
        quality(raw)
        return
    if not fields.get("evt_label"):
        st.warning("이벤트_라벨 맞춤 측정기준을 찾지 못했습니다. 메타데이터에서 등록 여부를 확인해 주세요.");return
    additional=f_and(f_eq("eventName","click_event"),f_eq("sessionDefaultChannelGroup","Direct"),
                     f_regex(fields["evt_label"],".*(쇼핑|shopping).*") )
    dims=[f for f in [fields.get("evt_page"),fields.get("evt_area"),fields.get("evt_label")] if f]
    if not dims:st.warning("이벤트 페이지·영역·라벨이 조회되지 않아 분석할 수 없어요.");return
    df,raw=safe_query(cfg,token,start,end,dims,["eventCount","activeUsers"],additional,limit=1000,sort="eventCount")
    if df is None:return
    if df.empty:st.info("쇼핑 관련 클릭 이벤트가 검색되지 않았어요.");return
    pretty={fields.get("evt_page"):"이벤트 페이지",fields.get("evt_area"):"이벤트 영역",fields.get("evt_label"):"이벤트 라벨"}
    left,right=st.columns([1,1.35])
    with left:
        if fields.get("evt_area") and fields["evt_area"] in df:
            d=df.groupby(fields["evt_area"],as_index=False)["eventCount"].sum().sort_values("eventCount",ascending=False)
            bar(d,fields["evt_area"],"eventCount",h=300,n=12)
        else:bar(df,dims[-1],"eventCount",h=300)
    with right:table(df.rename(columns={**pretty,"eventCount":"클릭 이벤트 수","activeUsers":"활성 사용자"}),f"ga4_inapp_{start}_{end}",height=300)
    st.caption("KT앱 내 쇼핑 관련 클릭이며, 실제 쇼핑라운지 도착 또는 구매 전환 건수가 아닙니다.")


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
    heading("페이지 분석","인기 페이지는 조회수 중심으로 정렬하고, 기획전·상품 상세는 각각 별도 메뉴에서 분석해요.")
    df,raw=page_data(cfg,token,start,end)
    if df is None or df.empty:return
    group=df.groupby("페이지 유형",as_index=False)["screenPageViews"].sum().sort_values("screenPageViews",ascending=False)
    left,right=st.columns([.8,1.35])
    with left:bar(group,"페이지 유형","screenPageViews",h=285)
    with right:
        keyword=st.text_input("페이지 검색",placeholder="예: 메인, 쿠폰, 장바구니",key="page_search")
        filtered=df[~df["페이지 유형"].isin(["기획전","상품 상세"])].copy()
        if keyword:filtered=filtered[filtered["unifiedScreenName"].str.contains(re.escape(keyword),case=False,na=False)]
        cols=filtered.rename(columns={"unifiedScreenName":"페이지","screenPageViews":"조회수","activeUsers":"활성 사용자"})
        table(cols[["페이지 유형","페이지","조회수","활성 사용자"]].head(100),f"ga4_pages_{start}_{end}",height=245)
    quality(raw)
    with st.expander("페이지 하나 선택해 일별·시간대별 조회",expanded=False):
        opts=filtered["unifiedScreenName"].head(300).tolist()
        if opts:
            chosen=st.selectbox("페이지 선택",opts,key="page_detail")
            df2,_=safe_query(cfg,token,start,end,["date"],["screenPageViews"],f_eq("unifiedScreenName",chosen),limit=365)
            if df2 is not None and not df2.empty:
                dd=df2.sort_values("date");trend_plot(dd,"screenPageViews","일별",f"single_page_{start}_{end}")


def promotion_id(title):
    s=str(title)
    found=re.search(r">기획전>(\d{4,})",s)
    return found.group(1) if found else ""


def promotions(cfg,token,start,end):
    heading("기획전 성과","기획전별 조회수·이용자를 표시합니다. 기획전별 상품·매출은 프로모션 ID 귀속 태깅 검증 전까지 표시하지 않아요.")
    pages_,raw=page_data(cfg,token,start,end)
    if pages_ is None:return
    p=pages_[pages_["페이지 유형"]=="기획전"].copy()
    if p.empty:st.info("해당 기간 기획전 페이지가 검색되지 않았어요.");return
    p["기획전 ID"]=p["unifiedScreenName"].apply(promotion_id)
    p=p.sort_values("screenPageViews",ascending=False)
    a,b=st.columns([1,1.2])
    with a:bar(p,"기획전 ID","screenPageViews",h=305,n=12)
    with b:
        show=p.rename(columns={"unifiedScreenName":"기획전 화면명","screenPageViews":"조회수","activeUsers":"활성 사용자"})
        table(show[["기획전 ID","기획전 화면명","조회수","활성 사용자"]].head(100),f"ga4_promo_{start}_{end}",height=305)
    with st.expander("기획전 클릭 → 일별·시간대별 상세",expanded=False):
        ids=p["unifiedScreenName"].head(150).tolist()
        chosen=st.selectbox("기획전 선택",ids,key="promo_selected")
        drill_filter=f_eq("unifiedScreenName",chosen)
        d,_=safe_query(cfg,token,start,end,["date"],["screenPageViews"],drill_filter,limit=365)
        if d is not None and not d.empty:
            d=d.sort_values("date").reset_index(drop=True)
            clicked=trend_plot(d,"screenPageViews","일별",f"promotion_day_{start}_{end}")
            dates=pd.to_datetime(d["date"],format="%Y%m%d").dt.strftime("%Y-%m-%d").tolist()
            if clicked:
                v=pd.to_datetime(clicked,format="%Y%m%d",errors="coerce")
                if pd.notna(v):st.session_state["promo_selected_date"]=v.strftime("%Y-%m-%d")
            if st.session_state.get("promo_selected_date") not in dates:st.session_state["promo_selected_date"]=dates[-1]
            selected=st.selectbox("시간별 상세 날짜",dates,key="promo_selected_date")
            h,_=safe_query(cfg,token,selected,selected,["hour"],["screenPageViews"],drill_filter,limit=24)
            if h is not None:
                base=pd.DataFrame({"hour":[f"{i:02d}" for i in range(24)]})
                h["hour"]=h["hour"].astype(str).str.zfill(2)
                base=base.merge(h,on="hour",how="left").fillna(0)
                fig=go.Figure(go.Bar(x=base["hour"]+"시",y=base["screenPageViews"],marker_color=TEAL))
                fig.update_layout(height=190,margin=dict(l=4,r=4,t=4,b=4))
                st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
    quality(raw)


def shop_item_ids(cfg,token,start,end):
    # IDs observed in lounge item detail screens. This is not a user/session revenue attribution.
    df,raw=page_data(cfg,token,start,end)
    if df is None or df.empty:return set(),raw
    p=df[df["페이지 유형"]=="상품 상세"]
    codes=p["unifiedScreenName"].str.extract(r">상품상세>(\d{6,})",expand=False).dropna()
    return set(codes.astype(str)),raw


def products(cfg,token,start,end):
    heading("상품·브랜드 분석","항목 ID와 쇼핑라운지 상품 상세 페이지 ID를 대조합니다. 매출은 전체 KT멤버십 상품 이벤트 기준이므로 쇼핑라운지 기여 매출이 아닙니다.")
    selected=st.radio("상품 분석 방식",["쇼핑라운지 상품 ID 매칭","브랜드 참고"],horizontal=True,label_visibility="collapsed")
    codes,source=shop_item_ids(cfg,token,start,end)
    if selected=="쇼핑라운지 상품 ID 매칭":
        st.caption(f"쇼핑라운지 상품 상세 화면에서 발견된 상품 ID {len(codes):,}개 · 상품 수익은 해당 ID의 전체 속성 수익")
        if not codes:st.warning("이 기간에 상품 상세 페이지 ID를 읽지 못했어요. 화면 이름 규칙을 먼저 확인해야 합니다.");return
    dims=["itemId","itemName","itemBrand"] if selected=="쇼핑라운지 상품 ID 매칭" else ["itemBrand"]
    mets=["itemsViewed","itemsCheckedOut","itemsPurchased","itemRevenue"]
    df,raw=safe_query(cfg,token,start,end,dims,mets,limit=10000,sort="itemRevenue")
    if df is None:return
    if df.empty:st.info("상품 데이터가 조회되지 않았어요.");return
    if selected=="쇼핑라운지 상품 ID 매칭":
        df=df[df["itemId"].astype(str).isin(codes)].copy()
        if df.empty:
            st.warning("가져온 상품 수익 상위 10,000행에서 쇼핑라운지 상품 ID를 찾지 못했어요. 매출 0원이 아니라 조회 범위 또는 ID 연결을 재검증해야 합니다.")
            quality(raw);return
        search=st.text_input("상품명·상품 ID 검색",placeholder="예: 종근당, 132237997")
        if search:df=df[df["itemName"].astype(str).str.contains(re.escape(search),case=False,na=False)|df["itemId"].astype(str).str.contains(re.escape(search),case=False,na=False)]
    df=df.sort_values("itemRevenue",ascending=False)
    display=(df.rename(columns={"itemId":"상품 ID","itemName":"상품명","itemBrand":"브랜드",
             "itemsViewed":"조회 수량","itemsCheckedOut":"결제 진행 수량","itemsPurchased":"구매 수량","itemRevenue":"상품 수익"}))
    a,b=st.columns([.85,1.4])
    with a:
        name="itemName" if "itemName" in df else "itemBrand"
        bar(df,name,"itemRevenue",h=310,n=9)
    with b:table(display.head(250),f"ga4_items_{start}_{end}",height=310)
    st.caption("상품 조회·구매는 수량(개)이며 거래/사용자 수와 다릅니다. 상품 수익은 해당 ID의 다른 경로 판매가 포함될 수 있어요.")
    quality(raw)
    if selected=="쇼핑라운지 상품 ID 매칭":product_detail(cfg,token,start,end,df)


def product_detail(cfg,token,start,end,df):
    if df.empty or "itemId" not in df:
        return
    with st.expander("🔎 상품 선택 → 일별·시간대별 상세 (전체 속성 수익)",expanded=False):
        names={str(row["itemId"]):str(row["itemName"]) for _,row in df.head(500).iterrows()}
        ids=list(names)
        if not ids:
            st.info("상품 ID를 가진 항목이 없어요.")
            return
        selected=st.selectbox("상품 선택",ids,format_func=lambda k:f"{names[k][:65]}  ·  {k}",key="detail_item_id")
        metric_title=st.segmented_control("상세 지표",["조회 수량","결제 진행 수량","구매 수량","상품 수익"],default="상품 수익")
        met={"조회 수량":"itemsViewed","결제 진행 수량":"itemsCheckedOut","구매 수량":"itemsPurchased","상품 수익":"itemRevenue"}[metric_title]
        condition=f_eq("itemId",selected)
        d,raw=safe_query(cfg,token,start,end,["date"],[met],condition,limit=700)
        if d is None or d.empty:
            st.info("해당 상품의 날짜별 이벤트가 조회되지 않아요.")
            return
        d=d.sort_values("date").reset_index(drop=True)
        selected_date=trend_plot(d,met,"일별",f"item_trend_{selected}_{met}_{start}_{end}")
        dates=pd.to_datetime(d["date"],format="%Y%m%d",errors="coerce").dropna().dt.strftime("%Y-%m-%d").tolist()
        if not dates:return
        key=f"item_detail_day_{start}_{end}"
        if key not in st.session_state or st.session_state[key] not in dates:
            st.session_state[key]=dates[-1]
        if selected_date:
            day=pd.to_datetime(selected_date,format="%Y%m%d",errors="coerce")
            if pd.notna(day) and day.strftime("%Y-%m-%d") in dates:
                st.session_state[key]=day.strftime("%Y-%m-%d")
        choice=st.selectbox("시간대별 상세 날짜",dates,key=key)
        h,hr=safe_query(cfg,token,choice,choice,["hour"],[met],condition,limit=24)
        if h is None:return
        base=pd.DataFrame({"hour":[f"{i:02d}" for i in range(24)]})
        if h.empty:base[met]=0
        else:
            h["hour"]=h["hour"].astype(str).str.zfill(2)
            base=base.merge(h[["hour",met]],on="hour",how="left").fillna(0)
        fig=go.Figure(go.Bar(x=base["hour"]+"시",y=base[met],marker_color=TEAL))
        fig.update_layout(height=190,margin=dict(l=4,r=4,t=6,b=4),paper_bgcolor="white",plot_bgcolor="white")
        st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
        st.caption("항목 ID 일치 기준의 KT멤버십 속성 전체 이벤트입니다. "
                   "쇼핑라운지 내부 구매/기여 매출로 귀속한 값이 아니에요.")
        quality(raw)


def purchase_actions(cfg,token,start,end,fields):
    heading("쿠폰·구매 행동","실제 주문완료 화면의 purchase 이벤트를 기준으로 쿠폰 수익을 집계합니다.")
    mode=st.radio("쿠폰·행동",["쿠폰별 구매","구매 전환 퍼널","멤버십 등급별 구매"],horizontal=True,label_visibility="collapsed")
    if mode=="구매 전환 퍼널":
        values=funnel(cfg,token,start,end,fields.get("page1"))
        if not values:
            st.warning("GA4 퍼널 API에서 이 조건의 결과를 가져오지 못했어요. 구매 건수/상품 조회 수량을 임의로 나눠 전환율을 만들지 않습니다.")
            return
        a,b,c=st.columns(3)
        with a:kpi("쇼핑라운지 진입",values[0],"screen_view 조건")
        with b:kpi("진입 후 구매 완료",values[1],"purchase 이벤트 사용자")
        with c:kpi("퍼널 완료율",values[1]/values[0] if values[0] else None,"닫힌 퍼널","rate")
        st.caption("퍼널 API(v1alpha) 기준. GA4 탐색의 단계 순서·세그먼트·기간에 따라 결과가 다를 수 있습니다.")
        return
    dim=fields.get("coupon") if mode=="쿠폰별 구매" else fields.get("grade")
    if not dim:
        st.warning("이 분석에 필요한 맞춤 측정기준이 GA4 메타데이터에 등록되어 있지 않거나 검색명과 일치하지 않아요.")
        return
    condition=order_filter()
    if mode=="쿠폰별 구매":condition=f_and(condition,f_not(f_eq(dim,"(not set)")))
    df,raw=safe_query(cfg,token,start,end,[dim],["transactions","purchaseRevenue"],condition,limit=500,sort="purchaseRevenue")
    if df is None:return
    if df.empty:st.info("현재 필터에서 구매 데이터가 없어요.");return
    show=df.rename(columns={dim:"쿠폰" if mode=="쿠폰별 구매" else "등급",
                            "transactions":"거래 수","purchaseRevenue":"구매 수익"})
    left,right=st.columns([.9,1.25])
    with left:bar(df,dim,"purchaseRevenue",h=300,n=10)
    with right:table(show,f"ga4_coupons_{mode}_{start}_{end}",height=300)
    st.caption("쿠폰 할인비용이 아닌 쿠폰이 기록된 구매 수익입니다. 쿠폰 중복 적용·환불·집계 시점에 유의하세요.")
    quality(raw)
    if mode=="쿠폰별 구매":
        with st.expander("🔎 쿠폰 선택 → 일별·시간대별 매출",expanded=False):
            coupon_names=df[dim].astype(str).tolist()
            name=st.selectbox("쿠폰명",coupon_names,key="coupon_selected")
            d_filter=f_and(condition,f_eq(dim,name))
            daily,dr=safe_query(cfg,token,start,end,["date"],["purchaseRevenue"],d_filter,limit=400)
            if daily is not None and not daily.empty:
                daily=daily.sort_values("date").reset_index(drop=True)
                picked=trend_plot(daily,"purchaseRevenue","일별",f"coupon_trend_{start}_{end}_{name}")
                dates=pd.to_datetime(daily["date"],format="%Y%m%d",errors="coerce").dropna().dt.strftime("%Y-%m-%d").tolist()
                if dates:
                    key=f"coupon_day_{start}_{end}"
                    if key not in st.session_state or st.session_state[key] not in dates:st.session_state[key]=dates[-1]
                    if picked:
                        parsed=pd.to_datetime(picked,format="%Y%m%d",errors="coerce")
                        if pd.notna(parsed) and parsed.strftime("%Y-%m-%d") in dates:
                            st.session_state[key]=parsed.strftime("%Y-%m-%d")
                    day=st.selectbox("시간대별 상세 날짜",dates,key=key)
                    hourly,hr=safe_query(cfg,token,day,day,["hour"],["purchaseRevenue"],d_filter,limit=24)
                    if hourly is not None:
                        base=pd.DataFrame({"hour":[f"{i:02d}" for i in range(24)]})
                        if hourly.empty:base["purchaseRevenue"]=0
                        else:
                            hourly["hour"]=hourly["hour"].astype(str).str.zfill(2)
                            base=base.merge(hourly[["hour","purchaseRevenue"]],on="hour",how="left").fillna(0)
                        fig=go.Figure(go.Bar(x=base["hour"]+"시",y=base["purchaseRevenue"],marker_color=TEAL))
                        fig.update_layout(height=190,margin=dict(l=4,r=4,t=6,b=4),paper_bgcolor="white",plot_bgcolor="white")
                        st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})
            else:st.info("이 쿠폰의 날짜별 실적을 조회할 수 없어요.")


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
    # Date presets use the business timezone even when Streamlit Cloud runs in UTC.
    today=dt.datetime.now(dt.timezone(dt.timedelta(hours=9))).date()
    end=today-dt.timedelta(days=1)
    default_start=end-dt.timedelta(days=29)
    with st.sidebar:
        st.markdown("**분석 기간**")
        preset=st.segmented_control("기간",["7일","30일","90일","직접"],default="30일",label_visibility="collapsed")
        if preset=="직접":
            d=st.date_input("기간 선택",(default_start,end),max_value=today,format="YYYY/MM/DD",key="v3_date")
            if not isinstance(d,(tuple,list)) or len(d)!=2:st.info("시작일과 종료일을 모두 지정해 주세요.");st.stop()
            start,last=d
        else:
            n={"7일":7,"30일":30,"90일":90}[preset]
            start,last=end-dt.timedelta(days=n-1),end
        if start>last:start,last=last,start
        st.caption(f"{start:%Y.%m.%d} ~ {last:%Y.%m.%d}")
    return start,last


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
        st.markdown("### 📊 쇼핑라운지")
        selected=st.radio("메뉴",MENUS,label_visibility="collapsed",key="menu_v3")
        st.caption(NAV_HELP[MENUS.index(selected)])
        st.divider()
    start,end=choose_dates()
    with st.sidebar:
        if st.button("↻ 데이터 다시 조회",use_container_width=True):
            st.session_state.pop("ga4_v3_cache",None);st.rerun()
        if st.button("로그아웃",use_container_width=True):
            logout();st.rerun()
        st.caption(f"GA4 속성 {cfg['property']} · 읽기 전용")
    st.markdown('<div class="header-title">쇼핑라운지 GA4 인사이트</div>',unsafe_allow_html=True)
    scope = ("KT멤버십 속성 · 상품 ID 매칭 기준 (수익은 속성 전체)" if selected == MENUS[5]
             else "KT멤버십 속성 · 분석별 필터 기준 적용")
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
    st.caption("GA4 실데이터 · 접근권한이 있는 계정에만 데이터가 표시됩니다. 외부 공유 전에 Google OAuth 클라이언트 비밀번호 및 조직 보안정책을 검토하세요.")

# alias avoids shadowing `pages` local while keeping readable handler map
pages_view=pages
if __name__=="__main__":main()
