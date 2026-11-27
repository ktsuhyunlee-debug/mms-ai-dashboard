# -*- coding: utf-8 -*-
"""쇼핑라운지 GA4 운영 인사이트 v8.0 · 상품/기획전/시간대 집중 분석 — 실제 수집값 검증 기반, 7메뉴 compact dashboard.

Streamlit Secrets:
[service_account] json = '''서비스 계정 JSON 전체 내용'''
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
from collections import OrderedDict

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from google.oauth2 import service_account
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.auth.exceptions import GoogleAuthError

st.set_page_config(page_title="쇼핑라운지 | GA4 인사이트", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

API = "https://analyticsdata.googleapis.com/v1beta"
ALPHA = "https://analyticsdata.googleapis.com/v1alpha"
SCOPE = "https://www.googleapis.com/auth/analytics.readonly"
DEFAULT_PROPERTY = "544313553"
LOUNGE_TITLE = "KT멤버십>쇼핑라운지"
# GA4 탐색의 「총 구매 수익」과 같은 grossPurchaseRevenue를 우선 사용합니다.
GROSS_REVENUE = "grossPurchaseRevenue"
PURCHASE_TITLE = "KT멤버십>쇼핑라운지>주문>주문완료"
MENUS = ["01 · 종합 현황", "02 · 이용 고객", "03 · 유입·시간대", "04 · 페이지", "05 · 기획전", "06 · 상품·브랜드", "07 · 쿠폰·구매 행동"]
NAV_HELP = ["9개 KPI · 주요 변화 · 인기 TOP", "성별·연령·기기·등급·로그인", "시간대별 방문·구매 · KT멤버십 쇼핑 관련 클릭", "자주 보는 화면 · 유형 · 상세", "기획전별 방문·고객 특성·구매 태깅 확인", "상품별 방문·구매 고객 성별/연령 비교", "쿠폰 사용 거래 · 수익 · 퍼널"]
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

# ------------------ Server-side service-account authentication ------------------
def config():
    """Read a server-side credential from Streamlit Secrets (never expose to clients)."""
    try:
        block = st.secrets.get("service_account", {})
        ga4 = st.secrets.get("ga4", {})
        raw = block.get("json")
        if isinstance(raw, str):
            info = json.loads(raw.strip())
        else:
            # Also supports a TOML service-account table, without a JSON string.
            info = dict(block)
        if not isinstance(info, dict):
            return None
        if (info.get("type") != "service_account" or
                not info.get("client_email") or
                not info.get("private_key") or
                not info.get("token_uri")):
            return None
        return {"property": str(ga4.get("property_id", DEFAULT_PROPERTY)),
                "title": str(ga4.get("page_keyword", LOUNGE_TITLE)),
                "service_account_info": info}
    except (KeyError, FileNotFoundError, TypeError, ValueError):
        return None


def active_token(cfg):
    """Generate/refresh a Google API access token server-side, per Streamlit session.

    No OAuth redirect, user sign-in, or token exposed to the browser.
    """
    info = cfg["service_account_info"]
    creds = st.session_state.get("ga4_service_credentials")
    if (creds is None or
            getattr(creds, "service_account_email", None) != info["client_email"]):
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=[SCOPE])
        st.session_state["ga4_service_credentials"] = creds
        # Old OAuth token/cache is not used by this release.
        st.session_state.pop("ga4_access_token", None)
        st.session_state.pop("ga4_token_expires_at", None)
    if not creds.valid:
        creds.refresh(GoogleAuthRequest())
    return creds.token

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


def custom_dimensions(meta):
    """Return every custom dimension advertised for this property.

    Don't rely only on customDefinition: older/variant metadata responses may
    carry an explicit custom apiName without that boolean.
    """
    return [d for d in meta.get('dimensions', []) if
            d.get('customDefinition') or str(d.get('apiName', '')).startswith(
                ('customUser:', 'customEvent:', 'customItem:'))]


def custom_field(meta, names, preferred_scope=None):
    """Find the *advertised* API field, never guess a GA4 property key.

    Search UI name and API name (including parameter suffix). The parameter
    sometimes has a usable label while uiName has been renamed in the UI.
    """
    dimensions=custom_dimensions(meta)
    lookup=[]
    for d in dimensions:
        api_name=str(d.get('apiName',''))
        labels=[str(d.get('uiName','')),api_name,
                api_name.partition(':')[2]]
        lookup.append((d, [canonical(v) for v in labels]))
    for exact in (True, False):
        for candidate in names:
            needle=canonical(candidate)
            if not needle: continue
            matches=[]
            for d,labels in lookup:
                if (any(needle == v for v in labels) if exact else
                    any(needle in v for v in labels if len(needle)>=3)):
                    matches.append(d)
            if matches:
                if preferred_scope:
                    matches.sort(key=lambda d: (
                        not str(d.get('apiName','')).startswith(preferred_scope+':'),
                        str(d.get('apiName',''))))
                return matches[0]['apiName']
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


# Actual names must be advertised by this property's metadata. The aliases
# only help identify supported fields; none is submitted to the API blindly.
CUSTOMER_FIELD_ALIASES={
    'gender': ['고객_성별','고객 성별','up_gender','up_sex','user_gender','user_sex',
               'customer_gender','customer_sex','gender','sex','성별'],
    'age': ['고객_연령대','고객 연령대','up_age','user_age','user_age_group',
            'customer_age','age_group','age','연령대','연령'],
    'grade': ['고객_등급','고객 등급','up_grade','up_membership_grade',
              'user_grade','membership_grade','membership_tier','member_grade',
              'membership_level','grade','tier','등급'],
    'login': ['고객_고객 로그인 유형','고객_로그인 유형','로그인 유형',
              'up_login_type','user_login_type','up_login','user_login',
              'login_type','login_method','auth_type','login','로그인'],
}


def customer_role_tags(meta, api_name):
    """Classify obvious names, to prevent age values being shown as gender."""
    item=next((d for d in custom_dimensions(meta)
               if d.get('apiName')==api_name), {})
    suffix=str(api_name).partition(':')[2]
    names=(suffix, str(item.get('uiName','')))
    joined=' '.join(names).casefold()
    parts=set(re.split(r'[^a-z0-9가-힣]+',joined))
    compact=canonical(suffix)
    tags=set()
    if ('연령' in joined or '나이' in joined or
        parts.intersection({'age','ages','agegroup','agebracket'}) or
        compact in {'upage','userage','customerage','custage','memberage'}):
        tags.add('age')
    if ('성별' in joined or parts.intersection({'gender','sex','gendergroup'}) or
        compact in {'upgender','upsex','usergender','usersex',
                    'customergender','customersex','custgender','custsex'}):
        tags.add('gender')
    if ('등급' in joined or parts.intersection({'grade','tier','membershiptier',
                                            'membershipgrade','membershiplevel'}) or
        compact in {'upgrade','upmembershipgrade','usergrade','membergrade'}):
        tags.add('grade')
    if ('로그인' in joined or parts.intersection({'login','logintype','loginmethod',
                                              'authtype','authmethod','signin'}) or
        compact in {'uplogin','uplogintype','userlogintype','userlogin'}):
        tags.add('login')
    return tags


def customer_field(meta, role):
    """Resolve advertised membership labels; never map age to gender.

    Custom event dimensions are permitted when correctly labeled, but unlike
    customUser dimensions they may only be present on certain events.
    """
    all_choices=[d for d in custom_dimensions(meta)
                 if str(d.get('apiName','')).startswith(('customUser:','customEvent:'))]
    matches=[]
    for item in all_choices:
        name=str(item.get('apiName',''))
        tags=customer_role_tags(meta,name)
        if tags != {role}:
            continue
        names=[item.get('uiName',''),name.partition(':')[2],name]
        if any(canonical(alias)==canonical(label)
               for alias in CUSTOMER_FIELD_ALIASES[role] for label in names):
            matches.append(name)
    if not matches:
        # Do not guess from an unrelated field with a similar substring.
        return None
    matches.sort(key=lambda x:(0 if x.startswith('customUser:') else 1,x))
    return matches[0]


def info_fields(meta):
    d={'page1': ['페이지_1Depth','페이지 1Depth','페이지 1뎁스'],
       'coupon': ['주문 쿠폰','주문_쿠폰','order_coupon'],
       'evt_page': ['이벤트_페이지','이벤트 페이지'],
       'evt_area': ['이벤트_영역','이벤트 영역'],
       'evt_label': ['이벤트_라벨','이벤트 라벨']}
    fields={key:custom_field(meta,names,preferred_scope='customEvent')
            for key,names in d.items()}
    for role in ('gender','age','grade','login'):
        fields[role]=customer_field(meta,role)
    return fields


def admin_custom_dimensions(cfg, token):
    """Read registered dimensions; optional Admin API may be disabled.

    The Admin list is for *diagnosis*, not direct runReport names. Only a field
    advertised in Data API metadata is allowed for the report query.
    """
    url=("https://analyticsadmin.googleapis.com/v1beta/properties/"
         +str(cfg['property'])+"/customDimensions")
    store=st.session_state.setdefault('ga4_v63_admin_cache',{})
    now=dt.datetime.now(dt.timezone.utc)
    key=str(cfg['property'])
    cached=store.get(key)
    if cached and (now-cached[0]).total_seconds()<300:
        return cached[1],cached[2]
    result=[]
    next_page=None
    try:
        for _ in range(5):
            params={'pageSize':200}
            if next_page:params['pageToken']=next_page
            resp=requests.get(url, headers={'Authorization':f'Bearer {token}'},
                              params=params,timeout=18)
            if not resp.ok:
                try: msg=resp.json().get('error',{}).get('message','')
                except ValueError:msg='요청 실패'
                status=(f'Google Analytics Admin API HTTP {resp.status_code}: '
                        +str(msg)[:170])
                store[key]=(now,None,status)
                return None,status
            doc=resp.json()
            result.extend(doc.get('customDimensions',[]))
            next_page=doc.get('nextPageToken')
            if not next_page:break
        store[key]=(now,result,None)
        return result,None
    except requests.RequestException as exc:
        reason='Admin API 연결 오류: '+type(exc).__name__
        store[key]=(now,None,reason)
        return None,reason


def scope_api_name(d):
    scope=str(d.get('scope',''))
    prefix={'USER':'customUser','EVENT':'customEvent','ITEM':'customItem'}.get(scope)
    param=str(d.get('parameterName',''))
    return f'{prefix}:{param}' if prefix and param else ''


def available_custom_labels(meta):
    ds=custom_dimensions(meta)
    labels={}
    for d in ds:
        api_name=d.get('apiName','')
        if api_name:
            ui=str(d.get('uiName') or '(표시 이름 없음)')
            labels[api_name]=f"{ui}  |  {api_name}"
    return labels


def customer_field_mapping_ui(meta,fields):
    """Do not allow a gender/age/grade/login cross-map, including stale picks."""
    labels=available_custom_labels(meta)
    customer_keys=[('gender','성별'),('age','연령'),('grade','멤버십 등급'),
                   ('login','로그인 유형')]
    detected=[]
    for key,title in customer_keys:
        session_key='v63_map_'+key
        current=st.session_state.get(session_key,'')
        roles=customer_role_tags(meta,current) if current in labels else set()
        if current and (current not in labels or (roles and key not in roles)):
            # This is run *before* the widget is instantiated, so Streamlit
            # allows the stale selection to be cleared safely.
            st.session_state[session_key]=''
            detected.append(f'{title} → {current}')
            current=''
        if fields.get(key):
            tags=customer_role_tags(meta,fields[key])
            if tags and key not in tags:
                detected.append(f'{title} → {fields[key]}')
                fields[key]=None

    missing=[(key,title) for key,title in customer_keys if not fields.get(key)]
    has_picks=any(st.session_state.get('v63_map_'+key) for key,_ in customer_keys)
    if not missing and not has_picks and not detected:
        return
    with st.expander('🔎 고객 특성 필드 연결 · 측정기준 확인',expanded=bool(missing or detected)):
        if detected:
            st.warning('잘못된 고객 정보 연결을 자동 해제했습니다: '+', '.join(detected)
                       +'. 성별에는 연령 필드를 사용할 수 없습니다.')
        st.caption('각 항목에 맞는 사용자/이벤트 범위 맞춤 측정기준만 연결합니다. '
                   '일반 Google 인구통계 값으로 자동 대체하지 않습니다.')
        show_other=st.checkbox('이름이 불분명한 맞춤 필드도 표시',value=False,
                               key='v64_show_ambiguous_user_fields',
                               help='연령·성별처럼 다른 항목임이 분명한 필드는 계속 제외합니다.')
        left,right=st.columns(2)
        for index,(key,title) in enumerate(customer_keys):
            with (left if index%2==0 else right):
                allowed=[]
                for api_name in labels:
                    if not api_name.startswith(('customUser:','customEvent:')):
                        continue
                    tags=customer_role_tags(meta,api_name)
                    if tags=={key} or (show_other and not tags):
                        allowed.append(api_name)
                allowed=sorted(allowed,key=lambda v:labels[v].casefold())
                options=['']+allowed
                current=st.session_state.get('v63_map_'+key,'')
                if current not in options:
                    st.session_state['v63_map_'+key]=''
                    current=''
                selected=st.selectbox(
                    title, options,index=options.index(current),
                    format_func=lambda x:'자동 검색 사용' if not x else labels[x],
                    key='v63_map_'+key,
                    help='성별에는 성별 필드, 연령에는 연령 필드만 선택할 수 있습니다.')
                if selected:
                    tags=customer_role_tags(meta,selected)
                    if tags and key not in tags:
                        st.error('다른 항목의 측정기준입니다. 연결하지 않았어요.')
                        fields[key]=None
                    else:
                        fields[key]=selected
                elif fields.get(key):
                    st.caption('자동 연결: '+str(fields[key]))
                elif not allowed:
                    st.caption('맞는 사용자 필드 후보 없음 · 아래 데이터 진단의 전체 목록을 확인해 주세요.')
                else:
                    st.caption('자동 연결 없음 · 위 후보에서 맞는 필드를 선택해 주세요.')
        assigned=[(k,fields.get(k)) for k,_ in customer_keys if fields.get(k)]
        duplicates={v for _,v in assigned if sum(a==v for _,a in assigned)>1}
        if duplicates:
            for key,title in customer_keys:
                if fields.get(key) in duplicates:
                    fields[key]=None
            st.error('하나의 GA4 측정기준을 여러 고객 항목에 중복 연결할 수 없습니다. '
                     '각 항목의 실제 필드를 다시 선택해 주세요.')
        st.caption('이벤트 범위(customEvent) 필드는 해당 이벤트에 값이 기록되어야 집계됩니다. '
                   '해당 항목을 발견하지 못한 경우 다른 값으로 대체하지 않습니다.')


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


def report(cfg,token,start,end,dimensions,metrics,filters=None,limit=250,sort=None,offset=0):
    req={"dateRanges":[{"startDate":str(start),"endDate":str(end)}],
         "dimensions":[{"name":d} for d in dimensions], "metrics":[{"name":m} for m in metrics],
         "limit":str(limit),"offset":str(offset),"keepEmptyRows":False}
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


def overview(cfg,token,start,end,fields,meta=None):
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
    with st.expander('고객 특성 요약 · 성별/연령 출처별 비교',expanded=False):
        st.caption('쇼핑라운지 화면의 screen_view 기준. 회원정보와 Google 추정 인구통계의 수치가 다를 수 있습니다.')
        if meta is not None:
            entries=list(demographic_source_fields(meta,fields))
            cols=st.columns(2)
            for i,kind in enumerate(['성별','연령']):
                with cols[i]:
                    st.markdown('**'+kind+'**')
                    for source,part,dim in entries:
                        if part!=kind:continue
                        d,r=safe_query(cfg,token,start,end,[dim],['activeUsers'],
                                       f_and(shop_filter(),f_eq('eventName','screen_view')),
                                       limit=100,show_error=False)
                        if d is None or d.empty:
                            st.caption(source+': 집계 불가');continue
                        d=d.rename(columns={dim:kind,'activeUsers':'활성 사용자 수'})
                        d=privacy_aggregate(d,'활성 사용자 수')
                        eligible=d.dropna(subset=['활성 사용자 수'])
                        st.caption(source)
                        if not eligible.empty:
                            st.dataframe(eligible.sort_values('활성 사용자 수',ascending=False).head(8),
                                         hide_index=True,use_container_width=True)
                        else:st.caption('표시 가능한 고객 집단 없음')
        else:st.caption('고객 특성 메타데이터를 확인할 수 없습니다.')

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


def lounge_visitor_segment_request(start, end, breakdown, page1):
    """GA4 v1alpha funnel query with a *user* inclusion segment.

    A membership criterion is satisfied by a screen_view carrying the exact
    page_1Depth value. Unlike a normal runReport event filter, the segment is
    user scoped. Funnel step 1 counts screen viewers inside that population.
    This is closer to a GA4 Exploration user segment, but it is not guaranteed
    to be identical in reporting identity, privacy thresholding, or sampling.
    """
    expr={"andGroup":{"expressions":[
        {"segmentEventFilter":{"eventName":"screen_view"}},
        {"segmentFilter":{"fieldName":page1,
                          "stringFilter":{"matchType":"EXACT","value":LOUNGE_TITLE}}}
    ]}}
    return {
        "dateRanges":[{"startDate":str(start),"endDate":str(end)}],
        "funnel":{"steps":[{"name":"화면 이용",
            "filterExpression":{"funnelEventFilter":{"eventName":"screen_view"}}}]},
        "funnelBreakdown":{"breakdownDimension":{"name":breakdown},"limit":"15"},
        "segments":[{"name":"쇼핑라운지 방문 고객",
            "userSegment":{"userInclusionCriteria":{"andConditionGroups":[{
                "conditionScoping":"USER_CRITERIA_WITHIN_SAME_EVENT",
                "segmentFilterExpression":expr
            }]}}}],
        "limit":"250"
    }


def visitor_breakdown(cfg, token, start, end, dim, fields):
    """Try the GA4 user segment; never replace its result without a scope label."""
    page1=fields.get('page1')
    if not page1:
        return None, None, ('쇼핑라운지 방문 조건인 페이지_1Depth의 API 필드를 '
                            '메타데이터에서 찾지 못했습니다.')
    try:
        response=api(cfg,token,':runFunnelReport',
                     lounge_visitor_segment_request(start,end,dim,page1), alpha=True)
        raw=response.get('funnelTable',{})
        df=frame(raw)
        if dim not in df.columns or 'activeUsers' not in df.columns:
            return None, raw, '퍼널 보고서 응답에 고객 분류 또는 활성 사용자 열이 없습니다.'
        # This funnel has exactly one step. GA4 versions may number it
        # differently; don't discard valid rows based on the step label.
        # The report may include a reserved total row; it must not be charted.
        df=df[~df[dim].fillna('').astype(str).str.startswith('RESERVED_')]
        df=df[[dim,'activeUsers']].copy()
        df=df.sort_values('activeUsers',ascending=False).reset_index(drop=True)
        if df.empty or not bool((df['activeUsers']>0).any()):
            return None, raw, '방문 고객 세그먼트의 분류별 활성 사용자 값이 없습니다.'
        return df, raw, None
    except GAError as exc:
        return None, None, str(exc)


def screen_activity_breakdown(cfg, token, start, end, dim, fields):
    """Event-scoped activity on a lounge screen, not a user segment."""
    page1=fields.get('page1')
    if page1:
        cond=f_and(f_eq('eventName','screen_view'),f_eq(page1,LOUNGE_TITLE))
        description='쇼핑라운지 화면 조회 이벤트 기준 (방문자 세그먼트 아님)'
    else:
        cond=f_and(f_eq('eventName','screen_view'),shop_filter())
        description='쇼핑라운지 화면 이름의 조회 이벤트 기준 (방문자 세그먼트 아님)'
    # Avoid requesting screenPageViews alongside custom dimensions when a
    # compatible dimension/metric pair has not been confirmed.
    df,raw=query(cfg,token,start,end,[dim],['activeUsers'],cond,limit=200,sort='activeUsers')
    return df,raw,description


def privacy_aggregate(df, metric, minimum=10):
    """Suppress small demographic cohorts in public dashboard and exports.

    Do not leak raw cohort counts through the chart, dataframe or XLSX.
    """
    df=df.copy()
    if metric in df.columns:
        df[metric]=pd.to_numeric(df[metric],errors='coerce')
        df.loc[df[metric]<minimum,metric]=float('nan')
    return df


def demographic_source_fields(meta,fields):
    choices=[('KT멤버십 회원 정보','성별','gender'),
             ('KT멤버십 회원 정보','연령','age')]
    for scope,kind,key in choices:
        api=fields.get(key)
        if api and available_meta(meta,api):
            yield (scope,kind,api)
    for kind,field in [('성별','userGender'),('연령','userAgeBracket')]:
        if available_meta(meta,field):yield ('GA4 추정 인구통계',kind,field)


def demographic_breakdown(cfg,token,start,end,meta,fields,filter_,label):
    """Query visit/purchase cohorts directly; never swap to entire property silently."""
    options=list(demographic_source_fields(meta,fields))
    if not options:
        st.warning('성별/연령 필드를 GA4에서 확인하지 못했어요.');return
    options_by_source=list(dict.fromkeys(x[0] for x in options))
    cols=st.columns([1,1])
    with cols[0]:
        source=st.selectbox('고객 정보 출처',options_by_source,key=f'v8_demo_src_{label}')
    with cols[1]:
        kind=st.selectbox('고객 특성',['성별','연령'],key=f'v8_demo_kind_{label}')
    dim=next((x[2] for x in options if x[0]==source and x[1]==kind),None)
    if not dim:
        st.info('이 출처의 측정기준이 GA4에 등록되어 있지 않아요.');return
    df,raw=safe_query(cfg,token,start,end,[dim],['activeUsers'],filter_,limit=200,sort='activeUsers')
    if df is None or df.empty:
        st.info('이 조합에서는 고객 특성값을 조회할 수 없어요. 상품/이벤트와 측정기준의 호환 여부를 확인해 주세요.');return
    df=df.rename(columns={dim:kind,'activeUsers':'활성 사용자 수'})
    vals=df[kind].astype(str).str.strip().str.casefold()
    known=vals[~vals.isin(['','(not set)','unknown','(other)'])]
    if kind=='성별' and not known.empty and known.str.fullmatch(r'\d{1,3}(?:대|세)?').mean()>.5:
        st.error('연령 값이 성별로 매핑되는 오류를 감지하여 표시하지 않습니다.');return
    if kind=='연령' and not known.empty and known.isin(['male','female','남성','여성','남','여']).mean()>.5:
        st.error('성별 값이 연령으로 매핑되는 오류를 감지하여 표시하지 않습니다.');return
    df=privacy_aggregate(df,'활성 사용자 수')
    shown=df.dropna(subset=['활성 사용자 수']).copy()
    a,b=st.columns([.85,1.3],gap='large')
    with a:
        if not shown.empty:bar(shown,kind,'활성 사용자 수',h=270,n=15)
        else:st.info('개인정보 보호 기준으로 표시 가능한 집단이 없어요.')
    with b:table(df,f'ga4_{label}_{source}_{kind}_{start}_{end}',height=275)
    st.caption(f'{source} · 측정기준 {dim} · 10명 미만 비공개. GA4에서 인구통계가 제한될 수 있습니다.')
    quality(raw)


def scoped_time_chart(cfg,token,start,end,filter_,metrics,key,description):
    df,raw=safe_query(cfg,token,start,end,['hour'],metrics,filter_,limit=24,show_error=False)
    if df is None:
        st.warning('시간대 조회가 지원되지 않거나 GA4 측정항목 조합이 호환되지 않습니다.');return
    time=pd.DataFrame({'hour':[f'{i:02d}' for i in range(24)]})
    if df.empty:
        for m in metrics:time[m]=0
    else:
        df['hour']=df['hour'].astype(str).str.zfill(2)
        time=time.merge(df[['hour']+metrics],on='hour',how='left').fillna(0)
    fig=go.Figure()
    for i,field in enumerate(metrics):
        fig.add_trace(go.Scatter(x=time['hour']+'시',y=time[field],name=field,mode='lines+markers',
                    marker=dict(color=BLUE if i%2==0 else TEAL)))
    fig.update_layout(height=300,margin=dict(l=6,r=10,t=12,b=8),
                      paper_bgcolor='white',plot_bgcolor='white',hovermode='x unified',
                      legend=dict(orientation='h',y=1.1))
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False},key=key)
    if description:st.caption(description)
    quality(raw)


def customers(cfg,token,start,end,fields,meta=None):
    heading('어떤 고객이 쇼핑라운지를 이용할까요?',
            'KT멤버십 등록 정보와 GA4 추정 인구통계를 서로 다른 데이터로 표시합니다. 분류되지 않은 고객도 확인할 수 있어요.')
    if meta is not None:
        customer_field_mapping_ui(meta,fields)
    selected=st.radio('고객 구분',['기기','성별','연령','멤버십 등급','로그인 유형','기기별 구매'],
                      horizontal=True,label_visibility='collapsed',key='v8_customer_select')
    if selected=='기기별 구매':
        st.caption('쇼핑라운지 주문완료 화면에 기록된 구매 이벤트 기준 · 운영체제별 구매 거래/수익')
        df,raw=safe_query(cfg,token,start,end,['operatingSystem'],['transactions',GROSS_REVENUE],
                          order_filter(),limit=100,sort=GROSS_REVENUE)
        if df is None or df.empty:return
        df['operatingSystem']=df['operatingSystem'].replace({'iOS':'아이폰(iOS)','Android':'안드로이드(Android)'})
        a,b=st.columns([1,1.3],gap='large')
        with a:bar(df,'operatingSystem','transactions')
        with b:table(df.rename(columns={'operatingSystem':'기기','transactions':'구매 거래 수',
                                        GROSS_REVENUE:'구매 수익'}),f'ga4_기기별구매_{start}_{end}')
        quality(raw);return
    if selected=='기기':
        source='기기(운영체제)'
        dim='operatingSystem'
    else:
        options=['KT멤버십 회원 정보','GA4 추정 인구통계'] if selected in ('성별','연령') else ['KT멤버십 회원 정보']
        source=st.radio('고객 정보 출처',options,horizontal=True,key=f'v8_source_{selected}')
        if source=='GA4 추정 인구통계':
            dim='userGender' if selected=='성별' else 'userAgeBracket'
        else:
            dim=fields.get({'성별':'gender','연령':'age','멤버십 등급':'grade','로그인 유형':'login'}[selected])
    if not dim:
        st.warning(f'{selected}에 해당하는 KT멤버십 고객 정보 필드를 GA4 메타데이터에서 찾지 못했어요. '
                   '페이지 하단의 데이터 진단에서 실제 측정기준을 확인해 주세요.')
        return
    if selected in ('성별','연령') and source=='GA4 추정 인구통계':
        st.info('Google에서 추정한 인구통계입니다. 미확인 고객과 개인정보 보호 제한으로 회원정보와 분포가 다를 수 있어요.')
    scope=st.radio('집계 기준',['쇼핑라운지 화면 이용','KT멤버십 전체 속성'],horizontal=True,
                   key='v8_customer_scope',help='화면 이용은 쇼핑라운지 screen_view 이벤트에 값이 기록된 경우를 집계합니다.')
    cond=(f_and(f_eq('eventName','screen_view'),shop_filter())
          if scope=='쇼핑라운지 화면 이용' else None)
    df,raw=safe_query(cfg,token,start,end,[dim],['activeUsers'],cond,limit=200,sort='activeUsers')
    if df is None or df.empty:return
    df=df.rename(columns={dim:selected,'activeUsers':'활성 사용자 수'})
    if selected=='성별':
        vals=df[selected].astype(str).str.strip()
        known=vals[~vals.isin(['','(not set)','unknown','(other)'])]
        if not known.empty and known.str.fullmatch(r'\d{1,3}(?:대|세)?').mean()>.5:
            st.error('성별 필드에서 연령 숫자가 반환되어 결과를 표시하지 않습니다. 필드 연결을 확인해 주세요.');return
    if selected=='연령':
        vals=df[selected].astype(str).str.casefold().str.strip()
        known=vals[~vals.isin(['','(not set)','unknown','(other)'])]
        if not known.empty and known.isin(['남','여','male','female','남성','여성']).mean()>.5:
            st.error('연령 필드에서 성별이 반환되어 결과를 표시하지 않습니다.');return
    df=privacy_aggregate(df,'활성 사용자 수')
    left,right=st.columns([1,1.3],gap='large')
    numeric=df[df['활성 사용자 수'].notna()].copy()
    with left:
        if numeric.empty: st.info('개인정보 보호 기준에 따라 표시 가능한 분류가 없어요.')
        else:bar(numeric,selected,'활성 사용자 수',h=330)
    with right:table(df,f'ga4_{selected}_{source}_{start}_{end}',height=315)
    st.caption(f'측정기준: {dim} · 범위: {scope}. 10명 미만의 세분화 수치는 공개 대시보드에서 숨깁니다. '
               '분류별 고유 사용자 수 합계는 기간 전체 고유 사용자 수와 다를 수 있습니다.')
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


def hourly_dashboard(cfg,token,start,end,fields):
    heading('시간대별 쇼핑라운지 방문 고객·세션·구매',
            '같은 시간에 방문한 고객 수와 세션 수를 나란히 보고, 구매 건수·금액을 함께 비교해요.')
    visit_filter=f_and(shop_filter(),f_eq('eventName','screen_view'))
    left,right=st.columns(2)
    with left:
        metric=st.radio('방문 지표',['활성 사용자','세션 수'],horizontal=True,key='v8_hour_traffic')
        visit_metric='activeUsers' if metric=='활성 사용자' else 'sessions'
        df,raw=safe_query(cfg,token,start,end,['hour'],[visit_metric],visit_filter,limit=24)
        if df is not None:
            chart_24hours(df,visit_metric,metric,'v8_visit_hour')
            quality(raw)
    with right:
        metric=st.radio('구매 지표',['구매 건수','구매 금액'],horizontal=True,key='v8_hour_purchase')
        purchase_metric='transactions' if metric=='구매 건수' else GROSS_REVENUE
        df,raw=safe_query(cfg,token,start,end,['hour'],[purchase_metric],order_filter(),limit=24)
        if df is not None:
            chart_24hours(df,purchase_metric,metric,'v8_purchase_hour')
            quality(raw)
    st.caption('방문은 쇼핑라운지 screen_view 이벤트, 구매는 주문완료 화면의 purchase 이벤트 기준입니다. '
               '서로 다른 시간대의 고객 수는 중복될 수 있으므로 24시간 합계를 전체 고유 고객 수로 해석하면 안 됩니다.')
    with st.expander('고객 수와 세션 수를 한 그래프에서 비교',expanded=False):
        visit,_=safe_query(cfg,token,start,end,['hour'],['activeUsers','sessions'],visit_filter,
                           limit=24,show_error=False)
        if visit is not None:
            fig=go.Figure()
            if not visit.empty:
                visit['hour']=visit['hour'].astype(str).str.zfill(2)
                visit=visit.sort_values('hour')
                for m,label,color in [('activeUsers','고객 수',BLUE),('sessions','세션 수',TEAL)]:
                    fig.add_trace(go.Scatter(x=visit['hour']+'시',y=visit[m],name=label,
                                             mode='lines+markers',line=dict(color=color)))
                fig.update_layout(height=290,hovermode='x unified',margin=dict(l=8,r=8,t=18,b=8))
                st.plotly_chart(fig,use_container_width=True,key='v8_hour_double')
            else:st.info('선택 기간 데이터가 없습니다.')
    with st.expander('요일 × 시간대별 방문 패턴',expanded=False):
        kind=st.radio('요일별 지표',['방문 고객 수','방문 세션 수'],horizontal=True,key='v8_weekhour_metric')
        chosen='activeUsers' if kind=='방문 고객 수' else 'sessions'
        w,wr=safe_query(cfg,token,start,end,['dayOfWeek','hour'],[chosen],visit_filter,
                        limit=300,show_error=False)
        if w is None:
            st.warning('요일×시간대 조합을 현재 GA4 API에서 조회하지 못했습니다.')
        elif w.empty:
            st.info('요일×시간대 기록이 없습니다.')
        else:
            w['d']=pd.to_numeric(w['dayOfWeek'],errors='coerce')
            w['h']=pd.to_numeric(w['hour'],errors='coerce')
            matrix=[[w.loc[(w['d']==day)&(w['h']==hh),chosen].sum() for hh in range(24)]
                    for day in [1,2,3,4,5,6,0]]
            fig=go.Figure(go.Heatmap(z=matrix,x=[f'{h:02d}시' for h in range(24)],
                y=['월','화','수','목','금','토','일'],colorscale='Blues'))
            fig.update_layout(height=315,margin=dict(l=5,r=5,t=8,b=7),
                              yaxis=dict(autorange='reversed'))
            st.plotly_chart(fig,use_container_width=True,key='v8_weekday_heatmap')
            st.caption('요일·시간 단위 집계값입니다. 여러 요일/시간의 고유 고객 수를 합쳐 '
                       '전체 방문 고객 수로 해석하면 안 됩니다.')
            quality(wr)
    with st.expander('특정 날짜의 시간대별 실적 보기',expanded=False):
        day=st.date_input('분석할 날짜',value=end,min_value=start,max_value=end,key='v8_hour_day')
        a,b=st.columns(2)
        with a:
            d,_=safe_query(cfg,token,day,day,['hour'],['activeUsers','sessions'],visit_filter,limit=24)
            if d is not None and not d.empty:
                chart_24hours(d,'activeUsers','방문 고객 수','v8_hour_day_visit')
        with b:
            d,_=safe_query(cfg,token,day,day,['hour'],['transactions',GROSS_REVENUE],order_filter(),limit=24)
            if d is not None and not d.empty:
                chart_24hours(d,'transactions','구매 건수','v8_hour_day_purchase')
    with st.expander('KT멤버십 앱 쇼핑 관련 클릭의 시간대',expanded=False):
        label=fields.get('evt_label')
        if not label:st.warning('이벤트_라벨 필드가 없어 클릭 시간대를 조회할 수 없습니다.')
        else:
            filt=f_and(f_eq('eventName','click_event'),
                       f_eq('sessionDefaultChannelGroup','Direct'),
                       f_regex(label,r'.*(쇼핑|shopping).*'))
            d,raw=safe_query(cfg,token,start,end,['hour'],['eventCount'],filt,limit=24)
            if d is not None:chart_24hours(d,'eventCount','쇼핑 관련 클릭 수','v8_hour_click')
            st.info('이 그래프는 KT멤버십 앱의 쇼핑 관련 클릭 시각이지 쇼핑라운지 진입 완료 시각이 아닙니다. '
                    '온사이트 메인배너·팝업·GNB 클릭 분석은 요청대로 제외했습니다.')


def chart_24hours(df,metric,title,key):
    full=pd.DataFrame({'hour':[f'{x:02d}' for x in range(24)]})
    if df is not None and not df.empty:
        sample=df.copy()
        sample['hour']=sample['hour'].astype(str).str.zfill(2)
        full=full.merge(sample[['hour',metric]],on='hour',how='left')
        full[metric]=full[metric].fillna(0)
    else:full[metric]=0
    fig=go.Figure(go.Bar(x=full['hour']+'시',y=full[metric],marker_color=TEAL,
                          hovertemplate='%{x}: %{y:,.0f}<extra></extra>'))
    fig.update_layout(height=260,margin=dict(l=6,r=8,t=10,b=10),
                      title=dict(text=title,font=dict(size=13)),
                      plot_bgcolor='white',paper_bgcolor='white')
    st.plotly_chart(fig,use_container_width=True,key=key)


def channels(cfg,token,start,end,fields):
    heading('고객은 어디에서 들어오고 무엇을 클릭할까요?',
            'KT멤버십 앱 내 쇼핑 관련 클릭 경로와 GA4 세션 유입 채널을 구분해 보여드려요.')
    mode=st.radio('유입 보고서',['시간대별 방문·구매','KT멤버십 앱 쇼핑 관련 클릭','세션 유입 채널'],
                  horizontal=True,label_visibility='collapsed',key='v8_acquisition')
    if mode=='시간대별 방문·구매':
        hourly_dashboard(cfg,token,start,end,fields)
        return
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


def promo_purchase_if_tagged(cfg,token,start,end,promo_id_value,meta):
    """Only count purchases explicitly tagged with this promotion ID."""
    if not promo_id_value or meta is None:return None
    candidates=['promotionId']
    candidates += [d['apiName'] for d in custom_dimensions(meta)
                   if any(v in canonical(d.get('uiName','')+' '+d.get('apiName',''))
                          for v in ('기획전코드','기획전id','promotionid','exhibitionid'))]
    for field in dict.fromkeys(candidates):
        if not available_meta(meta,field):continue
        df,raw=safe_query(cfg,token,start,end,[],['transactions',GROSS_REVENUE],
                          f_and(order_filter(),f_eq(field,promo_id_value)),show_error=False)
        if df is None or df.empty:continue
        transaction=one(df,'transactions')
        revenue=one(df,GROSS_REVENUE)
        if transaction is not None and transaction>0:
            return transaction,revenue,field
    return None


def promotions(cfg,token,start,end,fields=None,meta=None):
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
    st.markdown('#### 선택 기획전 종합 분석')
    choices=p['unifiedScreenName'].astype(str).head(200).tolist()
    if choices:
        selected=st.selectbox('기획전 선택',choices,format_func=lambda name:
            '기획전 '+(promotion_id(name) or '코드 미확인')+' · '+name[-38:],key='v8_promo_choice')
        promo_cond=f_and(f_eq('eventName','screen_view'),f_eq('unifiedScreenName',selected))
        promo_id_value=promotion_id(selected)
        st.caption('방문은 선택한 기획전 화면의 screen_view 이벤트 기준입니다.')
        a,b=st.columns(2)
        with a:
            d,_=safe_query(cfg,token,start,end,[],['activeUsers','screenPageViews'],
                            promo_cond,show_error=False)
            count=one(d,'activeUsers')
            st.metric('기획전 방문 활성 사용자',f'{count:,.0f}명' if count is not None and count>=10 else ('비공개(10명 미만)' if count is not None else '조회 불가'))
        with b:
            d,_=safe_query(cfg,token,start,end,[],['screenPageViews'],promo_cond,show_error=False)
            count=one(d,'screenPageViews')
            st.metric('기획전 화면 조회수',f'{count:,.0f}회' if count is not None else '조회 불가')
        with st.expander('기획전 방문 고객의 성별·연령',expanded=False):
            if meta is not None and fields is not None:
                demographic_breakdown(cfg,token,start,end,meta,fields,promo_cond,
                                       f'promotion_{promo_id_value}')
            else:st.info('고객 필드 정보를 불러오지 못했어요.')
        with st.expander('기획전 시간대별 유입',expanded=False):
            d,_=safe_query(cfg,token,start,end,['hour'],['activeUsers'],promo_cond,limit=24)
            if d is not None:chart_24hours(d,'activeUsers','시간대별 기획전 방문 고객',
                                          f'v8_promo_hour_{promo_id_value}')
        with st.expander('기획전 연결 구매 건수·수익 (수집 여부 확인)',expanded=False):
            st.caption('기획전 방문 이후 모든 주문을 자동으로 해당 기획전 매출로 간주하지 않습니다. '
                       '구매 이벤트 자체에 기획전 식별값이 있는 경우에만 표시합니다.')
            result=promo_purchase_if_tagged(cfg,token,start,end,promo_id_value,meta)
            if result is None:
                st.info('선택 기획전 ID와 구매를 연결하는 기록을 찾지 못했어요. '
                        '기획전별 구매 금액은 현재 확인 불가이며 전체 구매 금액으로 대체하지 않습니다.')
            else:
                tx,rev,field=result
                c1,c2=st.columns(2)
                with c1:st.metric('해당 ID가 기록된 구매 거래',f'{tx:,.0f}건')
                with c2:st.metric('해당 ID가 기록된 구매 수익',krw(rev))
                st.caption('연결 기준: '+field+' · 기록된 주문만 포함하며 기획전의 최종 기여 매출과 다를 수 있습니다.')
    quality(raw)


def shop_item_ids(cfg,token,start,end):
    # IDs observed in lounge item detail screens. This is not a user/session revenue attribution.
    df,raw=page_data(cfg,token,start,end)
    if df is None or df.empty:return set(),raw
    p=df[df["페이지 유형"]=="상품 상세"]
    codes=p["unifiedScreenName"].str.extract(r">상품상세>(\d{6,})",expand=False).dropna()
    return set(codes.astype(str)),raw


def fetch_items_all(cfg,token,start,end,max_rows=30000):
    dims=['itemId','itemName']
    measures=['itemsViewed','itemsPurchased','itemRevenue']
    chunks=[]
    offset=0
    count=None
    try:
        while offset<max_rows:
            batch=report(cfg,token,start,end,dims,measures,limit=5000,
                         sort='itemsViewed',offset=offset)
            df=frame(batch)
            if df.empty:break
            chunks.append(df)
            count=int(batch.get('rowCount',offset+len(df)))
            offset+=len(df)
            if count<=offset:break
        if not chunks:return pd.DataFrame(columns=dims+measures),False
        result=pd.concat(chunks,ignore_index=True)
        return result,(count is not None and count>offset)
    except GAError as exc:
        st.warning('상품 원본 목록 조회 실패: '+str(exc))
        return None,False


def merged_product_catalog(cfg,token,start,end):
    screens,raw=page_data(cfg,token,start,end)
    if screens is None:return None,False
    detail=screens[screens['페이지 유형']=='상품 상세'].copy()
    detail['상품 ID']=detail['unifiedScreenName'].astype(str).str.extract(r'>상품상세>(\d{6,})',expand=False)
    detail=detail.dropna(subset=['상품 ID'])
    summary=detail.groupby('상품 ID',as_index=False).agg({'screenPageViews':'sum',
                                                        'activeUsers':'sum'})
    # Same item may be split into multiple screen names; total activeUsers
    # by screen is not deduplicated across multiple screen-title variants.
    summary=summary.rename(columns={'screenPageViews':'상세 화면 조회수',
                                    'activeUsers':'상세 화면별 고객 수 합계(중복 가능)'})
    items,truncated=fetch_items_all(cfg,token,start,end)
    if items is None:return None,truncated
    items=items[~items['itemId'].astype(str).map(empty_label)].copy()
    # An item name can change. Choose the top-viewed name per ID and aggregate metrics.
    items=items.sort_values('itemsViewed',ascending=False)
    name=items.drop_duplicates('itemId')[['itemId','itemName']]
    amounts=items.groupby('itemId',as_index=False)[['itemsViewed','itemsPurchased','itemRevenue']].sum()
    items=name.merge(amounts,on='itemId',how='left').rename(columns={'itemId':'상품 ID','itemName':'상품명',
               'itemsViewed':'전체 속성 상품 조회 수량',
               'itemsPurchased':'전체 속성 구매 수량',
               'itemRevenue':'전체 속성 상품 수익'})
    result=summary.merge(items,on='상품 ID',how='outer',indicator=True)
    result['상품명']=result['상품명'].fillna('(상품명 미수집)')
    result['식별 출처']=result['_merge'].map({'both':'상세 화면+상품 이벤트',
                                          'left_only':'쇼핑라운지 상세 화면',
                                          'right_only':'전체 속성 상품 이벤트'})
    result=result.drop(columns='_merge').sort_values(
        ['상세 화면 조회수','전체 속성 상품 조회 수량'],ascending=False,na_position='last')
    return result.reset_index(drop=True),truncated


def products(cfg,token,start,end,fields=None,meta=None):
    heading('상품 상세 방문자와 구매자는 어떻게 다를까요?',
            '전체 상품 요약과 상품별 고객 성별·연령을 제공합니다. 미수집 상품은 목록에 나타나지 않을 수 있습니다.')
    st.info('상품 등록 DB가 연결되지 않았으므로 **GA4에서 식별 가능한 상품**만 표시합니다. '
            '조회·구매가 모두 0건인 미관측 등록 상품은 확인할 수 없습니다.')
    df,truncated=merged_product_catalog(cfg,token,start,end)
    if df is None or df.empty:
        st.info('선택 기간에 확인할 수 있는 상품이 없습니다.');return
    a,b,c=st.columns(3)
    with a:st.metric('GA4 식별 상품',f'{len(df):,}개')
    with b:st.metric('쇼핑라운지 상세 화면 식별 상품',f'{df["상세 화면 조회수"].notna().sum():,}개')
    with c:st.metric('전체 속성 구매 수량 기록 상품',f'{(df["전체 속성 구매 수량"].fillna(0)>0).sum():,}개')
    if truncated:st.warning('상품 이벤트 결과가 많아 상위 30,000행까지만 조회했습니다. 일부 상품이 빠질 수 있어요.')
    view=st.radio('상품 목록',['쇼핑라운지 상세 화면 식별','GA4 식별 상품 전체'],
                  horizontal=True,key='v8_item_scope')
    filtered=df.copy()
    if view=='쇼핑라운지 상세 화면 식별':
        filtered=filtered[filtered['상세 화면 조회수'].notna()].copy()
    text_query=st.text_input('상품명 또는 상품 ID로 검색',placeholder='상품명 / 상품코드',key='v8_item_search')
    if text_query:
        filtered=filtered[(filtered['상품명'].astype(str).str.contains(re.escape(text_query),case=False,na=False)) |
                          (filtered['상품 ID'].astype(str).str.contains(re.escape(text_query),case=False,na=False))]
    col1,col2=st.columns([.95,1.5],gap='large')
    with col1:
        rank_col=st.selectbox('순위 기준',['상세 화면 조회수','전체 속성 상품 조회 수량',
                                    '전체 속성 구매 수량','전체 속성 상품 수익'],key='v8_item_rank')
        ranked=filtered.dropna(subset=[rank_col]).sort_values(rank_col,ascending=False)
        if not ranked.empty:
            top=ranked.head(10).copy()
            top['상품 표시']=top.apply(lambda x:(str(x['상품명'])[:18] if str(x['상품명'])!='(상품명 미수집)'
                                           else str(x['상품 ID'])),axis=1)
            bar(top,'상품 표시',rank_col,h=365,n=10)
    with col2:
        cols=['상품 ID','상품명','식별 출처','상세 화면 조회수',
              '전체 속성 상품 조회 수량','전체 속성 구매 수량','전체 속성 상품 수익']
        table(filtered[cols].head(400),f'ga4_상품_요약_{start}_{end}',height=350)
    st.caption('상세 화면 조회수는 쇼핑라운지 화면 기준입니다. 나머지 상품 수량·수익은 KT멤버십 '
               '전체 속성 상품 이벤트 기준으로, 쇼핑라운지 단독 판매 실적이 아닙니다. '
               '화면별 고객 수 합계는 동일 고객 중복 가능성이 있어 통합 고객 수로 쓰지 않았습니다.')
    with st.expander('전체 상품 · 상세 방문 고객과 구매 고객 성별/연령 비교',expanded=False):
        if meta is None or fields is None:
            st.warning('현재 GA4 고객 필드 정보가 없어 분석할 수 없습니다.')
        else:
            visitor_scope=st.radio('전체 상품 상세 조회 기준',
                ['쇼핑라운지 화면 기록','KT멤버십 전체 속성'],
                horizontal=True,key='v8_allproduct_visitor_scope')
            all_view=f_eq('eventName','view_item')
            if visitor_scope=='쇼핑라운지 화면 기록':
                all_view=f_and(all_view,shop_filter())
            st.caption('상세 방문 고객: '+visitor_scope+'의 view_item 이벤트. '
                       '구매 고객: 쇼핑라운지 주문완료 화면 purchase 이벤트.')
            t1,t2=st.tabs(['전체 상품 상세 방문 고객','전체 상품 구매 고객'])
            with t1:
                demographic_breakdown(cfg,token,start,end,meta,fields,all_view,'all_product_view')
            with t2:
                demographic_breakdown(cfg,token,start,end,meta,fields,order_filter(),'all_product_buy')
            st.caption('두 집계는 이벤트 범위가 다르고 구매 이벤트에 itemId가 없을 수도 있습니다. '
                       '동일한 분모의 상품 전환율로 바로 나누지 않습니다.')
    if filtered.empty:return
    options=filtered['상품 ID'].astype(str).tolist()
    selection=st.selectbox('개별 상품 선택',options,
        format_func=lambda ident: f"{str(filtered.loc[filtered['상품 ID'].astype(str)==ident,'상품명'].iloc[0])[:65]} · {ident}",
        key='v8_product_choice')
    if not selection:return
    st.markdown('#### 선택 상품 상세 분석')
    id_cond=f_eq('itemId',selection)
    view_cond=f_and(id_cond,f_eq('eventName','view_item'))
    purchase_cond=f_and(id_cond,order_filter())
    st.caption('**상품 ID별 이벤트 조회** · 상세 방문(view_item)과 주문완료 화면 구매(purchase)를 '
               '서로 다른 조건으로 조회합니다.')
    t1,t2=st.columns(2)
    with t1:
        view_df,view_raw=safe_query(cfg,token,start,end,[],['activeUsers','eventCount'],
                                    view_cond,show_error=False)
        count=one(view_df,'activeUsers')
        st.metric('상품 상세 방문 활성 사용자 (전체 속성 view_item)',
                  f'{count:,.0f}명' if count is not None and count>=10 else ('비공개(10명 미만)' if count is not None else '조회 불가'))
    with t2:
        purchase_df,purchase_raw=safe_query(cfg,token,start,end,[],['activeUsers','transactions'],
                                            purchase_cond,show_error=False)
        purchase_count=one(purchase_df,'activeUsers')
        st.metric('해당 상품의 주문완료 화면 구매 활성 사용자',
                  f'{purchase_count:,.0f}명' if purchase_count is not None and purchase_count>=10 else ('비공개(10명 미만)' if purchase_count is not None else '조회 불가'))
    st.caption('상품 view_item 이벤트 자체에는 쇼핑라운지 진입 경로가 반드시 기록되는 것은 아닙니다. '
               '구매 고객 수는 itemId + 쇼핑라운지 주문완료 화면 필터가 호환될 때만 표시됩니다.')
    if meta is None or fields is None:
        st.warning('고객 성별·연령 분석에 필요한 필드 정보가 없습니다.');return
    source=st.radio('고객 비교 기준',['KT멤버십 회원 정보','GA4 추정 인구통계'],
                    horizontal=True,key='v8_product_customer_source')
    kind=st.radio('고객 특성',['성별','연령'],horizontal=True,key='v8_product_customer_kind')
    dim=(fields.get('gender' if kind=='성별' else 'age') if source=='KT멤버십 회원 정보'
         else 'userGender' if kind=='성별' else 'userAgeBracket')
    if not dim or not available_meta(meta,dim):
        st.warning(f'{source} {kind} 필드가 GA4 메타데이터에 없습니다. 값은 다른 항목으로 대체하지 않아요.')
        return
    p1,p2=st.columns(2,gap='large')
    for pane,title,cond in [(p1,'상품 상세 방문 고객',view_cond),(p2,'주문완료 구매 고객',purchase_cond)]:
        with pane:
            st.markdown(f'**{title} · {source} {kind}**')
            result,raw=safe_query(cfg,token,start,end,[dim],['activeUsers'],cond,
                                  limit=100,show_error=False)
            if result is None:
                st.warning('이벤트와 고객 측정기준을 함께 조회할 수 없어요. (GA4 API 호환성 확인 필요)')
                continue
            if result.empty:
                st.info('해당 상품/고객 정보가 집계되지 않았어요. 0명으로 단정하지 않습니다.')
                continue
            result=result.rename(columns={dim:kind,'activeUsers':'이용 고객 수'})
            if kind=='성별':
                valid=result[kind].astype(str).str.strip()
                known=valid[~valid.isin(['','(not set)','unknown','(other)'])]
                if not known.empty and known.str.fullmatch(r'\d{1,3}(?:대|세)?').mean()>.5:
                    st.error('연령 값이 성별로 연결되어 결과를 숨겼습니다.');continue
            result=privacy_aggregate(result,'이용 고객 수')
            public=result.dropna(subset=['이용 고객 수'])
            if not public.empty:bar(public,kind,'이용 고객 수',h=260,n=10)
            table(result,f'ga4_상품_{selection}_{title}_{source}_{kind}_{start}_{end}',height=235)
            quality(raw)
    st.caption('공개 화면의 작은 고객 집단(10명 미만)은 비공개 처리합니다. GA4 추정 인구통계는 '
               '정보가 없는 사용자가 많거나 임계값이 적용될 수 있습니다.')


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
        st.markdown('**고객 항목별 API 필드 후보**')
        rows=[]
        for key,name in [('gender','성별'),('age','연령'),('grade','멤버십 등급'),
                         ('login','로그인 유형')]:
            candidates=[d for d in custom_dimensions(meta)
                        if str(d.get('apiName','')).startswith('customUser:')
                        and customer_role_tags(meta,d.get('apiName',''))=={key}]
            rows.append({'항목':name,
                         '자동 연결된 필드':fields.get(key) or '없음',
                         '후보 목록':', '.join(d.get('apiName','') for d in candidates)
                                        or '이름으로 특정할 수 없음'})
        st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True)
        st.markdown('**① Data API에서 실제 조회 가능한 맞춤 측정기준**')
        all_custom=pd.DataFrame([
            {'표시 이름':d.get('uiName',''), 'API 이름':d.get('apiName',''),
             '설명':d.get('description','')} for d in custom_dimensions(meta)])
        if all_custom.empty:
            st.warning('이 GA4 속성에서 Data API가 반환한 맞춤 측정기준이 없습니다.')
        else:
            term=st.text_input('맞춤 필드 검색',placeholder='예: 성별, 연령, 고객, 쿠폰, 로그인, grade',
                               key='v63_search_meta')
            showing=all_custom
            if term.strip():
                mask=showing.astype(str).apply(
                    lambda c:c.str.contains(re.escape(term.strip()),case=False,regex=True)
                ).any(axis=1)
                showing=showing[mask]
            st.dataframe(showing,hide_index=True,use_container_width=True,height=250)
            st.download_button('전체 맞춤 필드 CSV 받기',
                               data=all_custom.to_csv(index=False).encode('utf-8-sig'),
                               file_name='ga4_data_api_custom_dimensions.csv',
                               mime='text/csv',key='v63_export_metadata')
        st.markdown('**② GA4 관리자에 등록된 맞춤 측정기준**')
        st.caption('Google Analytics Admin API의 읽기 전용 목록과 비교합니다. '
                   'Admin API가 비활성화돼 있다면 조회할 수 없지만 다른 메뉴는 정상 작동합니다.')
        admin_defs,admin_error=admin_custom_dimensions(cfg,token)
        if admin_error:
            st.info('관리자 목록 조회 불가: '+str(admin_error))
        else:
            admin_rows=[]
            for d in (admin_defs or []):
                api_name=scope_api_name(d)
                admin_rows.append({'표시 이름':d.get('displayName',''),
                                   '등록 파라미터':d.get('parameterName',''),
                                   '범위':d.get('scope',''),
                                   '예상 API 이름':api_name,
                                   'Data API 지원':('발견' if available_meta(meta,api_name)
                                                     else '미발견')})
            admin_df=pd.DataFrame(admin_rows)
            if admin_df.empty:
                st.info('Admin API에서 등록된 맞춤 측정기준 목록이 비어 있습니다.')
            else:
                st.dataframe(admin_df,hide_index=True,use_container_width=True,height=250)
                st.download_button('관리자 등록 목록 CSV 받기',
                    data=admin_df.to_csv(index=False).encode('utf-8-sig'),
                    file_name='ga4_admin_custom_dimensions.csv',
                    mime='text/csv',key='v63_export_admin')
            st.caption('Admin API에 등록돼 있어도 Data API에서 제공하지 않는 항목은 '
                       '조회 가능한 것으로 간주하지 않습니다.')
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
    cfg = config()
    if cfg is None:
        st.error("서비스 계정 설정을 확인해 주세요. Streamlit Secrets의 [service_account] json에 완전한 JSON 키가 필요합니다.")
        st.caption("서비스 계정 JSON은 Streamlit Secrets에만 저장하고 GitHub나 채팅에 올리지 마세요. ")
        st.stop()
    try:
        token = active_token(cfg)
    except (GoogleAuthError, ValueError, TypeError, KeyError, requests.RequestException) as exc:
        st.error("GA4 서비스 계정 인증 실패: 서비스 계정 키 유효성 및 Google Cloud API 설정을 확인해 주세요.")
        st.caption(f"오류 유형: {type(exc).__name__} (비공개 키는 표시하지 않음)")
        st.stop()
    with st.sidebar:
        st.markdown("### 📊 쇼핑라운지 운영 분석")
        selected=st.radio("메뉴",MENUS,label_visibility="collapsed",key="menu_v6")
        st.caption(NAV_HELP[MENUS.index(selected)])
        st.divider()
    start,end=choose_dates()
    with st.sidebar:
        if st.button("↻ 데이터 다시 조회",use_container_width=True):
            st.session_state.pop("ga4_v6_cache",None);st.session_state.pop("ga4_v63_admin_cache",None);st.rerun()
        st.caption(f"GA4 속성 {cfg['property']} · 서비스 계정 자동 조회 · 로그인 불필요")
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
        if handler in [customers,overview]:
            handler(cfg,token,start,end,fields,meta=meta)
        elif handler==pages_view:
            handler(cfg,token,start,end)
        elif handler in [promotions,products]:
            handler(cfg,token,start,end,fields,meta)
        else:
            handler(cfg,token,start,end,fields)
    diagnostic(cfg,token,meta,fields,start,end)
    st.caption("공개 대시보드 · GA4 집계 데이터만 제공 · 작은 고객 집단 비공개 · 지표별 집계 조건을 확인하세요. 데이터 처리 시점에 따라 GA4 탐색 결과와 차이가 날 수 있습니다.")

# alias avoids shadowing `pages` local while keeping readable handler map
pages_view=pages
if __name__=="__main__":main()
