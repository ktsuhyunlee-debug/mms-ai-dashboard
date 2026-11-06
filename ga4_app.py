# -*- coding: utf-8 -*-
"""GA4 Data API connection test and a read-only Streamlit dashboard.

OAuth 2.0 user-account flow: no GA4 admin permission or service account required.
The person signing in must already have GA4 Viewer or higher on the property.
Never commit OAuth client secrets or tokens.
"""

import datetime as dt
import secrets as pysecrets

import pandas as pd
import requests
import streamlit as st
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
from requests_oauthlib import OAuth2Session

st.set_page_config(
    page_title="쇼핑라운지 GA4 연결 테스트",
    page_icon="📊",
    layout="wide",
)

AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
GA4_API_URL = "https://analyticsdata.googleapis.com/v1beta"
READONLY_SCOPE = "https://www.googleapis.com/auth/analytics.readonly"
DEFAULT_PROPERTY_ID = "544313553"


def load_config():
    try:
        oauth = st.secrets["google_oauth"]
        property_id = str(st.secrets.get("ga4", {}).get("property_id", DEFAULT_PROPERTY_ID))
        vals = {
            "client_id": str(oauth["client_id"]),
            "client_secret": str(oauth["client_secret"]),
            "redirect_uri": str(oauth["redirect_uri"]),
            "state_signing_secret": str(oauth["state_signing_secret"]),
            "property_id": property_id,
        }
        if not all(vals.values()):
            return None
        return vals
    except (KeyError, FileNotFoundError):
        return None


def signer(secret):
    return URLSafeTimedSerializer(secret, salt="ga4-oauth-state-v1")


def auth_url(config):
    state = signer(config["state_signing_secret"]).dumps(
        {"nonce": pysecrets.token_urlsafe(24)}
    )
    client = OAuth2Session(
        client_id=config["client_id"],
        scope=[READONLY_SCOPE],
        redirect_uri=config["redirect_uri"],
        state=state,
    )
    uri, _ = client.authorization_url(
        AUTH_URL,
        access_type="online",
        prompt="select_account",
    )
    return uri


def receive_oauth_callback(config):
    params = st.query_params
    if "error" in params:
        err = str(params.get("error", "unknown"))
        st.query_params.clear()
        st.error(f"Google 로그인 승인이 완료되지 않았어요. 오류 유형: {err}")
        return

    code = params.get("code")
    state = params.get("state")
    if not code:
        return

    try:
        if not state:
            raise BadSignature("missing state")
        signer(config["state_signing_secret"]).loads(state, max_age=600)
        client = OAuth2Session(
            client_id=config["client_id"],
            scope=[READONLY_SCOPE],
            redirect_uri=config["redirect_uri"],
            state=state,
        )
        token = client.fetch_token(
            TOKEN_URL,
            code=code,
            client_secret=config["client_secret"],
            include_client_id=True,
            timeout=20,
        )
        if not token.get("access_token"):
            raise ValueError("empty access token")
        # The token lives in this user's Streamlit session only.
        st.session_state["ga4_access_token"] = token["access_token"]
        st.session_state["ga4_token_expires_at"] = (
            dt.datetime.now(dt.timezone.utc)
            + dt.timedelta(seconds=max(0, int(token.get("expires_in", 3600))-60))
        )
        st.query_params.clear()
        st.rerun()
    except (BadSignature, SignatureExpired):
        st.query_params.clear()
        st.error("로그인 검증 시간이 만료됐거나 요청이 올바르지 않아요. 로그인부터 다시 진행해 주세요.")
    except Exception as exc:
        st.query_params.clear()
        # Never print query values, auth codes, tokens or response bodies.
        st.error(f"로그인 연결에 실패했습니다. 오류 종류: {type(exc).__name__}")
        st.info("Google Cloud의 리디렉션 URI, 테스트 사용자 및 OAuth 인증 설정을 확인해 주세요.")


def current_token():
    access_token = st.session_state.get("ga4_access_token")
    expiry = st.session_state.get("ga4_token_expires_at")
    if not access_token or not expiry:
        return None
    if dt.datetime.now(dt.timezone.utc) >= expiry:
        st.session_state.pop("ga4_access_token", None)
        st.session_state.pop("ga4_token_expires_at", None)
        return None
    return access_token


def ga4_report(access_token, property_id, start, end,
               dims=None, mets=None, dimension_filter=None, limit=100):
    dims = dims or []
    mets = mets or ["activeUsers"]
    payload = {
        "dateRanges": [{"startDate": str(start), "endDate": str(end)}],
        "dimensions": [{"name": n} for n in dims],
        "metrics": [{"name": n} for n in mets],
        "limit": str(limit),
    }
    if dimension_filter:
        payload["dimensionFilter"] = dimension_filter

    response = requests.post(
        f"{GA4_API_URL}/properties/{property_id}:runReport",
        json=payload,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=30,
    )
    if not response.ok:
        try:
            data = response.json().get("error", {})
            msg = str(data.get("message", "API 응답 오류"))
        except ValueError:
            msg = "API 응답 오류"
        # Generic API errors only: no request headers/credentials.
        raise RuntimeError(f"GA4 API {response.status_code}: {msg[:350]}")
    return response.json()


def to_dataframe(result):
    dim_names = [x["name"] for x in result.get("dimensionHeaders", [])]
    met_names = [x["name"] for x in result.get("metricHeaders", [])]
    data = []
    for row in result.get("rows", []):
        values = [v.get("value", "") for v in row.get("dimensionValues", [])]
        nums = [v.get("value", "0") for v in row.get("metricValues", [])]
        data.append(dict(zip(dim_names + met_names, values + nums)))
    frame = pd.DataFrame(data, columns=dim_names + met_names)
    for k in met_names:
        if k in frame:
            frame[k] = pd.to_numeric(frame[k], errors="coerce")
    return frame


def display_report(access_token, property, start, end, dimensions, metrics,
                   filter_expr=None, limit=100):
    try:
        result = ga4_report(
            access_token, property, start, end,
            dims=dimensions, mets=metrics,
            dimension_filter=filter_expr, limit=limit
        )
    except RuntimeError as exc:
        st.error(str(exc))
        return
    df = to_dataframe(result)
    if df.empty:
        st.info("설정한 기간 또는 조건에 해당하는 데이터가 없습니다.")
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(
            "CSV 다운로드",
            df.to_csv(index=False).encode("utf-8-sig"),
            file_name="ga4_report.csv",
            mime="text/csv",
            key=f"csv_{'_'.join(dimensions or ['overall'])}_{'_'.join(metrics)}",
        )


st.title("📊 쇼핑라운지 GA4 연결 테스트")
st.caption("Google 계정 OAuth · 읽기 전용 · 속성 ID 544313553")

config = load_config()
if not config:
    st.warning("아직 Google OAuth 설정값이 등록되지 않았습니다.")
    st.markdown(
        """
        ### 시작하기
        1. [Google Cloud Console](https://console.cloud.google.com/)에서 프로젝트를 만들거나 선택합니다.
        2. **Google Analytics Data API**를 사용 설정합니다.
        3. Google 인증 플랫폼에서 OAuth 동의 화면을 설정하고, 테스트 사용자로 본인 계정을 추가합니다.
        4. OAuth 클라이언트 유형은 **웹 애플리케이션**으로 생성합니다.
        5. `secrets.example.toml`을 참고해 `.streamlit/secrets.toml`을 생성하고 값을 입력합니다.
        6. `streamlit run app.py`로 실행 후 **Google 계정으로 연결**을 누릅니다.

        자세한 절차는 ZIP 내부의 `README.md`에 있어요.
        """
    )
    st.stop()

if config["property_id"] != DEFAULT_PROPERTY_ID:
    st.caption(f"현재 연결 대상 속성: {config['property_id']}")

if not config["redirect_uri"].startswith(("https://", "http://localhost", "http://127.0.0.1")):
    st.error("리디렉션 URI는 HTTPS 또는 로컬호스트 주소여야 합니다.")
    st.stop()

receive_oauth_callback(config)

token = current_token()
if not token:
    st.info("GA4 접근 권한이 있는 Google 계정으로 인증해 주세요. OAuth 로그인만 수행하며 GA4 사용자 권한은 변경하지 않습니다.")
    st.link_button("Google 계정으로 GA4 연결", auth_url(config), type="primary")
    st.caption("회사 정책상 Google Cloud 프로젝트 생성·외부 API 사용 또는 배포가 제한될 수 있으니 사전 승인 여부를 확인하세요.")
    st.stop()

top_left, top_right = st.columns([4, 1])
with top_left:
    st.success("Google 계정 인증이 완료됐습니다. 아래에서 GA4 Data API 연결을 테스트할 수 있어요.")
with top_right:
    if st.button("연결 해제"):
        st.session_state.pop("ga4_access_token", None)
        st.session_state.pop("ga4_token_expires_at", None)
        st.rerun()

today = dt.date.today()
start, end = st.date_input(
    "조회 기간", value=(today - dt.timedelta(days=28), today - dt.timedelta(days=1)),
    max_value=today,
)
if not isinstance(start, dt.date) or not isinstance(end, dt.date):
    st.info("시작 날짜와 종료 날짜를 모두 선택하세요.")
    st.stop()

tab1, tab2, tab3, tab4 = st.tabs(
    ["① 연결 확인", "② 쇼핑라운지 페이지", "③ 유입 경로 (전체 속성)", "④ OS·고객 (전체 속성)"]
)

with tab1:
    st.subheader("GA4 API 호출 테스트")
    st.warning("이 탭의 숫자는 KT멤버십 **전체 GA4 속성** 기준이며 쇼핑라운지 단독 지표가 아닙니다.")
    display_report(token, config["property_id"], start, end,
                   dimensions=[], metrics=["activeUsers", "sessions", "screenPageViews"])

with tab2:
    st.subheader("쇼핑라운지 페이지별 조회수")
    st.caption("페이지 제목/화면 이름에 '쇼핑라운지'가 포함된 항목만 집계합니다. 제목에 해당 문구가 없는 관련 페이지는 누락될 수 있어요.")
    filter_expr = {
        "filter": {
            "fieldName": "unifiedScreenName",
            "stringFilter": {
                "matchType": "CONTAINS",
                "value": "쇼핑라운지",
                "caseSensitive": False,
            },
        }
    }
    display_report(token, config["property_id"], start, end,
                   dimensions=["unifiedScreenName"],
                   metrics=["screenPageViews", "activeUsers"],
                   filter_expr=filter_expr, limit=200)
    st.info("기획전별 '조회'는 살펴볼 수 있지만, 기획전별 '기여 매출'은 구매 이벤트와 기획전 식별정보가 연결돼 있어야 계산할 수 있어요.")

with tab3:
    st.subheader("전체 KT멤버십 속성 · 세션 유입 경로")
    st.warning("MMS와 KT멤버십 앱을 현재 데이터만으로 확실히 구분할 수 있다는 뜻은 아닙니다. UTM/딥링크 태깅을 점검해야 해요.")
    display_report(token, config["property_id"], start, end,
                   dimensions=["sessionSourceMedium"],
                   metrics=["sessions", "activeUsers"], limit=100)

with tab4:
    st.subheader("전체 KT멤버십 속성 · 운영체제")
    st.caption("쇼핑라운지 사용자만의 성별·연령·운영체제 비중은 별도 사용자 세그먼트 검증 후 구현해야 합니다.")
    display_report(token, config["property_id"], start, end,
                   dimensions=["operatingSystem"], metrics=["activeUsers"], limit=50)
    st.divider()
    st.subheader("전체 속성 성별 (제공되는 경우)")
    display_report(token, config["property_id"], start, end,
                   dimensions=["userGender"], metrics=["activeUsers"], limit=20)
    st.subheader("전체 속성 연령대 (제공되는 경우)")
    display_report(token, config["property_id"], start, end,
                   dimensions=["userAgeBracket"], metrics=["activeUsers"], limit=20)

st.divider()
st.caption("보안: OAuth 토큰은 브라우저 세션 종료 또는 만료 시 폐기하며 파일/GitHub에 저장하지 않습니다. Streamlit 서버에는 OAuth 요청과 응답 데이터가 처리되므로 회사 정책을 확인하세요.")
