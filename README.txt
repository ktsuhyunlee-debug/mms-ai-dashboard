MMS AI Dashboard V3.9 SSL Patch

목적
- 회사망의 자체 서명 인증서로 인해 Google Sheets 연결이 실패하는 문제 대응

동작
1. 정상 SSL 인증으로 Google Sheets 연결
2. SSLCertVerificationError가 발생한 경우에만 verify=False로 한 번 재시도
3. Google Sheets 다운로드 요청에만 적용

실행
1. 기존 실행창 Ctrl+C
2. 압축 해제
3. cd %USERPROFILE%\Desktop\MMS_AI_Dashboard_V3_9_SSL_PATCH
4. streamlit run app.py

주의
- 회사 PC 로컬 실행용 임시 호환 패치
- 외부 배포 시에는 회사 인증서를 정식으로 등록하는 방식이 권장됨
