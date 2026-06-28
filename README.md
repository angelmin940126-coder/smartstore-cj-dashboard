# 스마트스토어 CJ 출고 대시보드

스마트스토어 주문을 조회하고, 발주확인된 주문만 CJ LOIS 업로드용 엑셀로 내려받는 대시보드입니다.

## 환경변수

Vercel 프로젝트 환경변수에 아래 값을 등록해야 합니다.

- `NAVER_PROXY_URL`
- `NAVER_PROXY_SECRET`
- `NAVER_COMMERCE_CLIENT_ID`
- `NAVER_COMMERCE_CLIENT_SECRET`
- `NAVER_COMMERCE_ACCOUNT_TYPE`

AWS 고정 IP 프록시를 사용하는 경우 `NAVER_PROXY_URL`과 `NAVER_PROXY_SECRET`이 필수입니다.

## 로컬 실행

```powershell
python app.py
```

브라우저에서 `http://127.0.0.1:8765`로 접속합니다.
