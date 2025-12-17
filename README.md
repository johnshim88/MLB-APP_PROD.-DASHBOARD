# 26SS MLB 생산 스케줄 대시보드

OneDrive에서 실시간으로 동기화되는 Excel 데이터를 기반으로 한 웹 대시보드입니다.

## 아키텍처

```
[OneDrive] → [Railway/Render Backend] → [Vercel Frontend] → [Users]
     ↓              ↓                           ↓
  Excel 파일    주기적 동기화              정적 HTML
```

## 프로젝트 구조

```
DASHBOARD/
├── backend/
│   ├── server.py              # FastAPI 백엔드 서버
│   ├── onedrive_sync.py       # OneDrive 동기화 모듈
│   ├── requirements.txt       # Python 의존성
│   ├── Procfile              # Railway 배포 설정
│   └── render.yaml           # Render 배포 설정
├── frontend/
│   ├── index.html            # 대시보드 HTML 파일
│   └── vercel.json           # Vercel 배포 설정
└── .gitignore                # Git 제외 파일 목록
```

## 배포 가이드

### 1. 백엔드 배포 (Railway 또는 Render)

#### Railway 배포

1. [Railway](https://railway.app)에 가입하고 새 프로젝트 생성
2. GitHub 저장소 연결 또는 직접 배포
3. `backend` 폴더를 루트로 설정
4. 환경 변수 설정:
   - `ONEDRIVE_FILE_URL`: OneDrive 공유 링크
   - `DASHBOARD_PASSWORD`: 대시보드 접근 비밀번호
   - `SYNC_INTERVAL`: 동기화 주기 (초, 기본값: 3600)
   - `SUMMARY_EXCEL`: 엑셀 파일명 (기본값: "★26SS MLB 생산스케쥴_DASHBOARD.xlsx")
   - `SUMMARY_SHEET`: 시트 이름 (기본값: "수량 기준")
5. 배포 완료 후 백엔드 URL 확인 (예: `https://your-app.railway.app`)

#### Render 배포

1. [Render](https://render.com)에 가입하고 새 Web Service 생성
2. GitHub 저장소 연결
3. 설정:
   - **Root Directory**: `backend`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn server:app --host 0.0.0.0 --port $PORT`
4. 환경 변수 설정 (Railway와 동일)
5. 배포 완료 후 백엔드 URL 확인

### 2. 프론트엔드 배포 (Vercel)

1. [Vercel](https://vercel.com)에 가입하고 새 프로젝트 생성
2. GitHub 저장소 연결
3. 설정:
   - **Root Directory**: `frontend`
   - **Framework Preset**: Other
   - **Build Command**: (비워둠)
   - **Output Directory**: `.`
4. 환경 변수 설정:
   - `NEXT_PUBLIC_API_URL`: 백엔드 API URL (예: `https://your-app.railway.app`)
5. 배포 완료 후 프론트엔드 URL 확인

### 3. OneDrive 설정

1. Excel 파일을 OneDrive에 업로드
2. 파일을 우클릭하고 "공유" 선택
3. "모든 사용자가 이 링크를 사용할 수 있도록 허용" 설정
4. 생성된 공유 링크를 복사
5. 백엔드 환경 변수 `ONEDRIVE_FILE_URL`에 링크 설정

## 로컬 개발

### 백엔드 실행

```bash
cd backend
pip install -r requirements.txt
uvicorn server:app --reload
```

### 프론트엔드 실행

1. `frontend/index.html`을 브라우저에서 열기
2. 또는 간단한 HTTP 서버 사용:
   ```bash
   cd frontend
   python -m http.server 8080
   ```
3. 브라우저에서 `http://localhost:8080` 접속

## 환경 변수

### Backend

- `ONEDRIVE_FILE_URL`: OneDrive 파일 공유 링크 (필수)
- `DASHBOARD_PASSWORD`: 대시보드 접근 비밀번호 (필수)
- `SYNC_INTERVAL`: 동기화 주기 (초 단위, 기본값: 3600)
- `SUMMARY_EXCEL`: 엑셀 파일명 (기본값: "★26SS MLB 생산스케쥴_DASHBOARD.xlsx")
- `SUMMARY_SHEET`: 시트 이름 (기본값: "수량 기준")

### Frontend

- `NEXT_PUBLIC_API_URL`: 백엔드 API URL (Vercel 환경 변수로 설정)

## 기능

- ✅ 실시간 Excel 데이터 동기화 (OneDrive)
- ✅ 비밀번호 기반 인증
- ✅ 수량 기준 / 스타일수 기준 대시보드 전환
- ✅ 주차별 데이터 조회 (금주/차주)
- ✅ 국가별, 아이템별, 복종별 시각화
- ✅ 세부 복종별 차트 (드롭다운 필터)
- ✅ KPI 카드 (총 목표, 현재 완료, 진행률, 주차 목표진행률)

## 주의사항

- OneDrive 공유 링크는 "모든 사용자가 접근 가능"으로 설정해야 합니다
- 비밀번호는 환경 변수로 관리하며, 코드에 하드코딩하지 마세요
- Railway/Render 무료 플랜은 리소스 제한이 있을 수 있습니다
- 파일 크기가 크면 다운로드 시간이 오래 걸릴 수 있습니다

## 문제 해결

### 백엔드가 Excel 파일을 찾을 수 없음

- `ONEDRIVE_FILE_URL` 환경 변수가 올바르게 설정되었는지 확인
- OneDrive 공유 링크가 유효한지 확인
- 백엔드 로그에서 동기화 오류 확인

### 인증 실패

- `DASHBOARD_PASSWORD` 환경 변수가 올바르게 설정되었는지 확인
- 프론트엔드에서 올바른 비밀번호 입력 확인

### CORS 오류

- 백엔드의 CORS 설정에서 프론트엔드 도메인 허용 확인
- `server.py`의 `allow_origins` 설정 확인

