# 배포 가이드

이 문서는 대시보드를 클라우드에 배포하는 상세한 단계를 설명합니다.

## 사전 준비

1. **OneDrive 공유 링크 준비**
   - Excel 파일을 OneDrive에 업로드
   - 파일 우클릭 → "공유" → "모든 사용자가 이 링크를 사용할 수 있도록 허용"
   - 생성된 링크 복사 (예: `https://1drv.ms/x/...`)

2. **GitHub 저장소 준비** (선택사항)
   - 코드를 GitHub에 푸시하면 자동 배포가 쉬워집니다

## 1단계: 백엔드 배포

### Railway 사용 (권장)

1. [Railway](https://railway.app) 접속 및 가입
2. "New Project" 클릭
3. "Deploy from GitHub repo" 선택 (또는 "Empty Project" 후 수동 배포)
4. 저장소 선택 또는 코드 업로드
5. 프로젝트 설정:
   - **Root Directory**: `backend`
   - **Start Command**: 자동 감지됨 (Procfile 사용)
6. 환경 변수 설정 (Settings → Variables):
   ```
   ONEDRIVE_FILE_URL=https://1drv.ms/x/...
   DASHBOARD_PASSWORD=your-secure-password
   SYNC_INTERVAL=3600
   SUMMARY_EXCEL=★26SS MLB 생산스케쥴_DASHBOARD.xlsx
   SUMMARY_SHEET=수량 기준
   ```
7. 배포 완료 후 생성된 URL 확인 (예: `https://your-app.railway.app`)

### Render 사용

1. [Render](https://render.com) 접속 및 가입
2. "New +" → "Web Service" 클릭
3. GitHub 저장소 연결 또는 수동 배포
4. 서비스 설정:
   - **Name**: `dashboard-backend`
   - **Root Directory**: `backend`
   - **Environment**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn server:app --host 0.0.0.0 --port $PORT`
5. 환경 변수 설정 (Environment Variables):
   ```
   ONEDRIVE_FILE_URL=https://1drv.ms/x/...
   DASHBOARD_PASSWORD=your-secure-password
   SYNC_INTERVAL=3600
   SUMMARY_EXCEL=★26SS MLB 생산스케쥴_DASHBOARD.xlsx
   SUMMARY_SHEET=수량 기준
   ```
6. "Create Web Service" 클릭
7. 배포 완료 후 생성된 URL 확인 (예: `https://dashboard-backend.onrender.com`)

## 2단계: 프론트엔드 배포

### Vercel 사용

1. [Vercel](https://vercel.com) 접속 및 가입
2. "Add New..." → "Project" 클릭
3. GitHub 저장소 연결 또는 수동 업로드
4. 프로젝트 설정:
   - **Framework Preset**: Other
   - **Root Directory**: `frontend`
   - **Build Command**: (비워둠)
   - **Output Directory**: `.`
5. 환경 변수 설정 (Environment Variables):
   ```
   NEXT_PUBLIC_API_URL=https://your-backend-url.railway.app
   ```
   또는 `frontend/config.js` 파일 수정:
   ```javascript
   window.API_URL = "https://your-backend-url.railway.app";
   ```
6. "Deploy" 클릭
7. 배포 완료 후 생성된 URL 확인 (예: `https://your-app.vercel.app`)

### Vercel 환경 변수 주입 방법

Vercel에서 환경 변수를 HTML에 주입하려면:

1. **방법 1: config.js 파일 수정** (권장)
   - 배포 전 `frontend/config.js` 파일 수정
   - 또는 Vercel의 빌드 후처리 스크립트 사용

2. **방법 2: Vercel Serverless Function 사용**
   - `frontend/api/config.js` 생성:
   ```javascript
   export default function handler(req, res) {
     res.json({ apiUrl: process.env.NEXT_PUBLIC_API_URL });
   }
   ```
   - HTML에서 이 API를 호출하여 설정

3. **방법 3: 빌드 시 주입** (가장 간단)
   - `frontend/config.js`를 배포 전에 수정하거나
   - Vercel의 환경 변수를 사용하여 빌드 스크립트로 주입

## 3단계: 테스트

1. 프론트엔드 URL 접속
2. 비밀번호 입력 (백엔드에 설정한 `DASHBOARD_PASSWORD`)
3. 대시보드가 정상적으로 로드되는지 확인
4. 데이터가 올바르게 표시되는지 확인

## 문제 해결

### 백엔드가 시작되지 않음

- Railway/Render 로그 확인
- `requirements.txt`의 패키지가 올바르게 설치되었는지 확인
- 환경 변수가 올바르게 설정되었는지 확인

### OneDrive 동기화 실패

- `ONEDRIVE_FILE_URL`이 올바른지 확인
- OneDrive 링크가 "모든 사용자 접근 가능"으로 설정되었는지 확인
- 백엔드 로그에서 상세 오류 확인

### 프론트엔드에서 API 호출 실패

- 브라우저 콘솔에서 CORS 오류 확인
- 백엔드 URL이 올바른지 확인 (`frontend/config.js` 또는 환경 변수)
- 백엔드가 실행 중인지 확인

### 인증 실패

- 비밀번호가 올바른지 확인
- 백엔드의 `DASHBOARD_PASSWORD` 환경 변수 확인
- 브라우저 개발자 도구에서 네트워크 요청 확인

## 업데이트

### Excel 파일 업데이트

1. OneDrive에서 Excel 파일 수정
2. 백엔드가 자동으로 동기화 (설정된 `SYNC_INTERVAL` 주기)
3. 또는 백엔드를 재시작하여 강제 동기화

### 코드 업데이트

1. 코드 수정 후 GitHub에 푸시
2. Railway/Render와 Vercel이 자동으로 재배포
3. 또는 수동으로 재배포 트리거

## 비용

- **Railway**: 무료 플랜 500시간/월
- **Render**: 무료 플랜 (15분 비활성 시 슬립)
- **Vercel**: 무료 플랜 (제한적)
- **OneDrive**: 무료 (기본 저장 공간)

모든 서비스를 무료 플랜으로 사용 가능합니다.

