# ICM023 배포 안내

## 요구사항
- Python 3.9+
- Flask, openpyxl, gunicorn (배포 시)

## 1. Render.com (무료 티어)

1. [Render](https://render.com) 가입 후 **New → Web Service**
2. GitHub 저장소 `k30035600/icm023` 연결
3. 설정:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn -b 0.0.0.0:$PORT app:app`
   - **Root Directory**: 비워두거나 프로젝트 루트
4. **Create Web Service** 후 배포 완료 시 제공되는 URL로 접속

> **참고**: Render 무료 티어는 슬립 모드가 있어, 한동안 접속이 없으면 첫 요청이 느릴 수 있습니다. xlsx 파일은 서비스 재시작 시 초기화될 수 있으므로, 중요 데이터는 별도 백업을 권장합니다.

## 2. Railway 배포

### 전제
- GitHub에 `k30035600/icm023` 저장소가 푸시된 상태
- 프로젝트 루트에 `Procfile`, `requirements.txt` 있음 (현재 구성됨)

### 단계

1. **Railway 접속**  
   - https://railway.app 접속 후 **Login** (GitHub로 로그인 권장).

2. **새 프로젝트에서 GitHub 배포**
   - **New Project** 클릭.
   - **Deploy from GitHub repo** 선택.
   - GitHub 권한 허용 후 저장소 목록에서 **icm023** (또는 `k30035600/icm023`) 선택.

3. **서비스 설정**
   - 저장소 연결 후 Railway가 자동으로 빌드·배포 시도.
   - **Settings** 탭에서 확인:
     - **Build Command**: 비워두면 기본으로 `pip install -r requirements.txt` 등으로 빌드.
     - **Start Command**: 비워두면 **Procfile**의 `web: gunicorn -b 0.0.0.0:$PORT app:app` 사용.
   - `PORT`는 Railway가 자동으로 주입하므로 별도 설정 불필요.

4. **도메인 공개 (공개 URL 만들기)**
   - **방법 A** – 서비스가 정상 기동되면, **캔버스의 서비스 카드** 또는 **우측 서비스 패널**에 "Generate domain" / "도메인 생성" 안내가 뜰 수 있음. 그 안내를 따라 **Generate Domain** 실행.
   - **방법 B** – 서비스 클릭 → **Settings** 탭 → 아래로 내려서 **Networking** 섹션 찾기 → **Public Networking** 안에 **Generate Domain** 버튼이 있으면 클릭.
   - **방법 C** – 상단/좌측 메뉴에 **Networking**, **Domains**, **Public URL** 같은 항목이 있으면 들어가서 "Railway-provided domain" 또는 "Generate Domain" 선택.
   - 생성된 URL(예: `https://icm023-production-xxxx.up.railway.app`)로 접속.
   - **Generate Domain이 안 보일 때**: 이 서비스에 **TCP Proxy**가 붙어 있으면 **Generate Domain**이 숨겨질 수 있음. Settings → Networking에서 TCP Proxy가 있으면 삭제(휴지통 아이콘)한 뒤 다시 확인.

5. **재배포**
   - GitHub에 `git push` 하면 Railway가 자동으로 다시 빌드·배포.

### 문제 발생 시
- **Build 실패**: 터미널에서 로컬에선 `pip install -r requirements.txt` 가 정상 동작하는지 확인.
- **실행 후 즉시 종료**: **Start Command**가 `gunicorn -b 0.0.0.0:$PORT app:app` 인지 확인 (Procfile 그대로 사용).
- **502 Bad Gateway**: 첫 배포 후 1~2분 정도 기다린 뒤 다시 접속.

## 3. 로컬 / 자체 서버

```bash
# 개발
python app.py

# 프로덕션 (gunicorn)
pip install -r requirements.txt
gunicorn -b 0.0.0.0:5000 app:app
```

환경 변수 `PORT`가 있으면 해당 포트 사용 가능합니다.
