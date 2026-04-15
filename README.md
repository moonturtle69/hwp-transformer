# EBS 올림포스 HWP 자동 변환기

#### Streamlit Cloud 
1. https://github.com 에서 저장소 생성 후 이 폴더 파일 업로드
2. https://share.streamlit.io 접속 → GitHub 연결 → `app_streamlit.py` 선택
3. Deploy → URL 공유

#### Replit (코딩 없이 배포)
1. https://replit.com 접속 → "Import from GitHub" 또는 파일 업로드
2. Shell에서 `pip install -r requirements.txt` 실행
3. Run → URL 공유

## 📋 변환 규칙
- 영문 ❶ ↔ 한글 ❶ 자동 매칭
- 한글을 영문 위에 배치 (한글 앞 번호 제거)
- 하단 한글 번역 박스 자동 삭제
- 1~N 페이지 전체 자동 처리

## 📦 파일 목록
- `app.py` — Flask 웹앱
- `app_streamlit.py` — Streamlit 앱 (배포용)
- `requirements.txt` — 필요 패키지
- `시작.bat` — Windows 실행 스크립트
- `시작.sh` — Mac/Linux 실행 스크립트
