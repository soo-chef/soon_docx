# 배포 가이드

두 가지 상황에 대한 순서를 정리합니다.

---

## 상황 A. 다른 구글 계정으로 로컬에서 사용

> 내 PC에서 계속 실행하되, 구글 계정(시트)만 바꾸는 경우

### 순서

**1. 새 구글 클라우드 프로젝트 & 서비스 계정 생성**
- [console.cloud.google.com](https://console.cloud.google.com) → 새 계정으로 로그인
- 새 프로젝트 생성
- Google Sheets API + Google Drive API 사용 설정
- `IAM 및 관리자` → `서비스 계정` → 새 계정 생성 → JSON 키 다운로드

**2. JSON 파일 교체**
```
D:\j\soon\docx\
  ├── 기존파일.json          ← 삭제 또는 보관
  └── 새파일.json            ← 새로 다운받은 파일 저장
```

**3. config.json 수정**
```json
{
  "sheet_id": "새_구글시트_ID",
  "credentials_file": "새파일.json",
  "sheet_name": "입소자목록"
}
```

**4. 구글 시트 공유**
- 새 JSON 파일 안의 `client_email` 값 복사
- 사용할 구글 시트 → 공유 → 해당 이메일 추가 (편집자)

**5. 앱 실행**
```bash
streamlit run app.py
```

---

## 상황 B. Streamlit Cloud에 배포 (인터넷으로 접속)

> 여러 사람이 브라우저로 접속해서 사용하는 경우

### 사전 준비

| 필요한 것 | 설명 |
|---|---|
| GitHub 계정 | 코드를 올릴 저장소 |
| Streamlit Cloud 계정 | [share.streamlit.io](https://share.streamlit.io) (무료) |
| 구글 서비스 계정 JSON | 상황 A의 1번과 동일 |

---

### 순서

**1. 인증 (Secrets)**

이미 `sheets.py`의 `get_client`가 Streamlit Secrets의 `[gcp_service_account]`를 우선 사용하고, 없으면 로컬 `config.json`의 JSON 파일 경로를 씁니다. 별도 코드 수정은 필요 없습니다.

**2. PDF 변환 (Windows vs Linux)**

`filler.py`의 `build_merged_pdf`가 플랫폼에 따라 분기합니다.

- **Windows (로컬):** `docx2pdf` + Microsoft Word
- **Linux (Streamlit Cloud 등):** `libreoffice --headless`로 PDF 생성 후 `pypdf`로 병합

변환 중 생성되는 임시 디렉터리는 작업 후 삭제됩니다.

**3. 생성 파일 저장 위치**

`filler.py`의 `OUTPUT_DIR`은 Windows에서는 프로젝트의 `output/`, 그 외 OS에서는 `tempfile.gettempdir()`(일반적으로 `/tmp` 계열)을 사용합니다. 클라우드에서는 세션 종료 후 파일이 남지 않는 것이 정상입니다.

**4. 시스템 패키지 (LibreOffice)**

저장소 루트에 `packages.txt`가 있으며, 한 줄에 `libreoffice`를 적어 두었습니다. Streamlit Community Cloud가 배포 시 `apt`로 설치해 `libreoffice` 명령을 씁니다.

**5. GitHub에 코드 올리기**

> ⚠️ JSON 파일은 절대 GitHub에 올리지 마세요!

```bash
# .gitignore 파일에 추가
*.json          # credentials 파일 제외
output/         # 생성 파일 제외
__pycache__/
```

```bash
git init
git add .
git commit -m "영양사정기록지 자동 출력 시스템"
git remote add origin https://github.com/내계정/저장소명.git
git push -u origin main
```

**6. Streamlit Cloud 배포**

1. [share.streamlit.io](https://share.streamlit.io) 접속 → GitHub 로그인
2. `New app` → GitHub 저장소 선택 → `app.py` 선택 → Deploy

**7. Streamlit Secrets 등록**

배포된 앱 대시보드 → `Settings` → `Secrets` → 아래 내용 입력:

```toml
# Streamlit Secrets (TOML 형식)

sheet_id = "구글시트_ID"

[gcp_service_account]
type = "service_account"
project_id = "프로젝트_ID"
private_key_id = "키_ID"
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "서비스계정@프로젝트.iam.gserviceaccount.com"
client_id = "클라이언트_ID"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
```

> JSON 파일을 열면 위 항목들이 모두 있습니다. 복사해서 TOML 형식으로 붙여넣으세요.

**8. 구글 시트 공유**
- `client_email` 값을 구글 시트와 공유 (편집자)

---

## 로컬 vs 클라우드 기능 비교

| 기능 | 로컬 (현재) | Streamlit Cloud |
|---|---|---|
| 구글 시트 연동 | ✅ | ✅ |
| docx 생성 | ✅ | ✅ |
| ZIP 다운로드 | ✅ | ✅ |
| 합본 PDF | ✅ (Word 사용) | ✅ (LibreOffice 사용) |
| 자동 프린트 | ✅ | ❌ (직접 Ctrl+P) |
| 파일 영구 저장 | ✅ output/ 폴더 | ❌ 세션 종료 시 삭제 |

---

## 파일 구조 요약

```
D:\j\soon\docx\
├── app.py                    # Streamlit 앱
├── filler.py                 # docx 채우기 + PDF 변환
├── sheets.py                 # 구글 시트 연동
├── config.json               # 로컬 설정 (시트 ID, JSON 파일명)
├── soon-493508-da5109c2381f.json  # 구글 서비스 계정 키 (비공개)
├── 영양사정기록지_개정.docx   # 원본 템플릿
├── requirements.txt
├── packages.txt              # Streamlit Cloud: apt 패키지 (LibreOffice)
├── output/                   # 생성된 파일 저장 (Windows 로컬 전용)
├── README.md                 # 초기 설정 가이드
└── DEPLOY.md                 # 이 파일
```
