# 영양사정기록지 자동 출력 시스템

구글 시트에 입소자 데이터를 입력하고 버튼 한 번으로 영양사정기록지를 자동 생성·출력하는 시스템입니다.

---

## 전체 흐름

```
구글 시트 (데이터 입력)
    ↓ Streamlit 앱에서 "출력" 버튼 클릭
Python (docx 템플릿에 데이터 자동 입력)
    ↓
1번 입소자 → 파일 저장 후 미리보기
2번~ 입소자 → 자동 프린트
```

---

## 사용자가 직접 해야 할 설정 (최초 1회)

### 1단계. 구글 클라우드 API 설정

1. [https://console.cloud.google.com](https://console.cloud.google.com) 접속
   - jkj0601@gmail.com 계정으로 로그인

2. 상단 프로젝트 선택 → **새 프로젝트** 생성
   - 프로젝트 이름 예: `영양사정기록지`

3. 왼쪽 메뉴 → **API 및 서비스** → **라이브러리**
   - `Google Sheets API` 검색 → 사용 설정
   - `Google Drive API` 검색 → 사용 설정

4. 왼쪽 메뉴 → **API 및 서비스** → **사용자 인증 정보**
   - 상단 **+ 사용자 인증 정보 만들기** → **서비스 계정** 선택
   - **서비스 계정 이름**(또는 ID)은 임의의 영문 식별자면 됩니다. (예: `nutrition-sheet`)  
     **본인 Gmail 주소를 넣는 단계가 아닙니다.** 나중에 JSON에 생성되는 `...@프로젝트.iam.gserviceaccount.com` 형태의 **서비스 계정 전용 주소**가 따로 생깁니다.
   - 입력 후 **완료**

5. 생성된 서비스 계정 클릭 → **키** 탭 → **키 추가** → **새 키 만들기**
   - 형식: **JSON** 선택 → 생성
   - 다운로드된 파일을 아래 경로에 저장:
     ```
     D:\j\soon\docx\credentials.json
     ```

   **키 생성이 “조직 정책으로 사용 중지”될 때**  
   회사·학교 등 **Google Workspace / Cloud 조직**에서 `iam.disableServiceAccountKeyCreation`(서비스 계정 키 만들기 금지)을 켜 두면 JSON 키를 만들 수 없습니다. 선택지는 다음과 같습니다.

   - **조직 관리자에게 요청:** `조직 정책 관리자(Organization Policy Administrator)` 권한이 있는 분에게, 이 프로젝트(또는 폴더)에 한해 해당 제약을 끄거나 예외를 달아 달라고 요청합니다. (보안 정책상 거절될 수 있음.)
   - **정책이 없는 계정으로 프로젝트 만들기:** 개인 `@gmail.com`만 연결된 Google Cloud(조직에 소속되지 않은 경우)에서는 같은 메뉴에서 JSON 키를 만들 수 있는 경우가 많습니다. 시트·API는 그 프로젝트의 서비스 계정으로 다시 연결합니다.
   - 이 저장소의 앱은 현재 **서비스 계정 JSON**(로컬) 또는 **Streamlit Secrets의 `[gcp_service_account]`**(배포) 형태를 가정합니다. 키 없이만 쓰려면 OAuth 등 **인증 방식을 바꾸는 개발**이 필요합니다.

---

### 2단계. 구글 시트 생성 및 공유

1. [https://sheets.google.com](https://sheets.google.com) 에서 새 스프레드시트 생성
   - 제목: `영양사정기록지`

2. `credentials.json` 파일을 열어 `client_email` 값 복사
   - 형식 예: `nutrition-sheet@프로젝트명.iam.gserviceaccount.com`

3. 구글 시트 우상단 **공유** 버튼 클릭
   - 복사한 `client_email` 주소 입력
   - 권한: **편집자** 선택 → 완료

4. 시트 URL에서 시트 ID 복사해서 메모장에 저장
   - URL 예: `https://docs.google.com/spreadsheets/d/`**`여기가_시트_ID`**`/edit`
   - 나중에 `config.json` 파일을 만들 때 아래처럼 붙여 넣습니다:
     ```json
     {
       "sheet_id": "여기에_붙여넣기",
       "credentials_file": "credentials.json"
     }
     ```

---

### 3단계. Python 패키지 설치

터미널(cmd)에서 아래 명령 실행:

```bash
pip install gspread streamlit python-docx
```

---

### 4단계. 기본 프린터 확인

- Windows 설정 → 블루투스 및 디바이스 → 프린터 및 스캐너
- 사용할 프린터가 **기본 프린터**로 설정되어 있는지 확인

---

## 구글 시트 컬럼 구조

시트 1행에 아래 헤더를 그대로 붙여 넣으세요.

| 열 | 헤더명 | 입력 형식 |
|----|--------|----------|
| A | 성명 | 텍스트 |
| B | 입소일 | 날짜 |
| C | 생년월일 | 날짜 |
| D | 성별 | 남 / 여 |
| E | 신장(cm) | 숫자 |
| F | 평소체중(kg) | 숫자 |
| G | 등급 | 1~5 |
| H | 식사유형 | 텍스트 |
| I | 1일필요열량(kcal) | 숫자 |
| J | 1일필요단백질(g) | 숫자 |
| K | 식사방법 | 자립식사 / 부분도움 / 완전도움 |
| L | 식사섭취상태 | 양호 / 보통 / 불량 |
| M | 식사속도 | 양호 / 보통 / 불량 |
| N | 도구_젓가락 | TRUE / FALSE |
| O | 도구_숟가락 | TRUE / FALSE |
| P | 도구_포크숟가락 | TRUE / FALSE |
| Q | 도구_불가 | TRUE / FALSE |
| R | 문제_식욕저하 | TRUE / FALSE |
| S | 문제_저작곤란 | TRUE / FALSE |
| T | 문제_연하곤란 | TRUE / FALSE |
| U | 문제_소화불량 | TRUE / FALSE |
| V | 문제_구토 | TRUE / FALSE |
| W | 문제_없음 | TRUE / FALSE |
| X | 치아상태 | 양호 / 불량 / 의치(상악) / 의치(하악) / 잔존치아없음 |
| Y | 소화기능 | 정상 / 소화불량 / 역류 / 복부팽만 / 설사 / 변비 |
| Z | 배설양상 | 정상 / 설사 / 변비 / 복부팽만 |
| AA | 특이체질내용 | 텍스트 (없으면 빈칸) |
| AB | 주요진단명 | 텍스트 |
| AC | 질환_당뇨 | TRUE / FALSE |
| AD | 질환_고혈압 | TRUE / FALSE |
| AE | 질환_심장질환 | TRUE / FALSE |
| AF | 질환_뇌혈관질환 | TRUE / FALSE |
| AG | 질환_신장질환 | TRUE / FALSE |
| AH | 질환_간질환 | TRUE / FALSE |
| AI | 질환_암 | TRUE / FALSE |
| AJ | 질환_기타내용 | 텍스트 (없으면 빈칸) |
| AK | 복용약물내용 | 텍스트 (없으면 빈칸) |
| AL | 약물영향_식욕저하 | TRUE / FALSE |
| AM | 약물영향_구역구토 | TRUE / FALSE |
| AN | 약물영향_미각변화 | TRUE / FALSE |
| AO | 약물영향_흡수장애 | TRUE / FALSE |
| AP | 약물영향_기타내용 | 텍스트 (없으면 빈칸) |
| AQ | 종교 | 없음 / 기독교 / 천주교 / 불교 / 기타 |
| AR | 금식일기도시간 | 텍스트 (없으면 빈칸) |
| AS | 종교적식사제한 | 없음 / 육류제한 / 특정음식금기 |
| AT | 종교제한내용 | 텍스트 (없으면 빈칸) |
| AU | 문화적식습관내용 | 텍스트 (없으면 빈칸) |
| AV | 출신지역특성내용 | 텍스트 (없으면 빈칸) |
| AW | 선호음식 | 텍스트 (없으면 없음) |
| AX | 비선호음식 | 텍스트 (없으면 없음) |
| AY | 식품알러지내용 | 텍스트 (없으면 빈칸) |
| AZ | 수급자욕구 | 텍스트 |
| BA | 보호자욕구 | 텍스트 |
| BB | 영양사총평 | 텍스트 |
| BC | 작성일 | 날짜 |
| BD | 영양사이름 | 텍스트 (예: 제갈순) |

> **TRUE / FALSE 항목**은 구글 시트에서 셀 선택 후
> 삽입 → 체크박스 를 클릭하면 체크박스로 표시됩니다.

---

## 앱 실행 방법

설정 완료 후 터미널에서:

```bash
cd D:\j\soon\docx
streamlit run app.py
```

브라우저가 자동으로 열리며 앱이 실행됩니다.

---

## 파일 구조

```
D:\j\soon\docx\
├── app.py                        # Streamlit 메인 앱
├── filler.py                     # docx 템플릿 채우기
├── sheets.py                     # 구글 시트 연동
├── credentials.json              # 구글 API 키 (직접 저장 필요)
├── 영양사정기록지_개정.docx        # 원본 양식
├── 영양사정기록지_템플릿.docx      # 플레이스홀더 삽입된 템플릿 (자동 생성)
├── output/                       # 생성된 파일 저장 폴더
└── README.md                     # 이 파일
```

---

## 주의사항

- `credentials.json` 은 외부에 공유하지 마세요 (개인 API 키 포함).
- 구글 시트에서 날짜 형식은 `YYYY-MM-DD` 로 입력하세요.
- 출력 시 Windows 기본 프린터로 인쇄됩니다.
