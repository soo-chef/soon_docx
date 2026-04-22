"""
구글 시트 연동 모듈
"""
import json
import os
import re
import urllib.request
from datetime import date, timedelta
from typing import Optional
from urllib.parse import urlparse, parse_qs

import gspread
from google.oauth2.service_account import Credentials

# 구글 시트 날짜 시리얼 → 문자열 변환
_GOOGLE_EPOCH = date(1899, 12, 30)
_DATE_COLS = {'입소일', '생년월일', '작성일'}


def _fix_dates(record: dict) -> dict:
    for col in _DATE_COLS:
        val = record.get(col)
        if isinstance(val, (int, float)) and val > 100:
            record[col] = str(_GOOGLE_EPOCH + timedelta(days=int(val)))
    return record

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _is_cloud() -> bool:
    """
    Secrets에 서비스 계정 블록이 있으면 JSON 파일 대신 Secrets로 인증한다.
    (Streamlit Cloud뿐 아니라, 로컬에서 .streamlit/secrets.toml을 쓰는 경우도 동일.)
    """
    try:
        import streamlit as st
        return 'gcp_service_account' in st.secrets
    except Exception:
        return False


def load_config(config_path=None):
    if config_path is None:
        config_path = os.path.join(BASE_DIR, 'config.json')
    with open(config_path, encoding='utf-8') as f:
        return json.load(f)


def build_credentials(config=None):
    """시트·드라이브 API 공용 서비스 계정 Credentials."""
    if _is_cloud():
        import streamlit as st
        return Credentials.from_service_account_info(
            dict(st.secrets['gcp_service_account']),
            scopes=SCOPES,
        )

    if config is None:
        config = load_config()
    key = 'credentials_file'
    if key not in config or not str(config.get(key, '')).strip():
        raise ValueError(
            f'로컬(파일) 모드: config.json에 "{key}" 항목이 필요합니다 '
            f'(서비스 계정 JSON 파일명 또는 경로). '
            f'대신 .streamlit/secrets.toml에 [gcp_service_account]를 두면 Secrets로 인증합니다.'
        )
    creds_path = str(config[key]).strip()
    if not os.path.isabs(creds_path):
        creds_path = os.path.join(BASE_DIR, creds_path)
    if not os.path.isfile(creds_path):
        raise FileNotFoundError(
            f'인증 JSON을 찾을 수 없습니다: {creds_path}\n'
            f'config.json의 "{key}" 값을 확인하세요.'
        )
    return Credentials.from_service_account_file(creds_path, scopes=SCOPES)


def get_client(config=None):
    return gspread.authorize(build_credentials(config))


_DRIVE_FILE_ID_RE = re.compile(r'/file/d/([a-zA-Z0-9_-]+)')
_DRIVE_OPEN_ID_RE = re.compile(r'[?&]id=([a-zA-Z0-9_-]+)')


def _drive_file_id_from_url(url: str) -> Optional[str]:
    m = _DRIVE_FILE_ID_RE.search(url)
    if m:
        return m.group(1)
    m = _DRIVE_OPEN_ID_RE.search(url)
    if m:
        return m.group(1)
    parsed = urlparse(url)
    if parsed.netloc in ('drive.google.com', 'docs.google.com'):
        qs = parse_qs(parsed.query)
        if 'id' in qs and qs['id']:
            return qs['id'][0]
    return None


def fetch_image_bytes(url: str, creds: Optional[Credentials] = None, *, timeout=60) -> bytes:
    """
    공개 http(s) URL 또는 Google Drive 공유 링크에서 이미지 바이트를 가져온다.
    Drive는 서비스 계정으로 alt=media 다운로드 — 파일을 해당 client_email과 공유해야 한다.
    """
    url = (url or '').strip()
    if not url.startswith(('http://', 'https://')):
        raise ValueError('http(s) URL이 아닙니다.')

    fid = _drive_file_id_from_url(url)
    if fid:
        if creds is None:
            raise ValueError('Drive에서 받으려면 서비스 계정 인증이 필요합니다.')
        from google.auth.transport.requests import Request as GoogleAuthRequest

        creds.refresh(GoogleAuthRequest())
        media_url = f'https://www.googleapis.com/drive/v3/files/{fid}?alt=media'
        req = urllib.request.Request(
            media_url,
            headers={'Authorization': f'Bearer {creds.token}'},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.read()

    req = urllib.request.Request(
        url,
        headers={'User-Agent': 'Mozilla/5.0 (compatible; soon-docx/1.0)'},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _get_sheet_id(config=None) -> str:
    """sheet_id: Secrets 모드면 st.secrets, 아니면 config.json"""
    if _is_cloud():
        import streamlit as st
        if 'sheet_id' not in st.secrets:
            raise ValueError(
                'Secrets 모드인데 sheet_id가 없습니다. '
                'Streamlit Secrets(또는 secrets.toml)에 sheet_id를 추가하세요.'
            )
        return str(st.secrets['sheet_id'])
    if config is None:
        config = load_config()
    if 'sheet_id' not in config:
        raise ValueError('config.json에 sheet_id가 필요합니다.')
    return config['sheet_id']


_MEAL_PHOTO_HEADERS = ('식사사진첨부', '식사사진첨부2')


def _extract_url_from_sheet_formula(formula: str) -> Optional[str]:
    """=HYPERLINK(\"url\",...) / =IMAGE(\"url\") 등에서 첫 번째 URL 문자열 추출."""
    if not formula or not str(formula).strip().startswith('='):
        return None
    s = str(formula)
    for pat in (
        r'HYPERLINK\s*\(\s*"([^"]+)"',
        r"HYPERLINK\s*\(\s*'([^']+)'",
        r'=IMAGE\s*\(\s*"([^"]+)"',
        r"=IMAGE\s*\(\s*'([^']+)'",
    ):
        m = re.search(pat, s, re.I | re.DOTALL)
        if m:
            link = (m.group(1) or '').strip()
            if link.startswith(('http://', 'https://')):
                return link
    return None


def _enrich_one_photo_column(ws, records: list, header_name: str) -> None:
    """
    한 열에 대해 Drive 링크·HYPERLINK 수식 → 레코드에 실제 URL 반영.
    """
    if not records:
        return
    try:
        headers = ws.row_values(1)
    except Exception:
        return

    col_idx = None
    actual_key = None
    target = header_name.strip()
    for i, h in enumerate(headers):
        if h is None:
            continue
        if str(h).strip() == target:
            col_idx = i + 1
            actual_key = h
            break
    if col_idx is None:
        return

    from gspread.utils import ValueRenderOption, rowcol_to_a1

    n = len(records)
    rng = f"{rowcol_to_a1(2, col_idx)}:{rowcol_to_a1(n + 1, col_idx)}"
    try:
        formula_rows = ws.get(rng, value_render_option=ValueRenderOption.formula)
    except Exception:
        return
    try:
        formatted_rows = ws.get(rng, value_render_option=ValueRenderOption.formatted)
    except Exception:
        formatted_rows = None

    for i, rec in enumerate(records):
        if i >= len(formula_rows) or not formula_rows[i]:
            continue
        fcell = formula_rows[i][0]
        fcell = '' if fcell is None else str(fcell)
        url = _extract_url_from_sheet_formula(fcell)
        if not url and formatted_rows and i < len(formatted_rows) and formatted_rows[i]:
            fc = formatted_rows[i][0]
            if isinstance(fc, str) and fc.strip().startswith(('http://', 'https://')):
                url = fc.strip()
        if not url:
            cur = rec.get(actual_key, '')
            if isinstance(cur, str) and cur.strip().startswith(('http://', 'https://')):
                url = cur.strip()
        if url:
            rec[actual_key] = url


def enrich_meal_photo_urls(ws, records: list) -> None:
    """식사사진첨부, 식사사진첨부2 열 수식에서 URL 보강."""
    for hdr in _MEAL_PHOTO_HEADERS:
        _enrich_one_photo_column(ws, records, hdr)


def get_all_records(config=None):
    """구글 시트에서 모든 입소자 데이터 가져오기"""
    if config is None:
        config = load_config()
    gc = get_client(config)
    sheet_id = _get_sheet_id(config)
    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet(config.get('sheet_name', '입소자목록'))
    records = ws.get_all_records(
        expected_headers=None,
        value_render_option='UNFORMATTED_VALUE',
    )
    # Drive 링크·HYPERLINK 수식 → 실제 URL 반영 (성명 필터 전에 전 행 기준)
    enrich_meal_photo_urls(ws, records)
    # 빈 행 제거 + 날짜 시리얼 변환
    records = [_fix_dates(r) for r in records if r.get('성명', '')]
    return records


def test_connection(config=None):
    """
    연결 테스트 — 성공 시 (스프레드시트 파일 제목, 워크시트 탭 제목) 튜플 반환.

    구글 시트에서 '파일 제목'(브라우저 상단)과 '하단 탭 이름'은 서로 다른 값이다.
    데이터는 탭 단위로 읽으며, 파일 제목만으로는 탭을 구분하지 않는다.
    """
    if config is None:
        config = load_config()
    gc = get_client(config)
    sheet_id = _get_sheet_id(config)
    sh = gc.open_by_key(sheet_id)
    tab = (config.get('sheet_name') or '').strip() or '입소자목록'
    ws = sh.worksheet(tab)
    return sh.title, ws.title
