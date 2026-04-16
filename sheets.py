"""
구글 시트 연동 모듈
"""
import json
import os
from datetime import date, timedelta

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
    """Streamlit Secrets에 gcp_service_account가 있으면 클라우드 환경으로 판단"""
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


def get_client(config=None):
    # Streamlit Cloud: Secrets에서 인증 정보 읽기
    if _is_cloud():
        import streamlit as st
        creds = Credentials.from_service_account_info(
            dict(st.secrets['gcp_service_account']),
            scopes=SCOPES,
        )
        return gspread.authorize(creds)

    # 로컬: JSON 파일에서 읽기
    if config is None:
        config = load_config()
    creds_path = config['credentials_file']
    if not os.path.isabs(creds_path):
        creds_path = os.path.join(BASE_DIR, creds_path)
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return gspread.authorize(creds)


def _get_sheet_id(config=None) -> str:
    """sheet_id: 클라우드는 Secrets에서, 로컬은 config에서"""
    if _is_cloud():
        import streamlit as st
        return str(st.secrets['sheet_id'])
    if config is None:
        config = load_config()
    return config['sheet_id']


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
    # 빈 행 제거 + 날짜 시리얼 변환
    records = [_fix_dates(r) for r in records if r.get('성명', '')]
    return records


def test_connection(config=None):
    """연결 테스트 - 성공 시 시트 제목 반환"""
    if config is None:
        config = load_config()
    gc = get_client(config)
    sheet_id = _get_sheet_id(config)
    sh = gc.open_by_key(sheet_id)
    return sh.title
