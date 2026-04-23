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
    if parsed.netloc in (
        'drive.google.com',
        'docs.google.com',
        'drive.usercontent.google.com',
    ):
        qs = parse_qs(parsed.query)
        if 'id' in qs and qs['id']:
            return qs['id'][0]
    return None


def _http_get_bytes(url: str, *, timeout: int) -> tuple:
    """(bytes, email.message.Message | None) — urllib 응답 본문과 헤더."""
    req = urllib.request.Request(
        url,
        headers={'User-Agent': 'Mozilla/5.0 (compatible; soon-docx/1.0)'},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read(), resp.info()


def _bytes_look_like_image(data: bytes, headers=None) -> bool:
    """Drive HTML 안내·바이러스 스캔 페이지 등을 걸러낸다."""
    if not data or len(data) < 50:
        return False
    if headers is not None:
        ct = str(headers.get('Content-Type', '') or '').lower()
        if 'text/html' in ct:
            return False
        if 'image' in ct:
            return True
    head = data.lstrip()[:500].lower()
    if head.startswith(b'<!doctype') or head.startswith(b'<html'):
        return False
    if data[:3] == b'\xff\xd8\xff':
        return True
    if len(data) >= 8 and data[:8] == b'\x89PNG\r\n\x1a\n':
        return True
    if data[:6] in (b'GIF87a', b'GIF89a'):
        return True
    if len(data) >= 12 and data[:4] == b'RIFF' and data[8:12] == b'WEBP':
        return True
    return False


def fetch_image_bytes(url: str, creds: Optional[Credentials] = None, *, timeout=60) -> bytes:
    """
    공개 http(s) URL 또는 Google Drive 공유 링크에서 이미지 바이트를 가져온다.
    우선 Drive v3 alt=media(서비스 계정). 실패·무인증·usercontent 링크 등은
    원본 URL 및 uc?export=… 형태로 HTTP 직접 수신을 시도한다.
    """
    url = (url or '').strip()
    if not url.startswith(('http://', 'https://')):
        raise ValueError('http(s) URL이 아닙니다.')

    fid = _drive_file_id_from_url(url)

    if fid and creds is not None:
        try:
            from google.auth.transport.requests import Request as GoogleAuthRequest

            creds.refresh(GoogleAuthRequest())
            media_url = f'https://www.googleapis.com/drive/v3/files/{fid}?alt=media'
            req = urllib.request.Request(
                media_url,
                headers={'Authorization': f'Bearer {creds.token}'},
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = resp.read()
                api_hdrs = resp.info()
            if _bytes_look_like_image(data, api_hdrs):
                return data
        except Exception:
            pass

    if fid:
        candidates = []
        for u in (
            url,
            f'https://drive.google.com/uc?export=download&id={fid}',
            f'https://drive.google.com/uc?export=view&id={fid}',
        ):
            if u not in candidates:
                candidates.append(u)
        last_err = None
        for u in candidates:
            try:
                data, hdrs = _http_get_bytes(u, timeout=timeout)
                if _bytes_look_like_image(data, hdrs):
                    return data
            except Exception as e:
                last_err = e
                continue
        if last_err is not None:
            raise last_err
        raise ValueError('Drive 이미지 URL에서 유효한 바이너리를 받지 못했습니다.')

    data, _hdrs = _http_get_bytes(url, timeout=timeout)
    return data


def _get_sheet_id(config=None) -> str:
    """
    sheet_id:
    - Secrets 인증 모드: 전달된 config(예: app에서 병합한 값)의 sheet_id 우선,
      없으면 st.secrets['sheet_id'], 둘 다 없으면 오류.
    - 로컬 키 파일 모드: config의 sheet_id(config.json 또는 호출부에서 전달).
    """
    if _is_cloud():
        import streamlit as st
        sid = ''
        if config and str(config.get('sheet_id', '')).strip():
            sid = str(config['sheet_id']).strip()
        elif 'sheet_id' in st.secrets and str(st.secrets['sheet_id'] or '').strip():
            sid = str(st.secrets['sheet_id']).strip()
        if not sid:
            raise ValueError(
                'sheet_id가 비어 있습니다. config.json(저장소에 포함 시) 또는 '
                'Streamlit Secrets의 sheet_id를 설정하세요.'
            )
        return sid
    if config is None:
        config = load_config()
    if 'sheet_id' not in config:
        raise ValueError('config.json에 sheet_id가 필요합니다.')
    return config['sheet_id']


_MEAL_PHOTO_HEADERS = (
    '식사사진첨부',
    '식사사진첨부2',
    '식사사진등첨부',
    '식사사진등첨부2',
)


def meal_header_compact(s) -> str:
    """
    1행 헤더 비교용: NBSP·일반 공백 제거, 전각 숫자 → ASCII.
    열 삽입/삭제 후 '식사사진 첨부2'처럼 띄어쓴 헤더가 있어도 코드의 '식사사진첨부2'와 맞춘다.
    """
    if s is None:
        return ''
    t = (
        str(s)
        .replace('\u00a0', ' ')
        .replace('２', '2')
        .replace('１', '1')
        .replace('０', '0')
    )
    return re.sub(r'\s+', '', t)


def _extract_url_from_sheet_formula(formula: str) -> Optional[str]:
    """=HYPERLINK(\"url\",...) / =IMAGE(\"url\") 등에서 첫 번째 URL 문자열 추출."""
    if not formula or not str(formula).strip().startswith('='):
        return None
    s = str(formula)
    for pat in (
        r'HYPERLINK\s*\(\s*"([^"]+)"',
        r"HYPERLINK\s*\(\s*'([^']+)'",
        r'=\s*IMAGE\s*\(\s*"([^"]+)"',
        r"=\s*IMAGE\s*\(\s*'([^']+)'",
    ):
        m = re.search(pat, s, re.I | re.DOTALL)
        if m:
            link = (m.group(1) or '').strip()
            if link.startswith(('http://', 'https://')):
                return link
    return None


def _is_truncated_drive_view_url(url: str) -> bool:
    """=IMAGE(\"...id=\"&BW2) 처럼 첫 인자만 파싱되면 id 뒤 파일키가 비어 있는 URL."""
    u = (url or '').strip()
    if not u.startswith(('http://', 'https://')):
        return False
    if 'drive.google.com' not in u or 'id=' not in u:
        return False
    return not re.search(r'id=[A-Za-z0-9_-]{10,}', u)


def _resolve_image_formula_with_ampersand(formula: str, ws) -> Optional[str]:
    """
    =IMAGE(\"https://...id=\"&BW2) 형태: 따옴표 문자열 뒤에 & 로 셀 참조가 붙은 수식.
    참조 셀 값을 이어 붙여 완성 URL을 만든다. (시트 화면에는 IMAGE로 잘 보여도
    API는 첫 번째 문자열만 주는 경우가 많아 docx 쪽으로 URL이 넘어가지 않던 원인)
    """
    s = str(formula).strip()
    if 'IMAGE' not in s.upper() or '&' not in s:
        return None
    # IMAGE(url, mode, …) 처럼 &셀 뒤에 쉼표 인자가 오는 경우가 많아, 첫 인자 끝은 ',' 또는 ')' 로 본다.
    # 열 이름은 3글자 초과(예: AAAA)도 허용. '= IMAGE (' 처럼 공백도 허용.
    m = re.search(
        r'=\s*IMAGE\s*\(\s*"([^"]*)"\s*&\s*\$?([A-Za-z]+)\$?(\d+)\s*(?:,|\))',
        s,
        re.I,
    )
    if not m:
        m = re.search(
            r"=\s*IMAGE\s*\(\s*'([^']*)'\s*&\s*\$?([A-Za-z]+)\$?(\d+)\s*(?:,|\))",
            s,
            re.I,
        )
    if not m:
        return None
    prefix, col, row_s = m.group(1), m.group(2).upper(), m.group(3)
    a1 = f'{col}{row_s}'
    try:
        part2 = ws.acell(a1).value
    except Exception:
        return None
    if part2 is None:
        return None
    part2 = str(part2).strip()
    if not part2 or part2 in ('#N/A', '#REF!', '#ERROR!') or '찾을 수 없음' in part2:
        return None
    merged = prefix + part2
    if merged.startswith(('http://', 'https://')):
        return merged
    if part2.startswith(('http://', 'https://')):
        return part2
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
    header_canon = None
    header_raw = None
    target = header_name.strip()
    target_c = meal_header_compact(header_name)
    for i, h in enumerate(headers):
        if h is None:
            continue
        hs = str(h).strip()
        if hs == target or meal_header_compact(h) == target_c:
            col_idx = i + 1
            header_canon = hs
            header_raw = str(h) if h is not None else header_canon
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
        # usercontent.google.com/...&id="&BX2 처럼 첫 인자에 id= 만 있고 파일키는 &뒤 셀인 경우:
        # _is_truncated_drive_view_url 은 'drive.google.com' 문자열만 보더라 usercontent 는 병합을 건너뜀 → BU만 되고 BV 실패.
        incomplete_concat_id = (
            url
            and 'id=' in url
            and '&' in fcell
            and 'IMAGE' in fcell.upper()
            and not re.search(r'id=[A-Za-z0-9_-]{10,}', url)
        )
        if not url or _is_truncated_drive_view_url(url) or incomplete_concat_id:
            merged = _resolve_image_formula_with_ampersand(fcell, ws)
            if merged:
                url = merged
        if not url and formatted_rows and i < len(formatted_rows) and formatted_rows[i]:
            fc = formatted_rows[i][0]
            if isinstance(fc, str) and fc.strip().startswith(('http://', 'https://')):
                url = fc.strip()
        if not url:
            for key_try in (header_canon, header_raw):
                if key_try is None:
                    continue
                cur = rec.get(key_try, '')
                if isinstance(cur, str) and cur.strip().startswith(('http://', 'https://')):
                    url = cur.strip()
                    break
        if url:
            # 같은 열로 보이는 기존 dict 키에도 반영 + 1행 헤더 문자열 키에는 항상 기록한다.
            # get_all_records가 BV 열 키를 빼먹거나(병합/빈 헤더) 키 표기가 달라도 URL이 버려지지 않게.
            hk = meal_header_compact(header_canon)
            for rk in list(rec.keys()):
                if meal_header_compact(rk) == hk:
                    rec[rk] = url
            rec[header_canon] = url
            if header_raw and header_raw != header_canon:
                rec[header_raw] = url


def enrich_meal_photo_urls(ws, records: list) -> None:
    """식사사진첨부, 식사사진첨부2 열 수식에서 URL 보강."""
    for hdr in _MEAL_PHOTO_HEADERS:
        _enrich_one_photo_column(ws, records, hdr)


def _drive_uc_view_url_from_file_id(fid) -> Optional[str]:
    """BW/BX 등 image_id 셀 값 → 브라우저와 같은 uc?export=view URL."""
    if fid is None:
        return None
    s = str(fid).strip()
    if not s or s in ('#N/A', '#REF!', '#ERROR!'):
        return None
    if s.startswith(('http://', 'https://')):
        return s
    s = re.sub(r'\.0$', '', s)
    if len(s) < 12:
        return None
    return f'https://drive.google.com/uc?export=view&id={s}'


def _pick_first_nonempty(rec: dict, keys: tuple) -> Optional[str]:
    for k in keys:
        v = rec.get(k)
        if v is None:
            continue
        t = str(v).strip()
        if t:
            return t
    return None


def _slot_has_http_photo_url(rec: dict, header_names: tuple) -> bool:
    hset = {meal_header_compact(h) for h in header_names}
    for k, v in rec.items():
        if meal_header_compact(k) not in hset:
            continue
        s = '' if v is None else str(v).strip()
        if s.startswith(('http://', 'https://')):
            return True
    return False


def _apply_photo_url_to_header_group(rec: dict, header_names: tuple, url: str) -> None:
    if not url:
        return
    hset = {meal_header_compact(h) for h in header_names}
    for rk in list(rec.keys()):
        if meal_header_compact(rk) in hset:
            rec[rk] = url
    for h in header_names:
        rec[h] = url


def enrich_meal_photo_from_image_id_columns(records: list) -> None:
    """
    1행에 image_id / image2_id 같은 열이 있고 BW·BX에 파일 ID만 있을 때,
    IMAGE 수식 보강이 비어 있으면 uc?export=view&id=… URL을 채운다.
    (식사사진첨부 열과 동일한 형식으로 docx/filler가 읽을 수 있게)
    """
    id1_keys = ('image_id', 'IMAGE_ID', 'Image_ID', 'Image_id')
    id2_keys = ('image2_id', 'IMAGE2_ID', 'Image2_ID', 'Image2_id')
    primary = ('식사사진첨부', '식사사진등첨부')
    secondary = ('식사사진첨부2', '식사사진등첨부2')
    for rec in records:
        u1 = _drive_uc_view_url_from_file_id(_pick_first_nonempty(rec, id1_keys))
        u2 = _drive_uc_view_url_from_file_id(_pick_first_nonempty(rec, id2_keys))
        if u1 and not _slot_has_http_photo_url(rec, primary):
            _apply_photo_url_to_header_group(rec, primary, u1)
        if u2 and not _slot_has_http_photo_url(rec, secondary):
            _apply_photo_url_to_header_group(rec, secondary, u2)


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
    enrich_meal_photo_from_image_id_columns(records)
    # 빈 행 제거 + 날짜 시리얼 변환
    records = [_fix_dates(r) for r in records if r.get('성명', '')]
    limit = int(config.get('debug_person_limit') or 0)
    if limit > 0:
        records = records[:limit]
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
