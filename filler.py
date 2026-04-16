"""
영양사정기록지 docx 자동 채우기 모듈

테이블 행/셀 구조 (unique cells 기준):
  행01 (8셀): 성명 | [값] | 입소일 | [값] | 생년월일 | [값] | 성별 | [값]
  행02 (4셀): 작성일 | [값] | 작성자 | [값]
  행04 (8셀): 신장 | [값 cm] | 평소체중 | [값 kg] | 등급 | [값] | 식사유형 | [값]
  행06 (4셀): 1일필요열량 | [값 kcal] | 1일필요단백질 | [값 g]
  행08 (4셀): 식사방법 | [체크박스] | 식사섭취상태 | [체크박스]
  행09 (4셀): 식사속도 | [체크박스] | 도구사용 | [체크박스 다중]
  행10 (2셀): 식사시문제점 | [체크박스 다중]
  행12 (2셀): 치아상태 | [체크박스]
  행13 (2셀): 소화기능 | [체크박스]
  행14 (2셀): 배설양상 | [체크박스]
  행15 (2셀): 특이체질 | [체크박스+내용]
  행17 (4셀): 선호음식 | [값] | 비선호음식 | [값]
  행18 (2셀): 식품알러지 | [체크박스+내용]
  행20 (2셀): 주요진단명 | [값]
  행21 (2셀): 주요질환 | [체크박스 다중+기타내용]
  행22 (2셀): 현재복용약물 | [체크박스+내용]
  행23 (2셀): 영양관련약물영향 | [체크박스 다중+기타내용]
  행25 (4셀): 종교 | [체크박스] | 금식일기도시간 | [값]
  행26 (2셀): 종교적식사제한 | [체크박스+내용]
  행27 (2셀): 문화적식습관 | [체크박스+내용]
  행28 (2셀): 출신지역특성 | [체크박스+내용]
  행30 (3셀): 개별욕구제목 | 수급자 | [값]
  행31 (3셀): 개별욕구제목 | 보호자 | [값]
  행33 (2셀): 영양사총평 | [값]
"""
import os
import platform
import re
import shutil
import datetime
import tempfile

from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, '영양사정기록지_개정.docx')

# 로컬(Windows): output/ 폴더 사용 / 클라우드(Linux): 시스템 임시 폴더 사용
if platform.system() == 'Windows':
    OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
    os.makedirs(OUTPUT_DIR, exist_ok=True)
else:
    OUTPUT_DIR = tempfile.gettempdir()

# ─────────────────────────────────────────
# 내부 유틸
# ─────────────────────────────────────────

def _unique_cells(row):
    """병합 셀 제거 후 고유 셀만 반환"""
    seen = set()
    cells = []
    for cell in row.cells:
        if id(cell._tc) not in seen:
            seen.add(id(cell._tc))
            cells.append(cell)
    return cells


def _set_text(cell, value, left_align=False, center=False):
    """텍스트 셀에 값 설정 (기존 run 서식 유지)"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn

    if value is None:
        value = ''
    value = str(value).strip()
    para = cell.paragraphs[0]
    if para.runs:
        para.runs[0].text = value
        for r in para.runs[1:]:
            r.text = ''
    else:
        para.add_run(value)

    if left_align:
        # 셀 속성에서 tcFitText / noWrap 제거 (글자 균등배분 원인)
        tcPr = cell._tc.find(qn('w:tcPr'))
        if tcPr is not None:
            for tag in ('w:tcFitText', 'w:noWrap'):
                el = tcPr.find(qn(tag))
                if el is not None:
                    tcPr.remove(el)
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif center:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _set_unit(cell, value, unit):
    """단위 셀: 값이 있으면 '167 cm' 형식으로, 없으면 단위만"""
    if value is None or str(value).strip() == '':
        return
    para = cell.paragraphs[0]
    new_text = f'{value} {unit}'
    if para.runs:
        para.runs[0].text = new_text
        for r in para.runs[1:]:
            r.text = ''
    else:
        para.add_run(new_text)


def _check(cell, options_to_check):
    """
    체크박스 셀에서 선택된 옵션만 ☑, 나머지는 □
    options_to_check: list[str]
    """
    para = cell.paragraphs[0]
    if not para.runs:
        return
    text = para.runs[0].text
    # 기존 ☑ → □ 초기화
    text = text.replace('\u2611', '\u25a1')
    # 선택 옵션 체크
    for opt in options_to_check:
        text = text.replace(f'\u25a1 {opt}', f'\u2611 {opt}')
        text = text.replace(f'\u25a1{opt}', f'\u2611{opt}')
    para.runs[0].text = text


def _check_with_content(cell, selected_option, content=''):
    """
    체크박스 + 내용 텍스트 셀
    예: □ 없음   □ 있음 (내용:                )
    """
    para = cell.paragraphs[0]
    if not para.runs:
        return
    text = para.runs[0].text
    # 초기화
    text = text.replace('\u2611', '\u25a1')
    # 체크
    if selected_option:
        text = text.replace(f'\u25a1 {selected_option}', f'\u2611 {selected_option}')
        text = text.replace(f'\u25a1{selected_option}', f'\u2611{selected_option}')
    # 내용 삽입
    if content:
        content = str(content).strip()
        # 패턴: "내용 :", "내용:", "해당식품 :", "약물명 및 복용 이유 :" 뒤 공백을 내용으로 교체
        text = re.sub(r'(내용\s*:\s*)\s{2,}', f'\\g<1>{content}  ', text)
        text = re.sub(r'(해당식품\s*:\s*)\s{2,}', f'\\g<1>{content}  ', text)
        text = re.sub(r'(약물명 및 복용 이유\s*:\s*)\s{2,}', f'\\g<1>{content}  ', text)
        # 괄호 안 공백 채우기: 기타(   ) → 기타(내용)
        text = re.sub(r'기타\(\s+\)', f'기타({content})', text)
    para.runs[0].text = text


# ─────────────────────────────────────────
# 메인 채우기 함수
# ─────────────────────────────────────────

def fill_document(data: dict) -> Document:
    """
    data: 구글 시트 1행 데이터 dict
    반환: 채워진 Document 객체
    """
    doc = Document(TEMPLATE_PATH)
    table = doc.tables[0]
    rows = {i: _unique_cells(row) for i, row in enumerate(table.rows)}

    def v(key, default=''):
        val = data.get(key, default)
        if val is None:
            return default
        return val

    def is_true(key):
        val = data.get(key, False)
        if isinstance(val, bool):
            return val
        if isinstance(val, str):
            return val.strip().upper() in ('TRUE', '참', '1', 'Y', 'YES')
        return bool(val)

    # ── 행01: 성명 / 입소일 / 생년월일 / 성별 ──
    _set_text(rows[1][1], v('성명'), center=True)
    _set_text(rows[1][3], v('입소일'))
    _set_text(rows[1][5], v('생년월일'))
    _set_text(rows[1][7], v('성별'))

    # ── 행02: 작성일 / 작성자(영양사) ──
    작성일 = v('작성일') or datetime.date.today().strftime('%Y-%m-%d')
    _set_text(rows[2][1], 작성일)
    영양사 = v('영양사이름')
    if 영양사:
        _set_text(rows[2][3], f'{영양사} 영양사')

    # ── 행04: 신장 / 평소체중 / 등급 / 식사유형 ──
    _set_unit(rows[4][1], v('신장'), 'cm')
    _set_unit(rows[4][3], v('평소체중'), 'kg')
    _set_text(rows[4][5], v('등급'))
    _set_text(rows[4][7], v('식사유형'))

    # ── 행06: 1일필요열량 / 1일필요단백질 ──
    _set_unit(rows[6][1], v('1일필요열량'), 'kcal')
    _set_unit(rows[6][3], v('1일필요단백질'), 'g')

    # ── 행08: 식사방법 / 식사섭취상태 ──
    식사방법 = v('식사방법')
    if 식사방법:
        _check(rows[8][1], [식사방법])

    식사섭취상태 = v('식사섭취상태')
    if 식사섭취상태:
        _check(rows[8][3], [식사섭취상태])

    # ── 행09: 식사속도 / 도구사용 (다중) ──
    식사속도 = v('식사속도')
    if 식사속도:
        _check(rows[9][1], [식사속도])

    도구 = []
    if is_true('도구_젓가락'):     도구.append('젓가락')
    if is_true('도구_숟가락'):     도구.append('숟가락')
    if is_true('도구_포크숟가락'): 도구.append('포크숟가락')
    if is_true('도구_불가'):       도구.append('불가')
    if 도구:
        _check(rows[9][3], 도구)

    # ── 행10: 식사시 문제점 (다중) ──
    문제 = []
    if is_true('문제_식욕저하'): 문제.append('식욕저하')
    if is_true('문제_저작곤란'): 문제.append('저작곤란')
    if is_true('문제_연하곤란'): 문제.append('연하곤란')
    if is_true('문제_소화불량'): 문제.append('소화불량')
    if is_true('문제_구토'):     문제.append('구토')
    if is_true('문제_없음'):     문제.append('없음')
    if 문제:
        _check(rows[10][1], 문제)

    # ── 행12: 치아상태 ──
    치아 = v('치아상태')
    if 치아:
        _check(rows[12][1], [치아])

    # ── 행13: 소화기능 ──
    소화 = v('소화기능')
    if 소화:
        _check(rows[13][1], [소화])

    # ── 행14: 배설양상 ──
    배설 = v('배설양상')
    if 배설:
        _check(rows[14][1], [배설])

    # ── 행15: 특이체질 ──
    특이 = v('특이체질내용')
    if 특이:
        _check_with_content(rows[15][1], '있음', 특이)
    else:
        _check(rows[15][1], ['없음'])

    # ── 행17: 선호음식 / 비선호음식 ──
    _set_text(rows[17][1], v('선호음식', '없음') or '없음')
    _set_text(rows[17][3], v('비선호음식', '없음') or '없음')

    # ── 행18: 식품알러지 ──
    알러지 = v('식품알러지내용')
    if 알러지:
        _check_with_content(rows[18][1], '있음', 알러지)
    else:
        _check(rows[18][1], ['없음'])

    # ── 행20: 주요진단명 ──
    _set_text(rows[20][1], v('주요진단명'), left_align=True)

    # ── 행21: 주요질환 (다중) ──
    질환 = []
    if is_true('질환_당뇨'):       질환.append('당뇨')
    if is_true('질환_고혈압'):     질환.append('고혈압')
    if is_true('질환_심장질환'):   질환.append('심장질환')
    if is_true('질환_뇌혈관질환'): 질환.append('뇌혈관질환')
    if is_true('질환_신장질환'):   질환.append('신장질환')
    if is_true('질환_간질환'):     질환.append('간질환')
    if is_true('질환_암'):         질환.append('암')
    기타질환 = v('질환_기타내용')
    if 기타질환:
        질환.append('기타')
    if 질환:
        _check(rows[21][1], 질환)
    if 기타질환:
        para = rows[21][1].paragraphs[0]
        if para.runs:
            para.runs[0].text = re.sub(
                r'기타\(\s*\)', f'기타({기타질환})', para.runs[0].text
            )

    # ── 행22: 현재복용약물 ──
    약물 = v('복용약물내용')
    if 약물:
        _check_with_content(rows[22][1], '있음', 약물)
    else:
        _check(rows[22][1], ['없음'])

    # ── 행23: 영양관련약물영향 (다중) ──
    약물영향 = []
    if is_true('약물영향_식욕저하'): 약물영향.append('식욕저하')
    if is_true('약물영향_구역구토'): 약물영향.append('구역/구토')
    if is_true('약물영향_미각변화'): 약물영향.append('미각변화')
    if is_true('약물영향_흡수장애'): 약물영향.append('흡수장애')
    기타약물영향 = v('약물영향_기타내용')
    if 기타약물영향:
        약물영향.append('기타')
    if not 약물영향:
        약물영향.append('없음')
    _check(rows[23][1], 약물영향)
    if 기타약물영향:
        para = rows[23][1].paragraphs[0]
        if para.runs:
            para.runs[0].text = re.sub(
                r'기타\(\s*', f'기타({기타약물영향}', para.runs[0].text
            )

    # ── 행25: 종교 / 금식일기도시간 ──
    종교 = v('종교', '없음') or '없음'
    _check(rows[25][1], [종교])
    _set_text(rows[25][3], v('금식일기도시간'))

    # ── 행26: 종교적식사제한 ──
    종교제한 = v('종교적식사제한', '없음') or '없음'
    종교제한내용 = v('종교제한내용')
    _check_with_content(rows[26][1], 종교제한, 종교제한내용)

    # ── 행27: 문화적식습관 ──
    문화 = v('문화적식습관내용')
    if 문화:
        _check_with_content(rows[27][1], '있음', 문화)
    else:
        _check(rows[27][1], ['없음'])

    # ── 행28: 출신지역특성 ──
    출신 = v('출신지역특성내용')
    if 출신:
        _check_with_content(rows[28][1], '있음', 출신)
    else:
        _check(rows[28][1], ['해당없음'])

    # ── 행30/31: 개별 욕구 ──
    _set_text(rows[30][2], v('수급자욕구'), left_align=True)
    _set_text(rows[31][2], v('보호자욕구'), left_align=True)

    # ── 행33: 영양사 총평 ──
    _set_text(rows[33][1], v('영양사총평'), left_align=True)

    return doc


# ─────────────────────────────────────────
# 저장 / 출력
# ─────────────────────────────────────────

def save_document(doc: Document, name: str) -> str:
    """output/ 폴더에 저장 후 경로 반환"""
    # 파일 이름에 사용 불가 문자 제거
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', name)
    path = os.path.join(OUTPUT_DIR, f'{safe_name}.docx')
    doc.save(path)
    return path


def open_for_preview(path: str):
    """첫 번째 문서: Word로 열어서 미리보기"""
    import subprocess
    subprocess.Popen(['cmd', '/c', 'start', '', path], shell=False)


def print_document(path: str):
    """자동 프린트 (Windows 기본 프린터)"""
    import subprocess
    subprocess.Popen(
        ['cmd', '/c', 'start', '/MIN', '', '/PRINT', path],
        shell=False
    )


def generate_all(records: list) -> list:
    """
    모든 입소자 docx 생성 후 경로 목록 반환
    records: get_all_records() 결과
    """
    paths = []
    for rec in records:
        name = rec.get('성명', '미입력')
        doc = fill_document(rec)
        path = save_document(doc, name)
        paths.append(path)
    return paths


# ─────────────────────────────────────────
# ZIP + 합본 PDF 생성
# ─────────────────────────────────────────

def build_zip(paths: list) -> bytes:
    """개인별 docx 파일들을 ZIP으로 묶어 bytes 반환"""
    import io
    import zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for path in paths:
            zf.write(path, os.path.basename(path))
    return buf.getvalue()


def _build_pdf_windows(paths: list, tmp_dir: str) -> list:
    """Windows: Word COM(docx2pdf)으로 PDF 변환, PDF 경로 목록 반환"""
    import pythoncom
    from docx2pdf import convert

    pythoncom.CoInitialize()
    try:
        pdf_paths = []
        for docx_path in paths:
            pdf_path = os.path.join(
                tmp_dir,
                os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
            )
            convert(docx_path, pdf_path)
            pdf_paths.append(pdf_path)
        return pdf_paths
    finally:
        pythoncom.CoUninitialize()


def _build_pdf_libreoffice(paths: list, tmp_dir: str) -> list:
    """Linux/클라우드: LibreOffice headless로 PDF 변환, PDF 경로 목록 반환"""
    import subprocess

    pdf_paths = []
    for docx_path in paths:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf',
             '--outdir', tmp_dir, docx_path],
            check=True,
        )
        pdf_path = os.path.join(
            tmp_dir,
            os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
        )
        pdf_paths.append(pdf_path)
    return pdf_paths


def build_merged_pdf(paths: list) -> bytes:
    """
    개인별 docx → PDF 변환 후 하나로 합쳐 bytes 반환
    Windows: Word COM(docx2pdf) 사용
    Linux/클라우드: LibreOffice headless 사용
    """
    import io
    from pypdf import PdfWriter, PdfReader

    tmp_dir = tempfile.mkdtemp()
    try:
        if platform.system() == 'Windows':
            pdf_paths = _build_pdf_windows(paths, tmp_dir)
        else:
            pdf_paths = _build_pdf_libreoffice(paths, tmp_dir)

        writer = PdfWriter()
        for pdf_path in pdf_paths:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)

        buf = io.BytesIO()
        writer.write(buf)
        return buf.getvalue()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
