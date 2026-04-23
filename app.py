"""
영양사정기록지 자동 출력 시스템
Streamlit 앱

실행: streamlit run app.py
"""
import json
import os
import platform
import time

import streamlit as st

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

st.set_page_config(
    page_title='영양사정기록지 자동 출력',
    page_icon='📋',
    layout='wide',
)

# secrets에서만 설정 읽기
dietitian = st.secrets.get('dietitian', '')
sheet_id = st.secrets.get('sheet_id', '')
sheet_name = st.secrets.get('sheet_name', '입소자목록')

# config 대체용 dict
config = {
    'dietitian': dietitian,
    'sheet_id': sheet_id,
    'sheet_name': sheet_name,
}

# ─────────────────────────────────────────
# 사이드바: 설정(읽기 전용)
# ─────────────────────────────────────────
with st.sidebar:
    st.title('⚙️ 설정')

    st.text_input(
        '영양사 이름',
        value=dietitian,
        disabled=True,
        help='Streamlit secrets에서 읽어옵니다',
    )

    st.text_input(
        '구글 시트 ID',
        value=sheet_id,
        disabled=True,
        help='Streamlit secrets에서 읽어옵니다',
    )

    st.text_input(
        '시트 탭 이름',
        value=sheet_name,
        disabled=True,
        help='Streamlit secrets에서 읽어옵니다',
    )

    st.caption('설정값은 .streamlit/secrets.toml 에서 관리합니다.')
    st.divider()

    # 연결 테스트
    if st.button('🔗 연결 테스트'):
        try:
            import sheets
            cfg = {
                'dietitian': dietitian,
                'sheet_id': sheet_id,
                'sheet_name': sheet_name,
            }
            doc_title, tab_title = sheets.test_connection(cfg)
            st.success(
                f'연결 성공!\n'
                f'· 파일 제목(상단): {doc_title}\n'
                f'· 읽은 탭(하단): {tab_title}'
            )
        except Exception as e:
            st.error(f'연결 실패: {e}')

    st.divider()
    st.caption('📂 출력 파일 위치')
    st.code(os.path.join(BASE_DIR, 'output'))


# ─────────────────────────────────────────
# 메인 화면
# ─────────────────────────────────────────
st.title('📋 영양사정기록지 자동 출력 시스템')

# 데이터 불러오기
col1, col2 = st.columns([1, 4])
with col1:
    load_btn = st.button('🔄 데이터 불러오기', use_container_width=True)

if load_btn or 'records' in st.session_state:
    if load_btn:
        try:
            import sheets
            cfg = {
                'dietitian': dietitian,
                'sheet_id': sheet_id,
                'sheet_name': sheet_name,
            }
            with st.spinner('구글 시트에서 데이터 가져오는 중...'):
                records = sheets.get_all_records(cfg)
            st.session_state['records'] = records
            st.success(f'총 {len(records)}명 데이터 불러옴')
        # except Exception as e:
        #     st.error(f'데이터 불러오기 실패: {e}')
        #     st.stop()

        except Exception as e:
            import traceback
            st.error(f'데이터 불러오기 실패: {type(e).__name__}: {repr(e)}')
            st.code(traceback.format_exc())
            st.stop()



    records = st.session_state.get('records', [])

    if not records:
        st.warning('시트에 데이터가 없습니다.')
        st.stop()

    # 데이터 미리보기 테이블
    st.subheader(f'입소자 목록 ({len(records)}명)')
    preview_cols = ['성명', '입소일', '생년월일', '성별', '신장', '평소체중', '등급', '식사유형', '영양사이름']
    preview_data = [
        {k: r.get(k, '') for k in preview_cols if k in r}
        for r in records
    ]
    st.dataframe(preview_data, use_container_width=True, hide_index=True)

    st.divider()

    # 출력 버튼
    st.subheader('📄 파일 생성 및 출력')
    st.info(
        '**생성 결과물**\n'
        '- 📦 **개인별 ZIP** — 각 입소자 .docx 파일 모음 (기록 보관용)\n'
        '- 📄 **합본 PDF** — 전원 한 파일. 열어서 **Ctrl+P** 한 번으로 전체 인쇄'
    )

    col_a, col_b = st.columns([1, 3])
    with col_a:
        generate_btn = st.button('⚙️ 파일 생성', type='primary', use_container_width=True)

    if generate_btn:
        import filler

        override_dietitian = dietitian.strip()

        # ── 1단계: 개인별 docx 생성 ──
        progress = st.progress(0, text='문서 생성 중...')
        status = st.empty()
        total = len(records)
        paths = []

        for i, rec in enumerate(records):
            name = rec.get('성명', f'{i+1}번')
            status.text(f'생성 중: {name} ({i+1}/{total})')
            try:
                if override_dietitian:
                    rec = dict(rec)
                    rec['영양사이름'] = override_dietitian
                doc = filler.fill_document(rec, config)
                path = filler.save_document(doc, name)
                paths.append(path)
            except Exception as e:
                st.error(f'{name} 문서 생성 실패: {e}')
            progress.progress((i + 1) / total)

        status.text('')
        st.success(f'✅ {len(paths)}명 문서 생성 완료')

        if not paths:
            st.stop()

        # ── 2단계: ZIP 생성 → session_state 저장 ──
        with st.spinner('ZIP 파일 묶는 중...'):
            st.session_state['zip_bytes'] = filler.build_zip(paths)
            st.session_state['zip_paths'] = paths

        # ── 3단계: 합본 PDF 생성 → session_state 저장 ──
        _pdf_hint = (
            '합본 PDF 변환 중... (Word가 잠깐 실행됩니다)'
            if platform.system() == 'Windows'
            else '합본 PDF 변환 중... (LibreOffice)'
        )
        with st.spinner(_pdf_hint):
            try:
                st.session_state['pdf_bytes'] = filler.build_merged_pdf(paths)
                st.session_state['pdf_ok'] = True
            except Exception as e:
                _pdf_req = (
                    'Microsoft Word가 설치되어 있어야 합니다.'
                    if platform.system() == 'Windows'
                    else '서버에 LibreOffice가 설치되어 있어야 합니다. (Streamlit Cloud는 저장소 루트의 packages.txt 참고)'
                )
                st.warning(f'PDF 변환 실패: {e}\n\n{_pdf_req}')
                st.session_state['pdf_ok'] = False

    # ── 다운로드 버튼: session_state에 데이터 있으면 항상 표시 ──
    if 'zip_bytes' in st.session_state:
        import filler
        st.divider()
        st.subheader('⬇️ 다운로드')
        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label='📦 개인별 ZIP 다운로드',
                data=st.session_state['zip_bytes'],
                file_name='영양사정기록지_개인별.zip',
                mime='application/zip',
                use_container_width=True,
            )
            st.caption('각 입소자 .docx 파일 묶음')

        with col2:
            if st.session_state.get('pdf_ok'):
                st.download_button(
                    label='📄 합본 PDF 다운로드 (인쇄용)',
                    data=st.session_state['pdf_bytes'],
                    file_name='영양사정기록지_합본.pdf',
                    mime='application/pdf',
                    use_container_width=True,
                )
                st.caption('열어서 Ctrl+P → 전체 인쇄')
            else:
                st.button('📄 합본 PDF (변환 실패)', disabled=True, use_container_width=True)

        # ── 생성된 파일 목록 ──
        with st.expander('생성된 파일 목록 보기'):
            for path in st.session_state.get('zip_paths', []):
                st.write(f'- `{path}`')

else:
    st.info('왼쪽 상단 **데이터 불러오기** 버튼을 눌러 구글 시트에서 데이터를 가져오세요.')

    st.divider()
    st.subheader('📝 구글 시트 헤더 목록')
    st.write('시트 1행에 아래 헤더를 정확히 입력하세요:')

    headers = [
        '성명', '입소일', '생년월일', '성별', '작성일', '영양사이름', '신장', '평소체중', '등급', '식사유형',
        '1일필요열량', '1일필요단백질', '식사방법_자립식사', '식사방법_부분도움', '식사방법_완전도움', 
        '식사섭취상태_양호', '식사섭취상태_보통', '식사섭취상태_불량', 
        '식사속도_양호', '식사속도_보통', '식사속도_불량',
        '도구_젓가락', '도구_숟가락', '도구_포크숟가락', '도구_불가',
        '문제_식욕저하', '문제_저작곤란', '문제_연하곤란', '문제_소화불량', '문제_구토', '문제_없음',
        '치아상태', '소화기능', '배설양상', '특이체질_없음', '특이체질_있음', '특이체질내용', 
        '선호음식', '비선호음식', '식품알러지_없음', '식품알러지_있음', '식품알러지내용',
        '주요진단명', '질환_당뇨', '질환_고혈압', '질환_뇌혈관질환',
        '질환_신경질환', '질환_치매', '질환_암', '질환_기타', '질환_기타내용',
        '현재복용약물_없음', '현재복용약물_있음', '복용약물내용',
        '약물영향_없음', '약물영향_식욕저하', '약물영향_구역구토',
        '약물영향_흡수장애', '약물영향_기타', '약물영향_기타내용',
        '종교', '종교_기타내용', '금식일기도시간', '종교적식사제한', '종교제한내용',
        '문화적식습관', '문화적식습관내용', '출신지역국가특성', '출신지역특성내용',
        '수급자욕구', '보호자욕구', '영양사총평',
        '식사사진첨부', '식사사진첨부2',
    ]


    # 탭으로 구분하여 복사하기 쉽게 출력
    st.code('\t'.join(headers), language=None)
    st.caption('위 텍스트를 복사해서 구글 시트 A1 셀에 붙여넣으면 헤더가 한 번에 입력됩니다.')
