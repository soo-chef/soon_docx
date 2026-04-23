"""템플릿 docx의 첫 표 tblGrid를 갱신합니다. Word에서 해당 파일을 닫은 뒤 실행하세요.

성명 행(10열) 목표(cm) — 이름값·생년월일값은 2열 균등 분할:
  성명 2.49 | 이름값 2.25 | 입소일 2 | 입소일값 2.5 | 생년월일 2.5 | 생년월일값 2.5 | 성별 2 | 성별값 2.5
"""
import io
import os
import re
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(BASE, '영양사정기록지_개정.docx')

NEW = (
    '<w:tblGrid><w:gridCol w:w="1412"/><w:gridCol w:w="638"/><w:gridCol w:w="638"/>'
    '<w:gridCol w:w="1134"/><w:gridCol w:w="1417"/><w:gridCol w:w="1417"/>'
    '<w:gridCol w:w="709"/><w:gridCol w:w="709"/><w:gridCol w:w="1134"/>'
    '<w:gridCol w:w="1417"/></w:tblGrid>'
)


def main():
    with zipfile.ZipFile(PATH, 'r') as z:
        t = z.read('word/document.xml').decode('utf-8')
    if NEW in t:
        print('이미 목표 tblGrid 입니다:', PATH)
        return
    m = re.search(r'<w:tblGrid>.*?</w:tblGrid>', t)
    if not m:
        raise SystemExit('document.xml 에 tblGrid 가 없습니다.')
    t2 = t.replace(m.group(0), NEW, 1)
    buf = io.BytesIO()
    with zipfile.ZipFile(PATH, 'r') as zin, zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                data = t2.encode('utf-8')
            zout.writestr(item, data)
    with open(PATH, 'wb') as f:
        f.write(buf.getvalue())
    print('갱신 완료:', PATH)


if __name__ == '__main__':
    main()
