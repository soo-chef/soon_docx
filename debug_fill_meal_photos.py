"""
상위 N명만 시트에서 읽어 식사사진 URL·DOCX 생성을 점검합니다.

  python debug_fill_meal_photos.py
  python debug_fill_meal_photos.py 2

config.json 이 있어야 합니다. 출력은 output/debug_<성명>.docx
"""
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


def main():
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 2
    import sheets
    import filler

    config = sheets.load_config()
    config['debug_person_limit'] = n
    records = sheets.get_all_records(config)
    print('--- records:', len(records), '명 ---')
    for rec in records:

        def vg(k, d='', _r=rec):
            v = _r.get(k, d)
            return d if v is None else v

        name = str(rec.get('성명', '이름없음')).strip() or '이름없음'
        print('\n성명:', name)
        p1 = filler._meal_photo_raw_from_record(rec, filler._MEAL_PHOTO_NAMES, vg)
        p2 = filler._meal_photo_raw_from_record(rec, filler._MEAL_PHOTO2_NAMES, vg)
        u1 = filler._parse_meal_photo_urls(p1)
        u2 = filler._parse_meal_photo_urls(p2)
        print('  raw 식사사진첨부 앞 80자:', repr(str(p1)[:80]))
        print('  raw 식사사진첨부2 앞 80자:', repr(str(p2)[:80]))
        print('  parse URLs 1:', u1)
        print('  parse URLs 2:', u2)
        doc = filler.fill_document(rec, config)
        safe = __import__('re').sub(r'[\\/:*?"<>|]', '_', name)
        out = os.path.join(filler.OUTPUT_DIR, f'debug_{safe}.docx')
        os.makedirs(filler.OUTPUT_DIR, exist_ok=True)
        doc.save(out)
        print('  saved:', out)


if __name__ == '__main__':
    main()
