# =========================================================
# 압축(zip) 읽기 시험용 묶음(화면 주소 끝에 #demo4)
# - 작업계획서_묶음.zip: 작업계획서.docx + 하위 폴더의 작업계획도 그림 + 읽을 수 없는 .hwp(건너뛰어야 정상)
# - 특별교육_묶음.zip: 사진 없는 교육자료 pptx + 따로 찍은 교육사진 2장 + 압축 안의 압축(이수자 확인서 pdf)
# 한글 파일 이름은 한국어 Windows 압축처럼 CP949로 저장한다(UTF-8 표시 없음) → 화면에서 이름이 깨지지 않아야 정상
# 필요: make_office_samples.py를 먼저 실행(office 폴더)
# =========================================================

import io
import os
import zipfile

from pptx import Presentation
from pptx.util import Inches

import make_samples as M

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'zip')


class Cp949Info(zipfile.ZipInfo):
    """이름을 CP949 바이트로, UTF-8 표시 없이 기록한다(한국어 Windows '압축 폴더'와 같은 방식)."""

    def _encodeFilenameFlags(self):
        return self.filename.encode('cp949'), self.flag_bits & ~0x800


def cp949_info(name):
    info = Cp949Info(name)
    info.compress_type = zipfile.ZIP_DEFLATED
    return info


def png_bytes(color):
    return M.base64.b64decode(M.png_data_uri(color=color).split(',')[1])


def education_pptx_without_photo():
    p = Presentation()
    s = p.slides.add_slide(p.slide_layouts[6])
    s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8)).text_frame.text = '특별안전보건교육 실시 결과 (샘플·가상)'
    t = s.shapes.add_table(8, 2, Inches(0.5), Inches(1.2), Inches(9), Inches(4)).table
    for i, (k, v) in enumerate([('교육일시', '2026-11-09 13:00 ~ 15:00 (2시간)'), ('교육 대상', '일용근로자 — 34. 밀폐공간에서의 작업'),
                                ('교육 내용', '산소농도 측정, 환기, 송기마스크 착용'), ('강사', '(가상) 안전관리자 정하늘'),
                                ('성명', '서명'), ('김민수', '(서명)'), ('이준호', '(서명)'), ('박서연', '(서명)')]):
        t.cell(i, 0).text, t.cell(i, 1).text = k, v
    buf = io.BytesIO()
    p.save(buf)
    return buf.getvalue()


def main():
    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(HERE, 'office', '작업계획서.docx'), 'rb') as f:
        plan_docx = f.read()
    with zipfile.ZipFile(os.path.join(OUT, '작업계획서_묶음.zip'), 'w') as z:
        z.writestr(cp949_info('작업계획서.docx'), plan_docx)
        z.writestr(cp949_info('도면/작업계획도.png'), png_bytes((200, 200, 200)))
        z.writestr(cp949_info('참고/회사양식.hwp'), b'HWP Document File')

    inner = io.BytesIO()
    with zipfile.ZipFile(inner, 'w') as z:
        with open(os.path.join(HERE, '특별안전교육_밀폐공간_샘플.pdf'), 'rb') as f:
            z.writestr(cp949_info('이수자_확인서.pdf'), f.read())
    with zipfile.ZipFile(os.path.join(OUT, '특별교육_묶음.zip'), 'w') as z:
        z.writestr(cp949_info('특별교육_밀폐공간.pptx'), education_pptx_without_photo())
        z.writestr(cp949_info('교육사진/사진1.png'), png_bytes((120, 160, 200)))
        z.writestr(cp949_info('교육사진/사진2.png'), png_bytes((160, 190, 140)))
        z.writestr(cp949_info('확인서.zip'), inner.getvalue())
    print(sorted(os.listdir(OUT)))


if __name__ == '__main__':
    main()
