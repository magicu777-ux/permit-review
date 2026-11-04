# =========================================================
# Word(docx)·PowerPoint(pptx) 읽기 시험용 가상 서류(화면 주소 끝에 #demo3)
# 장면 1(맨홀 + 스카이)과 같은 내용을 작업계획서는 docx, 특별교육은 pptx로 만든다 → 검토 결과가 장면 1과 같아야 정상
# 필요: pip install python-docx python-pptx
# =========================================================

import os
import io

from docx import Document
from docx.enum.text import WD_BREAK
from pptx import Presentation
from pptx.util import Inches, Pt

import make_samples as M

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'office')


def make_plan_docx():
    d = Document()
    d.add_heading('작업계획서 (샘플·가상)', 1)
    rows = [
        ('작업명', '통신맨홀 내부 광케이블 접속 및 전주 인입선 정리 작업'),
        ('작업장소', '○○시 ○○구 ○○로 통신맨홀 M-123, 전주 #45'),
        ('작업기간', '2026-11-10 ~ 2026-11-11'),
        ('투입 인원', '4명 (김민수, 이준호, 박서연, 최도윤)'),
        ('투입 장비', '고소작업차(스카이) 1대 — 차량탑재형 고소작업대, 정격하중 200kg / 송풍기, 산소·유해가스 측정기'),
    ]
    t = d.add_table(rows=0, cols=2)
    for k, v in rows:
        cells = t.add_row().cells
        cells[0].text, cells[1].text = k, v
    d.add_heading('작업 방법 및 순서', 2)
    for step in ['1. 작업구역 출입 통제 및 맨홀 개방부 안전펜스 설치',
                 '2. 맨홀 진입 전 산소 및 유해가스 농도 측정',
                 '3. 송풍기로 강제 환기 후 진입',
                 '4. 고소작업차 아웃트리거 설치 후 전주 인입선 정리 (안전대 착용)']:
        d.add_paragraph(step)
    d.add_heading('위험요인별 안전대책', 2)
    d.add_paragraph('질식: 산소농도 18% 이상 확인, 강제 환기, 송기마스크 비치 / 추락: 안전대 착용, 하부 출입 통제')
    d.add_heading('비상조치', 2)
    d.add_paragraph('비상연락망: 현장 책임자 → 협력사 안전관리자 → 119')
    d.add_heading('작업계획도', 2)
    d.add_paragraph('(도면) 맨홀 위치, 고소작업차 진입경로, 안전펜스 위치, 작업지휘자 위치')
    s = d.add_table(rows=2, cols=4)
    for (r, c), text in {(0, 0): '작성자', (0, 1): '박서연 (서명)', (0, 2): '검토자', (0, 3): '김민수 (서명)',
                         (1, 0): '협력사 안전보건관리책임자', (1, 1): '(서명)', (1, 2): '도급사 안전관리총괄책임자', (1, 3): '(서명)'}.items():
        s.cell(r, c).text = text
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_heading('작업지휘자 지정서', 1)
    d.add_paragraph('작업지휘자: 김민수 (현장 책임자) — 위 사람을 작업지휘자로 지정합니다. (서명)')
    d.save(os.path.join(OUT, '작업계획서.docx'))


def make_education_pptx():
    p = Presentation()
    blank = p.slide_layouts[6]
    s1 = p.slides.add_slide(blank)
    s1.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8)).text_frame.text = '특별안전보건교육 실시 결과 (샘플·가상)'
    info = s1.shapes.add_table(4, 2, Inches(0.5), Inches(1.2), Inches(9), Inches(2)).table
    for i, (k, v) in enumerate([('교육일시', '2026-11-09 13:00 ~ 15:00 (2시간)'), ('교육 대상', '일용근로자 — 34. 밀폐공간에서의 작업'),
                                ('교육 내용', '산소농도 측정, 환기, 송기마스크 착용, 비상시 구출'), ('강사', '(가상) 안전관리자 정하늘')]):
        info.cell(i, 0).text, info.cell(i, 1).text = k, v
    names = s1.shapes.add_table(4, 3, Inches(0.5), Inches(3.6), Inches(9), Inches(1.6)).table
    for i, row in enumerate([('순번', '성명', '서명'), ('1', '김민수', '(서명)'), ('2', '이준호', '(서명)'), ('3', '박서연', '(서명)')]):
        for j, v in enumerate(row):
            names.cell(i, j).text = v
    s2 = p.slides.add_slide(blank)
    s2.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8)).text_frame.text = '교육 사진'
    png = io.BytesIO(M.base64.b64decode(M.png_data_uri().split(',')[1]))
    s2.shapes.add_picture(png, Inches(0.5), Inches(1.2), Inches(4))
    p.save(os.path.join(OUT, '특별안전교육_밀폐공간.pptx'))


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    make_plan_docx()
    make_education_pptx()
    print(sorted(os.listdir(OUT)))
