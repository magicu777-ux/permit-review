# =========================================================
# 체험·시험용 가상 서류를 만든다(실제 회사·인물과 무관한 샘플).
# 작업: 통신맨홀(밀폐공간) 광케이블 접속 + 고소작업차(스카이)로 전주 인입선 정리
# 일부러 빠뜨린 부분(검토에서 걸려야 정상):
#   - 작업계획서: 밀폐공간 감시인, 신호수 지정 없음
#   - 특별교육: 일용근로자 2시간인데 채용 시 교육 1시간 서류 없음, 이수자 명단에서 최도윤 빠짐
#   - 위험성평가: 참여자 명단에서 최도윤 빠짐
# 사용법: python make_samples.py   (엑셀은 openpyxl, PDF는 Chrome/Edge 인쇄로 만든다)
# =========================================================

import os
import base64
import shutil
import struct
import subprocess
import tempfile
import zlib

from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
CHROME_CANDIDATES = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
]
WORKERS = ['김민수', '이준호', '박서연', '최도윤']
STYLE = """<style>body{font-family:'Malgun Gothic',sans-serif;font-size:12pt;margin:24px} h1{font-size:18pt;text-align:center}
table{width:100%;border-collapse:collapse;margin:8px 0} td,th{border:1px solid #333;padding:6px;vertical-align:top}
th{background:#eee} .page{page-break-before:always}</style>"""


def png_data_uri(width=240, height=150, color=(120, 160, 200)):
    """교육 사진 대신 쓰는 단색 PNG(PDF 안에 이미지로 들어간다)."""
    raw = b''.join(b'\x00' + bytes(color) * width for _ in range(height))

    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xffffffff)
    png = (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
           + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))
    return 'data:image/png;base64,' + base64.b64encode(png).decode()


def make_permit():
    wb = Workbook()
    ws = wb.active
    ws.title = '작업허가서'
    rows = [
        ['작업허가서 (샘플·가상)'],
        ['작업명', '통신맨홀 내부 광케이블 접속 및 전주 인입선 정리 작업'],
        ['작업장소', '○○시 ○○구 ○○로 통신맨홀 M-123 및 인근 전주 #45'],
        ['작업일시', '2026-11-10 09:00 ~ 18:00'],
        ['협력사', '(가상) 한빛통신'],
        ['허가 작업 구분', '밀폐공간 작업(맨홀), 고소작업(전주, 고소작업차(스카이))'],
        [],
        ['순번', '성명', '직종', '연락처'],
    ]
    rows += [[i + 1, name, '통신공', '010-0000-000' + str(i)] for i, name in enumerate(WORKERS)]
    rows += [[], ['작업책임자', '김민수'], ['허가자', '(서명)']]
    for row in rows:
        ws.append(row)
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 50
    wb.save(os.path.join(HERE, '작업허가서_샘플.xlsx'))


def make_risk():
    wb = Workbook()
    ws = wb.active
    ws.title = '위험성평가'
    rows = [
        ['위험성평가표 (샘플·가상)'],
        ['평가일', '2026-11-05', '작업명', '통신맨홀 광케이블 접속 및 전주 인입선 정리'],
        [],
        ['공정', '유해·위험요인', '위험도', '감소대책', '담당자', '조치일'],
        ['맨홀 진입', '밀폐공간 산소결핍·유해가스로 인한 질식', '상', '진입 전 산소·유해가스 측정, 강제환기, 송기마스크 비치', '이준호', '2026-11-10'],
        ['맨홀 내부 접속', '맨홀 뚜껑 개방부 추락', '중', '개방부 안전펜스·라바콘 설치', '박서연', '2026-11-10'],
        ['전주 인입선 정리', '고소작업차(스카이) 버킷에서 추락', '상', '안전대 착용, 버킷 내 안전고리 체결', '최도윤', '2026-11-10'],
        ['전주 인입선 정리', '고소작업차 전도·협착', '중', '아웃트리거 설치, 작업반경 출입 통제', '김민수', '2026-11-10'],
        [],
        ['근로자 참여자', '성명'],
        ['', '김민수'], ['', '이준호'], ['', '박서연'],
    ]
    for row in rows:
        ws.append(row)
    for col, width in zip('ABCDEF', (16, 36, 8, 44, 10, 12)):
        ws.column_dimensions[col].width = width
    wb.save(os.path.join(HERE, '위험성평가_샘플.xlsx'))


PLAN_HTML = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">{STYLE}</head><body>
<h1>작업계획서 (샘플·가상)</h1>
<table>
<tr><th>작업명</th><td>통신맨홀 내부 광케이블 접속 및 전주 인입선 정리 작업</td></tr>
<tr><th>작업장소</th><td>○○시 ○○구 ○○로 통신맨홀 M-123, 전주 #45</td></tr>
<tr><th>작업기간</th><td>2026-11-10 ~ 2026-11-11</td></tr>
<tr><th>투입 인원</th><td>4명 (김민수, 이준호, 박서연, 최도윤)</td></tr>
<tr><th>투입 장비</th><td>고소작업차(스카이) 1대 — 차량탑재형 고소작업대, 작업높이 12m, 정격하중 200kg / 송풍기, 산소·유해가스 측정기</td></tr>
</table>
<h3>작업 방법 및 순서</h3>
<p>1. 작업구역 출입 통제 및 맨홀 개방부 안전펜스 설치<br>
2. 맨홀 진입 전 산소 및 유해가스(황화수소, 일산화탄소) 농도 측정<br>
3. 송풍기로 강제 환기 후 진입, 작업 중 연속 측정<br>
4. 광케이블 접속 후 퇴출, 맨홀 복구<br>
5. 고소작업차 아웃트리거 설치 후 전주 인입선 정리 (안전대 착용, 안전고리 체결)</p>
<h3>위험요인별 안전대책</h3>
<p>질식: 산소농도 18% 이상 확인, 강제 환기, 송기마스크 비치<br>
추락: 안전대 착용, 하부 출입 통제, 낙하물 방지<br>
전도: 아웃트리거 설치, 작업반경 통제</p>
<h3>비상조치</h3>
<p>비상연락망: 현장 책임자 → 협력사 안전관리자 → 119. 응급 시 구조 후 즉시 119 신고.</p>
<h3>작업계획도</h3>
<p>(도면) 맨홀 M-123 위치, 고소작업차 진입경로, 안전펜스·라바콘 위치, 작업지휘자 위치 표시</p>
<table>
<tr><th>작성자</th><td>박서연 (서명)</td><th>검토자</th><td>김민수 (서명)</td></tr>
<tr><th>협력사 안전보건관리책임자</th><td>(가상) 한빛통신 대표 (서명)</td><th>도급사 안전관리총괄책임자</th><td>(서명)</td></tr>
</table>
<div class="page">
<h1>작업지휘자 지정서</h1>
<table>
<tr><th>작업명</th><td>통신맨홀 내부 광케이블 접속 및 전주 인입선 정리 작업</td></tr>
<tr><th>작업지휘자</th><td>김민수 (현장 책임자)</td></tr>
<tr><th>지정일</th><td>2026-11-05</td></tr>
</table>
<p>위 사람을 작업지휘자로 지정합니다. 지정자: (가상) 한빛통신 대표 (서명) / 작업지휘자: 김민수 (서명)</p>
</div>
</body></html>"""

EDU_HTML = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">{STYLE}</head><body>
<h1>특별안전보건교육 실시 결과 (샘플·가상)</h1>
<table>
<tr><th>교육일시</th><td>2026-11-09 13:00 ~ 15:00</td><th>교육시간</th><td>2시간</td></tr>
<tr><th>교육 대상</th><td colspan="3">일용근로자 — 34. 밀폐공간에서의 작업(통신맨홀)</td></tr>
<tr><th>교육 내용</th><td colspan="3">산소농도 측정 및 작업환경 평가, 환기 방법, 보호구(송기마스크) 착용, 비상시 구출 및 응급조치</td></tr>
<tr><th>강사</th><td colspan="3">(가상) 안전관리자 정하늘</td></tr>
</table>
<h3>이수자 명단</h3>
<table>
<tr><th>순번</th><th>성명</th><th>소속</th><th>서명</th></tr>
<tr><td>1</td><td>김민수</td><td>한빛통신</td><td>(서명)</td></tr>
<tr><td>2</td><td>이준호</td><td>한빛통신</td><td>(서명)</td></tr>
<tr><td>3</td><td>박서연</td><td>한빛통신</td><td>(서명)</td></tr>
</table>
<p>위와 같이 교육을 실시하였음을 확인합니다.</p>
<h3>교육 사진</h3>
<img src="{png_data_uri()}" width="240"> <img src="{png_data_uri(color=(160, 190, 140))}" width="240">
</body></html>"""

LICENSE_HTML = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">{STYLE}</head><body>
<h1>차량탑재형 고소작업대 조종교육 이수증 (샘플·가상)</h1>
<table><tr><th>성명</th><td>최도윤</td></tr><tr><th>교육과정</th><td>차량탑재형 고소작업대 조종교육</td></tr>
<tr><th>이수일</th><td>2026-03-15</td></tr></table>
</body></html>"""


PERMIT_SCREEN_HTML = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">{STYLE}
<style>body{{background:#fff;font-size:15px}} .bar{{background:#1f4f8f;color:#fff;padding:8px 12px;font-weight:bold}}</style></head><body>
<div class="bar">안전보건시스템 — 작업허가서 조회 (샘플 화면·가상)</div>
<table>
<tr><th>작업명</th><td>통신맨홀 내부 광케이블 접속 및 전주 인입선 정리 작업</td></tr>
<tr><th>작업장소</th><td>○○시 ○○구 ○○로 통신맨홀 M-123 및 인근 전주 #45</td></tr>
<tr><th>작업일시</th><td>2026-11-10 09:00 ~ 18:00</td></tr>
<tr><th>허가 작업 구분</th><td>밀폐공간 작업(맨홀), 고소작업(전주, 고소작업차(스카이))</td></tr>
<tr><th>작업책임자</th><td>김민수</td></tr>
</table>
<table><tr><th>순번</th><th>성명</th><th>직종</th></tr>
{''.join(f'<tr><td>{i + 1}</td><td>{n}</td><td>통신공</td></tr>' for i, n in enumerate(WORKERS))}
</table></body></html>"""


def html_to_png(html_text, png_name, size='900,620'):
    """Safe_b 화면 캡처를 흉내 낸 그림(글자 인식 시험용)."""
    chrome = next(p for p in CHROME_CANDIDATES if os.path.exists(p))
    work = tempfile.mkdtemp()
    try:
        page = os.path.join(work, 'page.html')
        with open(page, 'w', encoding='utf-8') as f:
            f.write(html_text)
        out = os.path.join(work, 'out.png')
        subprocess.run([chrome, '--headless=new', '--disable-gpu', '--hide-scrollbars', '--virtual-time-budget=3000',
                        f'--window-size={size}', f'--user-data-dir={os.path.join(work, "profile")}', f'--screenshot={out}',
                        'file:///' + page.replace('\\', '/')],
                       stdin=subprocess.DEVNULL, capture_output=True, timeout=120)
        shutil.copyfile(out, os.path.join(HERE, png_name))
    finally:
        shutil.rmtree(work, ignore_errors=True)


def html_to_pdf(html_text, pdf_name):
    chrome = next(p for p in CHROME_CANDIDATES if os.path.exists(p))
    work = tempfile.mkdtemp()
    try:
        page = os.path.join(work, 'page.html')
        with open(page, 'w', encoding='utf-8') as f:
            f.write(html_text)
        out = os.path.join(work, 'out.pdf')
        # 시간 예산을 주지 않으면 이미지가 든 페이지에서 인쇄가 멈추는 경우가 있다
        subprocess.run([chrome, '--headless=new', '--disable-gpu', '--no-pdf-header-footer', '--virtual-time-budget=5000',
                        f'--user-data-dir={os.path.join(work, "profile")}', f'--print-to-pdf={out}',
                        'file:///' + page.replace('\\', '/')],
                       stdin=subprocess.DEVNULL, capture_output=True, timeout=120)
        shutil.copyfile(out, os.path.join(HERE, pdf_name))
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == '__main__':
    old = os.path.join(HERE, '특별안전교육_샘플.pdf')
    if os.path.exists(old):
        os.remove(old)
    make_permit()
    html_to_png(PERMIT_SCREEN_HTML, '작업허가서_캡처_샘플.png')
    make_risk()
    html_to_pdf(PLAN_HTML, '작업계획서_샘플.pdf')
    html_to_pdf(EDU_HTML, '특별안전교육_밀폐공간_샘플.pdf')
    html_to_pdf(LICENSE_HTML, '고소작업대_조종교육_이수증_샘플.pdf')
    print('샘플 생성:', ', '.join(sorted(n for n in os.listdir(HERE) if n.endswith(('.xlsx', '.pdf')))))
