# 시험용 두 번째 장면(화면 주소 끝에 #demo2): 분전반 차단기 교체(380V) + 브라켓 용접.
# Expected: 전기도면 없음(fail), 교육일지 1개로 전기·화기 합침(fail), 화기감시자가 교육 이수자 아님(fail), 16시간(ok)
#           Safe_B 화면: High Risk '예'(fail), 작업구분 전기작업인데 '전기작업 계획서' 제목의 계획서 없음(fail)
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_samples as M
from openpyxl import Workbook

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'scenario2')
os.makedirs(OUT, exist_ok=True)
M.HERE = OUT

wb = Workbook(); ws = wb.active
for row in [['작업허가서'], ['작업명', '지하 1층 공용 분전반 차단기 교체 및 케이블 트레이 브라켓 용접'], ['작업일시', '2026-11-20'],
            ['허가 작업 구분', '전기작업(380V 분전반 정전작업), 화기작업(전기용접)'], [], ['순번', '성명'], [1, '한지훈'], [2, '오세린'], [3, '정우진']]:
    ws.append(row)
wb.save(os.path.join(OUT, '작업허가서.xlsx'))

M.html_to_png(M.permit_screen_html(
    {'화기작업', '전기작업'},
    {'작업명': '지하 1층 공용 분전반 차단기 교체 및 케이블 트레이 브라켓 용접', '작업시작': '2026-11-20 09:00', '작업종료': '2026-11-20 17:00',
     '사업구분': '유선미디어', '관리팀': '유선미디어안전보건팀', '도급사': '(가상) 새빛전기', '작업장소 구분': '옥내', '작업장소(사옥)': '○○사옥',
     '작업장소(세부)': '지하 1층 전기실 공용 분전반', '작업내용': '분전반 차단기 교체(380V 정전작업), 케이블 트레이 브라켓 용접',
     '작업책임자': '한지훈', '작업인원': '3명'},
    ['한지훈', '오세린', '정우진'], high_risk_no=False), '작업허가서_캡처.png')

plan = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">{M.STYLE}</head><body><h1>작업계획서</h1>
<table><tr><th>작업명</th><td>지하 1층 공용 분전반 차단기 교체 및 브라켓 용접</td></tr>
<tr><th>작업기간</th><td>2026-11-20</td></tr><tr><th>투입 인원</th><td>3명 (한지훈, 오세린, 정우진)</td></tr></table>
<h3>작업 방법 및 순서</h3><p>1. 분전반 주차단기 개방, 잠금장치·꼬리표 부착(LOTO) 2. 검전기로 검전 후 접지 3. 차단기 교체 4. 용접기로 브라켓 용접(방화포 설치, 소화기 비치)</p>
<h3>안전대책</h3><p>절연장갑 착용, 불티 비산방지 방화포 설치</p><h3>비상조치</h3><p>비상연락 119</p>
<table><tr><th>작성자</th><td>오세린</td><th>검토자</th><td>한지훈</td></tr>
<tr><th>협력사 안전보건관리책임자</th><td>(서명)</td><th>도급사 안전관리총괄책임자</th><td>(서명)</td></tr></table>
<div class="page"><h1>작업지휘자 지정서</h1><p>작업지휘자: 한지훈 (서명) — 위 사람을 작업지휘자로 지정함</p></div></body></html>"""
M.html_to_pdf(plan, '작업계획서.pdf')

edu = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">{M.STYLE}</head><body><h1>특별안전보건교육 일지</h1>
<table><tr><th>교육일</th><td>2026-11-17</td><th>교육시간</th><td>16시간</td></tr>
<tr><th>교육 대상 작업</th><td colspan="3">17. 전압 50볼트 이상 정전 및 활선작업 / 38. 화재위험작업(용접)</td></tr>
<tr><th>교육 내용</th><td colspan="3">전기 위험성, 정전 절차, 화재위험작업 시 준수사항</td></tr><tr><th>강사</th><td colspan="3">안전관리자</td></tr></table>
<h3>이수자 명단</h3><table><tr><th>성명</th><th>서명</th></tr><tr><td>한지훈</td><td>(서명)</td></tr><tr><td>오세린</td><td>(서명)</td></tr></table>
<img src="{M.png_data_uri()}" width="200"></body></html>"""
M.html_to_pdf(edu, '특별교육_전기화기.pdf')

watcher = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">{M.STYLE}</head><body><h1>화기감시자 지정서</h1>
<p>화기감시자 성명: 정우진</p><p>위 사람을 화기감시자로 지정합니다.</p></body></html>"""
M.html_to_pdf(watcher, '화기감시자_지정서.pdf')
print(sorted(os.listdir(OUT)))
