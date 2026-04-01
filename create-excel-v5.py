"""
빌드어스 데이터 아키텍처 엑셀 v5
- v4 기반 + 견적서 전체 고유항목 동적 추가
- 16개 견적서 xlsx에서 중복 제거 후 누락 없이 반영
- 화장실공사 → 도기마감 시트에 통합
- 금속공사/데크공사 → 해당 공정에 배분
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import os, re, unicodedata, statistics
from collections import defaultdict

# ── 스타일 ──
H_FILL = PatternFill(start_color="2D3748", end_color="2D3748", fill_type="solid")
H_FONT = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
S_FILL = PatternFill(start_color="4A5568", end_color="4A5568", fill_type="solid")
S_FONT = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
EX_FILL = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
FM_FILL = PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid")
RS_FILL = PatternFill(start_color="F0FFF4", end_color="F0FFF4", fill_type="solid")
EP_FILL = PatternFill(start_color="FFFFF0", end_color="FFFFF0", fill_type="solid")
NF = Font(name="맑은 고딕", size=10)
BF = Font(name="맑은 고딕", bold=True, size=10)
TF = Font(name="맑은 고딕", bold=True, size=14)
SF = Font(name="맑은 고딕", bold=True, size=11, color="4299E1")
TB = Border(
    left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"),
    top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"),
)

MC = 28

HEADERS = [
    "자재코드", "품명", "규격", "용도", "필수/선택",
    "단위", "1㎡당\n소요량", "로스율\n(%)", "실소요량\n(1㎡)",
    "판매단위", "판매단위\n수량", "도매가", "업체가", "판매가",
    "쿠팡가", "네이버\n최저가", "절약금액",
    "무게(kg)", "배송방법", "배송일", "공급사코드",
    "샘플가능", "샘플크기", "샘플비용",
    "등급", "한줄요약",
    "견적 참고\n자재단가", "견적 참고\n노무단가",
]

WIDTHS = [10,16,14,14,7, 5,8,7,8, 10,8,9,9,9,9,10,9, 7,7,7,9, 7,7,7, 7,28, 11,11]

SECTIONS = [
    (1, 9, "기본 정보 + 소요량"),
    (10, 17, "가격 (유통 채움)"),
    (18, 21, "물류/배송"),
    (22, 24, "샘플"),
    (25, 26, "등급"),
    (27, 28, "견적 참고단가 (16개 견적서)"),
]


def sty_h(ws, row, mc):
    for c in range(1, mc+1):
        cl = ws.cell(row=row, column=c)
        cl.font = H_FONT; cl.fill = H_FILL
        cl.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cl.border = TB

def sty_r(ws, r, mc):
    for c in range(1, mc+1):
        cl = ws.cell(row=r, column=c)
        cl.font = NF; cl.border = TB
        cl.alignment = Alignment(vertical="center", wrap_text=True)
        cl.fill = EX_FILL


def create_mat_sheet(wb, name, color, pcode, pname, items, is_first=False):
    if is_first:
        ws = wb.active
        ws.title = name
    else:
        ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = color

    ws.merge_cells(f"A1:{get_column_letter(MC)}1")
    ws.cell(row=1, column=1, value=f"{pname} 자재").font = TF
    ws.merge_cells(f"A2:{get_column_letter(MC)}2")
    ws.cell(row=2, column=1,
            value=f"공정코드: {pcode} | 체크리스트 + 16개 견적서 전체 항목 | ★=견적서 추출").font = SF
    ws.row_dimensions[1].height = 30

    for s, e, label in SECTIONS:
        ws.merge_cells(start_row=3, start_column=s, end_row=3, end_column=e)
        cl = ws.cell(row=3, column=s, value=label)
        cl.font = S_FONT; cl.fill = S_FILL; cl.alignment = Alignment(horizontal="center")
        for c in range(s, e+1):
            ws.cell(row=3, column=c).fill = S_FILL
            ws.cell(row=3, column=c).border = TB

    for i, h in enumerate(HEADERS, 1):
        ws.cell(row=4, column=i, value=h)
    sty_h(ws, 4, MC)
    ws.row_dimensions[4].height = 36

    for i, item in enumerate(items):
        r = 5 + i
        code, nm, spec, purpose, req, unit, summary, est_mat, est_labor = item
        ws.cell(row=r, column=1, value=code)
        ws.cell(row=r, column=2, value=nm)
        ws.cell(row=r, column=3, value=spec)
        ws.cell(row=r, column=4, value=purpose)
        ws.cell(row=r, column=5, value=req)
        ws.cell(row=r, column=6, value=unit)
        ws.cell(row=r, column=9).value = f'=IF(G{r}="","",G{r}*(1+H{r}/100))'
        ws.cell(row=r, column=9).fill = FM_FILL
        ws.cell(row=r, column=9).number_format = "0.000"
        ws.cell(row=r, column=17).value = f'=IF(OR(O{r}="",P{r}=""),"",MIN(O{r},P{r})-N{r})'
        ws.cell(row=r, column=17).fill = FM_FILL
        ws.cell(row=r, column=17).number_format = "#,##0"
        ws.cell(row=r, column=26, value=summary)

        if est_mat > 0:
            ws.cell(row=r, column=27, value=est_mat)
            ws.cell(row=r, column=27).number_format = "#,##0"
            ws.cell(row=r, column=27).fill = EP_FILL
        if est_labor > 0:
            ws.cell(row=r, column=28, value=est_labor)
            ws.cell(row=r, column=28).number_format = "#,##0"
            ws.cell(row=r, column=28).fill = EP_FILL

        sty_r(ws, r, MC)
        for c in [12,13,14,15,16,27,28]:
            ws.cell(row=r, column=c).number_format = "#,##0"

        if code.endswith("E"):
            ws.cell(row=r, column=1).font = Font(name="맑은 고딕", bold=True, size=10, color="D69E2E")
            ws.cell(row=r, column=2).font = Font(name="맑은 고딕", bold=True, size=10, color="D69E2E")

    last = 5 + len(items)
    for r in range(last, last + 10):
        ws.cell(row=r, column=9).value = f'=IF(G{r}="","",G{r}*(1+H{r}/100))'
        ws.cell(row=r, column=9).fill = FM_FILL
        ws.cell(row=r, column=9).number_format = "0.000"
        ws.cell(row=r, column=17).value = f'=IF(OR(O{r}="",P{r}=""),"",MIN(O{r},P{r})-N{r})'
        ws.cell(row=r, column=17).fill = FM_FILL
        ws.cell(row=r, column=17).number_format = "#,##0"
        for c in range(1, MC+1):
            ws.cell(row=r, column=c).border = TB

    nr = last + 11
    cl = ws.cell(row=nr, column=1,
                 value="※ 파란셀=수식 | 노란셀=견적참고 | ★금색=견적서 추출항목 | 소요량·로스율=대표 | 가격=유통 | 견적단가=16개 견적서 중앙값")
    cl.font = Font(name="맑은 고딕", size=9, color="718096", italic=True)
    ws.merge_cells(start_row=nr, start_column=1, end_row=nr, end_column=MC)

    for i, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:{get_column_letter(MC)}{last+9}"
    return ws


# ══════════════════════════════════════════════════
# v4 기존 항목 (체크리스트 + 일부 견적서 항목)
# ══════════════════════════════════════════════════

mat_01_pre = [
    ("PRE-001","보양재","롤/시트","엘리베이터·공용부·에어컨·창호 보양","필수","M2","현장 손상 방지",2000,8000),
    ("PRE-002","소화기","ABC 3.3kg","현장 비치","필수","개","소방 안전",0,0),
    ("PRE-003","안전수칙 안내문","A3","현장 부착","필수","장","",0,0),
    ("PRE-004","공사 안내문","A4","이웃 배포","필수","장","",0,0),
    ("PRE-005","마대자루","PP","폐기물 담기","필수","장","",0,0),
    ("PRE-006","줄자","5m/7.5m","실측용","필수","개","",0,0),
    ("PRE-007","레이저 줄자","—","정밀 실측","선택","개","",0,0),
    ("PRE-008","레벨기","레이저","수평·수직 확인","필수","개","",0,0),
    ("PRE-009","수평계","자석형","수평 확인","필수","개","",0,0),
    ("PRE-010","안전모","—","작업자 보호","필수","개","",0,0),
    ("PRE-011","안전화","—","작업자 보호","필수","켤레","",0,0),
    ("PRE-012","빗자루/쓰레받이","—","현장 청소","필수","세트","",0,0),
    ("PRE-013","사다리차","—","자재 양중","필수","대","대운반",0,350000),
    ("PRE-014","작업등","LED","공사 기간 조명","필수","개","",0,0),
]

mat_02_demo = [
    ("DM-001","뿌레카(전동해머)","—","콘크리트·타일 파쇄","필수","대","",0,0),
    ("DM-002","그라인더","4인치","절단·연삭","필수","대","",0,0),
    ("DM-003","매직","유성","철거 범위 현장 표시","필수","개","",0,0),
    ("DM-004","보양재","롤/시트","파쇄 후 즉시 보양","필수","롤","",0,0),
    ("DM-005","레미탈","40kg","습식 자재","필수","포","",0,0),
    ("DM-006","벽돌","시멘트벽돌","조적 공사","선택","장","",0,0),
    ("DM-007","단열재","아이소핑크","확장부 철거 단면","선택","장","",0,0),
]

mat_03_plumb = [
    ("PL-001","동파이프","15A/20A","급수관","필수","M","",0,0),
    ("PL-002","PB관(엑셀파이프)","16mm/20mm","급수관(대체)","필수","M","",0,0),
    ("PL-003","PVC 배수관","50mm/75mm/100mm","배수관","필수","M","",0,0),
    ("PL-004","배수트랩","스텐/PVC","바닥 배수구","필수","개","",0,0),
    ("PL-005","밸브","앵글/볼밸브","급수 차단","필수","개","",0,0),
    ("PL-006","이음쇠(피팅)","각종","관 연결","필수","개","",0,0),
    ("PL-007","쉬라우드(커버)","도금/크롬","노출배관 마감","선택","개","",0,0),
    ("PL-008","세대 분배기","STS","급수 분기","선택","세트","",0,0),
    ("PL-009","감압밸브","—","수압 조절","선택","개","",0,0),
    ("PL-010","배관 보온재","PE폼","동결 방지","필수","M","",0,0),
    ("PL-011","연결 소켓","각종","관경 변환","필수","개","",0,0),
    ("PL-012","실리콘","배관용","접합부 밀봉","필수","개","",0,0),
    ("PL-013","행거/클램프","STS","배관 고정","필수","개","",0,0),
    ("PL-014","보일러 분배기","6~8구","난방 분기","선택","세트","",0,0),
    ("PL-015","온도조절기","디지털","난방 제어","선택","개","",0,0),
    ("PL-016","XL파이프","난방배관","난방관","선택","M","",0,0),
    ("PL-017","몰탈","시멘트사","바닥 미장","필수","포","",0,0),
    ("PL-018","세면대 급수관","STS 플렉시블","세면대 연결","필수","개","",0,0),
    ("PL-019","양변기 급수관","STS 플렉시블","양변기 연결","필수","개","",0,0),
    ("PL-020","배수호스","PVC","세탁기·싱크대","필수","M","",0,0),
    ("PL-021","오버플로우","PVC","싱크대 하부","필수","개","",0,0),
    ("PL-022","역류방지밸브","—","배수 역류 차단","필수","개","",0,0),
]

mat_04_elec = [
    ("EL-001","전선(HIV)","2.5sq","조명·콘센트","필수","M","",0,0),
    ("EL-002","전선(HIV)","4.0sq","에어컨·인덕션","필수","M","",0,0),
    ("EL-003","전선(HIV)","6.0sq","인입선","필수","M","",0,0),
    ("EL-004","CD관(가요관)","16mm","전선 보호","필수","M","",0,0),
    ("EL-005","분전반","ELB 포함","차단기 함체","필수","조","",0,0),
    ("EL-006","콘센트","매립형","전원 공급","필수","개","",0,0),
    ("EL-007","스위치","1구/2구/3구","조명 제어","필수","개","",0,0),
    ("EL-008","비닐테이프","전기용","절연","필수","롤","",0,0),
    ("EL-009","매입박스","—","벽체 매립","필수","개","",0,0),
    ("EL-010","접지선","녹색","접지용","필수","M","",0,0),
    ("EL-011","LED 다운라이트","3인치 6W","매립등","필수","개","",0,0),
    ("EL-012","LED 다운라이트","4인치 10W","매립등","필수","개","",0,0),
    ("EL-013","간접조명 LED바","5050 12V","간접등","선택","M","",0,0),
    ("EL-014","SMPS","12V/24V","LED 전원","선택","개","",0,0),
    ("EL-015","센서등","PIR","현관 자동","선택","개","",0,0),
    ("EL-016","비상등","LED","비상조명","선택","개","",0,0),
    ("EL-017","감지기","연기/열","화재 감지","필수","개","",0,0),
    ("EL-018","인터폰","비디오폰","방문자 확인","선택","대","",0,0),
    ("EL-019","도어락","디지털","현관문","선택","대","",0,0),
    ("EL-020","TV 단자","—","TV 연결","선택","개","",0,0),
    ("EL-021","LAN 단자","CAT6","인터넷","선택","개","",0,0),
    ("EL-022","전열교환기","—","환기","선택","대","",0,0),
]

mat_05_waterproof = [
    ("WP-001","도막방수재","우레탄/아크릴","바닥 방수","필수","kg","",0,0),
    ("WP-002","시트방수재","HDPE/PVC","벽·바닥 방수","선택","M2","",0,0),
    ("WP-003","프라이머","도막방수용","접착력 향상","필수","kg","",0,0),
    ("WP-004","실리콘","방수용","조인트","필수","개","",0,0),
    ("WP-005","부직포","—","보강","필수","M2","",0,0),
    ("WP-006","코너비드","방수용","모서리 보강","필수","M","",0,0),
    ("WP-007","방수테이프","부틸","이음새","필수","롤","",0,0),
    ("WP-008","배수트랩","스텐","바닥 배수구","필수","개","",0,0),
    ("WP-009","후렌치코트(배수판)","HDPE","바닥 배수","선택","M2","",0,0),
    ("WP-010","시멘트 방수제","혼입형","몰탈 혼합","선택","kg","",0,0),
    ("WP-011","레미탈","40kg","방수 위 미장","필수","포","",0,0),
    ("WP-012","방수 보호몰탈","—","방수층 보호","필수","kg","",0,0),
    ("WP-013","아덱스 도막방수제","2회","화장실 방수","필수","식","",240000,240000),
]

mat_06_window = [
    ("WN-001","시스템 창호","발코니","외부 창","필수","세트","",0,0),
    ("WN-002","이중창","거실/방","방음·단열","선택","세트","",0,0),
    ("WN-003","로이유리","복층","단열 유리","선택","M2","",0,0),
    ("WN-004","접합유리","안전·방음","방음 유리","선택","M2","",0,0),
    ("WN-005","방충망","알루미늄","방충","필수","세트","",0,0),
    ("WN-006","핸들","멀티포인트","개폐 장치","필수","개","",0,0),
    ("WN-007","경첩","스테인리스","연결 장치","필수","개","",0,0),
    ("WN-008","기밀재(모헤어)","—","기밀 유지","필수","M","",0,0),
    ("WN-009","실리콘","건축용","외부 코킹","필수","개","",0,0),
    ("WN-010","PVC 문틀","—","문틀 교체","선택","세트","",0,0),
    ("WN-011","현관문","—","현관 도어","선택","세트","",0,0),
    ("WN-012","방문","—","실내 도어","필수","세트","",0,0),
    ("WN-013","중문","슬라이딩/스윙","주방·현관","선택","세트","",0,0),
    ("WN-014","유리 파티션","강화유리","공간 분리","선택","M2","",0,0),
    ("WN-015","블라인드","롤/우드","차양","선택","세트","",0,0),
    ("WN-016","커튼레일","매립형","커튼 설치","선택","M","",0,0),
]

mat_07_wood = [
    ("WD-001","각재","30×30mm","골조","필수","M","",0,0),
    ("WD-002","합판","4.5T/9T","보강·하지","필수","장","",0,0),
    ("WD-003","석고보드","9.5T","벽체·천장","필수","장","",0,0),
    ("WD-004","고밀도 MDF","9T","벽체 마감","선택","장","",0,0),
    ("WD-005","실리콘","건축용","코킹","필수","개","",0,0),
    ("WD-006","못/나사","각종","고정","필수","박스","",0,0),
    ("WD-007","본드","목공용","접착","필수","통","",0,0),
    ("WD-008","단열재","아이소핑크 20T","발코니 단열","선택","장","",0,0),
    ("WD-009","우레탄폼","500ml","틈새 충진","필수","개","",0,0),
    ("WD-010","걸레받이","MDF/PVC","벽·바닥 경계","선택","M","",0,0),
    ("WD-011","문선","MDF","문틀 마감","필수","M","",0,0),
    ("WD-012","코너비드","PVC/알루미늄","모서리 보호","필수","M","",0,0),
    ("WD-013","경첩","—","도어 연결","필수","개","",0,0),
    ("WD-014","도어 핸들","—","도어 손잡이","필수","개","",0,0),
    ("WD-015","도어 스토퍼","—","벽 보호","선택","개","",0,0),
    ("WD-016","드라이월 테이프","—","석고 조인트","필수","롤","",0,0),
    ("WD-017","조인트 컴파운드","—","석고 메움","필수","통","",0,0),
    ("WD-018","천장틀(경량철골)","M-bar/C-bar","천장 골조","선택","M","",0,0),
    ("WD-019","행거볼트","—","천장 매달기","필수","개","",0,0),
    ("WD-020","방부목","SPF","외부 데크","선택","M","",0,0),
    ("WD-021","커튼박스","MDF","커튼레일 매립","선택","세트","",0,0),
    ("WD-022","흡음재","그라스울 24K","방음 충진","선택","장","",0,0),
]

mat_08_film = [
    ("FM-001","인테리어 필름","PVC","벽체·도어·가구","필수","M2","",0,0),
    ("FM-002","프라이머","필름용","접착력 향상","필수","리터","",0,0),
    ("FM-003","스퀴지","플라스틱","공기 제거","필수","개","",0,0),
    ("FM-004","히트건","—","열수축","필수","대","",0,0),
    ("FM-005","칼날","스냅오프","재단","필수","통","",0,0),
]

mat_09_tile = [
    ("TL-001","바닥타일","600×600","화장실·현관","필수","M2","",0,0),
    ("TL-002","벽타일","300×600","화장실","필수","M2","",0,0),
    ("TL-003","대형타일","600×1200","화장실·현관","선택","M2","",0,0),
    ("TL-004","모자이크타일","300×300","포인트","선택","장","",0,0),
    ("TL-005","타일접착제","X18","접착","필수","포","",0,0),
    ("TL-006","줄눈재(메지)","화이트/그레이","충진","필수","kg","",0,0),
    ("TL-007","레미탈","40kg","하지 미장","필수","포","",0,0),
    ("TL-008","십자 스페이서","2mm/3mm","줄눈 간격","필수","봉지","",0,0),
    ("TL-009","실리콘","타일용","모서리·코너","필수","개","",0,0),
    ("TL-010","코너비드","타일용","모서리 마감","선택","M","",0,0),
    ("TL-011","논슬립 타일","—","바닥 미끄럼방지","선택","M2","",0,0),
    ("TL-012","타일절단기","수동/전동","재단","필수","대","",0,0),
    ("TL-013","프라이머","타일하지","접착력","필수","리터","",0,0),
    ("TL-014","몰탈(벽돌쌓기)","시멘트사","벽돌 접합","선택","포","",0,0),
    ("TL-015","내수합판","12T","벽체 하지","선택","장","",0,0),
    ("TL-016","후크타일","300×300","셀프시공","선택","M2","",0,0),
    ("TL-017","조적벽돌","시멘트벽돌","화장실 벽체","필수","장","",0,0),
    ("TL-018","방수테이프","타일하부","방수 보강","필수","롤","",0,0),
    ("TL-019","헥사곤타일","—","포인트","선택","M2","",0,0),
    ("TL-020","데코타일(LVT)","—","바닥 대체","선택","M2","",0,0),
    ("TL-021","소형타일","100×100","벽체 포인트","선택","M2","",0,0),
]

mat_10_paint = [
    ("PT-001","수성페인트","KS","천장·벽체","필수","리터","",0,0),
    ("PT-002","유성페인트","KS","목재·금속","선택","리터","",0,0),
    ("PT-003","프라이머","수성/유성","초벌","필수","리터","",0,0),
    ("PT-004","퍼티","수성","면 고르기","필수","kg","",0,0),
    ("PT-005","사포","120~400방","연마","필수","장","",0,0),
    ("PT-006","롤러","9인치","넓은 면","필수","개","",0,0),
    ("PT-007","붓","3인치/4인치","좁은 면·보수","필수","개","",0,0),
    ("PT-008","마스킹테이프","—","비도장부 보호","필수","롤","",0,0),
    ("PT-009","비닐 보양재","PE","바닥 보호","필수","롤","",0,0),
    ("PT-010","조색제","유니버설","색상 배합","선택","병","",0,0),
    ("PT-011","방수페인트","욕실용","방수면 도장","선택","리터","",0,0),
    ("PT-012","곰팡이방지 페인트","—","욕실·다용도","선택","리터","",0,0),
    ("PT-013","우드스테인","수성/유성","목재 마감","선택","리터","",0,0),
    ("PT-014","에폭시 코팅","—","바닥 특수","선택","세트","",0,0),
    ("PT-015","무늬코트","질석","천장 텍스처","선택","포","",0,0),
    ("PT-016","천장보수재","—","석고보드 이음","필수","통","",0,0),
    ("PT-017","실리콘","조색실리콘","코킹 마감","필수","개","",0,0),
]

mat_11_floor = [
    ("FL-001","강마루","12T","거실·방","필수","PY","",0,0),
    ("FL-002","강화마루","8T","경제형","선택","PY","",0,0),
    ("FL-003","원목마루","15T","프리미엄","선택","PY","",0,0),
    ("FL-004","비닐타일(LVT)","3T","실용형","선택","M2","",0,0),
    ("FL-005","방음매트","EPE 2mm","층간 소음","필수","롤","",0,0),
    ("FL-006","PE 필름","0.1T","습기 차단","필수","롤","",0,0),
    ("FL-007","T-몰딩","알루미늄","마감 경계","필수","M","",0,0),
    ("FL-008","걸레받이","PVC/MDF","벽·바닥 경계","필수","M","",0,0),
    ("FL-009","접착제","마루용","시공 접착","선택","통","",0,0),
    ("FL-010","본드","SBR","하지 접착","선택","통","",0,0),
    ("FL-011","장판","PVC","방 바닥","선택","M2","",0,0),
    ("FL-012","데코타일","LVT","방·거실","선택","M2","",0,0),
    ("FL-013","디딤돌(논슬립)","—","계단","선택","M","",0,0),
]

mat_12_wallpaper = [
    ("WL-001","실크 벽지","합지","벽체","필수","롤","",0,0),
    ("WL-002","실크 벽지","방염","천장","필수","롤","",0,0),
    ("WL-003","풀","전분풀","벽지 접착","필수","kg","",0,0),
    ("WL-004","초배지","—","하지","필수","롤","",0,0),
    ("WL-005","정배지","—","마감","필수","롤","",0,0),
    ("WL-006","실리콘","백색","이음새","필수","개","",0,0),
    ("WL-007","퍼티","수성","면 고르기","필수","kg","",0,0),
    ("WL-008","사포","240방","벽면 연마","필수","장","",0,0),
    ("WL-009","커터칼","벽지용","재단","필수","개","",0,0),
    ("WL-010","스무서","벽지용","밀착","필수","개","",0,0),
    ("WL-011","합지 벽지","—","경제형","선택","롤","",0,0),
    ("WL-012","포인트 벽지","—","악센트","선택","롤","",0,0),
    ("WL-013","천연벽지","한지/마","프리미엄","선택","롤","",0,0),
    ("WL-014","단열벽지","—","결로 방지","선택","롤","",0,0),
    ("WL-015","몰딩","MDF/PS","천장 경계","선택","M","",0,0),
    ("WL-016","코너비드","PVC","모서리","필수","M","",0,0),
    ("WL-017","시트지","PVC","가구·문틀","선택","M","",0,0),
    ("WL-018","타카","스테이플","부착","필수","박스","",0,0),
    ("WL-019","탄성코트","세라믹코트","발코니 마감","선택","M2","",0,0),
]

mat_13_furniture = [
    ("FN-001","싱크대 상부장","—","주방 수납","필수","M","",275000,55000),
    ("FN-002","싱크대 하부장","—","주방 수납","필수","M","",275000,55000),
    ("FN-003","싱크대 상판","인조석/세라믹","주방 작업면","필수","M","",0,0),
    ("FN-004","싱크볼","스텐/화강석","주방 설거지","필수","SET","",350000,50000),
    ("FN-005","싱크 수전","—","주방 급수","필수","SET","",250000,50000),
    ("FN-006","후드","레인지후드","환기","필수","대","",0,0),
    ("FN-007","가스레인지","3구/인덕션","조리","선택","대","",0,0),
    ("FN-008","냉장고장","—","냉장고 매립","선택","자","",0,0),
    ("FN-009","아일랜드","—","주방 보조","선택","M","",0,0),
    ("FN-010","신발장","현관","현관 수납","필수","자","",130000,50000),
    ("FN-011","붙박이장","방","옷장","선택","자","",130000,50000),
    ("FN-012","키큰장","주방/세탁실","수납","선택","자","",130000,50000),
    ("FN-013","세면대","도기","화장실","필수","개","",0,0),
    ("FN-014","양변기","도기","화장실","필수","개","",0,0),
    ("FN-015","욕조","아크릴/FRP","화장실","선택","개","",0,0),
    ("FN-016","세면 수전","—","화장실","필수","SET","",0,0),
    ("FN-017","샤워 수전","—","화장실","필수","SET","",0,0),
    ("FN-018","거울","벽걸이","화장실","필수","개","",0,0),
    ("FN-019","수건걸이","스텐/크롬","화장실","필수","SET","",50000,50000),
    ("FN-020","휴지걸이","스텐/크롬","화장실","필수","SET","",50000,50000),
    ("FN-021","경첩","BLUM","가구 도어","필수","개","",0,0),
    ("FN-022","레일","BLUM","서랍","필수","세트","",0,0),
    ("FN-023","손잡이","—","가구 도어","필수","개","",0,0),
    ("FN-024","가구 도어","—","주방/수납","필수","EA","",140000,20000),
    ("FN-025","가구 EP(마감판)","측면/상판","마감","필수","EA","",120000,20000),
    ("FN-026","플랩장","냉장고/키큰장","상부 수납","선택","M","",160000,50000),
    ("FN-027","오픈장","벽걸이","개방 수납","선택","자","",160000,50000),
    ("FN-028","서랍장","BLUM","하부 수납","선택","SET","",0,0),
    ("FN-029","홈바 상부장","—","홈바 수납","선택","M","",160000,50000),
    ("FN-030","홈바 하부장","—","홈바 수납","선택","M","",160000,50000),
]

mat_14_fixture = [
    ("FX-001","샤워기","스텐/크롬","화장실","필수","SET","",0,0),
    ("FX-002","배수구 트랩","스텐/PVC","바닥","필수","개","",0,0),
    ("FX-003","타올바","스텐","화장실","선택","개","",0,0),
    ("FX-004","비누대","스텐/크롬","화장실","선택","개","",0,0),
    ("FX-005","샤워파티션","유리","화장실","선택","세트","",200000,50000),
    ("FX-006","SMC 천장재","—","화장실 천장","필수","SET","",300000,100000),
    ("FX-007","이노솔","—","화장실 천장","선택","M2","",40000,5000),
    ("FX-008","환풍기","—","화장실 환기","필수","EA","",0,0),
    ("FX-009","유가(욕실바닥재)","150×150","화장실 바닥","필수","식","",32000,10000),
    ("FX-010","조적 벽체/젠다이","신설","화장실 구조","선택","SET","",150000,250000),
    ("FX-011","거울장","제작","화장실 수납","선택","SET","",250000,50000),
    ("FX-012","욕조 수전","—","욕조 급수","선택","SET","",250000,50000),
    ("FX-013","코너선반","스텐","화장실 수납","선택","개","",0,0),
    ("FX-014","도어락","디지털","현관","선택","SET","",0,0),
    ("FX-015","인터폰","비디오폰","현관","선택","SET","",0,0),
]

mat_15_final = [
    ("FI-001","입주청소","전문업체","전체 청소","필수","PY","",0,0),
    ("FI-002","하자 보수재","각종","보수","필수","세트","",0,0),
    ("FI-003","실리콘","투명/백색","마감 코킹","필수","개","",0,0),
    ("FI-004","터치업 페인트","—","스크래치 보수","필수","개","",0,0),
    ("FI-005","여분 자재","벽지/타일/마루","보수용","필수","세트","",0,0),
    ("FI-006","자재 이력서","—","사용 자재 기록","필수","부","",0,0),
    ("FI-007","준공 사진","—","시공 기록","필수","세트","",0,0),
]


# ══════════════════════════════════════════════════
# 견적서에서 추출한 항목을 동적으로 추가
# ══════════════════════════════════════════════════

def safe_num(v):
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def normalize_process(name):
    """공정명을 표준화"""
    if not name:
        return None
    name = name.strip()
    name = re.sub(r'^[A-Za-z0-9]+\.\s*', '', name)
    if re.match(r'^[\d.]+$', name):
        return None
    PMAP = {
        "가설공사": "PRE", "철거 공사": "DM", "철거공사": "DM",
        "철거 / 설비 공사": "DM", "철거/설비 공사": "DM",
        "설비 공사": "PL", "설비공사": "PL",
        "필름 공사": "FM", "필름공사": "FM",
        "가구 공사": "FN", "가구공사": "FN",
        "전기 공사": "EL", "전기공사": "EL",
        "화장실 공사": "FX", "화장실공사": "FX",
        "타일 공사": "TL", "타일공사": "TL",
        "도배 공사": "WL", "도배공사": "WL",
        "도장 공사": "PT", "도장공사": "PT",
        "마루 공사": "FL", "마루공사": "FL",
        "목공사": "WD", "목 공사": "WD",
        "금속 공사": "WD", "금속공사": "WD",
        "데크 공사": "PRE", "데크공사": "PRE",
        "추가/ 변경공사": "PRE",
    }
    for key, val in PMAP.items():
        if key in name:
            return val
    return None


def extract_estimate_items():
    """16개 견적서 xlsx에서 전체 항목을 추출하여 공정코드별 고유항목 딕셔너리 반환"""
    FOLDER = "견적서"
    all_items = defaultdict(lambda: defaultdict(lambda: {
        "units": set(), "specs": set(), "locations": set(),
        "mat_prices": [], "lab_prices": [], "sites": set()
    }))

    for f in sorted(os.listdir(FOLDER)):
        if not f.endswith('.xlsx') or '빌드어스' in f or '정리' in f or f.startswith('.'):
            continue
        fn = unicodedata.normalize('NFC', f)
        # 약칭
        short = fn[:8]
        for key, val in {"푸르지오 월드마크 리모델링 공사 견적":"푸르지오","만현마을":"만현","금호어울림":"금호","김성용":"김성용","김은경":"김은경",
                         "린스트라우스":"린스","백현마을":"백현","산본하이어스":"산본","이선":"이선","이성숙":"이성숙",
                         "전기영":"전기영","진산마을":"진산","행당동":"행당동"}.items():
            if key in fn:
                short = val
                break

        filepath = os.path.join(FOLDER, f)
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
        except:
            continue

        # 상세 시트 찾기
        ws = None
        for sn in wb.sheetnames:
            if sn == "상세" or "동" in sn:
                ws = wb[sn]
                break
        if not ws:
            ws = wb[wb.sheetnames[-1]]

        # 상세 영역 시작점 찾기
        detail_start = None
        gaset_count = 0
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if b and "가설공사" in str(b).strip():
                gaset_count += 1
                if gaset_count == 2:
                    detail_start = r
                    break
        if not detail_start:
            detail_start = 50

        current_proc = None
        for r in range(detail_start, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, 20)]
            while len(row) < 19:
                row.append(None)

            col_b = row[1]
            if not col_b:
                continue
            col_b_str = str(col_b).strip()

            # 소계/합계 스킵
            if any(kw in col_b_str for kw in ["소계","합계","직접공사비","TOTAL","간접공사비","일반관리비","이윤","인테리어 공사"]):
                continue

            # 공정 헤더
            proc = normalize_process(col_b_str)
            if proc:
                current_proc = proc
                continue

            if not current_proc:
                continue

            # -1 서브카테고리 vs 데이터 판별
            a_str = str(row[0]).strip() if row[0] else ""
            if a_str in ["-1", "-2", "-3"]:
                has_data = False
                for ci in [5, 6, 7, 8, 9, 10, 11]:
                    v = safe_num(row[ci])
                    if v is not None and v > 0:
                        has_data = True
                        break
                if not has_data:
                    continue

            # 항목 수집
            name = col_b_str
            spec = str(row[2]).strip()[:30] if row[2] else ""
            loc = str(row[3]).strip()[:20] if row[3] else ""
            unit = str(row[4]).strip() if row[4] else ""
            mat_p = safe_num(row[6])  # 자재단가
            lab_p = safe_num(row[8])  # 노무단가

            item = all_items[current_proc][name]
            item["units"].add(unit)
            if spec:
                item["specs"].add(spec)
            if loc:
                item["locations"].add(loc)
            item["sites"].add(short)
            if mat_p and mat_p > 0:
                item["mat_prices"].append(mat_p)
            if lab_p and lab_p > 0:
                item["lab_prices"].append(lab_p)

        wb.close()

    return all_items


def get_existing_names(items_list):
    """기존 항목의 품명 set 반환"""
    return {item[1].strip() for item in items_list}


def add_estimate_items(existing_items, pcode, estimate_items):
    """견적서 항목 중 기존에 없는 것을 추가"""
    existing_names = get_existing_names(existing_items)
    added = []
    idx = 1

    for name, data in sorted(estimate_items.items()):
        # 기존 항목과 매칭 확인
        matched = False
        for en in existing_names:
            # 양방향 부분 매칭
            if name == en or name in en or en in name:
                matched = True
                break
            # 핵심 단어 매칭 (3글자 이상)
            name_words = set(w for w in re.split(r'[\s/·()]+', name) if len(w) >= 2)
            en_words = set(w for w in re.split(r'[\s/·()]+', en) if len(w) >= 2)
            if name_words and en_words and name_words & en_words:
                # 공통 단어가 있으면 매칭
                overlap = len(name_words & en_words) / min(len(name_words), len(en_words))
                if overlap >= 0.5:
                    matched = True
                    break
        if matched:
            continue

        # 단위 결정 (가장 많이 나온 단위)
        units = list(data["units"] - {""})
        unit = units[0] if units else "식"

        # 규격 결정
        specs = list(data["specs"] - {""})
        spec = specs[0] if specs else "—"
        if len(spec) > 25:
            spec = spec[:25]

        # 위치/용도
        locs = list(data["locations"] - {""})
        purpose = locs[0] if locs else ""
        if len(purpose) > 20:
            purpose = purpose[:20]

        # 단가 (중앙값)
        mat_price = round(statistics.median(data["mat_prices"])) if data["mat_prices"] else 0
        lab_price = round(statistics.median(data["lab_prices"])) if data["lab_prices"] else 0

        # 한줄요약
        sites = sorted(data["sites"])
        summary = f"{len(sites)}현장: {','.join(sites[:5])}"

        code = f"{pcode}-{idx:02d}E"
        added.append((code, name, spec, purpose, "필수", unit, summary, mat_price, lab_price))
        idx += 1

    return existing_items + added


# ══════════════════════════════════════════════════
# 메인 실행
# ══════════════════════════════════════════════════

print("📊 견적서 16개 파일에서 항목 추출 중...")
estimate_data = extract_estimate_items()

print("  추출 완료:")
for proc, items in sorted(estimate_data.items()):
    print(f"    {proc}: {len(items)}개 고유 항목")

# 공정별 견적 항목 추가
print("\n🔄 기존 항목과 병합 중...")
all_mat = {
    "PRE": mat_01_pre,
    "DM": mat_02_demo,
    "PL": mat_03_plumb,
    "EL": mat_04_elec,
    "WP": mat_05_waterproof,
    "WN": mat_06_window,
    "WD": mat_07_wood,
    "FM": mat_08_film,
    "TL": mat_09_tile,
    "PT": mat_10_paint,
    "FL": mat_11_floor,
    "WL": mat_12_wallpaper,
    "FN": mat_13_furniture,
    "FX": mat_14_fixture,
    "FI": mat_15_final,
}

for proc_code, est_items in estimate_data.items():
    if proc_code in all_mat:
        before = len(all_mat[proc_code])
        all_mat[proc_code] = add_estimate_items(all_mat[proc_code], proc_code, est_items)
        after = len(all_mat[proc_code])
        if after > before:
            print(f"  {proc_code}: {before}개 → {after}개 (+{after-before}개)")

# ── 엑셀 생성 ──
print("\n📝 엑셀 생성 중...")
wb = openpyxl.Workbook()

all_proc = [
    ("자재_01_공사전",  "718096","PRE","공사전 준비",     all_mat["PRE"], True),
    ("자재_02_철거",    "A0AEC0","DM","철거",            all_mat["DM"], False),
    ("자재_03_설비",    "4299E1","PL","설비(배관)",       all_mat["PL"], False),
    ("자재_04_전기",    "ECC94B","EL","전기",            all_mat["EL"], False),
    ("자재_05_방수",    "4299E1","WP","습식/방수",       all_mat["WP"], False),
    ("자재_06_창호",    "38B2AC","WN","창호",            all_mat["WN"], False),
    ("자재_07_목공",    "9F7AEA","WD","목공",            all_mat["WD"], False),
    ("자재_08_필름",    "ED8936","FM","필름",            all_mat["FM"], False),
    ("자재_09_타일",    "ED8936","TL","타일",            all_mat["TL"], False),
    ("자재_10_도장",    "48BB78","PT","도장",            all_mat["PT"], False),
    ("자재_11_마루",    "9F7AEA","FL","마루/바닥",       all_mat["FL"], False),
    ("자재_12_도배",    "E53E3E","WL","도배",            all_mat["WL"], False),
    ("자재_13_가구",    "D69E2E","FN","가구",            all_mat["FN"], False),
    ("자재_14_도기마감","FC8181","FX","도기/마감기구",    all_mat["FX"], False),
    ("자재_15_준공",    "38B2AC","FI","준공청소/점검",    all_mat["FI"], False),
]

total_items = 0
new_items = 0
priced_items = 0
for name, color, pcode, pname, items, is_first in all_proc:
    create_mat_sheet(wb, name, color, pcode, pname, items, is_first)
    total_items += len(items)
    for item in items:
        if item[0].endswith("E"):
            new_items += 1
        if item[7] > 0 or item[8] > 0:
            priced_items += 1
    print(f"  {name}: {len(items)}개 자재 (견적추출 {sum(1 for i in items if i[0].endswith('E'))}개)")


# ── 공정 마스터 ──
ws_p = wb.create_sheet("공정마스터")
ws_p.sheet_properties.tabColor = "ED8936"
ws_p.merge_cells("A1:O1")
ws_p.cell(row=1, column=1, value="공정 마스터").font = TF
ph = ["공정코드","공정명","순서","카테고리","기본공기(일)","선행공정","병행가능","필수여부",
      "난이도","감리포인트","자재수","도면필요","날씨영향","소음영향","시즌"]
for i, h in enumerate(ph, 1):
    ws_p.cell(row=2, column=i, value=h)
sty_h(ws_p, 2, 15)
procs = [
    ("PRE","공사전 준비",1,"사전","3~5","—","—","필수","하","보양 상태 확인",len(all_mat["PRE"]),"X","X","X","—"),
    ("DM","철거",2,"해체","3~7","PRE","—","필수","중","구조벽 확인",len(all_mat["DM"]),"△","X","O","—"),
    ("PL","설비(배관)",3,"습식","5~10","DM","EL","필수","상","배관 압력 테스트",len(all_mat["PL"]),"O","X","X","동절기 주의"),
    ("EL","전기",4,"건식","5~10","DM","PL","필수","상","절연 저항 측정",len(all_mat["EL"]),"O","X","X","—"),
    ("WP","방수",5,"습식","3~5","PL","—","필수","상","48시간 담수 시험",len(all_mat["WP"]),"△","O","X","우기 주의"),
    ("WN","창호",6,"건식","2~3","DM","WD","선택","중","기밀·수밀 테스트",len(all_mat["WN"]),"△","X","X","—"),
    ("WD","목공",7,"건식","10~15","WP","—","필수","상","수직·수평 검사",len(all_mat["WD"]),"O","X","O","—"),
    ("FM","필름",8,"건식","3~5","WD","—","선택","중","기포·들뜸 확인",len(all_mat["FM"]),"X","X","X","—"),
    ("TL","타일",9,"습식","5~7","WP","—","필수","상","줄눈·들뜸 확인",len(all_mat["TL"]),"△","O","X","동절기 주의"),
    ("PT","도장",10,"습식","3~5","WD","—","필수","중","도막 두께 확인",len(all_mat["PT"]),"X","O","X","동·우기 주의"),
    ("FL","마루",11,"건식","2~3","PT","—","필수","중","들뜸·단차 확인",len(all_mat["FL"]),"X","X","X","—"),
    ("WL","도배",12,"습식","2~3","PT","—","필수","중","이음매·기포 확인",len(all_mat["WL"]),"X","O","X","동절기 주의"),
    ("FN","가구",13,"건식","5~10","WD","—","필수","중","수평·여닫힘 확인",len(all_mat["FN"]),"O","X","X","—"),
    ("FX","도기/마감",14,"건식","2~3","TL","—","필수","중","고정·누수 확인",len(all_mat["FX"]),"X","X","X","—"),
    ("FI","준공",15,"—","1~2","ALL","—","필수","하","전체 하자 점검",len(all_mat["FI"]),"X","X","X","—"),
]
for i, p in enumerate(procs, 3):
    for j, v in enumerate(p, 1):
        ws_p.cell(row=i, column=j, value=v)
    sty_r(ws_p, i, 15)
for c, w in enumerate([8,12,5,6,9,8,8,7,5,18,7,7,7,7,10], 1):
    ws_p.column_dimensions[get_column_letter(c)].width = w


# ── 호환성·공급사·면적참조 ──
ws_c = wb.create_sheet("호환성")
ws_c.sheet_properties.tabColor = "38B2AC"
ch = ["브랜드","호환 브랜드","카테고리","호환 등급","비고"]
ws_c.merge_cells("A1:E1"); ws_c.cell(1,1,value="브랜드 호환성").font = TF
for i, h in enumerate(ch,1): ws_c.cell(2,i,value=h)
sty_h(ws_c,2,5)
compat = [("LX하우시스","한화 L&C","바닥재","A","두께·클릭 호환"),("한샘","리바트","가구","B","하드웨어 유사"),
          ("대림","아메리칸 스탠다드","도기","A","배관 규격 동일"),("KCC","LX Z:IN","창호","B","프로파일 유사"),
          ("영림","한솔","필름","A","폭·두께 호환")]
for i, r in enumerate(compat,3):
    for j, v in enumerate(r,1): ws_c.cell(i,j,value=v)
    sty_r(ws_c,i,5)

ws_s = wb.create_sheet("공급사")
ws_s.sheet_properties.tabColor = "38B2AC"

ws_a = wb.create_sheet("면적참조")
ws_a.sheet_properties.tabColor = "38B2AC"
ah = ["공간","면적(㎡)","바닥","벽체","천장","비고"]
ws_a.merge_cells("A1:F1"); ws_a.cell(1,1,value="표준 면적 참조").font = TF
for i, h in enumerate(ah,1): ws_a.cell(2,i,value=h)
sty_h(ws_a,2,6)
areas = [("거실",20,"O","O","O",""),("안방",12,"O","O","O",""),("방1",8,"O","O","O",""),
         ("주방",6,"O","O","O",""),("화장실(거실)",3.5,"O","O","O",""),("현관",2,"O","X","O","")]
for i, r in enumerate(areas,3):
    for j, v in enumerate(r,1): ws_a.cell(i,j,value=v)
    sty_r(ws_a,i,6)


# ── 견적서 요약 ──
ws_e = wb.create_sheet("견적서요약")
ws_e.sheet_properties.tabColor = "D69E2E"
ws_e.merge_cells("A1:R1")
ws_e.cell(1,1,value="17개 견적서 프로젝트 비교").font = TF
eh = ["프로젝트","총액(만원)","평당가(만원)","평수","가설","철거/설비","필름","가구","전기",
      "화장실","타일","도배","도장","마루","목공","금속","기타","비고"]
for i, h in enumerate(eh,1):
    ws_e.cell(2,i,value=h)
sty_h(ws_e,2,18)
est_summary = [
    ("푸르지오 월드마크(본)",13370,256,52,"701","1366","342","1997","700","826","1137","895","120","759","1870","0","0",""),
    ("푸르지오 월드마크(추가)",39,0,0,"0","0","0","21","0","10","0","0","0","0","0","0","0","변동분"),
    ("만현마을 9단지(타일ver)",18476,430,43,"585","1758","263","2227","1116","981","4263","657","160","0","2764","0","0","전면 타일"),
    ("만현마을 9단지(필름ver)",17494,407,43,"585","1758","720","2333","1116","981","1082","657","160","973","3700","0","0","필름+마루"),
    ("금호어울림",10270,320,32,"497","830","420","1495","461","465","668","484","119","322","1251","0","0",""),
    ("김성용 태영데시앙",12240,374,33,"612","1052","373","2154","778","710","718","552","110","550","2019","0","0","에어컨 포함"),
    ("김은경",11848,406,29,"443","930","355","1816","671","653","689","469","120","498","1760","0","0",""),
    ("린스트라우스",18564,420,44,"695","1770","345","2697","851","826","1137","853","120","1449","3050","0","0",""),
    ("백현마을",16700,390,43,"620","1700","400","2500","850","800","1100","800","120","1200","2800","0","0",""),
    ("산본하이어스",13200,380,35,"550","1200","300","2200","750","750","900","600","100","800","2500","0","0",""),
    ("이선",12800,370,35,"580","1100","350","2100","700","700","850","550","120","700","2400","100","0","금속공사 포함"),
    ("이성숙",8500,330,26,"400","700","200","1500","500","550","500","350","80","450","1600","0","0",""),
    ("전기영 판교해링턴",7150,310,23,"215","468","382","1048","492","410","435","310","102","426","1649","0","0",""),
    ("진산마을 삼성레미안",18607,447,42,"1471","2173","383","2868","996","1004","989","657","168","1338","2955","500","0","금속공사 포함"),
    ("행당동 대림",7918,344,23,"437","491","298","951","461","465","668","484","135","322","1251","0","0",""),
    ("전기영 해링턴",7150,310,23,"215","468","382","1048","492","410","435","310","102","426","1649","0","0",""),
    ("신봉동 삼성쉐르빌",23000,442,52,"4476","2252","562","1903","1180","1093","1333","900","208","1310","3414","0","0","PDF 견적"),
]
for i, r in enumerate(est_summary, 3):
    for j, v in enumerate(r, 1):
        ws_e.cell(i, j, value=v)
    sty_r(ws_e, i, 18)
    for c in range(2, 18):
        ws_e.cell(i, c).number_format = "#,##0"
for c, w in enumerate([18,10,10,6,7,9,7,7,7,7,7,7,7,7,7,7,7,15], 1):
    ws_e.column_dimensions[get_column_letter(c)].width = w


# ── 저장 ──
OUTPUT = "빌드어스_데이터아키텍처_v5.xlsx"
wb.save(OUTPUT)
print(f"\n✅ {OUTPUT} 생성 완료!")
print(f"   총 {total_items}개 항목 ({new_items}개 견적서 추출)")
print(f"   견적 참고단가: {priced_items}개")
print(f"   시트: {len(wb.sheetnames)}개")
