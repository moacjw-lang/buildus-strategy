"""
빌드어스 데이터 아키텍처 엑셀 v2
- 공정별 자재 시트 분리
- 각 공정 시트 = 자재 목록 + 소요량/로스율 + 가격 통합
- 견적 시뮬레이터에서 각 공정 시트 참조
"""
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── 스타일 ──
HEADER_FILL = PatternFill(start_color="2D3748", end_color="2D3748", fill_type="solid")
HEADER_FONT = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
SECTION_FILL = PatternFill(start_color="4A5568", end_color="4A5568", fill_type="solid")
SECTION_FONT = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
EXAMPLE_FILL = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
FORMULA_FILL = PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid")
RESULT_FILL = PatternFill(start_color="F0FFF4", end_color="F0FFF4", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFFFF0", end_color="FFFFF0", fill_type="solid")
NORMAL = Font(name="맑은 고딕", size=10)
BOLD = Font(name="맑은 고딕", bold=True, size=10)
TITLE = Font(name="맑은 고딕", bold=True, size=14)
SUBTITLE = Font(name="맑은 고딕", bold=True, size=11, color="4299E1")
THIN = Border(
    left=Side(style="thin", color="E2E8F0"),
    right=Side(style="thin", color="E2E8F0"),
    top=Side(style="thin", color="E2E8F0"),
    bottom=Side(style="thin", color="E2E8F0"),
)
INPUT_BORDER = Border(
    left=Side(style="medium", color="D69E2E"),
    right=Side(style="medium", color="D69E2E"),
    top=Side(style="medium", color="D69E2E"),
    bottom=Side(style="medium", color="D69E2E"),
)

def style_header(ws, row, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN

def style_section(ws, row, max_col, label):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    cell = ws.cell(row=row, column=1, value=label)
    cell.font = SECTION_FONT
    cell.fill = SECTION_FILL
    cell.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(1, max_col + 1):
        ws.cell(row=row, column=c).fill = SECTION_FILL
        ws.cell(row=row, column=c).border = THIN

def style_rows(ws, start, end, max_col, example=False):
    for r in range(start, end + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = NORMAL
            cell.border = THIN
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if example:
                cell.fill = EXAMPLE_FILL

def note(ws, row, col, text, end_col=None):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = Font(name="맑은 고딕", size=9, color="718096", italic=True)
    cell.alignment = Alignment(wrap_text=True)
    if end_col:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=end_col)

def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ══════════════════════════════════
wb = openpyxl.Workbook()


# ─────────────────────────────────
# 공정별 자재 시트 생성 함수
# ─────────────────────────────────
def create_material_sheet(wb, sheet_name, tab_color, proc_code, proc_name, materials):
    """
    materials: list of dict
      code, name, spec, brand, maker, unit, pack_unit, pack_qty,
      wholesale, dealer, retail, coupang, naver, survey_date,
      weight, shipping, ship_days, ship_fee, supplier,
      sample_yn, sample_size, sample_free, sample_cost,
      summary, features, caution,
      style, grade, spaces,
      # 매핑 (소요량/로스율)
      usage_purpose, required, per_sqm, loss_pct, usage_note
    """
    ws = wb.create_sheet(sheet_name) if sheet_name != wb.sheetnames[0] else wb.active
    ws.sheet_properties.tabColor = tab_color

    MAX_COL = 26

    # 제목
    ws.merge_cells(f"A1:{get_column_letter(MAX_COL)}1")
    ws.cell(row=1, column=1, value=f"{proc_name} 자재").font = TITLE
    ws.merge_cells(f"A2:{get_column_letter(MAX_COL)}2")
    ws.cell(row=2, column=1, value=f"공정코드: {proc_code} | 자재 + 소요량 + 가격 + 배송 통합").font = SUBTITLE
    ws.row_dimensions[1].height = 30

    # 섹션 구분 (3행)
    row = 3
    sections = [
        (1, 9, "기본 정보 + 소요량"),
        (10, 17, "가격"),
        (18, 21, "물류/배송"),
        (22, 24, "샘플"),
        (25, 26, "등급"),
    ]
    for s, e, label in sections:
        ws.merge_cells(start_row=row, start_column=s, end_row=row, end_column=e)
        cell = ws.cell(row=row, column=s, value=label)
        cell.font = SECTION_FONT
        cell.fill = SECTION_FILL
        cell.alignment = Alignment(horizontal="center")
        for c in range(s, e + 1):
            ws.cell(row=row, column=c).fill = SECTION_FILL
            ws.cell(row=row, column=c).border = THIN

    # 헤더 (4행)
    headers = [
        # 기본 + 소요량 (1-9)
        "자재코드", "품명", "규격", "용도", "필수/선택",
        "단위", "1㎡당\n소요량", "로스율\n(%)", "실소요량\n(1㎡)",
        # 가격 (10-17)
        "판매단위", "판매단위\n수량", "도매가", "업체가", "판매가",
        "쿠팡가", "네이버\n최저가", "절약금액",
        # 물류 (18-21)
        "무게(kg)", "배송방법", "배송일", "공급사코드",
        # 샘플 (22-24)
        "샘플가능", "샘플크기", "샘플비용",
        # 등급 (25-26)
        "등급", "한줄요약",
    ]
    row = 4
    for i, h in enumerate(headers, 1):
        ws.cell(row=row, column=i, value=h)
    style_header(ws, row, MAX_COL)
    ws.row_dimensions[4].height = 36

    # 데이터
    for i, m in enumerate(materials):
        r = 5 + i
        ws.cell(row=r, column=1, value=m["code"])
        ws.cell(row=r, column=2, value=m["name"])
        ws.cell(row=r, column=3, value=m["spec"])
        ws.cell(row=r, column=4, value=m["purpose"])
        ws.cell(row=r, column=5, value=m["required"])
        ws.cell(row=r, column=6, value=m["unit"])
        ws.cell(row=r, column=7, value=m["per_sqm"])
        ws.cell(row=r, column=7).number_format = "0.00"
        ws.cell(row=r, column=8, value=m["loss_pct"])
        # 실소요량 수식
        ws.cell(row=r, column=9).value = f"=G{r}*(1+H{r}/100)"
        ws.cell(row=r, column=9).fill = FORMULA_FILL
        ws.cell(row=r, column=9).number_format = "0.000"

        ws.cell(row=r, column=10, value=m["pack_unit"])
        ws.cell(row=r, column=11, value=m["pack_qty"])
        ws.cell(row=r, column=12, value=m["wholesale"])
        ws.cell(row=r, column=12).number_format = "#,##0"
        ws.cell(row=r, column=13, value=m["dealer"])
        ws.cell(row=r, column=13).number_format = "#,##0"
        ws.cell(row=r, column=14, value=m["retail"])
        ws.cell(row=r, column=14).number_format = "#,##0"
        ws.cell(row=r, column=15, value=m.get("coupang", ""))
        ws.cell(row=r, column=15).number_format = "#,##0"
        ws.cell(row=r, column=16, value=m.get("naver", ""))
        ws.cell(row=r, column=16).number_format = "#,##0"
        # 절약금액 수식
        ws.cell(row=r, column=17).value = f'=IF(OR(O{r}="",P{r}=""),"",MIN(O{r},P{r})-N{r})'
        ws.cell(row=r, column=17).fill = FORMULA_FILL
        ws.cell(row=r, column=17).number_format = "#,##0"

        ws.cell(row=r, column=18, value=m.get("weight", ""))
        ws.cell(row=r, column=19, value=m.get("shipping", ""))
        ws.cell(row=r, column=20, value=m.get("ship_days", ""))
        ws.cell(row=r, column=21, value=m.get("supplier", ""))

        ws.cell(row=r, column=22, value=m.get("sample_yn", "N"))
        ws.cell(row=r, column=23, value=m.get("sample_size", "—"))
        ws.cell(row=r, column=24, value=m.get("sample_cost", "—"))

        ws.cell(row=r, column=25, value=m.get("grade", "표준"))
        ws.cell(row=r, column=26, value=m.get("summary", ""))

        style_rows(ws, r, r, MAX_COL, example=True)

    # 빈 행 20줄 추가 (입력용)
    last_data = 5 + len(materials)
    for r in range(last_data, last_data + 20):
        # 실소요량 수식 미리
        ws.cell(row=r, column=9).value = f'=IF(G{r}="","",G{r}*(1+H{r}/100))'
        ws.cell(row=r, column=9).fill = FORMULA_FILL
        ws.cell(row=r, column=9).number_format = "0.000"
        # 절약금액 수식 미리
        ws.cell(row=r, column=17).value = f'=IF(OR(O{r}="",P{r}=""),"",MIN(O{r},P{r})-N{r})'
        ws.cell(row=r, column=17).fill = FORMULA_FILL
        ws.cell(row=r, column=17).number_format = "#,##0"
        for c in range(1, MAX_COL + 1):
            ws.cell(row=r, column=c).border = THIN

    # 노트
    nr = last_data + 21
    note(ws, nr, 1, f"※ 파란 셀(실소요량, 절약금액)은 수식 자동계산 — 건드리지 마세요", MAX_COL)
    note(ws, nr+1, 1, f"※ 자재를 추가할 때: 빈 행에 입력하면 수식이 자동 적용됩니다", MAX_COL)
    note(ws, nr+2, 1, f"※ 행이 부족하면 마지막 행의 수식을 아래로 드래그하세요", MAX_COL)

    set_widths(ws, [
        10, 16, 12, 12, 7,     # 기본
        5, 8, 7, 8,            # 소요량
        10, 8, 9, 9, 9, 9, 10, 9,  # 가격
        7, 7, 7, 9,            # 물류
        7, 7, 7,               # 샘플
        7, 25,                 # 등급
    ])
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:{get_column_letter(MAX_COL)}{last_data + 19}"

    return ws


# ─────────────────────────────────
# 공정별 자재 데이터
# ─────────────────────────────────

# 방수
materials_waterproof = [
    dict(code="WP-001", name="수성 방수액", spec="18L", purpose="방수층 형성", required="필수",
         unit="L", per_sqm=0.6, loss_pct=10, pack_unit="통(18L)", pack_qty=18,
         wholesale=35000, dealer=42000, retail=45000, coupang=58000, naver=52000,
         weight=20, shipping="화물", ship_days="2~3일", supplier="SUP-001",
         sample_yn="N", grade="표준", summary="욕실 바닥/벽 방수. 2회 도포로 완벽 방수."),
    dict(code="WP-002", name="부직포", spec="1m×50m", purpose="보강재", required="필수",
         unit="㎡", per_sqm=1.05, loss_pct=5, pack_unit="롤(50㎡)", pack_qty=50,
         wholesale=12000, dealer=14000, retail=15000, coupang=22000, naver=18000,
         weight=3, shipping="택배", ship_days="1~2일", supplier="SUP-001",
         sample_yn="N", grade="표준", summary="방수 보강용 부직포. 겹침 5cm."),
    dict(code="WP-003", name="프라이머", spec="4L", purpose="바탕 처리", required="선택",
         unit="L", per_sqm=0.1, loss_pct=5, pack_unit="캔(4L)", pack_qty=4,
         wholesale=8000, dealer=10000, retail=12000, coupang=18000, naver=15000,
         weight=5, shipping="택배", ship_days="1~2일", supplier="SUP-001",
         sample_yn="N", grade="표준", summary="구축 바닥 바탕 처리용."),
    dict(code="WP-004", name="우레탄 방수재", spec="18L", purpose="방수층 형성", required="필수",
         unit="L", per_sqm=0.5, loss_pct=10, pack_unit="통(18L)", pack_qty=18,
         wholesale=55000, dealer=65000, retail=72000, coupang=95000, naver=88000,
         weight=22, shipping="화물", ship_days="2~3일", supplier="SUP-001",
         sample_yn="N", grade="프리미엄", summary="고내구성 우레탄 방수. 옥상/베란다용."),
]

# 타일
materials_tile = [
    dict(code="TL-001", name="포세린 타일", spec="600×600", purpose="바닥 마감", required="필수",
         unit="장", per_sqm=2.8, loss_pct=10, pack_unit="박스(8장)", pack_qty=8,
         wholesale=4500, dealer=5500, retail=6200, coupang=9000, naver=8200,
         weight=25, shipping="화물", ship_days="3~5일", supplier="SUP-002",
         sample_yn="Y", sample_size="1장", sample_cost="무료", grade="표준",
         summary="내구성 높은 포세린. 욕실/거실 겸용."),
    dict(code="TL-002", name="논슬립 타일", spec="300×300", purpose="욕실 바닥", required="필수",
         unit="장", per_sqm=11.1, loss_pct=8, pack_unit="박스(18장)", pack_qty=18,
         wholesale=1800, dealer=2200, retail=2500, coupang=4000, naver=3500,
         weight=20, shipping="화물", ship_days="3~5일", supplier="SUP-002",
         sample_yn="Y", sample_size="1장", sample_cost="무료", grade="경제",
         summary="미끄럼 방지 타일. 욕실 바닥 필수."),
    dict(code="TL-003", name="대리석 타일", spec="600×600", purpose="바닥/벽 마감", required="필수",
         unit="장", per_sqm=2.8, loss_pct=12, pack_unit="박스(6장)", pack_qty=6,
         wholesale=12000, dealer=15000, retail=18000, coupang=28000, naver=25000,
         weight=30, shipping="화물", ship_days="5~7일", supplier="SUP-002",
         sample_yn="Y", sample_size="1장", sample_cost="무료", grade="프리미엄",
         summary="이탈리아산 대리석 질감. 고급 인테리어."),
    dict(code="BA-001", name="타일 본드", spec="20kg", purpose="접착", required="필수",
         unit="kg", per_sqm=4.5, loss_pct=5, pack_unit="포(20kg)", pack_qty=20,
         wholesale=8000, dealer=10000, retail=12000, coupang=18000, naver=15000,
         weight=20, shipping="화물", ship_days="2~3일", supplier="SUP-002",
         sample_yn="N", grade="표준", summary="시멘트 기반 타일 접착제."),
    dict(code="GR-001", name="줄눈재", spec="5kg", purpose="줄눈 마감", required="필수",
         unit="kg", per_sqm=0.5, loss_pct=10, pack_unit="포(5kg)", pack_qty=5,
         wholesale=4000, dealer=5000, retail=6000, coupang=9000, naver=8000,
         weight=5, shipping="택배", ship_days="1~2일", supplier="SUP-002",
         sample_yn="N", grade="표준", summary="백색/그레이. 무기질 줄눈재."),
    dict(code="TL-SP-001", name="십자 스페이서", spec="2mm", purpose="줄눈 간격", required="필수",
         unit="개", per_sqm=8, loss_pct=20, pack_unit="봉(100개)", pack_qty=100,
         wholesale=1500, dealer=2000, retail=2500, coupang=4000, naver=3000,
         weight=0.5, shipping="택배", ship_days="1~2일", supplier="SUP-002",
         sample_yn="N", grade="표준", summary="2mm 줄눈 간격 유지용."),
]

# 도장
materials_paint = [
    dict(code="PT-001", name="수성 페인트 (백색)", spec="18L", purpose="벽면/천장 도장", required="필수",
         unit="L", per_sqm=0.15, loss_pct=5, pack_unit="통(18L)", pack_qty=18,
         wholesale=45000, dealer=55000, retail=62000, coupang=85000, naver=78000,
         weight=22, shipping="화물", ship_days="2~3일", supplier="SUP-003",
         sample_yn="Y", sample_size="100ml", sample_cost="무료", grade="표준",
         summary="KCC 수성 페인트. 2회 도장 기준."),
    dict(code="PT-002", name="수성 페인트 (컬러)", spec="4L", purpose="벽면 포인트", required="선택",
         unit="L", per_sqm=0.15, loss_pct=5, pack_unit="캔(4L)", pack_qty=4,
         wholesale=18000, dealer=22000, retail=25000, coupang=35000, naver=32000,
         weight=5, shipping="택배", ship_days="1~2일", supplier="SUP-003",
         sample_yn="Y", sample_size="100ml", sample_cost="무료", grade="표준",
         summary="포인트 벽 컬러. 조색 가능."),
    dict(code="PT-003", name="퍼티", spec="10kg", purpose="면 고르기", required="필수",
         unit="kg", per_sqm=0.3, loss_pct=10, pack_unit="포(10kg)", pack_qty=10,
         wholesale=6000, dealer=8000, retail=10000, coupang=15000, naver=12000,
         weight=10, shipping="택배", ship_days="1~2일", supplier="SUP-003",
         sample_yn="N", grade="표준", summary="벽면 평활 작업용 퍼티."),
    dict(code="PT-004", name="사포", spec="P180", purpose="면 정리", required="필수",
         unit="장", per_sqm=0.3, loss_pct=30, pack_unit="묶음(50장)", pack_qty=50,
         wholesale=100, dealer=150, retail=200, coupang=300, naver=250,
         weight=0.5, shipping="택배", ship_days="1~2일", supplier="SUP-003",
         sample_yn="N", grade="표준", summary="퍼티 후 면 정리용 사포."),
    dict(code="PT-005", name="마스킹 테이프", spec="24mm×40m", purpose="양생", required="필수",
         unit="롤", per_sqm=0.1, loss_pct=20, pack_unit="롤(1개)", pack_qty=1,
         wholesale=1500, dealer=2000, retail=2500, coupang=3500, naver=3000,
         weight=0.1, shipping="택배", ship_days="1~2일", supplier="SUP-003",
         sample_yn="N", grade="표준", summary="도장 경계 양생용."),
]

# 마루
materials_floor = [
    dict(code="FL-001", name="강마루", spec="1210×193×8T", purpose="바닥 마감", required="필수",
         unit="장", per_sqm=5.35, loss_pct=5, pack_unit="박스(8장/1.87㎡)", pack_qty=8,
         wholesale=4500, dealer=5500, retail=6200, coupang=9000, naver=8200,
         weight=12, shipping="화물", ship_days="3~5일", supplier="SUP-004",
         sample_yn="Y", sample_size="1장", sample_cost="3,000원", grade="표준",
         summary="클릭 방식 강마루. 셀프 시공 가능."),
    dict(code="FL-002", name="강화마루", spec="1210×193×12T", purpose="바닥 마감", required="필수",
         unit="장", per_sqm=5.35, loss_pct=5, pack_unit="박스(8장/1.87㎡)", pack_qty=8,
         wholesale=7000, dealer=8500, retail=9500, coupang=14000, naver=12000,
         weight=15, shipping="화물", ship_days="3~5일", supplier="SUP-004",
         sample_yn="Y", sample_size="1장", sample_cost="3,000원", grade="프리미엄",
         summary="12T 고급 강화마루. 충격 흡수."),
    dict(code="FL-003", name="걸레받이", spec="60mm×2.4m", purpose="벽 하단 마감", required="필수",
         unit="m", per_sqm=0.4, loss_pct=10, pack_unit="본(2.4m)", pack_qty=1,
         wholesale=2000, dealer=2500, retail=3000, coupang=5000, naver=4000,
         weight=0.5, shipping="택배", ship_days="1~2일", supplier="SUP-004",
         sample_yn="N", grade="표준", summary="마루 벽 접합부 마감용."),
    dict(code="FL-004", name="PE 폼 언더레이", spec="2mm×50m", purpose="하부 완충", required="필수",
         unit="㎡", per_sqm=1.0, loss_pct=5, pack_unit="롤(50㎡)", pack_qty=50,
         wholesale=600, dealer=800, retail=1000, coupang=1800, naver=1500,
         weight=3, shipping="택배", ship_days="1~2일", supplier="SUP-004",
         sample_yn="N", grade="표준", summary="바닥 소음 완충 + 습기 차단."),
]

# 도배
materials_wallpaper = [
    dict(code="WL-001", name="합지벽지", spec="롤(16.5㎡)", purpose="벽면 마감", required="필수",
         unit="롤", per_sqm=0.07, loss_pct=10, pack_unit="롤(1개)", pack_qty=1,
         wholesale=8000, dealer=10000, retail=12000, coupang=18000, naver=15000,
         weight=2, shipping="택배", ship_days="1~2일", supplier="SUP-005",
         sample_yn="Y", sample_size="A4", sample_cost="무료", grade="경제",
         summary="경제적 합지벽지. 풀 도포 필요."),
    dict(code="WL-002", name="실크벽지", spec="롤(15.6㎡)", purpose="벽면 마감", required="필수",
         unit="롤", per_sqm=0.075, loss_pct=10, pack_unit="롤(1개)", pack_qty=1,
         wholesale=15000, dealer=18000, retail=22000, coupang=32000, naver=28000,
         weight=2.5, shipping="택배", ship_days="1~2일", supplier="SUP-005",
         sample_yn="Y", sample_size="A4", sample_cost="무료", grade="표준",
         summary="PVC 코팅 실크벽지. 세척 가능."),
    dict(code="WL-003", name="벽지풀", spec="5kg", purpose="접착", required="필수",
         unit="kg", per_sqm=0.15, loss_pct=5, pack_unit="통(5kg)", pack_qty=5,
         wholesale=5000, dealer=6000, retail=7000, coupang=10000, naver=9000,
         weight=5, shipping="택배", ship_days="1~2일", supplier="SUP-005",
         sample_yn="N", grade="표준", summary="전분 기반 벽지풀."),
]

# ── 시트 생성 ──
ws_first = wb.active
ws_first.title = "자재_방수"
processes = [
    ("자재_방수", "4299E1", "PROC-04", "방수", materials_waterproof),
    ("자재_타일", "ED8936", "PROC-08", "타일", materials_tile),
    ("자재_도장", "48BB78", "PROC-09", "도장", materials_paint),
    ("자재_마루", "9F7AEA", "PROC-10", "마루/바닥", materials_floor),
    ("자재_도배", "E53E3E", "PROC-11", "도배", materials_wallpaper),
]

for i, (name, color, pcode, pname, mats) in enumerate(processes):
    if i == 0:
        ws = wb.active
        ws.title = name
        ws.sheet_properties.tabColor = color
        # 제목
        ws.merge_cells("A1:Z1")
        ws.cell(row=1, column=1, value=f"{pname} 자재").font = TITLE
        ws.merge_cells("A2:Z2")
        ws.cell(row=2, column=1, value=f"공정코드: {pcode} | 자재 + 소요량 + 가격 + 배송 통합").font = SUBTITLE
        ws.row_dimensions[1].height = 30
        MAX_COL = 26

        sections = [
            (1, 9, "기본 정보 + 소요량"),
            (10, 17, "가격"),
            (18, 21, "물류/배송"),
            (22, 24, "샘플"),
            (25, 26, "등급"),
        ]
        row = 3
        for s, e, label in sections:
            ws.merge_cells(start_row=row, start_column=s, end_row=row, end_column=e)
            cell = ws.cell(row=row, column=s, value=label)
            cell.font = SECTION_FONT
            cell.fill = SECTION_FILL
            cell.alignment = Alignment(horizontal="center")
            for c in range(s, e + 1):
                ws.cell(row=row, column=c).fill = SECTION_FILL
                ws.cell(row=row, column=c).border = THIN

        headers = [
            "자재코드", "품명", "규격", "용도", "필수/선택",
            "단위", "1㎡당\n소요량", "로스율\n(%)", "실소요량\n(1㎡)",
            "판매단위", "판매단위\n수량", "도매가", "업체가", "판매가",
            "쿠팡가", "네이버\n최저가", "절약금액",
            "무게(kg)", "배송방법", "배송일", "공급사코드",
            "샘플가능", "샘플크기", "샘플비용",
            "등급", "한줄요약",
        ]
        row = 4
        for j, h in enumerate(headers, 1):
            ws.cell(row=row, column=j, value=h)
        style_header(ws, row, MAX_COL)
        ws.row_dimensions[4].height = 36

        for mi, m in enumerate(mats):
            r = 5 + mi
            vals = [
                m["code"], m["name"], m["spec"], m["purpose"], m["required"],
                m["unit"], m["per_sqm"], m["loss_pct"], None,
                m["pack_unit"], m["pack_qty"], m["wholesale"], m["dealer"], m["retail"],
                m.get("coupang", ""), m.get("naver", ""), None,
                m.get("weight", ""), m.get("shipping", ""), m.get("ship_days", ""), m.get("supplier", ""),
                m.get("sample_yn", "N"), m.get("sample_size", "—"), m.get("sample_cost", "—"),
                m.get("grade", "표준"), m.get("summary", ""),
            ]
            for j, v in enumerate(vals):
                if v is not None:
                    ws.cell(row=r, column=j+1, value=v)
            ws.cell(row=r, column=7).number_format = "0.00"
            ws.cell(row=r, column=9).value = f"=G{r}*(1+H{r}/100)"
            ws.cell(row=r, column=9).fill = FORMULA_FILL
            ws.cell(row=r, column=9).number_format = "0.000"
            ws.cell(row=r, column=17).value = f'=IF(OR(O{r}="",P{r}=""),"",MIN(O{r},P{r})-N{r})'
            ws.cell(row=r, column=17).fill = FORMULA_FILL
            ws.cell(row=r, column=17).number_format = "#,##0"
            for c in [12,13,14,15,16]:
                ws.cell(row=r, column=c).number_format = "#,##0"
            style_rows(ws, r, r, MAX_COL, example=True)

        last = 5 + len(mats)
        for r in range(last, last + 20):
            ws.cell(row=r, column=9).value = f'=IF(G{r}="","",G{r}*(1+H{r}/100))'
            ws.cell(row=r, column=9).fill = FORMULA_FILL
            ws.cell(row=r, column=9).number_format = "0.000"
            ws.cell(row=r, column=17).value = f'=IF(OR(O{r}="",P{r}=""),"",MIN(O{r},P{r})-N{r})'
            ws.cell(row=r, column=17).fill = FORMULA_FILL
            ws.cell(row=r, column=17).number_format = "#,##0"
            for c in range(1, MAX_COL + 1):
                ws.cell(row=r, column=c).border = THIN

        nr = last + 21
        note(ws, nr, 1, "※ 파란 셀(실소요량, 절약금액)은 수식 자동계산", MAX_COL)
        note(ws, nr+1, 1, "※ 빈 행에 자재 추가 → 수식 자동 적용", MAX_COL)

        set_widths(ws, [10,16,12,12,7,5,8,7,8,10,8,9,9,9,9,10,9,7,7,7,9,7,7,7,7,25])
        ws.freeze_panes = "C5"
    else:
        ws = wb.create_sheet(name)
        ws.sheet_properties.tabColor = color
        MAX_COL = 26
        ws.merge_cells(f"A1:{get_column_letter(MAX_COL)}1")
        ws.cell(row=1, column=1, value=f"{pname} 자재").font = TITLE
        ws.merge_cells(f"A2:{get_column_letter(MAX_COL)}2")
        ws.cell(row=2, column=1, value=f"공정코드: {pcode} | 자재 + 소요량 + 가격 + 배송 통합").font = SUBTITLE
        ws.row_dimensions[1].height = 30

        sections = [
            (1, 9, "기본 정보 + 소요량"),
            (10, 17, "가격"),
            (18, 21, "물류/배송"),
            (22, 24, "샘플"),
            (25, 26, "등급"),
        ]
        row = 3
        for s, e, label in sections:
            ws.merge_cells(start_row=row, start_column=s, end_row=row, end_column=e)
            cell = ws.cell(row=row, column=s, value=label)
            cell.font = SECTION_FONT
            cell.fill = SECTION_FILL
            cell.alignment = Alignment(horizontal="center")
            for c in range(s, e + 1):
                ws.cell(row=row, column=c).fill = SECTION_FILL
                ws.cell(row=row, column=c).border = THIN

        headers = [
            "자재코드", "품명", "규격", "용도", "필수/선택",
            "단위", "1㎡당\n소요량", "로스율\n(%)", "실소요량\n(1㎡)",
            "판매단위", "판매단위\n수량", "도매가", "업체가", "판매가",
            "쿠팡가", "네이버\n최저가", "절약금액",
            "무게(kg)", "배송방법", "배송일", "공급사코드",
            "샘플가능", "샘플크기", "샘플비용",
            "등급", "한줄요약",
        ]
        row = 4
        for j, h in enumerate(headers, 1):
            ws.cell(row=row, column=j, value=h)
        style_header(ws, row, MAX_COL)
        ws.row_dimensions[4].height = 36

        for mi, m in enumerate(mats):
            r = 5 + mi
            vals = [
                m["code"], m["name"], m["spec"], m["purpose"], m["required"],
                m["unit"], m["per_sqm"], m["loss_pct"], None,
                m["pack_unit"], m["pack_qty"], m["wholesale"], m["dealer"], m["retail"],
                m.get("coupang", ""), m.get("naver", ""), None,
                m.get("weight", ""), m.get("shipping", ""), m.get("ship_days", ""), m.get("supplier", ""),
                m.get("sample_yn", "N"), m.get("sample_size", "—"), m.get("sample_cost", "—"),
                m.get("grade", "표준"), m.get("summary", ""),
            ]
            for j, v in enumerate(vals):
                if v is not None:
                    ws.cell(row=r, column=j+1, value=v)
            ws.cell(row=r, column=7).number_format = "0.00"
            ws.cell(row=r, column=9).value = f"=G{r}*(1+H{r}/100)"
            ws.cell(row=r, column=9).fill = FORMULA_FILL
            ws.cell(row=r, column=9).number_format = "0.000"
            ws.cell(row=r, column=17).value = f'=IF(OR(O{r}="",P{r}=""),"",MIN(O{r},P{r})-N{r})'
            ws.cell(row=r, column=17).fill = FORMULA_FILL
            ws.cell(row=r, column=17).number_format = "#,##0"
            for c in [12,13,14,15,16]:
                ws.cell(row=r, column=c).number_format = "#,##0"
            style_rows(ws, r, r, MAX_COL, example=True)

        last = 5 + len(mats)
        for r in range(last, last + 20):
            ws.cell(row=r, column=9).value = f'=IF(G{r}="","",G{r}*(1+H{r}/100))'
            ws.cell(row=r, column=9).fill = FORMULA_FILL
            ws.cell(row=r, column=9).number_format = "0.000"
            ws.cell(row=r, column=17).value = f'=IF(OR(O{r}="",P{r}=""),"",MIN(O{r},P{r})-N{r})'
            ws.cell(row=r, column=17).fill = FORMULA_FILL
            ws.cell(row=r, column=17).number_format = "#,##0"
            for c in range(1, MAX_COL + 1):
                ws.cell(row=r, column=c).border = THIN

        nr = last + 21
        note(ws, nr, 1, "※ 파란 셀(실소요량, 절약금액)은 수식 자동계산", MAX_COL)
        note(ws, nr+1, 1, "※ 빈 행에 자재 추가 → 수식 자동 적용", MAX_COL)

        set_widths(ws, [10,16,12,12,7,5,8,7,8,10,8,9,9,9,9,10,9,7,7,7,9,7,7,7,7,25])
        ws.freeze_panes = "C5"


# ─────────────────────────────────
# 공정 마스터
# ─────────────────────────────────
ws_proc = wb.create_sheet("공정마스터")
ws_proc.sheet_properties.tabColor = "ED8936"

ws_proc.merge_cells("A1:N1")
ws_proc.cell(row=1, column=1, value="공정 마스터").font = TITLE
ws_proc.merge_cells("A2:N2")
ws_proc.cell(row=2, column=1, value="공정 1개 = 1행 | 순서·단가·가이드 통합").font = SUBTITLE

sections3 = [(1, 7, "기본 + 순서"), (8, 12, "적정 단가"), (13, 14, "가이드")]
for s, e, label in sections3:
    ws_proc.merge_cells(start_row=3, start_column=s, end_row=3, end_column=e)
    cell = ws_proc.cell(row=3, column=s, value=label)
    cell.font = SECTION_FONT; cell.fill = SECTION_FILL; cell.alignment = Alignment(horizontal="center")
    for c in range(s, e + 1):
        ws_proc.cell(row=3, column=c).fill = SECTION_FILL
        ws_proc.cell(row=3, column=c).border = THIN

h3 = ["공정코드","공정명","순서","선행공정","시공일","양생일","합계일",
      "자재비(㎡당)","인건비(㎡당)","적정합계","기능공일당","㎡당공수",
      "셀프난이도","자재시트"]
for i, h in enumerate(h3, 1):
    ws_proc.cell(row=4, column=i, value=h)
style_header(ws_proc, 4, 14)

proc_data = [
    ["PROC-01","철거",1,"—",2,0,"—","—","—","20만/일",0,"하","—"],
    ["PROC-02","설비(배관)",2,"PROC-01",2,0,"—","—","—","25만/일",0,"상","—"],
    ["PROC-03","전기",3,"PROC-01",1,0,"—","—","—","25만/일",0,"상","—"],
    ["PROC-04","방수",4,"PROC-02",1,2,"1~2만","2~3만","3~5만/㎡","25만/일",0.1,"중","자재_방수"],
    ["PROC-05","창호",5,"PROC-01",1,0,"—","—","—","—",0,"상","—"],
    ["PROC-06","목공",6,"PROC-04",3,1,"2~5만","3~5만","5~10만/㎡","28만/일",0.12,"상","—"],
    ["PROC-07","필름",7,"PROC-06",1,0,"1~3만","1~2만","2~5만/㎡","20만/일",0.05,"중","—"],
    ["PROC-08","타일",8,"PROC-04",3,1,"3~8만","3~5만","6~13만/㎡","30만/일",0.15,"상","자재_타일"],
    ["PROC-09","도장",9,"PROC-06",2,1,"0.3~1만","0.5~1만","1~2만/㎡","20만/일",0.05,"하","자재_도장"],
    ["PROC-10","마루/바닥",10,"PROC-09",2,0,"2~6만","1~3만","3~9만/㎡","25만/일",0.08,"중","자재_마루"],
    ["PROC-11","도배",11,"PROC-09",2,0,"0.5~2만","0.5~1만","1~3만/㎡","22만/일",0.06,"하","자재_도배"],
    ["PROC-12","가구/도기",12,"PROC-11",2,0,"—","—","—","—",0,"중","—"],
    ["PROC-13","준공점검",13,"PROC-12",1,0,"—","—","—","—",0,"하","—"],
]
for i, row_data in enumerate(proc_data):
    r = 5 + i
    for j, val in enumerate(row_data):
        ws_proc.cell(row=r, column=j+1, value=val)
    ws_proc.cell(row=r, column=7).value = f"=E{r}+F{r}"
    ws_proc.cell(row=r, column=7).fill = FORMULA_FILL
    style_rows(ws_proc, r, r, 14, example=True)

note(ws_proc, 18, 1, "※ 자재시트 열: 해당 공정의 자재 데이터가 있는 시트명. 클릭 시 이동 가능.", 14)
set_widths(ws_proc, [11,11,6,10,7,7,7,11,11,12,10,9,8,10])
ws_proc.freeze_panes = "C5"


# ─────────────────────────────────
# 호환성
# ─────────────────────────────────
ws_compat = wb.create_sheet("호환성")
ws_compat.sheet_properties.tabColor = "FC8181"

ws_compat.merge_cells("A1:G1")
ws_compat.cell(row=1, column=1, value="자재 호환성").font = TITLE
ws_compat.merge_cells("A2:G2")
ws_compat.cell(row=2, column=1, value="NG/조건부만 기록 (OK는 기록 안 함)").font = SUBTITLE

h4 = ["자재코드A","자재명A","자재코드B","자재명B","판정","사유","대체추천"]
for i, h in enumerate(h4, 1):
    ws_compat.cell(row=4, column=i, value=h)
style_header(ws_compat, 4, 7)

compat_data = [
    ["WP-004","우레탄 방수재","BA-001","타일 본드(시멘트)","NG","우레탄 위 시멘트 본드 부착력 없음","에폭시 본드"],
    ["WP-001","수성 방수액","TL-001","포세린 타일","조건부","완전 경화(24~48h) 후에만 가능","—"],
    ["PT-001","수성 페인트","PT-002","수성 페인트(컬러)","조건부","하도 건조 후 상도 가능. 동시 도포 불가","—"],
]
for i, d in enumerate(compat_data):
    r = 5 + i
    for j, v in enumerate(d):
        ws_compat.cell(row=r, column=j+1, value=v)
    if d[4] == "NG":
        ws_compat.cell(row=r, column=5).font = Font(name="맑은 고딕", bold=True, color="E53E3E", size=10)
    else:
        ws_compat.cell(row=r, column=5).font = Font(name="맑은 고딕", bold=True, color="D69E2E", size=10)
    style_rows(ws_compat, r, r, 7, example=True)

dv = DataValidation(type="list", formula1='"NG,조건부"', allow_blank=True)
ws_compat.add_data_validation(dv)
dv.add("E5:E100")
set_widths(ws_compat, [11,15,11,15,8,30,18])
ws_compat.freeze_panes = "A5"


# ─────────────────────────────────
# 공급사
# ─────────────────────────────────
ws_sup = wb.create_sheet("공급사")
ws_sup.sheet_properties.tabColor = "9F7AEA"

ws_sup.merge_cells("A1:H1")
ws_sup.cell(row=1, column=1, value="공급사 테이블").font = TITLE
ws_sup.merge_cells("A2:H2")
ws_sup.cell(row=2, column=1, value="공급사 1개 = 1행 | 공정별 자재 시트에서 FK로 연결").font = SUBTITLE

h5 = ["공급사코드","공급사명","취급카테고리","리드타임","최소주문","반품조건","결제조건","연락처"]
for i, h in enumerate(h5, 1):
    ws_sup.cell(row=4, column=i, value=h)
style_header(ws_sup, 4, 8)

sup_data = [
    ["SUP-001","OO건자재","방수, 도장","2영업일","10만원","미개봉 7일","월말정산","010-XXXX-XXXX"],
    ["SUP-002","XX세라믹","타일","3~5영업일","30만원","파손만 교환","선결제","010-XXXX-XXXX"],
    ["SUP-003","KCC대리점","도장","1~2영업일","5만원","미개봉 14일","월말정산","010-XXXX-XXXX"],
    ["SUP-004","한화마루","마루","3~5영업일","20만원","미개봉 7일","선결제","010-XXXX-XXXX"],
    ["SUP-005","OO벽지","도배","2~3영업일","10만원","미개봉 7일","월말정산","010-XXXX-XXXX"],
]
for i, d in enumerate(sup_data):
    r = 5 + i
    for j, v in enumerate(d):
        ws_sup.cell(row=r, column=j+1, value=v)
    style_rows(ws_sup, r, r, 8, example=True)
set_widths(ws_sup, [11,14,14,10,10,14,10,16])
ws_sup.freeze_panes = "A5"


# ─────────────────────────────────
# 면적 참조
# ─────────────────────────────────
ws_area = wb.create_sheet("면적참조")
ws_area.sheet_properties.tabColor = "38B2AC"

ws_area.merge_cells("A1:H1")
ws_area.cell(row=1, column=1, value="공간 유형별 표준 면적").font = TITLE
ws_area.merge_cells("A2:H2")
ws_area.cell(row=2, column=1, value="견적 시뮬레이터 기본값 | 사용자 미입력 시 참조").font = SUBTITLE

h6 = ["평형대","전용면적(㎡)","욕실(㎡)","주방(㎡)","거실(㎡)","안방(㎡)","작은방(㎡)","베란다(㎡)"]
for i, h in enumerate(h6, 1):
    ws_area.cell(row=4, column=i, value=h)
style_header(ws_area, 4, 8)

area_data = [
    ["18평",59,3.5,5,14,10,7,5],
    ["25평",84,4.5,7,20,13,9,7],
    ["32평",105,5,8,25,15,10,8],
    ["34평",112,5.5,9,28,16,11,9],
    ["43평",142,6,11,35,20,13,11],
]
for i, d in enumerate(area_data):
    r = 5 + i
    for j, v in enumerate(d):
        ws_area.cell(row=r, column=j+1, value=v)
    style_rows(ws_area, r, r, 8, example=True)
set_widths(ws_area, [10,12,10,10,10,10,10,10])


# ─────────────────────────────────
# 견적 시뮬레이터
# ─────────────────────────────────
ws_est = wb.create_sheet("견적시뮬레이터")
ws_est.sheet_properties.tabColor = "48BB78"

ws_est.merge_cells("A1:M1")
ws_est.cell(row=1, column=1, value="빌드어스 자동 견적 시뮬레이터").font = Font(name="맑은 고딕", bold=True, size=16)
ws_est.merge_cells("A2:M2")
ws_est.cell(row=2, column=1, value="노란 셀에 면적 입력 → 각 공정 시트에서 자재 가져와 자동 계산").font = SUBTITLE

# 입력 영역
style_section(ws_est, 4, 13, "▼ 면적 입력 (노란 셀)")
for i, h in enumerate(["공간", "면적(㎡)"], 1):
    ws_est.cell(row=5, column=i, value=h)
    ws_est.cell(row=5, column=i).font = HEADER_FONT
    ws_est.cell(row=5, column=i).fill = HEADER_FILL
    ws_est.cell(row=5, column=i).border = THIN

spaces = [("욕실 바닥",4.5),("욕실 벽",12),("주방",7),("거실",20),("안방",13),("작은방",9),("베란다",7)]
for i,(sp,val) in enumerate(spaces):
    r = 6+i
    ws_est.cell(row=r, column=1, value=sp).font = NORMAL
    ws_est.cell(row=r, column=1).border = THIN
    c = ws_est.cell(row=r, column=2, value=val)
    c.fill = INPUT_FILL; c.border = INPUT_BORDER
    c.font = Font(name="맑은 고딕", bold=True, size=11)
    c.number_format = "0.0"

# 견적 결과
est_row = 14
style_section(ws_est, est_row, 13, "▼ 견적 결과 — 자동 계산 (면적 변경 시 자동 업데이트)")

est_headers = [
    "공정", "자재명", "규격", "단위",
    "적용면적(㎡)", "1㎡당소요량", "로스율(%)", "실소요량(총)",
    "판매단위\n수량", "필요수량\n(올림)", "단가", "자재비", "시트참조",
]
for i, h in enumerate(est_headers, 1):
    ws_est.cell(row=est_row+1, column=i, value=h)
style_header(ws_est, est_row+1, 13)

# 견적 데이터 — 각 공정 시트 참조
est_items = [
    # (공정, 자재명, 규격, 단위, 면적수식, 소요량, 로스율, 판매단위수량, 단가, 시트명)
    ("방수", "수성 방수액", "18L", "L", "=B6", 0.6, 10, 18, 45000, "자재_방수"),
    ("방수", "부직포", "1m×50m", "㎡", "=B6", 1.05, 5, 50, 15000, "자재_방수"),
    ("방수", "프라이머", "4L", "L", "=B6", 0.1, 5, 4, 12000, "자재_방수"),
    ("타일", "포세린 타일", "600×600", "장", "=B6+B7", 2.8, 10, 8, 6200, "자재_타일"),
    ("타일", "타일 본드", "20kg", "kg", "=B6+B7", 4.5, 5, 20, 12000, "자재_타일"),
    ("타일", "줄눈재", "5kg", "kg", "=B6+B7", 0.5, 10, 5, 6000, "자재_타일"),
    ("타일", "십자 스페이서", "2mm", "개", "=B6+B7", 8, 20, 100, 2500, "자재_타일"),
    ("도장", "수성 페인트", "18L", "L", "=B8+B9+B10", 0.15, 5, 18, 62000, "자재_도장"),
    ("도장", "퍼티", "10kg", "kg", "=B8+B9+B10", 0.3, 10, 10, 10000, "자재_도장"),
    ("도장", "사포", "P180", "장", "=B8+B9+B10", 0.3, 30, 50, 200, "자재_도장"),
    ("도장", "마스킹테이프", "24mm", "롤", "=B8+B9+B10", 0.1, 20, 1, 2500, "자재_도장"),
    ("마루", "강마루", "1210×193", "장", "=B8+B9+B10", 5.35, 5, 8, 6200, "자재_마루"),
    ("마루", "PE 폼 언더레이", "2mm", "㎡", "=B8+B9+B10", 1.0, 5, 50, 1000, "자재_마루"),
    ("마루", "걸레받이", "60mm", "m", "=B8+B9+B10", 0.4, 10, 1, 3000, "자재_마루"),
    ("도배", "합지벽지", "롤(16.5㎡)", "롤", "=B8+B9+B10", 0.07, 10, 1, 12000, "자재_도배"),
    ("도배", "벽지풀", "5kg", "kg", "=B8+B9+B10", 0.15, 5, 5, 7000, "자재_도배"),
]

for i, (proc, name, spec, unit, area_f, usage, loss, pack, price, sheet) in enumerate(est_items):
    r = est_row + 2 + i
    ws_est.cell(row=r, column=1, value=proc).font = NORMAL
    ws_est.cell(row=r, column=2, value=name).font = NORMAL
    ws_est.cell(row=r, column=3, value=spec).font = NORMAL
    ws_est.cell(row=r, column=4, value=unit).font = NORMAL

    ws_est.cell(row=r, column=5).value = area_f
    ws_est.cell(row=r, column=5).fill = FORMULA_FILL
    ws_est.cell(row=r, column=5).number_format = "0.0"

    ws_est.cell(row=r, column=6, value=usage).number_format = "0.00"
    ws_est.cell(row=r, column=7, value=loss)

    # 실소요량(총)
    ws_est.cell(row=r, column=8).value = f"=E{r}*F{r}*(1+G{r}/100)"
    ws_est.cell(row=r, column=8).fill = FORMULA_FILL
    ws_est.cell(row=r, column=8).number_format = "0.00"

    ws_est.cell(row=r, column=9, value=pack)

    # 필요수량(올림)
    ws_est.cell(row=r, column=10).value = f"=CEILING(H{r}/I{r},1)"
    ws_est.cell(row=r, column=10).fill = FORMULA_FILL

    ws_est.cell(row=r, column=11, value=price).number_format = "#,##0"

    # 자재비
    ws_est.cell(row=r, column=12).value = f"=J{r}*K{r}"
    ws_est.cell(row=r, column=12).fill = RESULT_FILL
    ws_est.cell(row=r, column=12).font = Font(name="맑은 고딕", bold=True, size=10)
    ws_est.cell(row=r, column=12).number_format = "#,##0"

    ws_est.cell(row=r, column=13, value=sheet).font = Font(name="맑은 고딕", size=9, color="718096")

    style_rows(ws_est, r, r, 13)

# 공정별 소계
total_start = est_row + 2
total_end = total_start + len(est_items) - 1

# 합계
sum_row = total_end + 1
ws_est.cell(row=sum_row, column=1, value="합계").font = Font(name="맑은 고딕", bold=True, size=13)
ws_est.cell(row=sum_row, column=12).value = f"=SUM(L{total_start}:L{total_end})"
ws_est.cell(row=sum_row, column=12).fill = PatternFill(start_color="C6F6D5", end_color="C6F6D5", fill_type="solid")
ws_est.cell(row=sum_row, column=12).font = Font(name="맑은 고딕", bold=True, size=14, color="22543D")
ws_est.cell(row=sum_row, column=12).number_format = "#,##0"
ws_est.cell(row=sum_row, column=12).border = Border(top=Side(style="double", color="22543D"), bottom=Side(style="double", color="22543D"))

# 공정별 소계
subtotal_row = sum_row + 2
style_section(ws_est, subtotal_row, 13, "▼ 공정별 소계")

procs_for_subtotal = ["방수", "타일", "도장", "마루", "도배"]
for pi, pname in enumerate(procs_for_subtotal):
    r = subtotal_row + 1 + pi
    ws_est.cell(row=r, column=1, value=pname).font = BOLD
    ws_est.cell(row=r, column=1).border = THIN
    formula = f'=SUMIF(A{total_start}:A{total_end},"{pname}",L{total_start}:L{total_end})'
    ws_est.cell(row=r, column=2).value = formula
    ws_est.cell(row=r, column=2).number_format = "#,##0"
    ws_est.cell(row=r, column=2).font = BOLD
    ws_est.cell(row=r, column=2).fill = RESULT_FILL
    ws_est.cell(row=r, column=2).border = THIN

# 비교 표
comp_row = subtotal_row + 1 + len(procs_for_subtotal) + 1
style_section(ws_est, comp_row, 13, "▼ 비용 비교")
ws_est.cell(row=comp_row+1, column=1, value="구분").font = HEADER_FONT
ws_est.cell(row=comp_row+1, column=1).fill = HEADER_FILL
ws_est.cell(row=comp_row+1, column=1).border = THIN
ws_est.cell(row=comp_row+1, column=2, value="금액").font = HEADER_FONT
ws_est.cell(row=comp_row+1, column=2).fill = HEADER_FILL
ws_est.cell(row=comp_row+1, column=2).border = THIN
ws_est.cell(row=comp_row+1, column=3, value="절약").font = HEADER_FONT
ws_est.cell(row=comp_row+1, column=3).fill = HEADER_FILL
ws_est.cell(row=comp_row+1, column=3).border = THIN

r1 = comp_row + 2
ws_est.cell(row=r1, column=1, value="빌드어스 (셀프 자재비)").font = BOLD
ws_est.cell(row=r1, column=2).value = f"=L{sum_row}"
ws_est.cell(row=r1, column=2).font = Font(name="맑은 고딕", bold=True, size=12, color="22543D")
ws_est.cell(row=r1, column=2).number_format = "#,##0"
ws_est.cell(row=r1, column=3, value="—")

r2 = comp_row + 3
ws_est.cell(row=r2, column=1, value="인터넷 구매 시 (추정 +15%)").font = NORMAL
ws_est.cell(row=r2, column=2).value = f"=L{sum_row}*1.15"
ws_est.cell(row=r2, column=2).number_format = "#,##0"
ws_est.cell(row=r2, column=3).value = f"=B{r2}-B{r1}"
ws_est.cell(row=r2, column=3).number_format = "#,##0"
ws_est.cell(row=r2, column=3).fill = RESULT_FILL

r3 = comp_row + 4
ws_est.cell(row=r3, column=1, value="업체 의뢰 시 (자재+인건비)").font = NORMAL
ws_est.cell(row=r3, column=2).value = f"=L{sum_row}*2.5"
ws_est.cell(row=r3, column=2).number_format = "#,##0"
ws_est.cell(row=r3, column=3).value = f"=B{r3}-B{r1}"
ws_est.cell(row=r3, column=3).number_format = "#,##0"
ws_est.cell(row=r3, column=3).fill = RESULT_FILL

for rr in [r1, r2, r3]:
    style_rows(ws_est, rr, rr, 3)

# 사용법 노트
nr = r3 + 2
note(ws_est, nr, 1, "사용법", 13)
ws_est.cell(row=nr, column=1).font = Font(name="맑은 고딕", bold=True, size=12)
notes = [
    "1. 노란 셀에 실제 면적(㎡)을 입력하세요 (기본값: 25평 기준)",
    "2. 자재를 추가/삭제하려면 견적 결과 표에 행을 추가하고 수식을 복사하세요",
    "3. 자재 단가를 변경하려면 해당 공정 시트(자재_방수 등)에서 수정하세요",
    "4. 파란 셀 = 수식(자동계산) | 노란 셀 = 입력 | 초록 셀 = 결과",
    "5. 시트참조 열: 해당 자재의 상세 정보가 있는 시트명",
]
for i, n in enumerate(notes):
    note(ws_est, nr+1+i, 1, n, 13)

set_widths(ws_est, [12, 15, 12, 6, 11, 10, 8, 10, 8, 9, 10, 14, 10])
ws_est.freeze_panes = "A15"


# ─────────────────────────────────
# 저장
# ─────────────────────────────────
filepath = "/Users/kimjiwoong/Projects/buildus-strategy/빌드어스_데이터아키텍처_v2.xlsx"
wb.save(filepath)
print(f"저장 완료: {filepath}")
print(f"시트 수: {len(wb.sheetnames)}")
for name in wb.sheetnames:
    print(f"  - {name}")
