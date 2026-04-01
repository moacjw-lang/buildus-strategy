"""
빌드어스 데이터 아키텍처 엑셀 생성
- 5개 마스터 시트 + 보조 시트 + 견적 시뮬레이터
- 시트 간 수식 연결
"""
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy

# ── 스타일 정의 ──
COLORS = {
    "header_bg": "2D3748",    # 다크 헤더
    "header_font": "FFFFFF",
    "section_bg": "4A5568",   # 섹션 구분
    "section_font": "FFFFFF",
    "example_bg": "F7FAFC",   # 예시 행
    "formula_bg": "EBF8FF",   # 수식 셀
    "result_bg": "F0FFF4",    # 결과 셀
    "warning_bg": "FFF5F5",   # 경고
    "accent": "4299E1",
    "green": "48BB78",
    "orange": "ED8936",
    "red": "FC8181",
}

header_font = Font(name="맑은 고딕", bold=True, color=COLORS["header_font"], size=10)
header_fill = PatternFill(start_color=COLORS["header_bg"], end_color=COLORS["header_bg"], fill_type="solid")
section_font = Font(name="맑은 고딕", bold=True, color=COLORS["section_font"], size=10)
section_fill = PatternFill(start_color=COLORS["section_bg"], end_color=COLORS["section_bg"], fill_type="solid")
example_fill = PatternFill(start_color=COLORS["example_bg"], end_color=COLORS["example_bg"], fill_type="solid")
formula_fill = PatternFill(start_color=COLORS["formula_bg"], end_color=COLORS["formula_bg"], fill_type="solid")
result_fill = PatternFill(start_color=COLORS["result_bg"], end_color=COLORS["result_bg"], fill_type="solid")
warning_fill = PatternFill(start_color=COLORS["warning_bg"], end_color=COLORS["warning_bg"], fill_type="solid")
normal_font = Font(name="맑은 고딕", size=10)
bold_font = Font(name="맑은 고딕", bold=True, size=10)
title_font = Font(name="맑은 고딕", bold=True, size=14)
subtitle_font = Font(name="맑은 고딕", bold=True, size=11, color=COLORS["accent"])

thin_border = Border(
    left=Side(style="thin", color="E2E8F0"),
    right=Side(style="thin", color="E2E8F0"),
    top=Side(style="thin", color="E2E8F0"),
    bottom=Side(style="thin", color="E2E8F0"),
)

def style_header_row(ws, row, max_col):
    """헤더 행 스타일"""
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

def style_section_row(ws, row, max_col, label):
    """섹션 구분 행"""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    cell = ws.cell(row=row, column=1, value=label)
    cell.font = section_font
    cell.fill = section_fill
    cell.alignment = Alignment(horizontal="left", vertical="center")
    for col in range(1, max_col + 1):
        ws.cell(row=row, column=col).fill = section_fill
        ws.cell(row=row, column=col).border = thin_border

def style_data_rows(ws, start_row, end_row, max_col, is_example=False):
    """데이터 행 스타일"""
    for row in range(start_row, end_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = normal_font
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if is_example:
                cell.fill = example_fill

def write_note(ws, row, col, text, merge_end=None):
    """노트 작성"""
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = Font(name="맑은 고딕", size=9, color="718096", italic=True)
    cell.alignment = Alignment(wrap_text=True)
    if merge_end:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=merge_end)

def set_col_widths(ws, widths):
    """열 너비 설정"""
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def add_number_format(ws, row, cols, fmt="#,##0"):
    """숫자 포맷"""
    for col in cols:
        ws.cell(row=row, column=col).number_format = fmt


# ══════════════════════════════════════
wb = openpyxl.Workbook()

# ── 시트 1: 자재 마스터 ──
ws1 = wb.active
ws1.title = "1_자재마스터"
ws1.sheet_properties.tabColor = "4299E1"

headers1 = [
    "자재코드", "품명", "규격", "브랜드", "제조사", "카테고리",
    "단위", "판매단위", "판매단위수량",
    # 가격
    "도매가", "업체가", "판매가", "쿠팡가", "네이버최저가",
    "절약금액", "절약률(%)", "시장가조사일",
    # 물류
    "무게(kg)", "배송방법", "예상배송일", "배송비기준", "공급사코드",
    # 샘플
    "샘플가능", "샘플크기", "샘플무료유료", "샘플비용",
    # 콘텐츠
    "한줄요약", "특징", "주의사항",
    # 등급
    "스타일태그", "등급", "추천공간",
]
max_col1 = len(headers1)

# 제목
ws1.merge_cells("A1:AF1")
ws1.cell(row=1, column=1, value="자재 마스터").font = title_font
ws1.merge_cells("A2:AF2")
ws1.cell(row=2, column=1, value="자재 1개 = 1행 | 가격·배송·샘플·설명·등급 통합").font = subtitle_font
ws1.row_dimensions[1].height = 30
ws1.row_dimensions[2].height = 20

# 섹션 구분 행 (3행)
row = 3
# 기본정보(1-8), 판매단위수량(9), 가격(10-17), 물류(18-22), 샘플(23-26), 콘텐츠(27-29), 등급(30-32)
sections = [
    (1, 9, "기본 정보"),
    (10, 17, "가격"),
    (18, 22, "물류/배송"),
    (23, 26, "샘플"),
    (27, 29, "상품 설명"),
    (30, 32, "스타일/등급"),
]
for start, end, label in sections:
    ws1.merge_cells(start_row=row, start_column=start, end_row=row, end_column=end)
    cell = ws1.cell(row=row, column=start, value=label)
    cell.font = section_font
    cell.fill = section_fill
    cell.alignment = Alignment(horizontal="center")
    for c in range(start, end + 1):
        ws1.cell(row=row, column=c).fill = section_fill
        ws1.cell(row=row, column=c).border = thin_border

# 헤더 행 (4행)
row = 4
for i, h in enumerate(headers1, 1):
    ws1.cell(row=row, column=i, value=h)
style_header_row(ws1, row, max_col1)

# 예시 데이터
examples1 = [
    ["WP-001", "수성 방수액", "18L", "아이나비", "OO화학", "방수",
     "L", "통(18L)", 18,
     35000, 42000, 45000, 58000, 52000, None, None, "2026-04-01",
     20, "화물", "2~3일", "착불/5만이상무료", "SUP-001",
     "N", "—", "—", "",
     "욕실 바닥/벽 방수. 2회 도포로 완벽 방수.", "수성/친환경, 2회도포, 부직포보강 권장", "5도 이하 시공 불가",
     "—", "표준", "욕실, 베란다"],
    ["WP-002", "부직포", "1m폭×50m", "—", "OO섬유", "방수",
     "㎡", "롤(50㎡)", 50,
     12000, 14000, 15000, 22000, 18000, None, None, "2026-04-01",
     3, "택배", "1~2일", "3만이상무료", "SUP-001",
     "N", "—", "—", "",
     "방수 보강용 부직포. 코너/배관부 집중 시공.", "겹침5cm, 코너보강", "접히지 않게 시공",
     "—", "표준", "욕실"],
    ["TL-012", "폴리싱 타일", "600×600", "동서", "동서세라믹", "타일",
     "장", "박스(8장)", 8,
     3500, 4200, 4800, 6500, 5800, None, None, "2026-04-01",
     25, "화물", "3~5일", "착불", "SUP-002",
     "Y", "1장", "무료", 0,
     "광택 폴리싱 타일. 모던/미니멀 욕실에 적합.", "내구성높음, 미끄럼주의", "욕실바닥은 논슬립 권장",
     "모던", "표준", "욕실 벽, 거실"],
    ["PT-005", "수성 페인트", "4L", "KCC", "KCC", "도장",
     "L", "캔(4L)", 4,
     18000, 22000, 25000, 35000, 32000, None, None, "2026-04-01",
     5, "택배", "1~2일", "3만이상무료", "SUP-003",
     "Y", "100ml", "무료", 0,
     "실내 벽면/천장용 수성 페인트.", "냄새적음, 빠른건조, 셀프도장적합", "환기 필수",
     "—", "경제", "거실, 안방, 작은방"],
    ["FL-003", "강마루", "1210×193×8T", "한화", "한화L&C", "마루",
     "장", "박스(8장/1.87㎡)", 8,
     4500, 5500, 6200, 9000, 8200, None, None, "2026-04-01",
     12, "화물", "3~5일", "착불", "SUP-004",
     "Y", "1장", "유료", 3000,
     "클릭 방식 강마루. 셀프 시공 가능.", "클릭시공, 고내구성", "습기 많은 곳 불가",
     "모던", "표준", "거실, 안방"],
]

for i, ex in enumerate(examples1):
    r = 5 + i
    for j, val in enumerate(ex):
        ws1.cell(row=r, column=j + 1, value=val)
    # 수식: 절약금액 = MIN(쿠팡,네이버) - 판매가
    ws1.cell(row=r, column=15).value = f"=MIN(M{r},N{r})-L{r}"
    ws1.cell(row=r, column=15).fill = formula_fill
    ws1.cell(row=r, column=15).number_format = "#,##0"
    # 수식: 절약률 = 절약금액 / MIN(쿠팡,네이버)
    ws1.cell(row=r, column=16).value = f'=IF(MIN(M{r},N{r})=0,"",O{r}/MIN(M{r},N{r})*100)'
    ws1.cell(row=r, column=16).fill = formula_fill
    ws1.cell(row=r, column=16).number_format = "0.0"

    style_data_rows(ws1, r, r, max_col1, is_example=True)
    # 숫자 포맷
    for c in [10, 11, 12, 13, 14]:
        ws1.cell(row=r, column=c).number_format = "#,##0"

# 노트
write_note(ws1, 10, 1, "※ 절약금액/절약률은 수식 자동 계산 (파란 셀)", 32)
write_note(ws1, 11, 1, "※ 판매단위수량: 올림 계산에 사용 (예: 18L통 → 수량 2.97L면 1통)", 32)
write_note(ws1, 12, 1, "※ 시장가조사일: 월 1회 업데이트 권장", 32)

set_col_widths(ws1, [
    11, 15, 12, 10, 10, 8,  # 기본
    5, 12, 10,  # 판매단위
    9, 9, 9, 9, 11, 10, 9, 12,  # 가격
    8, 8, 9, 15, 10,  # 물류
    8, 8, 10, 8,  # 샘플
    30, 25, 18,  # 콘텐츠
    10, 8, 14,  # 등급
])

ws1.freeze_panes = "C5"
ws1.auto_filter.ref = f"A4:AF{4 + len(examples1)}"


# ── 시트 2: 공정-자재 매핑 ──
ws2 = wb.create_sheet("2_공정자재매핑")
ws2.sheet_properties.tabColor = "48BB78"

ws2.merge_cells("A1:J1")
ws2.cell(row=1, column=1, value="공정-자재 매핑").font = title_font
ws2.merge_cells("A2:J2")
ws2.cell(row=2, column=1, value="공정 × 자재 조합 | 소요량·로스율 포함 → 물량 자동 산출").font = subtitle_font

headers2 = [
    "공정코드", "공정명", "자재코드", "자재명", "용도",
    "필수/선택", "단위", "1㎡당 소요량", "로스율(%)",
    "실소요량(1㎡)",  # 수식
]
row = 4
for i, h in enumerate(headers2, 1):
    ws2.cell(row=row, column=i, value=h)
style_header_row(ws2, row, len(headers2))

examples2 = [
    ["PROC-04", "방수", "WP-001", "수성 방수액", "방수층 형성", "필수", "L", 0.6, 10],
    ["PROC-04", "방수", "WP-002", "부직포", "보강재", "필수", "㎡", 1.05, 5],
    ["PROC-04", "방수", "WP-003", "프라이머", "바탕 처리", "선택", "L", 0.1, 5],
    ["PROC-08", "타일", "TL-012", "폴리싱 타일 600×600", "바닥 마감", "필수", "장", 2.8, 10],
    ["PROC-08", "타일", "BA-001", "타일 본드", "접착", "필수", "kg", 4.5, 5],
    ["PROC-08", "타일", "GR-001", "줄눈재", "줄눈", "필수", "kg", 0.5, 10],
    ["PROC-09", "도장", "PT-005", "수성 페인트", "벽면 도장", "필수", "L", 0.15, 5],
    ["PROC-09", "도장", "PT-010", "퍼티", "면 고르기", "필수", "kg", 0.3, 10],
    ["PROC-11", "도배", "WP-050", "합지벽지", "벽면 마감", "필수", "롤", 0.12, 8],
    ["PROC-10", "마루", "FL-003", "강마루", "바닥 마감", "필수", "장", 5.35, 5],
]

for i, ex in enumerate(examples2):
    r = 5 + i
    for j, val in enumerate(ex):
        ws2.cell(row=r, column=j + 1, value=val)
    # 실소요량 = 소요량 × (1 + 로스율/100)
    ws2.cell(row=r, column=10).value = f"=H{r}*(1+I{r}/100)"
    ws2.cell(row=r, column=10).fill = formula_fill
    ws2.cell(row=r, column=10).number_format = "0.000"
    style_data_rows(ws2, r, r, len(headers2), is_example=True)

write_note(ws2, 15, 1, "※ 실소요량 = 소요량 × (1 + 로스율%) — 자동 계산", 10)
write_note(ws2, 16, 1, "※ 견적 시뮬레이터에서 면적 × 실소요량으로 총 물량 산출", 10)

set_col_widths(ws2, [11, 8, 11, 18, 12, 8, 6, 12, 9, 13])
ws2.freeze_panes = "A5"
ws2.auto_filter.ref = f"A4:J{4 + len(examples2)}"


# ── 시트 3: 공정 마스터 ──
ws3 = wb.create_sheet("3_공정마스터")
ws3.sheet_properties.tabColor = "ED8936"

ws3.merge_cells("A1:N1")
ws3.cell(row=1, column=1, value="공정 마스터").font = title_font
ws3.merge_cells("A2:N2")
ws3.cell(row=2, column=1, value="공정 1개 = 1행 | 순서·단가·가이드 통합").font = subtitle_font

# 섹션
row = 3
sections3 = [(1, 7, "기본 + 순서"), (8, 12, "적정 단가"), (13, 14, "가이드")]
for s, e, label in sections3:
    ws3.merge_cells(start_row=row, start_column=s, end_row=row, end_column=e)
    cell = ws3.cell(row=row, column=s, value=label)
    cell.font = section_font
    cell.fill = section_fill
    cell.alignment = Alignment(horizontal="center")
    for c in range(s, e + 1):
        ws3.cell(row=row, column=c).fill = section_fill
        ws3.cell(row=row, column=c).border = thin_border

headers3 = [
    "공정코드", "공정명", "순서", "선행공정", "시공일", "양생일", "합계일",
    "자재비(㎡당)", "인건비(㎡당)", "적정합계범위", "기능공일당", "㎡당공수(인·일)",
    "셀프난이도", "가이드문서",
]
row = 4
for i, h in enumerate(headers3, 1):
    ws3.cell(row=row, column=i, value=h)
style_header_row(ws3, row, len(headers3))

examples3 = [
    ["PROC-01", "철거", 1, "—", 2, 0, None, "—", "—", "—", "20만/일", 0, "하", ""],
    ["PROC-02", "설비(배관)", 2, "PROC-01", 2, 0, None, "—", "—", "—", "25만/일", 0, "상", ""],
    ["PROC-03", "전기", 3, "PROC-01", 1, 0, None, "—", "—", "—", "25만/일", 0, "상", ""],
    ["PROC-04", "방수", 4, "PROC-02", 1, 2, None, "1~2만", "2~3만", "3~5만/㎡", "25만/일", 0.1, "중", "방수_시공가이드.md"],
    ["PROC-05", "창호", 5, "PROC-01", 1, 0, None, "—", "—", "—", "—", 0, "상", ""],
    ["PROC-06", "목공", 6, "PROC-04", 3, 1, None, "2~5만", "3~5만", "5~10만/㎡", "28만/일", 0.12, "상", ""],
    ["PROC-07", "필름", 7, "PROC-06", 1, 0, None, "1~3만", "1~2만", "2~5만/㎡", "20만/일", 0.05, "중", ""],
    ["PROC-08", "타일", 8, "PROC-04", 3, 1, None, "3~8만", "3~5만", "6~13만/㎡", "30만/일", 0.15, "상", "타일_시공가이드.md"],
    ["PROC-09", "도장", 9, "PROC-06", 2, 1, None, "0.3~1만", "0.5~1만", "1~2만/㎡", "20만/일", 0.05, "하", "도장_시공가이드.md"],
    ["PROC-10", "마루/바닥", 10, "PROC-09", 2, 0, None, "2~6만", "1~3만", "3~9만/㎡", "25만/일", 0.08, "중", "마루_시공가이드.md"],
    ["PROC-11", "도배", 11, "PROC-09", 2, 0, None, "0.5~2만", "0.5~1만", "1~3만/㎡", "22만/일", 0.06, "하", "도배_시공가이드.md"],
    ["PROC-12", "가구/도기", 12, "PROC-11", 2, 0, None, "—", "—", "—", "—", 0, "중", ""],
    ["PROC-13", "준공점검", 13, "PROC-12", 1, 0, None, "—", "—", "—", "—", 0, "하", ""],
]

for i, ex in enumerate(examples3):
    r = 5 + i
    for j, val in enumerate(ex):
        ws3.cell(row=r, column=j + 1, value=val)
    # 합계일 수식
    ws3.cell(row=r, column=7).value = f"=E{r}+F{r}"
    ws3.cell(row=r, column=7).fill = formula_fill
    style_data_rows(ws3, r, r, len(headers3), is_example=True)

set_col_widths(ws3, [11, 11, 6, 10, 7, 7, 7, 11, 11, 12, 10, 12, 8, 18])
ws3.freeze_panes = "C5"


# ── 시트 4: 호환성 ──
ws4 = wb.create_sheet("4_호환성")
ws4.sheet_properties.tabColor = "FC8181"

ws4.merge_cells("A1:F1")
ws4.cell(row=1, column=1, value="자재 호환성").font = title_font
ws4.merge_cells("A2:F2")
ws4.cell(row=2, column=1, value="NG/조건부만 기록 (OK는 기록 안 함)").font = subtitle_font

headers4 = ["자재코드A", "자재명A", "자재코드B", "자재명B", "판정", "사유", "대체추천"]
row = 4
for i, h in enumerate(headers4, 1):
    ws4.cell(row=row, column=i, value=h)
style_header_row(ws4, row, len(headers4))

examples4 = [
    ["WP-010", "유성 방수재", "BA-003", "수성 타일본드", "NG", "유성 위에 수성 부착력 없음", "BA-007 (에폭시)"],
    ["WP-020", "시멘트 방수재", "TL-050", "PVC 타일", "조건부", "완전 경화(72시간) 후에만 가능", "—"],
    ["PT-005", "수성 페인트", "PT-020", "유성 페인트", "NG", "같은 면에 혼용 불가", "수성으로 통일"],
]
for i, ex in enumerate(examples4):
    r = 5 + i
    for j, val in enumerate(ex):
        ws4.cell(row=r, column=j + 1, value=val)
    # NG 빨간색
    if ex[4] == "NG":
        ws4.cell(row=r, column=5).font = Font(name="맑은 고딕", bold=True, color="E53E3E", size=10)
    elif ex[4] == "조건부":
        ws4.cell(row=r, column=5).font = Font(name="맑은 고딕", bold=True, color="D69E2E", size=10)
    style_data_rows(ws4, r, r, len(headers4), is_example=True)

# 판정 드롭다운
dv = DataValidation(type="list", formula1='"NG,조건부"', allow_blank=True)
dv.error = "NG 또는 조건부만 입력"
ws4.add_data_validation(dv)
dv.add(f"E5:E100")

set_col_widths(ws4, [11, 15, 11, 15, 8, 30, 18])
ws4.freeze_panes = "A5"


# ── 시트 5: 공급사 ──
ws5 = wb.create_sheet("5_공급사")
ws5.sheet_properties.tabColor = "9F7AEA"

ws5.merge_cells("A1:H1")
ws5.cell(row=1, column=1, value="공급사 테이블").font = title_font
ws5.merge_cells("A2:H2")
ws5.cell(row=2, column=1, value="공급사 1개 = 1행 | 자재마스터에서 공급사코드(FK)로 연결").font = subtitle_font

headers5 = ["공급사코드", "공급사명", "취급카테고리", "리드타임", "최소주문", "반품조건", "결제조건", "연락처"]
row = 4
for i, h in enumerate(headers5, 1):
    ws5.cell(row=row, column=i, value=h)
style_header_row(ws5, row, len(headers5))

examples5 = [
    ["SUP-001", "OO건자재", "방수, 도장", "2영업일", "10만원", "미개봉 7일", "월말정산", "010-XXXX-XXXX"],
    ["SUP-002", "XX세라믹", "타일", "3~5영업일", "30만원", "파손만 교환", "선결제", "010-XXXX-XXXX"],
    ["SUP-003", "KCC대리점", "도장", "1~2영업일", "5만원", "미개봉 14일", "월말정산", "010-XXXX-XXXX"],
    ["SUP-004", "한화마루", "마루", "3~5영업일", "20만원", "미개봉 7일", "선결제", "010-XXXX-XXXX"],
]
for i, ex in enumerate(examples5):
    r = 5 + i
    for j, val in enumerate(ex):
        ws5.cell(row=r, column=j + 1, value=val)
    style_data_rows(ws5, r, r, len(headers5), is_example=True)

set_col_widths(ws5, [11, 14, 14, 10, 10, 14, 10, 16])
ws5.freeze_panes = "A5"


# ── 시트 6: 면적 참조 ──
ws6 = wb.create_sheet("6_면적참조")
ws6.sheet_properties.tabColor = "38B2AC"

ws6.merge_cells("A1:H1")
ws6.cell(row=1, column=1, value="공간 유형별 표준 면적").font = title_font
ws6.merge_cells("A2:H2")
ws6.cell(row=2, column=1, value="비용 분석기 기본값 | 사용자가 직접 입력 안 할 때 참조").font = subtitle_font

headers6 = ["평형대", "전용면적(㎡)", "욕실(㎡)", "주방(㎡)", "거실(㎡)", "안방(㎡)", "작은방(㎡)", "베란다(㎡)"]
row = 4
for i, h in enumerate(headers6, 1):
    ws6.cell(row=row, column=i, value=h)
style_header_row(ws6, row, len(headers6))

examples6 = [
    ["18평", 59, 3.5, 5, 14, 10, 7, 5],
    ["25평", 84, 4.5, 7, 20, 13, 9, 7],
    ["32평", 105, 5, 8, 25, 15, 10, 8],
    ["34평", 112, 5.5, 9, 28, 16, 11, 9],
    ["43평", 142, 6, 11, 35, 20, 13, 11],
]
for i, ex in enumerate(examples6):
    r = 5 + i
    for j, val in enumerate(ex):
        ws6.cell(row=r, column=j + 1, value=val)
    style_data_rows(ws6, r, r, len(headers6), is_example=True)

set_col_widths(ws6, [10, 12, 10, 10, 10, 10, 10, 10])


# ══════════════════════════════════════
# 시트 7: 견적 시뮬레이터 (핵심!)
# ══════════════════════════════════════
ws7 = wb.create_sheet("견적시뮬레이터")
ws7.sheet_properties.tabColor = "48BB78"

ws7.merge_cells("A1:L1")
c = ws7.cell(row=1, column=1, value="빌드어스 자동 견적 시뮬레이터")
c.font = Font(name="맑은 고딕", bold=True, size=16)

ws7.merge_cells("A2:L2")
ws7.cell(row=2, column=1, value="면적 입력 → 자재 물량·비용 자동 산출 | 자재마스터·매핑 테이블 연동").font = subtitle_font

# ── 입력 영역 ──
row = 4
style_section_row(ws7, row, 12, "▼ 입력 영역 — 아래 노란 셀에 면적(㎡)을 입력하세요")

ws7.cell(row=5, column=1, value="공간").font = bold_font
ws7.cell(row=5, column=2, value="면적(㎡)").font = bold_font
ws7.cell(row=5, column=1).fill = header_fill
ws7.cell(row=5, column=1).font = header_font
ws7.cell(row=5, column=2).fill = header_fill
ws7.cell(row=5, column=2).font = header_font
for c in [1,2]:
    ws7.cell(row=5, column=c).border = thin_border

input_fill = PatternFill(start_color="FFFFF0", end_color="FFFFF0", fill_type="solid")
input_border = Border(
    left=Side(style="medium", color="D69E2E"),
    right=Side(style="medium", color="D69E2E"),
    top=Side(style="medium", color="D69E2E"),
    bottom=Side(style="medium", color="D69E2E"),
)

spaces = ["욕실 바닥", "욕실 벽", "주방", "거실", "안방", "작은방", "베란다"]
space_defaults = [4.5, 12, 7, 20, 13, 9, 7]

for i, (space, default) in enumerate(zip(spaces, space_defaults)):
    r = 6 + i
    ws7.cell(row=r, column=1, value=space).font = normal_font
    ws7.cell(row=r, column=1).border = thin_border
    cell = ws7.cell(row=r, column=2, value=default)
    cell.fill = input_fill
    cell.border = input_border
    cell.font = Font(name="맑은 고딕", bold=True, size=11)
    cell.number_format = "0.0"

# ── 공정 선택 ──
row = 14
style_section_row(ws7, row, 12, "▼ 공정 선택 — 시공할 공정에 O 입력")

ws7.cell(row=15, column=1, value="공정").font = header_font
ws7.cell(row=15, column=1).fill = header_fill
ws7.cell(row=15, column=2, value="시공 여부").font = header_font
ws7.cell(row=15, column=2).fill = header_fill
ws7.cell(row=15, column=3, value="적용 공간").font = header_font
ws7.cell(row=15, column=3).fill = header_fill
for c in [1,2,3]:
    ws7.cell(row=15, column=c).border = thin_border

processes = [
    ("방수", "O", "욕실 바닥"),
    ("타일", "O", "욕실 바닥+벽"),
    ("도장", "O", "거실+안방+작은방"),
    ("마루", "O", "거실+안방+작은방"),
    ("도배", "O", "거실+안방+작은방"),
]
for i, (proc, yn, space) in enumerate(processes):
    r = 16 + i
    ws7.cell(row=r, column=1, value=proc).font = normal_font
    ws7.cell(row=r, column=1).border = thin_border
    cell = ws7.cell(row=r, column=2, value=yn)
    cell.fill = input_fill
    cell.border = input_border
    cell.font = Font(name="맑은 고딕", bold=True, size=11)
    cell.alignment = Alignment(horizontal="center")
    ws7.cell(row=r, column=3, value=space).font = normal_font
    ws7.cell(row=r, column=3).border = thin_border

# ── 견적 결과 ──
row = 22
style_section_row(ws7, row, 12, "▼ 견적 결과 — 자동 계산")

result_headers = [
    "공정", "자재명", "규격", "단위",
    "적용면적(㎡)", "1㎡당소요량", "로스율(%)", "실소요량(총)",
    "판매단위", "필요수량(올림)", "단가", "자재비",
]
row = 23
for i, h in enumerate(result_headers, 1):
    ws7.cell(row=row, column=i, value=h)
style_header_row(ws7, row, len(result_headers))

# 견적 상세 — 방수
estimate_rows = [
    # (공정, 자재명, 규격, 단위, 면적참조, 소요량, 로스율, 판매단위수량, 단가)
    ("방수", "수성 방수액", "18L", "L", "B6", 0.6, 10, 18, 45000),
    ("방수", "부직포", "1m×50m", "㎡", "B6", 1.05, 5, 50, 15000),
    ("타일", "폴리싱 타일", "600×600", "장", "B6+B7", 2.8, 10, 8, 4800),
    ("타일", "타일 본드", "20kg", "kg", "B6+B7", 4.5, 5, 20, 18000),
    ("타일", "줄눈재", "5kg", "kg", "B6+B7", 0.5, 10, 5, 8000),
    ("도장", "수성 페인트", "4L", "L", "B8+B9+B10", 0.15, 5, 4, 25000),
    ("도장", "퍼티", "10kg", "kg", "B8+B9+B10", 0.3, 10, 10, 12000),
    ("마루", "강마루", "1210×193", "장", "B8+B9+B10", 5.35, 5, 8, 6200),
    ("도배", "합지벽지", "롤(16.5㎡)", "롤", "B8+B9+B10", 0.12, 8, 1, 15000),
]

for i, (proc, name, spec, unit, area_ref, usage, loss, pack_qty, price) in enumerate(estimate_rows):
    r = 24 + i
    ws7.cell(row=r, column=1, value=proc).font = normal_font
    ws7.cell(row=r, column=2, value=name).font = normal_font
    ws7.cell(row=r, column=3, value=spec).font = normal_font
    ws7.cell(row=r, column=4, value=unit).font = normal_font

    # E: 적용면적
    ws7.cell(row=r, column=5).value = f"={area_ref}"
    ws7.cell(row=r, column=5).fill = formula_fill
    ws7.cell(row=r, column=5).number_format = "0.0"

    # F: 소요량
    ws7.cell(row=r, column=6, value=usage).font = normal_font
    ws7.cell(row=r, column=6).number_format = "0.00"

    # G: 로스율
    ws7.cell(row=r, column=7, value=loss).font = normal_font

    # H: 실소요량(총) = 면적 × 소요량 × (1+로스율/100)
    ws7.cell(row=r, column=8).value = f"=E{r}*F{r}*(1+G{r}/100)"
    ws7.cell(row=r, column=8).fill = formula_fill
    ws7.cell(row=r, column=8).number_format = "0.00"

    # I: 판매단위수량
    ws7.cell(row=r, column=9, value=pack_qty).font = normal_font

    # J: 필요수량(올림) = CEILING(실소요량 / 판매단위수량, 1)
    ws7.cell(row=r, column=10).value = f"=CEILING(H{r}/I{r},1)"
    ws7.cell(row=r, column=10).fill = formula_fill

    # K: 단가
    ws7.cell(row=r, column=11, value=price).font = normal_font
    ws7.cell(row=r, column=11).number_format = "#,##0"

    # L: 자재비 = 필요수량 × 단가
    ws7.cell(row=r, column=12).value = f"=J{r}*K{r}"
    ws7.cell(row=r, column=12).fill = result_fill
    ws7.cell(row=r, column=12).font = Font(name="맑은 고딕", bold=True, size=10)
    ws7.cell(row=r, column=12).number_format = "#,##0"

    style_data_rows(ws7, r, r, len(result_headers))

# 합계 행
total_row = 24 + len(estimate_rows)
ws7.cell(row=total_row, column=1, value="합계").font = Font(name="맑은 고딕", bold=True, size=12)
ws7.cell(row=total_row, column=12).value = f"=SUM(L24:L{total_row - 1})"
ws7.cell(row=total_row, column=12).fill = PatternFill(start_color="C6F6D5", end_color="C6F6D5", fill_type="solid")
ws7.cell(row=total_row, column=12).font = Font(name="맑은 고딕", bold=True, size=14, color="22543D")
ws7.cell(row=total_row, column=12).number_format = "#,##0"
ws7.cell(row=total_row, column=12).border = Border(
    top=Side(style="double", color="22543D"),
    bottom=Side(style="double", color="22543D"),
)

# 비교 표
comp_row = total_row + 2
style_section_row(ws7, comp_row, 12, "▼ 비용 비교")

headers_comp = ["구분", "금액", "절약"]
for i, h in enumerate(headers_comp, 1):
    ws7.cell(row=comp_row+1, column=i, value=h)
style_header_row(ws7, comp_row+1, 3)

r = comp_row + 2
ws7.cell(row=r, column=1, value="빌드어스 (셀프 자재비)").font = bold_font
ws7.cell(row=r, column=2).value = f"=L{total_row}"
ws7.cell(row=r, column=2).font = Font(name="맑은 고딕", bold=True, size=12, color="22543D")
ws7.cell(row=r, column=2).number_format = "#,##0"
ws7.cell(row=r, column=3, value="—").font = normal_font

r2 = comp_row + 3
ws7.cell(row=r2, column=1, value="인터넷 최저가 구매 시 (추정)").font = normal_font
ws7.cell(row=r2, column=2).value = f"=L{total_row}*1.15"
ws7.cell(row=r2, column=2).number_format = "#,##0"
ws7.cell(row=r2, column=3).value = f"=B{r2}-B{r}"
ws7.cell(row=r2, column=3).number_format = "#,##0"
ws7.cell(row=r2, column=3).fill = result_fill

r3 = comp_row + 4
ws7.cell(row=r3, column=1, value="업체 의뢰 시 (자재+인건비)").font = normal_font
ws7.cell(row=r3, column=2).value = f"=L{total_row}*2.5"
ws7.cell(row=r3, column=2).number_format = "#,##0"
ws7.cell(row=r3, column=3).value = f"=B{r3}-B{r}"
ws7.cell(row=r3, column=3).number_format = "#,##0"
ws7.cell(row=r3, column=3).fill = result_fill

for rr in [r, r2, r3]:
    style_data_rows(ws7, rr, rr, 3)

# 사용법 노트
note_row = comp_row + 7
ws7.merge_cells(f"A{note_row}:L{note_row}")
ws7.cell(row=note_row, column=1, value="사용법").font = Font(name="맑은 고딕", bold=True, size=12)
notes = [
    "1. 노란색 셀에 실제 면적(㎡)을 입력하세요 (기본값: 25평 기준)",
    "2. 공정 선택에서 시공할 공정에 O를 입력하세요",
    "3. 견적 결과가 자동으로 계산됩니다",
    "4. 자재 단가를 변경하면 자재마스터 시트의 판매가를 수정하세요",
    "5. 새로운 자재를 추가하려면 견적 결과 표에 행을 추가하고 수식을 복사하세요",
    "※ 파란 셀 = 수식(자동계산) | 노란 셀 = 입력 | 초록 셀 = 결과",
]
for i, note in enumerate(notes):
    write_note(ws7, note_row + 1 + i, 1, note, 12)

set_col_widths(ws7, [14, 15, 12, 6, 11, 11, 9, 11, 10, 11, 10, 14])
ws7.freeze_panes = "A23"


# ── 시트 8: Q&A (보조) ──
ws8 = wb.create_sheet("8_QA")
ws8.sheet_properties.tabColor = "38B2AC"
ws8.merge_cells("A1:E1")
ws8.cell(row=1, column=1, value="체크리스트 Q&A 변환").font = title_font
headers8 = ["공정", "질문", "답변", "관련자재코드", "출처"]
for i, h in enumerate(headers8, 1):
    ws8.cell(row=3, column=i, value=h)
style_header_row(ws8, 3, len(headers8))
examples8 = [
    ["방수", "욕실 방수 꼭 해야 하나요?", "네, 필수. 안 하면 아래층 누수 → 보상 문제.", "WP-001", "체크리스트04 #3"],
    ["방수", "방수 위에 타일 바로 붙여도 되나요?", "경화(24~48시간) 후 가능. 덜 마른 상태 시공 시 방수층 손상.", "BA-001", "체크리스트04 #7"],
    ["타일", "타일 본드 종류가 중요한가요?", "매우 중요. 방수 위는 반드시 시멘트 기반 or 에폭시 본드 사용.", "BA-001", "체크리스트08 #5"],
]
for i, ex in enumerate(examples8):
    r = 4 + i
    for j, val in enumerate(ex):
        ws8.cell(row=r, column=j + 1, value=val)
    style_data_rows(ws8, r, r, len(headers8), is_example=True)
set_col_widths(ws8, [8, 30, 45, 12, 14])


# ── 시트 9: SEO (보조) ──
ws9 = wb.create_sheet("9_SEO")
ws9.sheet_properties.tabColor = "38B2AC"
ws9.merge_cells("A1:E1")
ws9.cell(row=1, column=1, value="SEO 키워드 리스트").font = title_font
headers9 = ["키워드", "월간검색량", "경쟁도", "검색의도", "연결기능"]
for i, h in enumerate(headers9, 1):
    ws9.cell(row=3, column=i, value=h)
style_header_row(ws9, 3, len(headers9))
examples9 = [
    ["래미안 인테리어 비용", 1200, "중", "정보탐색", "아파트별 비용 페이지"],
    ["셀프 방수 방법", 800, "중", "가이드", "시공가이드 + 마켓플레이스"],
    ["인테리어 견적 적정가", 2400, "높", "비교", "견적 비교기"],
    ["욕실 타일 셀프 시공", 600, "낮", "가이드", "시공가이드"],
    ["25평 인테리어 비용", 3200, "높", "정보탐색", "비용 분석기"],
]
for i, ex in enumerate(examples9):
    r = 4 + i
    for j, val in enumerate(ex):
        ws9.cell(row=r, column=j + 1, value=val)
    style_data_rows(ws9, r, r, len(headers9), is_example=True)
    ws9.cell(row=r, column=2).number_format = "#,##0"
set_col_widths(ws9, [22, 12, 8, 10, 20])


# ══════════════════════════════════════
# 저장
# ══════════════════════════════════════
filepath = "/Users/kimjiwoong/Projects/buildus-strategy/빌드어스_데이터아키텍처.xlsx"
wb.save(filepath)
print(f"저장 완료: {filepath}")
print(f"시트 수: {len(wb.sheetnames)}")
for name in wb.sheetnames:
    print(f"  - {name}")
