#!/usr/bin/env python3
"""
견적서 17개 파일 → 공정별 정리 엑셀 생성
기존 v4 엑셀에 엎지 않고, 추출 데이터만 공정별로 깔끔하게 정리
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, re, unicodedata
from collections import defaultdict

# ── 설정 ──────────────────────────────────────────────
FOLDER = "견적서"
OUTPUT = "견적서_공정별정리.xlsx"

# 표준 공정 매핑 (견적서에 나오는 다양한 이름 → 표준 공정명)
PROCESS_MAP = {
    "가설": "00_가설공사",
    "가설공사": "00_가설공사",
    "철거": "01_철거공사",
    "철거 공사": "01_철거공사",
    "철거공사": "01_철거공사",
    "철거 / 설비 공사": "01_철거_설비공사",
    "철거/설비 공사": "01_철거_설비공사",
    "철거 / 설비": "01_철거_설비공사",
    "설비": "02_설비공사",
    "설비 공사": "02_설비공사",
    "설비공사": "02_설비공사",
    "필름": "03_필름공사",
    "필름 공사": "03_필름공사",
    "필름공사": "03_필름공사",
    "가구": "04_가구공사",
    "가구 공사": "04_가구공사",
    "가구공사": "04_가구공사",
    "전기": "05_전기공사",
    "전기 공사": "05_전기공사",
    "전기공사": "05_전기공사",
    "화장실": "06_화장실공사",
    "화장실 공사": "06_화장실공사",
    "화장실공사": "06_화장실공사",
    "타일": "07_타일공사",
    "타일 공사": "07_타일공사",
    "타일공사": "07_타일공사",
    "도배": "08_도배공사",
    "도배 공사": "08_도배공사",
    "도배공사": "08_도배공사",
    "도장": "09_도장공사",
    "도장 공사": "09_도장공사",
    "도장공사": "09_도장공사",
    "마루": "10_마루공사",
    "마루 공사": "10_마루공사",
    "마루공사": "10_마루공사",
    "목공": "11_목공사",
    "목공사": "11_목공사",
    "목 공사": "11_목공사",
    "금속": "12_금속공사",
    "금속 공사": "12_금속공사",
    "금속공사": "12_금속공사",
    "데크": "13_데크공사",
    "데크 공사": "13_데크공사",
    "데크공사": "13_데크공사",
    "추가": "14_추가변경",
    "추가/ 변경공사": "14_추가변경",
    "추가/변경공사": "14_추가변경",
}

# 대표 공정 이름 (시트 탭용, 공정 목록에 나오는 순서대로)
MAIN_PROCESSES = [
    "00_가설공사",
    "01_철거공사", "01_철거_설비공사",
    "02_설비공사",
    "03_필름공사",
    "04_가구공사",
    "05_전기공사",
    "06_화장실공사",
    "07_타일공사",
    "08_도배공사",
    "09_도장공사",
    "10_마루공사",
    "11_목공사",
    "12_금속공사", "13_데크공사", "14_추가변경",
]

# 프로젝트 약칭 매핑
SHORT_NAMES = {
    "251230_푸르지오": "푸르지오본",
    "260127_푸르지오": "푸르지오추가",
    "260305_만현마을": None,  # 아래에서 타일/필름 구분
    "금호어울림": "금호어울림",
    "김성용": "김성용",
    "김은경": "김은경",
    "린스트라우스": "린스트라우스",
    "백현마을": "백현마을",
    "산본하이어스": "산본하이어스",
    "이선": "이선",
    "이성숙": "이성숙",
    "전기영": "전기영",
    "진산마을": "진산마을",
    "행당동": "행당동",
}


def get_short_name(filename):
    """파일명에서 프로젝트 약칭 추출 (macOS NFD 대응)"""
    fn = unicodedata.normalize('NFC', filename)
    for key, val in SHORT_NAMES.items():
        if key in fn:
            if val is not None:
                return val
            # 만현마을: 타일/필름 구분
            if "타일" in fn:
                return "만현타일"
            elif "필름" in fn:
                return "만현필름"
            return "만현마을"
    return fn[:10]


def normalize_process(name):
    """공정명을 표준화"""
    if not name:
        return None
    name = name.strip()
    # 번호 접두사 제거: "1. 철거 공사" → "철거 공사"
    name = re.sub(r'^[A-Za-z0-9]+\.\s*', '', name)
    # 숫자만 있는 경우 스킵
    if re.match(r'^[\d.]+$', name):
        return None
    if name in PROCESS_MAP:
        return PROCESS_MAP[name]
    # 부분 매칭 시도
    for key, val in PROCESS_MAP.items():
        if key in name:
            return val
    return None


def is_section_header(row_vals, col_b):
    """갑지-1 요약 영역의 공정 헤더인지 판별 (실제 항목이 아닌 소계 행)"""
    # 컬럼 B에 "공사" 포함되면서 숫자가 아닌 항목
    if not col_b:
        return False
    col_b = str(col_b).strip()
    # 갑지-1 스타일 소계 행: "1. 철거 공사" 같은 패턴
    if normalize_process(col_b):
        return True
    return False


def is_subcategory_header(col_a, col_b):
    """서브 카테고리 헤더인지 판별 (예: "주방 가구", "천장 공사")"""
    if not col_b:
        return False
    a = str(col_a).strip() if col_a else ""
    b = str(col_b).strip()
    # -1 형식의 서브 카테고리
    if a in ["-1", "-2", "-3"]:
        return True
    # "소계", "합계", "직접공사비" 등
    if any(kw in b for kw in ["소계", "합계", "직접공사비 계", "TOTAL", "간접공사비", "일반관리비", "이윤"]):
        return True
    return False


def is_data_row(row_vals):
    """실제 데이터 행인지 판별"""
    # 최소 항목명(col 2)이 있어야 함
    if not row_vals[1]:
        return False
    # 숫자 컬럼 중 하나라도 값이 있어야 함 (단가/금액/수량)
    has_number = False
    for i in [5, 6, 7, 8, 9, 10, 11]:
        if i < len(row_vals) and row_vals[i] is not None:
            try:
                v = float(row_vals[i])
                if v != 0:
                    has_number = True
                    break
            except (ValueError, TypeError):
                pass
    return True  # 단가가 0이라도 항목이 있으면 포함 (비교 참고용)


def safe_num(v):
    """안전하게 숫자 변환"""
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def parse_detail_sheet(filepath, short_name):
    """상세 시트를 파싱하여 공정별 항목 리스트 반환"""
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # 상세 시트 찾기 (보통 마지막 시트 or "상세" or "117동 204호")
    detail_sheet = None
    for sn in wb.sheetnames:
        if sn == "상세" or "동" in sn:
            detail_sheet = wb[sn]
            break
    if not detail_sheet:
        # 마지막 시트 사용
        detail_sheet = wb[wb.sheetnames[-1]]

    ws = detail_sheet
    max_row = ws.max_row
    max_col = ws.max_column

    # 프로젝트명 추출
    project_name = str(ws.cell(1, 1).value or "").replace("PROJECT:", "").strip()

    # 갑지-1 영역 vs 상세 영역 구분
    # 상세 영역: 첫 번째 공정의 실제 항목이 시작되는 곳
    # 패턴: 갑지-1 영역에서 가설공사 / 인테리어 공사 소계가 나온 후,
    #        다시 "가설공사" 헤더가 나오면 그게 상세 영역 시작

    detail_start = None
    gaset_count = 0

    for r in range(1, max_row + 1):
        b = ws.cell(r, 2).value
        if b and "가설공사" in str(b).strip():
            gaset_count += 1
            if gaset_count == 2:
                # 두 번째 "가설공사" = 상세 영역의 가설공사 섹션 시작
                detail_start = r
                break

    if not detail_start:
        # fallback: row 50 이후부터 시작
        detail_start = 50

    # 상세 영역에서 공정별 항목 추출
    items_by_process = defaultdict(list)
    current_process = None
    current_subcat = ""

    for r in range(detail_start, max_row + 1):
        # 행 데이터 읽기
        row_vals = []
        for c in range(1, min(max_col + 1, 20)):
            row_vals.append(ws.cell(r, c).value)

        # 패딩
        while len(row_vals) < 19:
            row_vals.append(None)

        col_a = row_vals[0]  # No/구분
        col_b = row_vals[1]  # Description/항목명

        if not col_b:
            continue

        col_b_str = str(col_b).strip()

        # 소계/합계/총합 행 스킵
        if any(kw in col_b_str for kw in ["소계", "합계", "직접공사비", "TOTAL", "간접공사비", "일반관리비", "이윤", "설계비", "공사보험"]):
            continue

        # "인테리어 공사" 상위 헤더 스킵
        if col_b_str == "인테리어 공사":
            continue

        # 공정 섹션 헤더 감지
        proc = normalize_process(col_b_str)
        if proc:
            # 이 행이 공정 섹션의 시작인지 확인
            # 소계 금액이 있는 행은 갑지-1 스타일 요약
            col_h = safe_num(row_vals[7])  # Amount
            col_j = safe_num(row_vals[9])
            col_l = safe_num(row_vals[11])

            # 갑지-1 스타일 요약행 (큰 소계가 있는)이면 공정 전환
            if col_l and col_l > 100000:
                # 이 행은 갑지-1 소계이므로 공정만 전환
                # 단, 이미 상세 영역에 들어와 있으면 공정 전환 후 스킵
                current_process = proc
                current_subcat = ""
                continue

            # 소계가 없는 공정 헤더 (순수 섹션 시작)
            current_process = proc
            current_subcat = ""
            continue

        # 서브카테고리 헤더 vs 데이터 행 구분
        # -1 행이라도 단가/금액 데이터가 있으면 데이터 행
        a_str = str(col_a).strip() if col_a else ""
        if a_str in ["-1", "-2", "-3"]:
            # 숫자 데이터가 하나라도 있으면 데이터 행 (서브카테고리 아님)
            has_price = False
            for ci in [6, 7, 8, 9, 10, 11]:
                v = safe_num(row_vals[ci]) if ci < len(row_vals) else None
                if v is not None and v != 0:
                    has_price = True
                    break
            # 수량(col 5)만 있어도 데이터 행
            qty = safe_num(row_vals[5])
            if qty is not None and qty > 0:
                has_price = True
            if not has_price:
                # 단가가 있지만 수량이 0인 경우도 데이터로 취급 (비교 참고용)
                for ci in [6, 8, 10]:
                    v = safe_num(row_vals[ci]) if ci < len(row_vals) else None
                    if v is not None and v > 0:
                        has_price = True
                        break
            if not has_price:
                current_subcat = col_b_str
                continue

        # 현재 공정이 없으면 스킵
        if not current_process:
            continue

        # 실제 데이터 행 추출
        item = {
            "프로젝트": short_name,
            "프로젝트명": project_name,
            "구분": str(col_a).strip() if col_a else "",
            "항목명": col_b_str,
            "규격": str(row_vals[2]).strip() if row_vals[2] else "",
            "위치": str(row_vals[3]).strip() if row_vals[3] else "",
            "단위": str(row_vals[4]).strip() if row_vals[4] else "",
            "수량": safe_num(row_vals[5]),
            "자재단가": safe_num(row_vals[6]),
            "자재금액": safe_num(row_vals[7]),
            "노무단가": safe_num(row_vals[8]),
            "노무금액": safe_num(row_vals[9]),
            "합계단가": safe_num(row_vals[10]),
            "합계금액": safe_num(row_vals[11]),
            "비고": str(row_vals[12]).strip() if row_vals[12] else "",
            "서브카테고리": current_subcat,
        }

        items_by_process[current_process].append(item)

    wb.close()
    return items_by_process


def write_output(all_data):
    """공정별 정리 엑셀 생성"""
    wb = openpyxl.Workbook()

    # 스타일 정의
    header_font = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2B3A4E")
    data_font = Font(name="맑은 고딕", size=9)
    num_fmt = '#,##0'
    thin_border = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0'),
    )
    # 프로젝트별 색상
    project_fills = {}
    colors = ["E8F4FD", "FFF3E0", "E8F5E9", "FCE4EC", "F3E5F5",
              "E0F7FA", "FFF8E1", "E8EAF6", "FBE9E7", "F1F8E9",
              "EDE7F6", "E0F2F1", "FFEBEE", "E3F2FD", "F9FBE7", "EFEBE9"]

    # 요약 시트 먼저 생성
    ws_summary = wb.active
    ws_summary.title = "요약"

    # 공정별 시트 생성
    sorted_processes = sorted(all_data.keys())

    summary_rows = []

    for proc_name in sorted_processes:
        items = all_data[proc_name]
        if not items:
            continue

        # 시트명 (31자 제한)
        sheet_name = proc_name[:31]
        ws = wb.create_sheet(title=sheet_name)

        # 헤더
        headers = [
            "프로젝트", "항목명", "규격/제품", "위치/사양", "서브카테고리",
            "단위", "수량", "자재단가", "자재금액", "노무단가", "노무금액",
            "합계단가", "합계금액", "비고"
        ]

        for c, h in enumerate(headers, 1):
            cell = ws.cell(1, c, h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

        ws.freeze_panes = "B2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

        # 프로젝트별 색상 할당
        projects_in_proc = sorted(set(it["프로젝트"] for it in items))
        for i, proj in enumerate(projects_in_proc):
            if proj not in project_fills:
                project_fills[proj] = PatternFill("solid", fgColor=colors[len(project_fills) % len(colors)])

        # 데이터 행 (프로젝트별 정렬)
        items_sorted = sorted(items, key=lambda x: (x["프로젝트"], x.get("서브카테고리", "")))

        for r, item in enumerate(items_sorted, 2):
            vals = [
                item["프로젝트"],
                item["항목명"],
                item["규격"],
                item["위치"],
                item["서브카테고리"],
                item["단위"],
                item["수량"],
                item["자재단가"],
                item["자재금액"],
                item["노무단가"],
                item["노무금액"],
                item["합계단가"],
                item["합계금액"],
                item["비고"],
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(r, c, v)
                cell.font = data_font
                cell.border = thin_border
                # 숫자 서식
                if c >= 7 and c <= 13 and isinstance(v, (int, float)):
                    cell.number_format = num_fmt
                # 프로젝트별 배경색
                cell.fill = project_fills.get(item["프로젝트"], PatternFill())

        # 열 너비
        widths = [10, 25, 20, 15, 12, 5, 7, 10, 12, 10, 12, 10, 12, 12]
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w

        # 요약 정보 수집
        total_items = len(items)
        proj_count = len(projects_in_proc)
        total_amount = sum(it["합계금액"] or 0 for it in items)
        summary_rows.append((proc_name, total_items, proj_count, total_amount, projects_in_proc))

    # ── 요약 시트 ──
    ws_summary.cell(1, 1, "공정별 견적 데이터 요약").font = Font(name="맑은 고딕", bold=True, size=14)
    ws_summary.cell(2, 1, f"견적서 {len(set(it['프로젝트'] for items in all_data.values() for it in items))}개 현장 / 총 {sum(len(v) for v in all_data.values())}개 항목").font = Font(name="맑은 고딕", size=10, color="666666")

    sum_headers = ["공정", "항목 수", "현장 수", "합계금액 합산", "포함 현장"]
    for c, h in enumerate(sum_headers, 1):
        cell = ws_summary.cell(4, c, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border

    for r, (proc, cnt, pcnt, amt, projs) in enumerate(summary_rows, 5):
        ws_summary.cell(r, 1, proc).font = data_font
        ws_summary.cell(r, 1).border = thin_border
        ws_summary.cell(r, 2, cnt).font = data_font
        ws_summary.cell(r, 2).border = thin_border
        ws_summary.cell(r, 2).number_format = num_fmt
        ws_summary.cell(r, 3, pcnt).font = data_font
        ws_summary.cell(r, 3).border = thin_border
        ws_summary.cell(r, 4, amt).font = data_font
        ws_summary.cell(r, 4).border = thin_border
        ws_summary.cell(r, 4).number_format = '#,##0'
        ws_summary.cell(r, 5, ", ".join(projs)).font = Font(name="맑은 고딕", size=8)
        ws_summary.cell(r, 5).border = thin_border

    ws_summary.column_dimensions['A'].width = 18
    ws_summary.column_dimensions['B'].width = 10
    ws_summary.column_dimensions['C'].width = 10
    ws_summary.column_dimensions['D'].width = 18
    ws_summary.column_dimensions['E'].width = 60

    # ── 단가비교 시트 (주요 항목별 현장 간 단가 비교) ──
    ws_compare = wb.create_sheet(title="단가비교")

    # 모든 항목에서 주요 항목 식별 (3개 이상 현장에서 등장하는 항목)
    item_by_name = defaultdict(list)
    for proc, items in all_data.items():
        for it in items:
            key = (it["항목명"], it["단위"])
            item_by_name[key].append({
                "프로젝트": it["프로젝트"],
                "공정": proc,
                "규격": it["규격"],
                "자재단가": it["자재단가"],
                "노무단가": it["노무단가"],
                "합계단가": it["합계단가"],
            })

    # 3개 이상 현장에서 등장하는 항목만 추출
    common_items = []
    for (name, unit), entries in item_by_name.items():
        projects = set(e["프로젝트"] for e in entries)
        if len(projects) >= 3:
            # 단가 통계
            prices = [e["합계단가"] for e in entries if e["합계단가"] and e["합계단가"] > 0]
            if prices:
                common_items.append({
                    "항목명": name,
                    "단위": unit,
                    "공정": entries[0]["공정"],
                    "현장수": len(projects),
                    "최저단가": min(prices),
                    "최고단가": max(prices),
                    "평균단가": sum(prices) / len(prices),
                    "편차율": (max(prices) - min(prices)) / min(prices) * 100 if min(prices) > 0 else 0,
                    "현장목록": ", ".join(sorted(projects)),
                })

    common_items.sort(key=lambda x: (-x["현장수"], x["공정"]))

    comp_headers = ["항목명", "단위", "공정", "현장수", "최저단가", "최고단가", "평균단가", "편차율(%)", "포함 현장"]
    for c, h in enumerate(comp_headers, 1):
        cell = ws_compare.cell(1, c, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border

    for r, ci in enumerate(common_items, 2):
        vals = [ci["항목명"], ci["단위"], ci["공정"], ci["현장수"],
                ci["최저단가"], ci["최고단가"], round(ci["평균단가"]), round(ci["편차율"], 1),
                ci["현장목록"]]
        for c, v in enumerate(vals, 1):
            cell = ws_compare.cell(r, c, v)
            cell.font = data_font
            cell.border = thin_border
            if c >= 5 and c <= 7:
                cell.number_format = num_fmt
            if c == 8 and isinstance(v, (int, float)) and v > 50:
                cell.font = Font(name="맑은 고딕", size=9, color="FF0000", bold=True)

    ws_compare.freeze_panes = "A2"
    ws_compare.auto_filter.ref = f"A1:I1"
    comp_widths = [25, 6, 18, 8, 12, 12, 12, 10, 50]
    for c, w in enumerate(comp_widths, 1):
        ws_compare.column_dimensions[get_column_letter(c)].width = w

    # 저장
    wb.save(OUTPUT)
    print(f"\n✅ {OUTPUT} 생성 완료")
    print(f"   시트: {len(wb.sheetnames)}개 (요약 + 공정별 + 단가비교)")
    for proc, cnt, pcnt, amt, _ in summary_rows:
        print(f"   {proc}: {cnt}개 항목 ({pcnt}개 현장)")
    print(f"   단가비교: {len(common_items)}개 공통 항목")


def main():
    all_data = defaultdict(list)
    file_count = 0

    for f in sorted(os.listdir(FOLDER)):
        if not f.endswith('.xlsx'):
            continue
        if f.startswith('빌드어스') or f.startswith('.') or f.startswith('견적서_'):
            continue

        filepath = os.path.join(FOLDER, f)
        short = get_short_name(f)

        print(f"📄 파싱: {f[:40]}... → [{short}]")

        try:
            items = parse_detail_sheet(filepath, short)
            for proc, proc_items in items.items():
                all_data[proc].extend(proc_items)
                print(f"   {proc}: {len(proc_items)}개")
            file_count += 1
        except Exception as e:
            print(f"   ⚠️ 오류: {e}")
            import traceback
            traceback.print_exc()

    print(f"\n총 {file_count}개 파일 파싱 완료")
    total_items = sum(len(v) for v in all_data.values())
    print(f"총 {total_items}개 항목 추출")

    write_output(all_data)


if __name__ == "__main__":
    main()
