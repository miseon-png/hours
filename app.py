from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def create_attendance_sheet(
    year, month, employee_data, output_filename="출퇴근기록부.xlsx"
):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}년 {month}월"

    # --- 색상 및 스타일 정의 ---
    font_bold = Font(name="맑은 고딕", size=10, bold=True)
    font_regular = Font(name="맑은 고딕", size=9)
    font_title = Font(name="맑은 고딕", size=14, bold=True)

    align_center = Alignment(horizontal="center", vertical="center")

    border_thin = Side(border_style="thin", color="000000")
    box_border = Border(
        left=border_thin, right=border_thin, top=border_thin, bottom=border_thin
    )

    fill_header = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    fill_sat = PatternFill(
        start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"
    )  # 토요일 파란색 계열
    fill_sun = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )  # 일요일/공휴일 주황색 계열

    # --- 1. 제목 열 작성 ---
    ws.merge_cells("A1:AJ1")
    ws["A1"] = f"팜360닷에이아이 익산지점 생산파트 {year}년 {month}월 출퇴근기록부"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    # --- 2. 헤더 (이름, 계약형태, 시급, 날짜, 요일) 작성 ---
    headers_left = ["이름", "계약형태", "시급\n(급여/근무시간)", "출/퇴"]
    for i, h in enumerate(headers_left, 1):
        ws.merge_cells(start_row=2, start_column=i, end_row=3, end_column=i)
        cell = ws.cell(row=2, column=i, value=h)
        cell.font = font_bold
        cell.alignment = align_center
        cell.fill = fill_header

    # 해당 월의 일수 및 요일 계산
    import calendar

    _, last_day = calendar.monthrange(year, month)
    days_kr = ["월", "화", "수", "목", "금", "토", "일"]

    for d in range(1, last_day + 1):
        col_idx = 4 + d
        dt = datetime(year, month, d)
        day_name = days_kr[dt.weekday()]

        # 2행: 요일
        cell_day = ws.cell(row=2, column=col_idx, value=day_name)
        # 3행: 일자
        cell_date = ws.cell(row=3, column=col_idx, value=d)

        for c in [cell_day, cell_date]:
            c.font = font_bold
            c.alignment = align_center
            if dt.weekday() == 5:  # 토요일
                c.fill = fill_sat
            elif dt.weekday() == 6:  # 일요일
                c.fill = fill_sun
            else:
                c.fill = fill_header

    # --- 3. 데이터 및 계산 작성 ---
    start_row = 4

    for emp in employee_data:
        # 직원 기본정보
        ws.merge_cells(
            start_row=start_row,
            start_column=1,
            end_row=start_row + 3,
            end_column=1,
        )
        ws.merge_cells(
            start_row=start_row,
            start_column=2,
            end_row=start_row + 3,
            end_column=2,
        )
        ws.merge_cells(
            start_row=start_row,
            start_column=3,
            end_row=start_row + 3,
            end_column=3,
        )

        ws.cell(row=start_row, column=1, value=emp["name"])
        ws.cell(row=start_row, column=2, value=emp["contract_type"])
        ws.cell(row=start_row, column=3, value=emp["wage"])

        # 출/퇴 구분 라벨
        type_labels = ["출근", "퇴근", "근무시간", "OT시간"]
        for idx, label in enumerate(type_labels):
            ws.cell(row=start_row + idx, column=4, value=label)

        # 일자별 출퇴근 기록 및 계산
        for d in range(1, last_day + 1):
            col_idx = 4 + d
            records = emp["records"].get(d, {})

            in_time_str = records.get("in", "")
            out_time_str = records.get("out", "")

            c_in = ws.cell(row=start_row, column=col_idx, value=in_time_str)
            c_out = ws.cell(row=start_row + 1, column=col_idx, value=out_time_str)

            work_str = ""
            ot_str = ""

            # 시간 계산 (HH:MM 형식 가정)
            if in_time_str and out_time_str:
                try:
                    fmt = "%H:%M"
                    t_in = datetime.strptime(in_time_str, fmt)
                    t_out = datetime.strptime(out_time_str, fmt)

                    # 총 근무 분 계산
                    diff_minutes = (t_out - t_in).seconds // 60

                    # 4시간 이상 근무 시 점심/휴게시간 1시간(60분) 차감 로직 예시
                    if diff_minutes >= 480:
                        diff_minutes -= 60

                    work_hours = diff_minutes // 60
                    work_mins = diff_minutes % 60
                    work_str = f"{work_hours}:{work_mins:02d}"

                    # 8시간 초과시 OT 시간 계산
                    if diff_minutes > 480:
                        ot_minutes = diff_minutes - 480
                        ot_str = f"{ot_minutes // 60}:{ot_minutes % 60:02d}"
                    else:
                        ot_str = "0:00"

                except ValueError:
                    pass

            ws.cell(row=start_row + 2, column=col_idx, value=work_str)
            ws.cell(row=start_row + 3, column=col_idx, value=ot_str)

        start_row += 4

    # --- 4. 전체 테두리 및 정렬 적용 ---
    for row in ws.iter_rows(
        min_row=2, max_row=start_row - 1, min_col=1, max_col=4 + last_day
    ):
        for cell in row:
            cell.border = box_border
            if not cell.alignment.horizontal:
                cell.alignment = align_center

    wb.save(output_filename)
    print(f"엑셀 파일 저장 완료: {output_filename}")


# 샘플 테스트 데이터 (OCR로 파싱되어 들어올 데이터 구조 예시)
sample_employees = [
    {
        "name": "천근하",
        "contract_type": "정규",
        "wage": 10320,
        "records": {
            3: {"in": "08:30", "out": "12:30"},
            4: {"in": "08:30", "out": "12:30"},
            11: {"in": "08:30", "out": "13:00"},  # 4.5시간 근무 예시
        },
    }
]

# 실행
create_attendance_sheet(2026, 8, sample_employees)
