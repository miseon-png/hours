import calendar
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def transform_attendance_excel(input_file, output_file):
    # 1. 원본 데이터 읽기
    wb_in = openpyxl.load_workbook(input_file)
    ws_in = wb_in.active

    # 직원별 일자별 출/퇴근 기록 정리
    # data_map[emp_name][day] = {'in': '08:30', 'out': '12:30'}
    data_map = {}
    for row in ws_in.iter_rows(min_row=4, values_only=True):
        if not row[0] or not row[2]:
            continue
        date_val, day_name, emp_name, in_time, out_time = row[:5]

        # 날짜 파싱 (datetime 객체 또는 문자열)
        if isinstance(date_val, datetime):
            day_num = date_val.day
        else:
            day_num = int(str(date_val).split("-")[-1])

        emp_name = str(emp_name).strip()
        in_str = str(in_time).strip() if in_time else ""
        out_str = str(out_time).strip() if out_time else ""

        if emp_name not in data_map:
            data_map[emp_name] = {}
        data_map[emp_name][day_num] = {"in": in_str, "out": out_str}

    # 직원 기본정보 매핑 (계약형태, 시급)
    emp_info_dict = {
        "천근하": ("정규", 10320),
        "김란님": ("정규", 11279),
        "김완수": ("정규", 11962),
        "심동희": ("계약\n월급계약", 10320),
        "하영실": ("계약\n시급계약", 10320),
        "민순기": ("계약\n시급계약", 10320),
        "도용엔 프엄안": ("계약\n시급계약", 10320),
        "VAN DUNG": ("계약\n월급계약", 10320),
        "VAN HA": ("정규", 10320),
    }

    # 2. 새로운 양식 엑셀 워크북 생성
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "2026년 9월"

    # 스타일 정의
    font_bold = Font(name="맑은 고딕", size=9, bold=True)
    font_normal = Font(name="맑은 고딕", size=9)
    font_title = Font(name="맑은 고딕", size=14, bold=True)
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    border_thin = Side(border_style="thin", color="000000")
    box_border = Border(
        left=border_thin, right=border_thin, top=border_thin, bottom=border_thin
    )

    fill_header = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    fill_sat = PatternFill(
        start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"
    )  # 연한 파랑
    fill_sun = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )  # 연한 주황

    year, month = 2026, 9
    _, last_day = calendar.monthrange(year, month)

    # 3. 메인 제목 생성 (A1:AJ1)
    ws_out.merge_cells(
        start_row=1, start_column=1, end_row=1, end_column=4 + last_day
    )
    ws_out["A1"] = (
        f"팜360닷에이아이 익산지점 생산파트 {year}년 {month}월 출퇴근기록부"
    )
    ws_out["A1"].font = font_title
    ws_out["A1"].alignment = align_center

    # 4. 헤더 레이아웃 구성
    headers_left = ["이름", "계약형태", "시급\n(급여/근무시간)", "출/퇴"]
    for i, h in enumerate(headers_left, 1):
        ws_out.merge_cells(start_row=2, start_column=i, end_row=3, end_column=i)
        cell = ws_out.cell(row=2, column=i, value=h)
        cell.font = font_bold
        cell.alignment = align_center
        cell.fill = fill_header

    days_kr = ["월", "화", "수", "목", "금", "토", "일"]
    for d in range(1, last_day + 1):
        col_idx = 4 + d
        dt = datetime(year, month, d)
        day_name = days_kr[dt.weekday()]

        c_day = ws_out.cell(row=2, column=col_idx, value=day_name)
        c_date = ws_out.cell(row=3, column=col_idx, value=d)

        for c in [c_day, c_date]:
            c.font = font_bold
            c.alignment = align_center
            c.fill = (
                fill_sat
                if dt.weekday() == 5
                else (fill_sun if dt.weekday() == 6 else fill_header)
            )

    # 5. 직원별 출퇴근 데이터 작성
    start_row = 4
    for emp_name, (contract_type, wage) in emp_info_dict.items():
        # 왼쪽 프로필 열 병합 (4행씩)
        ws_out.merge_cells(
            start_row=start_row,
            start_column=1,
            end_row=start_row + 3,
            end_column=1,
        )
        ws_out.merge_cells(
            start_row=start_row,
            start_column=2,
            end_row=start_row + 3,
            end_column=2,
        )
        ws_out.merge_cells(
            start_row=start_row,
            start_column=3,
            end_row=start_row + 3,
            end_column=3,
        )

        ws_out.cell(row=start_row, column=1, value=emp_name).font = font_bold
        ws_out.cell(row=start_row, column=2, value=contract_type).font = (
            font_normal
        )
        ws_out.cell(row=start_row, column=3, value=f"{wage:,}").font = font_normal

        type_labels = ["출근", "퇴근", "근무시간", "OT시간"]
        for idx, label in enumerate(type_labels):
            c = ws_out.cell(row=start_row + idx, column=4, value=label)
            c.font = font_bold
            c.alignment = align_center

        emp_records = data_map.get(emp_name, {})

        for d in range(1, last_day + 1):
            col_idx = 4 + d
            day_data = emp_records.get(d, {})
            in_t = day_data.get("in", "")
            out_t = day_data.get("out", "")

            # 출/퇴근 시간 기입
            c_in = ws_out.cell(row=start_row, column=col_idx, value=in_t)
            c_out = ws_out.cell(row=start_row + 1, column=col_idx, value=out_t)

            for c in [c_in, c_out]:
                c.font = font_normal
                c.alignment = align_center

            # 근무시간 / OT시간 계산
            work_str, ot_str = "", ""
            if in_t and out_t and ":" in in_t and ":" in out_t:
                try:
                    t1 = datetime.strptime(in_t, "%H:%M")
                    t2 = datetime.strptime(out_t, "%H:%M")
                    diff = (t2 - t1).seconds // 60

                    # 8시간(480분) 이상 시 1시간(60분) 휴게시간 공제
                    if diff >= 480:
                        diff -= 60

                    work_str = f"{diff // 60}:{diff % 60:02d}"

                    if diff > 480:
                        ot_m = diff - 480
                        ot_str = f"{ot_m // 60}:{ot_m % 60:02d}"
                    else:
                        ot_str = "0:00"
                except Exception:
                    pass

            c_work = ws_out.cell(
                row=start_row + 2, column=col_idx, value=work_str
            )
            c_ot = ws_out.cell(row=start_row + 3, column=col_idx, value=ot_str)

            for c in [c_work, c_ot]:
                c.font = font_normal
                c.alignment = align_center

        start_row += 4

    # 전체 격자 테두리 및 정렬 적용
    for row in ws_out.iter_rows(
        min_row=2, max_row=start_row - 1, min_col=1, max_col=4 + last_day
    ):
        for cell in row:
            cell.border = box_border
            if not cell.alignment.horizontal:
                cell.alignment = align_center

    wb_out.save(output_file)
    print(f"변환 완료! 파일 저장 위치: {output_file}")


# 실행
transform_attendance_excel(
    "2026년_9월_출퇴근기록.xlsx", "2026년_9월_출퇴근기록부_완성본.xlsx"
)
