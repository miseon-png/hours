import calendar
from datetime import datetime, date
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import streamlit as st

st.set_page_config(
    page_title="출퇴근기록부 엑셀 서식 자동 변환기",
    page_icon="📊",
    layout="wide",
)

st.title("📊 출퇴근기록부 엑셀 서식 자동 변환기")
st.write(
    "Gemini에서 추출한 원본 엑셀 파일(.xlsx)을 업로드하시면 서식 정리 및 근무/OT시간(저녁 휴게시간 30분 포함)이 자동 계산된 엑셀로 변환해 드립니다."
)


# 오늘 기준 지난달(전월) 연도 및 월 자동 산출
today = date.today()
first_day_this_month = today.replace(day=1)
last_month_date = first_day_this_month - calendar.timedelta(days=1) if hasattr(calendar, 'timedelta') else (first_day_this_month - date.resolution).replace(day=1) - date.resolution
# 정확한 전월 연/월 계산
if today.month == 1:
    default_year = today.year - 1
    default_month = 12
else:
    default_year = today.year
    default_month = today.month - 1


def process_excel(file_bytes, year, month):
    wb_in = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws_in = wb_in.active

    # 직원별 일자별 출/퇴근 데이터 추출
    data_map = {}
    for row in ws_in.iter_rows(min_row=4, values_only=True):
        if not row[0] or not row[2]:
            continue
        date_val, day_name, emp_name, in_time, out_time = row[:5]

        if isinstance(date_val, datetime):
            day_num = date_val.day
        else:
            try:
                day_num = int(str(date_val).split("-")[-1])
            except ValueError:
                continue

        emp_name = str(emp_name).strip()
        in_str = str(in_time).strip() if in_time and str(in_time) != "None" else ""
        out_str = str(out_time).strip() if out_time and str(out_time) != "None" else ""

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
    ws_out.title = f"{year}년 {month}월"

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
    )
    fill_sun = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )

    _, last_day = calendar.monthrange(year, month)

    # 상단 메인 제목 생성 (A1:AJ1)
    ws_out.merge_cells(
        start_row=1, start_column=1, end_row=1, end_column=4 + last_day
    )
    ws_out["A1"] = (
        f"팜360닷에이아이 익산지점 생산파트 {year}년 {month}월 출퇴근기록부"
    )
    ws_out["A1"].font = font_title
    ws_out["A1"].alignment = align_center

    # 헤더 작성 (이름, 계약형태, 시급, 출/퇴)
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

    # 본문 영역 작성 및 시간 자동 계산
    start_row = 4
    for emp_name, (contract_type, wage) in emp_info_dict.items():
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

            c_in = ws_out.cell(row=start_row, column=col_idx, value=in_t)
            c_out = ws_out.cell(row=start_row + 1, column=col_idx, value=out_t)

            for c in [c_in, c_out]:
                c.font = font_normal
                c.alignment = align_center

            work_str, ot_str = "", ""
            if in_t and out_t and ":" in in_t and ":" in out_t:
                try:
                    t1 = datetime.strptime(in_t, "%H:%M")
                    t2 = datetime.strptime(out_t, "%H:%M")
                    diff_min = (t2 - t1).seconds // 60

                    # 1. 점심 휴게시간 차감 (8시간 이상 근무 시 60분 공제)
                    if diff_min >= 480:
                        diff_min -= 60

                    # 2. 저녁 연장근무 휴게시간 차감 (17:30 이후 퇴근 시 저녁식사 30분 추가 공제)
                    if diff_min > 420 and t2.hour >= 18:
                        diff_min -= 30

                    work_str = f"{diff_min // 60}:{diff_min % 60:02d}"

                    # OT(연장근무) 시간 산출 (정규 8시간=480분 초과분)
                    if diff_min > 480:
                        ot_min = diff_min - 480
                        ot_str = f"{ot_min // 60}:{ot_min % 60:02d}"
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

    # 전체 격자 테두리 적용
    for row in ws_out.iter_rows(
        min_row=2, max_row=start_row - 1, min_col=1, max_col=4 + last_day
    ):
        for cell in row:
            cell.border = box_border
            if not cell.alignment.horizontal:
                cell.alignment = align_center

    output = io.BytesIO()
    wb_out.save(output)
    return output.getvalue()


# Streamlit UI
uploaded_file = st.file_uploader(
    "Gemini에서 추출한 엑셀 파일(.xlsx) 업로드", type=["xlsx"]
)

col1, col2 = st.columns([1, 1])
with col1:
    year = st.number_input("연도 선택", value=default_year)
    month = st.number_input("월 선택", value=default_month, min_value=1, max_value=12)

if uploaded_file is not None:
    if st.button("🚀 서식 자동 생성 및 계산 실행"):
        try:
            excel_bytes = process_excel(uploaded_file.read(), year, month)
            st.success("🎉 서식 변환 및 근무/OT시간 자동 계산이 완료되었습니다!")
            st.download_button(
                label="📥 변환된 출퇴근기록부 다운로드",
                data=excel_bytes,
                file_name=f"{year}년_{month}월_출퇴근기록부_완성본.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"오류 발생: {e}")
