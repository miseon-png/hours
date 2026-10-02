from datetime import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image
import streamlit as st

# 페이지 기본 설정
st.set_page_config(
    page_title="출퇴근기록부 엑셀 변환기", page_icon="📊", layout="wide"
)


# --- 엑셀 생성 함수 ---
def create_excel_bytes(year, month, employee_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}년 {month}월"

    # 스타일 설정
    font_bold = Font(name="맑은 고딕", size=10, bold=True)
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
    )
    fill_sun = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )

    # 1. 제목
    ws.merge_cells("A1:AJ1")
    ws["A1"] = f"팜360닷에이아이 익산지점 생산파트 {year}년 {month}월 출퇴근기록부"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    # 2. 헤더
    headers_left = ["이름", "계약형태", "시급\n(급여/근무시간)", "출/퇴"]
    for i, h in enumerate(headers_left, 1):
        ws.merge_cells(start_row=2, start_column=i, end_row=3, end_column=i)
        cell = ws.cell(row=2, column=i, value=h)
        cell.font = font_bold
        cell.alignment = align_center
        cell.fill = fill_header

    import calendar

    _, last_day = calendar.monthrange(year, month)
    days_kr = ["월", "화", "수", "목", "금", "토", "일"]

    for d in range(1, last_day + 1):
        col_idx = 4 + d
        dt = datetime(year, month, d)
        day_name = days_kr[dt.weekday()]

        c_day = ws.cell(row=2, column=col_idx, value=day_name)
        c_date = ws.cell(row=3, column=col_idx, value=d)

        for c in [c_day, c_date]:
            c.font = font_bold
            c.alignment = align_center
            c.fill = (
                fill_sat
                if dt.weekday() == 5
                else (fill_sun if dt.weekday() == 6 else fill_header)
            )

    # 3. 데이터 작성 및 시간/OT 자동 계산
    start_row = 4
    for emp in employee_data:
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

        type_labels = ["출근", "퇴근", "근무시간", "OT시간"]
        for idx, label in enumerate(type_labels):
            ws.cell(row=start_row + idx, column=4, value=label)

        for d in range(1, last_day + 1):
            col_idx = 4 + d
            records = emp["records"].get(d, {})

            in_time_str = records.get("in", "")
            out_time_str = records.get("out", "")

            ws.cell(row=start_row, column=col_idx, value=in_time_str)
            ws.cell(row=start_row + 1, column=col_idx, value=out_time_str)

            work_str, ot_str = "", ""
            if in_time_str and out_time_str:
                try:
                    t_in = datetime.strptime(in_time_str, "%H:%M")
                    t_out = datetime.strptime(out_time_str, "%H:%M")
                    diff_min = (t_out - t_in).seconds // 60

                    if diff_min >= 480:
                        diff_min -= 60  # 휴게시간 1시간 차감

                    work_str = f"{diff_min // 60}:{diff_min % 60:02d}"
                    if diff_min > 480:
                        ot_min = diff_min - 480
                        ot_str = f"{ot_min // 60}:{ot_min % 60:02d}"
                    else:
                        ot_str = "0:00"
                except ValueError:
                    pass

            ws.cell(row=start_row + 2, column=col_idx, value=work_str)
            ws.cell(row=start_row + 3, column=col_idx, value=ot_str)

        start_row += 4

    # 테두리 적용
    for row in ws.iter_rows(
        min_row=2, max_row=start_row - 1, min_col=1, max_col=4 + last_day
    ):
        for cell in row:
            cell.border = box_border
            if not cell.alignment.horizontal:
                cell.alignment = align_center

    # 메모리 버퍼로 저장
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# --- Streamlit UI 구성 ---
st.title("📋 출퇴근기록부 스캔본 ➡️️ 엑셀 변환기")
st.write(
    "스캔한 출퇴근기록부 이미지를 업로드하면 근무시간과 OT시간을 자동 계산하여 동일한 양식의 엑셀 파일로 생성합니다."
)

col1, col2 = st.columns([1, 1])

with col1:
    year = st.number_input("연도 선택", min_value=2020, max_value=2030, value=2026)
    month = st.selectbox("월 선택", list(range(1, 13)), index=7)  # 기본 8월

uploaded_file = st.file_uploader(
    "출퇴근기록부 스캔 이미지 업로드", type=["png", "jpg", "jpeg"]
)

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    st.image(image, caption="업로드된 스캔 이미지", use_column_width=True)

    if st.button("🚀 엑셀 파일 생성하기"):
        with st.spinner("OCR 인식 및 시간/OT 자동 계산 중..."):
            # TODO: 실제 OCR 연동 시 이곳에서 uploaded_file을 OCR API(CLOVA/Vision 등)에 전달하여 parsed_data를 만듭니다.
            # 임시 샘플 데이터
            sample_parsed_data = [
                {
                    "name": "천근하",
                    "contract_type": "정규",
                    "wage": 10320,
                    "records": {
                        3: {"in": "08:30", "out": "12:30"},
                        4: {"in": "08:30", "out": "12:30"},
                        11: {"in": "08:30", "out": "13:00"},
                    },
                },
                {
                    "name": "김란남",
                    "contract_type": "정규",
                    "wage": 11279,
                    "records": {
                        3: {"in": "08:30", "out": "12:30"},
                        11: {"in": "12:30", "out": "17:30"},
                    },
                },
            ]

            excel_data = create_excel_bytes(year, month, sample_parsed_data)

            st.success("엑셀 파일이 성공적으로 생성되었습니다!")
            st.download_button(
                label="📥 엑셀 파일 다운로드",
                data=excel_data,
                file_name=f"출퇴근기록부_{year}년_{month}월.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
