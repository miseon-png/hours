import calendar
from datetime import datetime
import io
import json
import fitz  # PyMuPDF
from google import genai
from google.genai import types
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image
import streamlit as st

st.set_page_config(
    page_title="출퇴근기록부 PDF/스캔본 ➡ 엑셀 변환기 (Gemini AI)",
    page_icon="📊",
    layout="wide",
)

# API Key 설정: Streamlit Secrets에서 가져오거나 사이드바에서 입력받기
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")

if not GEMINI_API_KEY:
    GEMINI_API_KEY = st.sidebar.text_input(
        "Google Gemini API Key 입력", type="password"
    )


# --- Gemini Vision API 연동 파싱 함수 ---
def analyze_image_with_gemini(pil_img, api_key):
    client = genai.Client(api_key=api_key)

    prompt = """
    이 이미지는 출퇴근기록부 표입니다.
    이미지에서 각 직원별로 이름(name), 계약형태(contract_type), 시급(wage), 그리고 1일부터 31일까지의 일자별 출근시간(in)과 퇴근시간(out)을 정확히 인식하여 정규 JSON 배열 형식으로만 응답하세요.

    [작성 규칙]
    1. 시간 포맷은 반드시 "HH:MM" 형태로 작성하세요. (예: 08:30, 12:30, 17:30)
    2. 연차, 반차, 결근 등의 문구가 적혀있거나 근무 기록이 없는 날은 출퇴근시간을 빈 문자열("")로 처리하세요.
    3. 반차, 연차 문구가 적혀 있는 날이어도 출퇴근시간은 빈 문자열("")로 설정하고 다른 어떠한 글자도 포함하지 마세요.

    [JSON 응답 포맷 구조 예시]
    [
      {
        "name": "천근하",
        "contract_type": "정규",
        "wage": 10320,
        "records": {
          "1": {"in": "", "out": ""},
          "2": {"in": "08:30", "out": "12:30"},
          "3": {"in": "08:30", "out": "12:30"}
        }
      }
    ]
    """

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[pil_img, prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.1,
        ),
    )

    result_text = response.text
    data = json.loads(result_text)

    if isinstance(data, dict):
        for k in data:
            if isinstance(data[k], list):
                return data[k]
    return data


# --- 엑셀 서식 및 시간 계산 함수 ---
def create_excel_bytes(year, month, employee_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}년 {month}월"

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

    ws.merge_cells("A1:AJ1")
    ws["A1"] = f"팜360닷에이아이 익산지점 생산파트 {year}년 {month}월 출퇴근기록부"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    headers_left = ["이름", "계약형태", "시급\n(급여/근무시간)", "출/퇴"]
    for i, h in enumerate(headers_left, 1):
        ws.merge_cells(start_row=2, start_column=i, end_row=3, end_column=i)
        cell = ws.cell(row=2, column=i, value=h)
        cell.font = font_bold
        cell.alignment = align_center
        cell.fill = fill_header

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

        ws.cell(row=start_row, column=1, value=emp.get("name", ""))
        ws.cell(row=start_row, column=2, value=emp.get("contract_type", ""))
        ws.cell(row=start_row, column=3, value=emp.get("wage", ""))

        type_labels = ["출근", "퇴근", "근무시간", "OT시간"]
        for idx, label in enumerate(type_labels):
            ws.cell(row=start_row + idx, column=4, value=label)

        records_data = emp.get("records", {})

        for d in range(1, last_day + 1):
            col_idx = 4 + d
            day_record = records_data.get(d) or records_data.get(str(d)) or {}

            in_time_str = day_record.get("in", "")
            out_time_str = day_record.get("out", "")

            ws.cell(row=start_row, column=col_idx, value=in_time_str)
            ws.cell(row=start_row + 1, column=col_idx, value=out_time_str)

            work_str, ot_str = "", ""
            if in_time_str and out_time_str:
                try:
                    t_in = datetime.strptime(in_time_str, "%H:%M")
                    t_out = datetime.strptime(out_time_str, "%H:%M")
                    diff_min = (t_out - t_in).seconds // 60

                    if diff_min >= 480:
                        diff_min -= 60

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

    for row in ws.iter_rows(
        min_row=2, max_row=start_row - 1, min_col=1, max_col=4 + last_day
    ):
        for cell in row:
            cell.border = box_border
            if not cell.alignment.horizontal:
                cell.alignment = align_center

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# --- Streamlit UI ---
st.title("📋 출퇴근기록부 PDF/스캔본 ➡ 엑셀 변환기 (Gemini AI)")

col1, col2 = st.columns([1, 1])
with col1:
    year = st.number_input("연도 선택", min_value=2020, max_value=2030, value=2026)
    month = st.selectbox("월 선택", list(range(1, 13)), index=7)

uploaded_file = st.file_uploader(
    "출퇴근기록부 PDF 또는 스캔 이미지 업로드",
    type=["pdf", "png", "jpg", "jpeg"],
)

if uploaded_file is not None:
    doc = None
    total_pages = 1

    if uploaded_file.type == "application/pdf":
        pdf_bytes = uploaded_file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        total_pages = len(doc)
        st.info(f"📄 총 {total_pages}페이지의 PDF 문서입니다.")

        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=150)
        preview_image = Image.open(io.BytesIO(pix.tobytes("png")))
    else:
        preview_image = Image.open(uploaded_file)

    st.image(
        preview_image, caption="업로드된 문서 미리보기", use_container_width=True
    )

    if st.button("🚀 Gemini AI로 문서 분석 및 엑셀 생성"):
        if not GEMINI_API_KEY:
            st.error(
                "Gemini API Key가 설정되지 않았습니다. 사이드바에 키를 입력해 주세요."
            )
        else:
            all_parsed_employees = []
            with st.spinner("Gemini AI가 업로드된 문서 표를 인식 및 파싱하는 중..."):
                try:
                    if doc is not None:
                        for p in range(total_pages):
                            page = doc.load_page(p)
                            pix = page.get_pixmap(dpi=150)
                            img = Image.open(io.BytesIO(pix.tobytes("png")))
                            emp_list = analyze_image_with_gemini(
                                img, GEMINI_API_KEY
                            )
                            all_parsed_employees.extend(emp_list)
                    else:
                        emp_list = analyze_image_with_gemini(
                            preview_image, GEMINI_API_KEY
                        )
                        all_parsed_employees.extend(emp_list)

                    excel_data = create_excel_bytes(
                        year, month, all_parsed_employees
                    )
                    st.success("Gemini AI 분석 완료! 엑셀 파일이 준비되었습니다.")
                    st.download_button(
                        label="📥 엑셀 파일 다운로드",
                        data=excel_data,
                        file_name=f"출퇴근기록부_{year}년_{month}월.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"Gemini API 호출 중 에러 발생: {e}")
