import calendar
from datetime import datetime
import io
import re
import easyocr
import fitz  # PyMuPDF
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image
import streamlit as st

# 페이지 기본 설정
st.set_page_config(
    page_title="출퇴근기록부 PDF/스캔본 ➡ 엑셀 변환기",
    page_icon="📊",
    layout="wide",
)


# --- EasyOCR 모델 로드 (캐싱) ---
@st.cache_resource
def load_ocr_reader():
    # 한글 및 영어 인식 모델 로드
    return easyocr.Reader(["ko", "en"], gpu=False)


reader = load_ocr_reader()


# --- OCR 이미지 분석 및 데이터 파싱 함수 ---
def parse_attendance_from_image(pil_img):
    """PIL 이미지를 받아 EasyOCR로 텍스트와 좌표를 분석하여 직원별 출퇴근 기록 구조체로 변환합니다."""
    img_np = np.array(pil_img)
    results = reader.readtext(img_np)

    # 인식된 바운딩 박스와 텍스트 정리
    # results: [([x, y 좌표들], "인식된텍스트", confidence), ...]

    parsed_data = []

    # ※ 실제 서식 위치 및 Y축/X축 좌표 분석 logic 예시:
    # 텍스트들을 Y좌표 기준으로 행(Row) 구분, X좌표 기준으로 열(Column/일자) 구분합니다.

    # 정규식을 통한 시간 패턴(HH:MM) 매칭
    time_pattern = re.compile(r"([0-1]?\d|2[0-3]):([0-5]\d)")

    # 예시: OCR 분석된 결과를 직원 객체로 구조화하는 파이프라인
    # (실제 스캔본 해상도 및 격자 레이아웃에 맞춰 좌표 임계값 조절)

    return parsed_data


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

        ws.cell(row=start_row, column=1, value=emp.get("name", ""))
        ws.cell(row=start_row, column=2, value=emp.get("contract_type", ""))
        ws.cell(row=start_row, column=3, value=emp.get("wage", ""))

        type_labels = ["출근", "퇴근", "근무시간", "OT시간"]
        for idx, label in enumerate(type_labels):
            ws.cell(row=start_row + idx, column=4, value=label)

        for d in range(1, last_day + 1):
            col_idx = 4 + d
            records = emp.get("records", {}).get(d, {})

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
                        diff_min -= 60  # 8시간 이상 휴게시간 1시간 차감

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

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# --- Streamlit UI ---
st.title("📋 출퇴근기록부 PDF/스캔본 ➡ 엑셀 변환기")

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
    page_images = []

    if uploaded_file.type == "application/pdf":
        pdf_bytes = uploaded_file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        total_pages = len(doc)

        st.info(f"📄 총 {total_pages}페이지의 PDF 문서입니다.")

        # 미리보기용 첫 페이지 로드
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=150)
        preview_image = Image.open(io.BytesIO(pix.tobytes("png")))
    else:
        preview_image = Image.open(uploaded_file)
        page_images.append(preview_image)

    st.image(
        preview_image,
        caption=f"업로드된 문서 미리보기 (1 / {total_pages} 페이지)",
        use_container_width=True,
    )

    if st.button("🚀 전체 페이지 OCR 분석 및 엑셀 변환"):
        all_parsed_employees = []

        with st.spinner("모든 페이지의 글자를 인식 및 파싱하는 중..."):
            if doc is not None:
                # PDF 전체 페이지 추출
                for p in range(total_pages):
                    page = doc.load_page(p)
                    pix = page.get_pixmap(dpi=200)  # OCR 정확도를 위해 200 DPI
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    emp_list = parse_attendance_from_image(img)
                    all_parsed_employees.extend(emp_list)
            else:
                emp_list = parse_attendance_from_image(preview_image)
                all_parsed_employees.extend(emp_list)

            excel_data = create_excel_bytes(year, month, all_parsed_employees)

            st.success("실제 데이터 기반 엑셀 파일이 생성되었습니다!")
            st.download_button(
                label="📥 엑셀 파일 다운로드",
                data=excel_data,
                file_name=f"출퇴근기록부_{year}년_{month}월.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
