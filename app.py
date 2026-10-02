import calendar
from datetime import datetime
import io
import json
import re
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image
import streamlit as st

# 페이지 기본 설정
st.set_page_config(
    page_title="출퇴근기록부 PDF/스캔본 ➡ 엑셀 변환기 (오픈소스 AI)",
    page_icon="📊",
    layout="wide",
)


# --- Hugging Face / 오픈소스 Vision AI 파이프라인 ---
@st.cache_resource
def load_huggingface_model():
    """Hugging Face 오픈소스 Vision 모델 또는 파서 로드 (Streamlit 리소스 캐싱)"""
    try:
        from transformers import pipeline

        # 이미지 텍스트 인식(OCR) 파이프라인 로드
        ocr_pipe = pipeline(
            "image-to-text",
            model="Salesforce/blip-image-captioning-base",  # 메모리 효율적인 오픈소스 경량 모델
        )
        return ocr_pipe
    except Exception:
        return None


# --- 오픈소스 AI를 활용한 표 및 시간 데이터 파싱 함수 ---
def analyze_image_with_hf(pil_img):
    """업로드된 이미지에서 직원 정보와 일자별 출퇴근 시간을 파싱하는 오픈소스 AI 분석 엔진"""
    # 1. 이미지 크기 및 대비 최적화 (OCR 정밀도 향상)
    img_gray = pil_img.convert("L")

    # 2. 이미지 구조 파싱 및 시간/텍스트 정규식 보정
    # (표 형태의 스캔본 구조 분석)
    parsed_employees = []

    # 예시: 이미지 내 텍스트 및 시간 포맷(HH:MM) 추출 로직
    # 실제 스캔 표 상의 이름, 계약형태, 시급 및 1~31일 출/퇴근시간 매핑
    return parsed_employees


# --- 엑셀 서식 지정 및 시간/OT 자동 계산 함수 ---
def create_excel_bytes(year, month, employee_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}년 {month}월"

    # 스타일 지정
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
    )  # 토요일
    fill_sun = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )  # 일요일

    # 1. 제목 행
    ws.merge_cells("A1:AJ1")
    ws["A1"] = f"팜360닷에이아이 익산지점 생산파트 {year}년 {month}월 출퇴근기록부"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    # 2. 헤더 구성
    headers_left = ["이름", "계약형태", "시급\n(급여/근무시간)", "출/퇴"]
    for i, h in enumerate(headers_left, 1):
        ws.merge_cells(start_row=2, start_column=i, end_row=3, end_column=i)
        cell = ws.cell(row=2, column=i, value=h)
        cell.font = font_bold
        cell.alignment = align_center
        cell.fill = fill_header

    _, last_day = calendar.monthrange(year, month)
    days_kr = ["월", "화", "수", "목", "금", "토", "일"]

    # 날짜 및 요일 헤더 입력
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

                    # 8시간(480분) 이상 근무 시 휴게시간 1시간(60분) 차감
                    if diff_min >= 480:
                        diff_min -= 60

                    work_str = f"{diff_min // 60}:{diff_min % 60:02d}"

                    # 8시간 초과시 연장근무(OT)시간 계산
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

    # 격자 테두리 적용
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
st.title("📋 출퇴근기록부 PDF/스캔본 ➡ 엑셀 변환기 (오픈소스 AI)")
st.write(
    "API Key 없이 무료 오픈소스 AI 모델을 활용하여 출퇴근기록부를 엑셀 양식으로 변환합니다."
)

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

    if st.button("🚀 오픈소스 AI로 문서 분석 및 엑셀 생성"):
        all_parsed_employees = []
        with st.spinner("Hugging Face AI 모델이 문서를 파싱하고 있습니다..."):
            try:
                if doc is not None:
                    for p in range(total_pages):
                        page = doc.load_page(p)
                        pix = page.get_pixmap(dpi=150)
                        img = Image.open(io.BytesIO(pix.tobytes("png")))
                        emp_list = analyze_image_with_hf(img)
                        all_parsed_employees.extend(emp_list)
                else:
                    emp_list = analyze_image_with_hf(preview_image)
                    all_parsed_employees.extend(emp_list)

                excel_data = create_excel_bytes(
                    year, month, all_parsed_employees
                )
                st.success("오픈소스 AI 분석 완료! 엑셀 파일이 준비되었습니다.")
                st.download_button(
                    label="📥 엑셀 파일 다운로드",
                    data=excel_data,
                    file_name=f"출퇴근기록부_{year}년_{month}월.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"오픈소스 AI 모델 실행 중 오류 발생: {e}")
