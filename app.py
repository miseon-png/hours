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


# --- EasyOCR 모델 캐싱 로드 (무료 오프라인 모델) ---
@st.cache_resource
def load_ocr_reader():
    # GPU 없이 CPU 모드로 한글/영어 인식기 로드
    return easyocr.Reader(["ko", "en"], gpu=False)


# --- 실제 표 이미지 파싱 및 데이터 추출 로직 ---
def analyze_image_and_extract_data(pil_img):
    reader = load_ocr_reader()
    img_np = np.array(pil_img)

    # 1. EasyOCR 실행 (텍스트, 좌표, 신뢰도 추출)
    results = reader.readtext(img_np)

    if not results:
        return []

    # 2. 추출된 결과를 Y좌표 기준(행)으로 정렬
    # result 구조: [ ( [[x1,y1], [x2,y2], [x3,y3], [x4,y4]], "텍스트", confidence ), ... ]
    ocr_items = []
    for bbox, text, prob in results:
        text_clean = text.strip()
        if not text_clean:
            continue
        # BBox의 중앙 Y좌표 및 X좌표 계산
        top_left, bottom_right = bbox[0], bbox[2]
        center_x = (top_left[0] + bottom_right[0]) / 2
        center_y = (top_left[1] + bottom_right[1]) / 2
        ocr_items.append(
            {
                "text": text_clean,
                "x": center_x,
                "y": center_y,
                "box": bbox,
            }
        )

    # Y좌표 순 정렬
    ocr_items.sort(key=lambda item: item["y"])

    # 3. 시간 포맷(HH:MM 또는 H:MM) 정규식
    time_re = re.compile(r"(\d{1,2})[:;.]?(\d{2})")

    # 인식된 전체 텍스트 중 시간 및 이름 관련 블록 묶기
    employees = []
    current_emp = None

    # 이름 및 계약형태 후보 찾기
    known_contracts = ["정규", "계약", "월급계약", "시급계약", "파트"]

    for item in ocr_items:
        txt = item["text"]

        # 직원 이름/계약 형태 감지 (시간 형식이 아니고 길이가 짧은 한글)
        is_time = time_re.search(txt)
        if not is_time and len(txt) <= 6 and re.search(r"[가-힣]", txt):
            if any(c in txt for c in known_contracts) or len(txt) >= 2:
                # 새로운 직원 라인 시작 가능성 체크
                if current_emp is None or (
                    current_emp and len(current_emp["records"]) > 0
                ):
                    if current_emp:
                        employees.append(current_emp)
                    current_emp = {
                        "name": txt,
                        "contract_type": "정규",
                        "wage": 10320,
                        "records": {},
                    }
                    continue

        # 시간 패턴 발견 시 처리
        if is_time and current_emp:
            m = time_re.search(txt)
            h, mins = m.group(1), m.group(2)
            if len(h) == 1:
                h = "0" + h
            formatted_time = f"{h}:{mins}"

            # X 좌표를 기반으로 1일~31일 위치 추정 (이미지 폭 기준 31등분)
            img_width = pil_img.width
            # 왼쪽 정보영역(약 15%)을 제외한 영역을 31일로 분할
            margin_left = img_width * 0.12
            day_width = (img_width - margin_left) / 31.0

            if item["x"] > margin_left:
                day_idx = int((item["x"] - margin_left) // day_width) + 1
                day_idx = max(1, min(31, day_idx))

                if day_idx not in current_emp["records"]:
                    current_emp["records"][day_idx] = {
                        "in": formatted_time,
                        "out": "",
                    }
                elif not current_emp["records"][day_idx]["out"]:
                    current_emp["records"][day_idx]["out"] = formatted_time

    if current_emp:
        employees.append(current_emp)

    # 데이터가 없을 경우 기본 안전 구조 반환
    if not employees:
        employees = [
            {
                "name": "인식필요(직원1)",
                "contract_type": "정규",
                "wage": 10320,
                "records": {},
            }
        ]

    return employees


# --- 엑셀 파일 생성 및 양식 구성 ---
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

    # 3. 데이터 및 시간/OT 계산
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

                    # 8시간 이상 근무 시 1시간 휴게시간 자동 차감
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


# --- UI ---
st.title("📋 출퇴근기록부 PDF/스캔본 ➡ 엑셀 변환기")
st.write("문서를 업로드하면 OCR 엔진이 시간 및 직원 데이터를 파싱하여 엑셀을 추출합니다.")

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

    if st.button("🚀 문서 OCR 읽기 및 엑셀 변환"):
        all_parsed_employees = []
        with st.spinner("OCR 엔진이 표와 시간 데이터 글자를 분석 중입니다..."):
            try:
                if doc is not None:
                    for p in range(total_pages):
                        page = doc.load_page(p)
                        pix = page.get_pixmap(dpi=200)
                        img = Image.open(io.BytesIO(pix.tobytes("png")))
                        emp_list = analyze_image_and_extract_data(img)
                        all_parsed_employees.extend(emp_list)
                else:
                    emp_list = analyze_image_and_extract_data(preview_image)
                    all_parsed_employees.extend(emp_list)

                excel_data = create_excel_bytes(
                    year, month, all_parsed_employees
                )
                st.success("문서 읽기 완료! 엑셀 다운로드 버튼이 아래에 준비되었습니다.")
                st.download_button(
                    label="📥 엑셀 파일 다운로드",
                    data=excel_data,
                    file_name=f"출퇴근기록부_{year}년_{month}월.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"문서 데이터 파싱 중 에러 발생: {e}")
