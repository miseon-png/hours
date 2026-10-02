import calendar
from datetime import datetime
import io
import re
import cv2
import fitz  # PyMuPDF
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pdfplumber
from PIL import Image
import pytesseract
import streamlit as st

st.set_page_config(
    page_title="출퇴근기록부 변환기 (로컬 파싱)", page_icon="📊", layout="wide"
)


def extract_table_from_pdf_digitally(pdf_bytes):
    """PDF 내부에 텍스트가 묻어있는 '텍스트 기반 PDF'인 경우 pdfplumber로 직접 추출"""
    employees = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                # 추출된 2차원 표 배열 데이터를 직원 구조체로 파싱
                for i in range(0, len(table)):
                    row = table[i]
                    # 첫 행에 이름이 포함된 4행 단위 구성 분석
                    pass
    return employees


def preprocess_and_ocr_image(pil_img):
    """이미지 전처리(이진화, 노이즈 제거) 후 Tesseract로 텍스트 추출"""
    # 1. OpenCV 이미지 변환 및 그레이스케일
    img_np = np.array(pil_img)
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)

    # 2. 이진화 (Thresholding) 처리하여 텍스트 선명화
    _, thresh = cv2.threshold(
        gray, 150, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU
    )

    # 3. Tesseract OCR 실행 (한글 + 숫자)
    # --psm 6: 단일 균일 텍스트 블록 가정
    config = "--psm 6 -c preserve_interword_spaces=1"
    raw_text = pytesseract.image_to_string(thresh, lang="kor+eng", config=config)

    return raw_text


# --- 엑셀 생성 및 다운로드 UI 부분은 기존과 동일 ---
