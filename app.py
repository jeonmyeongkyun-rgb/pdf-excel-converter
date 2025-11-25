import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import pandas as pd
import os
import io
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# ==============================================================================
# [필수] 여기에 발급받은 API 키를 따옴표 안에 붙여넣으세요!
# ==============================================================================
GOOGLE_API_KEY = "AIzaSyBQjCBOwYNjiy5Z-Ej_OQR8XSUHsbfvKPk"
# ==============================================================================

# Gemini 설정
genai.configure(api_key=GOOGLE_API_KEY)

# 화면 설정 (넓게 보기)
st.set_page_config(page_title="사내용 PDF 변환기", page_icon="🏢", layout="wide")
st.title("🏢 사내용 PDF ➡️ 엑셀 변환기")
st.markdown("""
- **여러 파일을 한 번에** 올릴 수 있습니다.
- 파일 이름은 **원본 그대로 유지**됩니다.
- 보안을 위해 외부 공유 시 주의해주세요.
""")

# API 키 누락 방지
if "여기에" in GOOGLE_API_KEY:
    st.error("🚨 코드 15번째 줄에 API 키를 입력하고 저장해주세요!")
    st.stop()

# 모델 설정 (가장 빠르고 저렴한 모델)
try:
    model = genai.GenerativeModel('gemini-2.5-flash')
except:
    st.error("모델 로딩 실패. 잠시 후 다시 시도해주세요.")

# --- 다중 파일 업로더 ---
uploaded_files = st.file_uploader(
    "변환할 PDF 파일들을 여기에 모두 드래그하세요 (여러 개 가능)", 
    type="pdf", 
    accept_multiple_files=True
)

# --- 변환 처리 함수 ---
def process_pdf(file_bytes, original_name):
    input_pdf = f"temp_{original_name}"
    # 확장자만 .xlsx로 변경
    output_xls = os.path.splitext(original_name)[0] + ".xlsx"
    
    with open(input_pdf, "wb") as f:
        f.write(file_bytes)

    try:
        doc = fitz.open(input_pdf)
        all_dfs = []
        
        # 엑셀 컬럼 정의
        columns = ["거래일자", "거래시간", "상태", "거래구분", "거래금액", "표면잔액", "취급점", "적요", "은행명", "계좌번호"]

        for i, page in enumerate(doc):
            # 이미지 변환
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            
            # Gemini에게 보낼 데이터
            image_parts = [{"mime_type": "image/png", "data": img_data}]
            
            # 강력한 프롬프트
            prompt = """
            이 이미지의 은행 거래내역 표를 파이프(|) 기호로 구분된 텍스트로 추출해.
            
            [규칙]
            1. 각 줄은 10개 항목: 날짜|시간|상태|구분|거래금액|표면잔액|취급점|적요|은행명|계좌번호
            2. '표면잔액'과 '취급점'이 붙어있으면 반드시 구분선(|)으로 나눠.
            3. 금액의 쉼표(,)는 유지하고, 계좌번호는 숫자만 남겨.
            4. 헤더와 배경 글자(KB 등)는 무시해.
            """
            
            response = model.generate_content([prompt, image_parts[0]])
            raw_text = response.text.strip().replace("```", "")
            
            data_rows = []
            for line in raw_text.split('\n'):
                if "|" in line:
                    parts = line.split('|')
                    # 칸 개수 맞추기 (오류 방지)
                    if len(parts) < 10: parts += [""] * (10 - len(parts))
                    if len(parts) > 10: parts = parts[:10]
                    parts = [p.strip() for p in parts]
                    data_rows.append(parts)
            
            if data_rows:
                df = pd.DataFrame(data_rows, columns=columns)
                all_dfs.append(df)

        if all_dfs:
            final_df = pd.concat(all_dfs, ignore_index=True)
            
            # 숫자 데이터 정리 (금액 콤마 제거 후 숫자로 변환)
            for col in ["거래금액", "표면잔액"]:
                final_df[col] = final_df[col].astype(str).str.replace(',', '').str.replace('원', '')
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            final_df.to_excel(output_xls, index=False)

            # --- 엑셀 디자인 적용 ---
            wb = load_workbook(output_xls)
            ws = wb.active
            
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            right_align = Alignment(horizontal='right', vertical='center')
            header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
            header_font = Font(bold=True)

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = thin_border
                    if cell.row == 1:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = center_align
                    else:
                        if cell.column in [5, 6]: # 금액 열
                            cell.number_format = '#,##0'
                            cell.alignment = right_align
                        elif cell.column == 10: # 계좌번호 열
                            cell.number_format = '@' # 텍스트 강제
                            cell.value = str(cell.value)
                            cell.alignment = center_align
                        else:
                            cell.alignment = center_align
            
            # A4 용지 설정 (9)
            ws.page_setup.paperSize = 9
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = False

            # 칸 너비 자동 조절
            for column_cells in ws.columns:
                length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                ws.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 4

            wb.save(output_xls)
            
            # 결과물 읽기
            with open(output_xls, "rb") as f:
                data = f.read()
            
            # 청소
            if os.path.exists(input_pdf): os.remove(input_pdf)
            if os.path.exists(output_xls): os.remove(output_xls)
            
            return data, output_xls
            
    except Exception as e:
        return None, str(e)
    
    return None, "변환 실패"

# --- 메인 실행 로직 ---
if uploaded_files:
    st.write(f"✅ **{len(uploaded_files)}개**의 파일이 선택되었습니다.")
    
    if st.button("🚀 일괄 변환 시작 (클릭)"):
        status_area = st.container()
        
        # 파일 하나씩 순서대로 처리
        for idx, file in enumerate(uploaded_files):
            with status_area:
                with st.expander(f"🔄 처리 중... {file.name}", expanded=True):
                    excel_data, result_name = process_pdf(file.getbuffer(), file.name)
                    
                    if excel_data and isinstance(excel_data, bytes):
                        st.success(f"완료! ({result_name})")
                        
                        # [핵심] 원본 파일명으로 다운로드 버튼 생성
                        st.download_button(
                            label=f"📥 {result_name} 다운로드",
                            data=excel_data,
                            file_name=result_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"down_{idx}"
                        )
                    else:
                        st.error(f"실패: {file.name} / 사유: {result_name}")
        
        st.success("🎉 모든 작업이 끝났습니다!")