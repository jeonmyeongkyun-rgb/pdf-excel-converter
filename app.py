import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import pandas as pd
import os
import io
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# ==============================================================================
# [필수] 구글 AI Studio에서 새로 발급받은 키를 여기에 넣으세요!
GOOGLE_API_KEY = "AIzaSyBQjCBOwYNjiy5Z-Ej_OQR8XSUHsbfvKPk"
# ==============================================================================

# Gemini 설정
genai.configure(api_key=GOOGLE_API_KEY)

# 모델 설정
try:
    model = genai.GenerativeModel('gemini-2.5-flash')
except:
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
    except:
        st.error("모델 로딩 실패. API 키를 확인해주세요.")

st.set_page_config(page_title="Gemini PDF 변환기", page_icon="⚡️", layout="wide")
st.title("⚡️ Gemini PDF ➡️ 엑셀 변환기 (완벽판)")

if "여기에" in GOOGLE_API_KEY:
    st.error("🚨 코드 16번째 줄에 API 키를 입력해주세요!")
    st.stop()

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []

uploaded_files = st.file_uploader(
    "변환할 PDF 파일들을 드래그하세요", 
    type="pdf", 
    accept_multiple_files=True
)

def process_pdf_universal(file_bytes, original_name):
    input_pdf = f"temp_{original_name}"
    
    # [수정된 부분] 파일 이름 강제로 .xlsx로 바꾸기 (대소문자 상관없이)
    file_root = os.path.splitext(original_name)[0] # 확장자 떼어내기
    output_xls = f"{file_root}.xlsx" # 뒤에 .xlsx 붙이기
    
    with open(input_pdf, "wb") as f:
        f.write(file_bytes)

    try:
        doc = fitz.open(input_pdf)
        all_dfs = []
        
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            image_parts = [{"mime_type": "image/png", "data": img_data}]
            
            prompt = """
            이 이미지에서 '표(Table)' 데이터를 찾아서 CSV 형식으로 변환해줘.
            배경의 워터마크는 무시하고, 표 구조를 그대로 유지해.
            숫자의 쉼표는 유지해. 오직 CSV 데이터만 출력해.
            """
            
            response = model.generate_content([prompt, image_parts[0]])
            csv_text = response.text.strip().replace("```csv", "").replace("```", "")
            
            try:
                df = pd.read_csv(io.StringIO(csv_text))
                if not df.empty:
                    all_dfs.append(df)
            except:
                pass

        if all_dfs:
            final_df = pd.concat(all_dfs, ignore_index=True)
            final_df.to_excel(output_xls, index=False)

            # 엑셀 디자인
            wb = load_workbook(output_xls)
            ws = wb.active
            
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
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
                        cell.alignment = center_align

            ws.page_setup.paperSize = 9
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = False

            for column_cells in ws.columns:
                try:
                    length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                    ws.column_dimensions[get_column_letter(column_cells[0].column)].width = max(10, min(length + 4, 60))
                except:
                    pass

            wb.save(output_xls)
            
            with open(output_xls, "rb") as f:
                data = f.read()
            
            if os.path.exists(input_pdf): os.remove(input_pdf)
            if os.path.exists(output_xls): os.remove(output_xls)
            
            return data, output_xls
            
    except Exception as e:
        return None, str(e)
    
    return None, "표 없음"


if uploaded_files:
    if st.button("🚀 일괄 변환 시작"):
        st.session_state.processed_files = []
        progress_bar = st.progress(0, text="작업 시작...")
        total = len(uploaded_files)
        
        for idx, file in enumerate(uploaded_files):
            progress_bar.progress(int((idx / total) * 100), text=f"🔄 변환 중 ({idx+1}/{total}): {file.name}")
            
            excel_data, result_name = process_pdf_universal(file.getbuffer(), file.name)
            
            if excel_data:
                st.session_state.processed_files.append({
                    "name": result_name,
                    "data": excel_data
                })
        
        progress_bar.progress(100, text="✅ 완료!")

if st.session_state.processed_files:
    st.success(f"🎉 총 {len(st.session_state.processed_files)}개 변환 완료!")
    
    # ZIP 다운로드
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in st.session_state.processed_files:
            zf.writestr(f['name'], f['data'])
            
    st.download_button(
        label="📦 전체 압축 다운로드 (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="PDF변환결과.zip",
        mime="application/zip"
    )
    
    st.divider()
    
    # 개별 다운로드
    cols = st.columns(3)
    for i, f in enumerate(st.session_state.processed_files):
        with cols[i % 3]:
            st.download_button(
                label=f"📥 {f['name']}",
                data=f['data'],
                # [중요] 파일 이름을 여기서 .xlsx로 확실하게 지정
                file_name=f['name'], 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"down_{i}"
            )