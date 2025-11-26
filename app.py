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
# [필수] 구글 AI Studio에서 발급받은 키를 여기에 넣으세요
GOOGLE_API_KEY = "AIzaSyBQjCBOwYNjiy5Z-Ej_OQR8XSUHsbfvKPk"
# ==============================================================================

genai.configure(api_key=GOOGLE_API_KEY)

try:
    model = genai.GenerativeModel('gemini-2.5-flash')
except:
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
    except:
        st.error("❌ 모델 로딩 실패.")

# 페이지 설정
st.set_page_config(page_title="Premium PDF Converter", page_icon="🥂", layout="wide")

# 🎨 디자인
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=Noto+Sans+KR:wght@300;400;700&display=swap');
    .stApp { background-color: #F9FAFB; color: #111827; font-family: 'Noto Sans KR', sans-serif; }
    h1 { font-family: 'Playfair Display', serif; color: #111827 !important; text-align: center; font-weight: 800; }
    .subtitle { text-align: center; color: #6B7280; margin-bottom: 2rem; }
    [data-testid='stFileUploader'] { background-color: #FFFFFF; border: 2px dashed #E5E7EB; border-radius: 16px; padding: 30px; }
    div.stButton > button { background-color: #2563EB; color: white; border: none; border-radius: 10px; padding: 0.6rem 1.5rem; font-weight: 600; width: 100%; }
    div.stButton > button:hover { background-color: #1D4ED8; transform: translateY(-1px); }
    .stSuccess, .stInfo, .stError { border-radius: 10px; border: none; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>PDF to Excel Converter</h1>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>깔끔하고 강력한 AI 문서 변환기</div>", unsafe_allow_html=True)

if "여기에" in GOOGLE_API_KEY:
    st.error("🚨 API 키 설정을 확인해주세요.")
    st.stop()

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []
if 'last_uploaded_ids' not in st.session_state:
    st.session_state.last_uploaded_ids = ""

uploaded_files = st.file_uploader("변환할 PDF 파일을 드래그 앤 드롭하세요", type="pdf", accept_multiple_files=True)

def process_pdf_universal(file_bytes, original_name):
    temp_input = f"temp_{original_name}"
    file_root = os.path.splitext(original_name)[0]
    final_output_xls = f"{file_root}.xlsx"
    
    # 파일 생성
    with open(temp_input, "wb") as f:
        f.write(file_bytes)

    try:
        doc = fitz.open(temp_input)
        all_dfs = []
        
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            image_parts = [{"mime_type": "image/png", "data": img_data}]
            
            prompt = """
            이 이미지에서 '표(Table)' 데이터를 찾아서 CSV 형식으로 변환해줘.
            배경의 워터마크나 로고는 무시해. 표 구조를 유지해.
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
            final_df.to_excel(final_output_xls, index=False)

            wb = load_workbook(final_output_xls)
            ws = wb.active
            
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            header_fill = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
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

            wb.save(final_output_xls)
            
            with open(final_output_xls, "rb") as f:
                data = f.read()
            
            return data, final_output_xls
            
    except Exception as e:
        return None, str(e)
    
    finally:
        # [핵심 수정] 성공하든 실패하든 무조건 임시 파일 삭제!
        if os.path.exists(temp_input): 
            try: os.remove(temp_input)
            except: pass
        if os.path.exists(final_output_xls): 
            try: os.remove(final_output_xls)
            except: pass

    return None, "표 없음"

# --- 자동 실행 로직 ---
if uploaded_files:
    current_file_ids = "".join([f.name + str(f.size) for f in uploaded_files])
    
    if current_file_ids != st.session_state.last_uploaded_ids:
        st.session_state.processed_files = []
        st.session_state.last_uploaded_ids = current_file_ids
        
        progress_text = "문서를 분석하고 있습니다..."
        my_bar = st.progress(0, text=progress_text)
        total = len(uploaded_files)
        
        for idx, file in enumerate(uploaded_files):
            my_bar.progress(int((idx / total) * 100), text=f"🔄 변환 중... ({idx+1}/{total}) : {file.name}")
            excel_data, result_name = process_pdf_universal(file.getbuffer(), file.name)
            
            if excel_data:
                st.session_state.processed_files.append({
                    "name": result_name,
                    "data": excel_data
                })
        
        my_bar.progress(100, text="완료!")
        st.success("모든 변환이 완료되었습니다.")

# --- 결과 화면 ---
if st.session_state.processed_files:
    st.markdown("<br>", unsafe_allow_html=True)
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in st.session_state.processed_files:
            zf.writestr(f['name'], f['data'])
            
    st.download_button(
        label="📦 전체 파일 한 번에 다운로드 (.ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Converted_Files.zip",
        mime="application/zip",
        use_container_width=True
    )
    
    st.markdown("---")
    st.markdown("#### 📂 개별 파일 목록")
    
    for i, f in enumerate(st.session_state.processed_files):
        with st.container():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"""
                <div style="padding: 15px; background: white; border-radius: 10px; border: 1px solid #E5E7EB; display: flex; align-items: center; margin-bottom: 10px;">
                    <span style="font-size: 1.2rem; margin-right: 10px;">📄</span>
                    <span style="font-weight: 600; color: #374151;">{f['name']}</span>
                </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown('<div style="height: 5px;"></div>', unsafe_allow_html=True)
                st.download_button(
                    label="다운로드",
                    data=f['data'],
                    file_name=f['name'],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"down_{i}",
                    use_container_width=True
                )