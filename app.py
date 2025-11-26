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

# Gemini 설정
genai.configure(api_key=GOOGLE_API_KEY)

try:
    model = genai.GenerativeModel('gemini-2.5-flash')
except:
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
    except:
        st.error("❌ 모델 로딩 실패. API 키를 확인해주세요.")

# 페이지 기본 설정
st.set_page_config(page_title="Premium PDF Converter", page_icon="🥂", layout="wide")

# --------------------------------------------------------------------------------
# 🎨 [디자인 핵심] 커스텀 CSS (호텔 라운지 스타일)
# --------------------------------------------------------------------------------
st.markdown("""
<style>
    /* 1. 폰트 가져오기 (Google Fonts: Playfair Display - 우아한 명조 느낌) */
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=Noto+Sans+KR:wght@300;400;700&display=swap');

    /* 2. 전체 배경 (깊은 차콜 블랙) */
    .stApp {
        background-color: #121212;
        color: #E0E0E0;
    }

    /* 3. 헤더/제목 스타일 (골드 & 명조체) */
    h1, h2, h3 {
        font-family: 'Playfair Display', serif;
        color: #D4AF37 !important; /* 샴페인 골드 */
        font-weight: 700;
        text-align: center;
        letter-spacing: 1px;
    }
    
    /* 부제목 스타일 */
    .subtitle {
        text-align: center;
        color: #A0A0A0;
        font-family: 'Noto Sans KR', sans-serif;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }

    /* 4. 파일 업로더 스타일 (심플하고 모던하게) */
    [data-testid='stFileUploader'] {
        background-color: #1E1E1E;
        border: 1px solid #333;
        border-radius: 10px;
        padding: 20px;
    }
    [data-testid='stFileUploader'] section {
        background-color: #1E1E1E;
    }
    
    /* 5. 버튼 스타일 (골드 그라데이션) */
    div.stButton > button {
        background: linear-gradient(135deg, #D4AF37 0%, #C5A059 100%);
        color: #000000;
        font-family: 'Noto Sans KR', sans-serif;
        font-weight: bold;
        border: none;
        border-radius: 30px; /* 둥근 캡슐 모양 */
        padding: 0.6rem 2rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(212, 175, 55, 0.3);
    }
    div.stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(212, 175, 55, 0.5);
        color: #000000;
        border: none;
    }

    /* 6. 성공/에러 메시지 박스 스타일 */
    .stSuccess, .stInfo, .stWarning {
        background-color: #1E1E1E !important;
        color: #D4AF37 !important;
        border-left: 5px solid #D4AF37 !important;
    }
    
    /* 7. 진행바 색상 변경 */
    .stProgress > div > div > div > div {
        background-color: #D4AF37;
    }

    /* 8. 구분선 */
    hr {
        border-color: #333;
    }
</style>
""", unsafe_allow_html=True)
# --------------------------------------------------------------------------------

# 타이틀 섹션 (가운데 정렬)
st.markdown("<h1>PREMIUM PDF CONVERTER</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitle'>Gemini 2.5 AI가 제공하는 고품격 문서 변환 서비스</p>", unsafe_allow_html=True)
st.markdown("---")

if "여기에" in GOOGLE_API_KEY:
    st.error("🚨 API 키가 설정되지 않았습니다. 코드 17번째 줄을 확인해주세요.")
    st.stop()

# 세션 상태 초기화
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []
if 'last_uploaded_ids' not in st.session_state:
    st.session_state.last_uploaded_ids = ""

# 파일 업로더
uploaded_files = st.file_uploader(
    "변환할 PDF 문서를 이곳에 놓아주세요.", 
    type="pdf", 
    accept_multiple_files=True
)

# --- 변환 함수 (기능은 동일) ---
def process_pdf_universal(file_bytes, original_name):
    temp_input_pdf = f"temp_{original_name}"
    file_root = os.path.splitext(original_name)[0]
    final_output_xls = f"{file_root}.xlsx"
    
    with open(temp_input_pdf, "wb") as f:
        f.write(file_bytes)

    try:
        doc = fitz.open(temp_input_pdf)
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
            final_df.to_excel(final_output_xls, index=False)

            wb = load_workbook(final_output_xls)
            ws = wb.active
            
            # 엑셀 디자인 (심플)
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

            wb.save(final_output_xls)
            
            with open(final_output_xls, "rb") as f:
                data = f.read()
            
            if os.path.exists(temp_input_pdf): os.remove(temp_input_pdf)
            if os.path.exists(final_output_xls): os.remove(final_output_xls)
            
            return data, final_output_xls
            
    except Exception as e:
        return None, str(e)
    return None, "표 없음"


# --- 자동 실행 로직 ---
if uploaded_files:
    current_file_ids = "".join([f.name + str(f.size) for f in uploaded_files])
    
    if current_file_ids != st.session_state.last_uploaded_ids:
        st.session_state.processed_files = []
        st.session_state.last_uploaded_ids = current_file_ids
        
        # 진행바 컨테이너 (깔끔하게 보이기 위함)
        with st.container():
            st.write(" ") # 여백
            progress_bar = st.progress(0, text="AI가 문서를 분석하고 있습니다...")
            total = len(uploaded_files)
            
            for idx, file in enumerate(uploaded_files):
                progress_bar.progress(int((idx / total) * 100), text=f"Processing... ({idx+1}/{total}) : {file.name}")
                
                excel_data, result_name = process_pdf_universal(file.getbuffer(), file.name)
                
                if excel_data:
                    st.session_state.processed_files.append({
                        "name": result_name,
                        "data": excel_data
                    })
            
            progress_bar.progress(100, text="Completed.")
            st.success("모든 변환이 완료되었습니다.")


# --- 결과 화면 ---
if st.session_state.processed_files:
    st.markdown("---")
    st.markdown("### 📥 Download Results")
    
    # 3열 레이아웃으로 버튼 정렬
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # ZIP 다운로드 버튼을 가장 크게/눈에 띄게 배치
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for f in st.session_state.processed_files:
                zf.writestr(f['name'], f['data'])
                
        st.download_button(
            label="📦 전체 일괄 다운로드 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Converted_Files.zip",
            mime="application/zip",
            use_container_width=True # 버튼 꽉 차게
        )
    
    st.write(" ") # 여백
    st.write("**개별 파일 다운로드:**")
    
    # 개별 파일 리스트업
    for i, f in enumerate(st.session_state.processed_files):
        col_a, col_b = st.columns([4, 1])
        with col_a:
            st.info(f"📄 {f['name']}") # 파일명 예쁘게 표시
        with col_b:
            st.download_button(
                label="다운로드",
                data=f['data'],
                file_name=f['name'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"down_{i}",
                use_container_width=True
            )