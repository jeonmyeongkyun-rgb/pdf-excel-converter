import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import pandas as pd
import os
import io
import zipfile
import glob # 파일 청소용 도구
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# ==============================================================================
# [필수] API 키 입력
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

st.set_page_config(page_title="Premium PDF Converter", page_icon="🥂", layout="wide")

# ----------------------------------------------------------------
# 🧹 [초강력 청소] 시작할 때 temp 파일이나 xlsx 파일이 보이면 다 지움
# ----------------------------------------------------------------
def clean_up_trash():
    trash_files = glob.glob("temp_*.pdf") + glob.glob("*.xlsx")
    for f in trash_files:
        try:
            os.remove(f)
        except:
            pass
# 앱이 실행될 때마다 청소 한 번 하고 시작
clean_up_trash()
# ----------------------------------------------------------------

# 스타일 설정
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;700&display=swap');
    .stApp { background-color: #F3F4F6; color: #1F2937; font-family: 'Noto Sans KR', sans-serif; }
    h1 { color: #111827 !important; text-align: center; font-weight: 800; margin-bottom: 0px; }
    .subtitle { text-align: center; color: #6B7280; margin-bottom: 30px; }
    div.stButton > button { background-color: #2563EB; color: white; border-radius: 8px; font-weight: bold; border: none; padding: 0.5rem 1rem; width: 100%; }
    div.stButton > button:hover { background-color: #1D4ED8; }
    [data-testid='stFileUploader'] { background: white; border: 2px dashed #D1D5DB; border-radius: 12px; padding: 20px; }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>PDF ➡️ Excel 변환기</h1>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>금융 거래내역, 견적서, 표 완벽 변환</div>", unsafe_allow_html=True)

if "여기에" in GOOGLE_API_KEY:
    st.error("🚨 API 키를 입력해주세요.")
    st.stop()

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []
if 'last_uploaded_ids' not in st.session_state:
    st.session_state.last_uploaded_ids = ""

uploaded_files = st.file_uploader("파일을 드래그하세요 (PDF)", type="pdf", accept_multiple_files=True)

def process_pdf_universal(file_bytes, original_name):
    # 1. 임시 PDF 저장
    temp_input = f"temp_{original_name}"
    
    # 2. 결과 파일명 강제 지정 (.xlsx)
    file_root = os.path.splitext(original_name)[0]
    final_output_xls = f"{file_root}.xlsx"
    
    # PDF 파일 생성
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
            이미지 속 '표(Table)' 데이터를 CSV로 변환해.
            규칙: 배경 글자 무시, 표 구조 유지, 숫자 쉼표 유지. 오직 CSV만 출력.
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

            # 엑셀 꾸미기
            wb = load_workbook(final_output_xls)
            ws = wb.active
            
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            header_fill = PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid")
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
        # [무조건 실행] 작업이 끝나면 임시 파일들은 즉시 삭제
        if os.path.exists(temp_input): os.remove(temp_input)
        if os.path.exists(final_output_xls): os.remove(final_output_xls)

    return None, "표 없음"

# --- 자동 실행 ---
if uploaded_files:
    current_file_ids = "".join([f.name + str(f.size) for f in uploaded_files])
    
    if current_file_ids != st.session_state.last_uploaded_ids:
        st.session_state.processed_files = []
        st.session_state.last_uploaded_ids = current_file_ids
        
        # 기존 쓰레기 파일 한번 더 청소
        clean_up_trash()
        
        progress_bar = st.progress(0, text="분석 중...")
        total = len(uploaded_files)
        
        for idx, file in enumerate(uploaded_files):
            progress_bar.progress(int((idx / total) * 100), text=f"변환 중... ({idx+1}/{total}) : {file.name}")
            
            # 변환 실행
            excel_data, result_name = process_pdf_universal(file.getbuffer(), file.name)
            
            if excel_data:
                st.session_state.processed_files.append({
                    "name": result_name, # 이게 바로 .xlsx 이름입니다
                    "data": excel_data
                })
        
        progress_bar.progress(100, text="완료!")

# --- 결과 표시 ---
if st.session_state.processed_files:
    st.success(f"총 {len(st.session_state.processed_files)}개 변환 완료")
    st.markdown("---")
    
    # ZIP 다운로드
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in st.session_state.processed_files:
            zf.writestr(f['name'], f['data'])
            
    st.download_button(
        label="📦 전체 압축 다운로드 (.ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Excel_Files.zip",
        mime="application/zip",
        use_container_width=True
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 개별 다운로드 (카드형)
    for i, f in enumerate(st.session_state.processed_files):
        with st.container():
            col1, col2 = st.columns([3, 1])
            with col1:
                # 여기서 엑셀 아이콘(📊)과 .xlsx 이름을 확인하세요!
                st.markdown(f"**📊 {f['name']}**")
            with col2:
                st.download_button(
                    label="다운로드",
                    data=f['data'],
                    file_name=f['name'], # 여기서 강제로 .xlsx로 다운로드됨
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"down_{i}",
                    use_container_width=True
                )