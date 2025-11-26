import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import pandas as pd
import os
import io
import zipfile
import glob
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# ==============================================================================
# [필수] 구글 AI Studio에서 발급받은 키를 여기에 넣으세요
GOOGLE_API_KEY = "AIzaSyDAGuC0v4hhdwegQhlxNWwAPwe3Vaym0rQ"
# ==============================================================================

genai.configure(api_key=GOOGLE_API_KEY)

try:
    model = genai.GenerativeModel('gemini-2.5-flash')
except:
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
    except:
        st.error("❌ 모델 로딩 실패.")

st.set_page_config(page_title="Smart PDF Converter", page_icon="🧠", layout="wide")

# --- 파일 청소 ---
def clean_up_trash():
    for f in glob.glob("temp_*.pdf") + glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass
clean_up_trash()

# --- 디자인 ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;700&display=swap');
    .stApp { background-color: #F8FAFC; font-family: 'Noto Sans KR', sans-serif; color: #334155; }
    h1 { color: #1E293B; font-weight: 800; text-align: center; }
    .subtitle { text-align: center; color: #64748B; margin-bottom: 2rem; }
    [data-testid='stFileUploader'] { background: white; border: 2px dashed #CBD5E1; border-radius: 12px; padding: 30px; }
    div.stButton > button { background-color: #3B82F6; color: white; border: none; border-radius: 8px; padding: 0.6rem; width: 100%; font-weight: bold; }
    div.stButton > button:hover { background-color: #2563EB; }
    .stSuccess, .stError { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>Smart PDF ➡️ Excel 변환기</h1>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>금융 문서(콤마 포함)와 일반 표를 모두 똑똑하게 처리합니다.</div>", unsafe_allow_html=True)

if "여기에" in GOOGLE_API_KEY:
    st.error("🚨 API 키를 입력해주세요.")
    st.stop()

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []
if 'last_uploaded_ids' not in st.session_state:
    st.session_state.last_uploaded_ids = ""

uploaded_files = st.file_uploader("PDF 파일을 드래그하세요", type="pdf", accept_multiple_files=True)

def process_pdf_smart(file_bytes, original_name):
    temp_input = f"temp_{original_name}"
    file_root = os.path.splitext(original_name)[0]
    final_output_xls = f"{file_root}.xlsx"
    
    with open(temp_input, "wb") as f:
        f.write(file_bytes)

    try:
        doc = fitz.open(temp_input)
        all_dfs = []
        
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            image_parts = [{"mime_type": "image/png", "data": img_data}]
            
            # --- [핵심 수정] 프롬프트 전략 변경 ---
            # CSV(콤마) 대신 파이프(|)를 사용하라고 강제합니다.
            # 이렇게 하면 금액(1,000)의 콤마 때문에 칸이 찢어지는 것을 막을 수 있습니다.
            prompt = """
            이 이미지에서 '표(Table)' 데이터를 추출해줘.
            
            [엄격한 규칙]
            1. 각 칸(Cell)의 데이터는 반드시 파이프기호(|)로 구분해. (콤마 쓰지 마)
            2. 예시 포맷: 날짜|내용|금액|비고
            3. 배경에 있는 워터마크(KB, 로고 등)나 희미한 글자는 절대 읽지 마.
            4. 표의 헤더(제목 줄)가 있다면 포함해.
            5. 금액의 쉼표(,)는 그대로 유지해.
            6. 오직 데이터만 출력해. (마크다운 태그 없이)
            """
            
            response = model.generate_content([prompt, image_parts[0]])
            
            if not response.text:
                continue # 응답 없으면 다음 페이지로

            # 결과 텍스트 정제
            raw_text = response.text.strip().replace("```", "").replace("csv", "").replace("txt", "")
            
            # 파이프(|)로 된 데이터를 판다스로 변환
            try:
                # sep='|' 옵션이 핵심입니다!
                df = pd.read_csv(io.StringIO(raw_text), sep='|', engine='python')
                
                # 데이터가 너무 적거나(1줄 이하) 깨진 경우 무시
                if len(df) > 0 and len(df.columns) > 1:
                    # 컬럼 이름 앞뒤 공백 제거
                    df.columns = df.columns.str.strip()
                    # 데이터 앞뒤 공백 제거 (문자열인 경우만)
                    df = df.apply(lambda x: x.strip() if isinstance(x, str) else x)
                    all_dfs.append(df)
            except:
                pass

        if not all_dfs:
            return None, "표를 인식하지 못했습니다. (워터마크가 너무 진하거나 표가 아닐 수 있습니다)"

        final_df = pd.concat(all_dfs, ignore_index=True)
        final_df.to_excel(final_output_xls, index=False)

        # 엑셀 디자인
        wb = load_workbook(final_output_xls)
        ws = wb.active
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        header_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
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
                    # 숫자인 경우 엑셀이 숫자로 인식하도록 처리 시도
                    try:
                        if isinstance(cell.value, str):
                            clean_val = cell.value.replace(',', '').replace('.', '').strip()
                            if clean_val.isdigit():
                                # 값은 그대로 두되(텍스트 유지), 오른쪽 정렬 등 서식만 적용 가능
                                pass
                    except:
                        pass

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
        return None, f"에러 발생: {str(e)}"
    
    finally:
        if os.path.exists(temp_input): 
            try: os.remove(temp_input)
            except: pass
        if os.path.exists(final_output_xls): 
            try: os.remove(final_output_xls)
            except: pass

# --- 실행 로직 ---
if uploaded_files:
    current_file_ids = "".join([f.name + str(f.size) for f in uploaded_files])
    
    if current_file_ids != st.session_state.last_uploaded_ids:
        st.session_state.processed_files = []
        st.session_state.last_uploaded_ids = current_file_ids
        
        clean_up_trash()
        
        progress_bar = st.progress(0, text="분석 중...")
        total = len(uploaded_files)
        
        for idx, file in enumerate(uploaded_files):
            progress_bar.progress(int((idx / total) * 100), text=f"변환 중... ({idx+1}/{total}) : {file.name}")
            
            excel_data, result_msg = process_pdf_smart(file.getbuffer(), file.name)
            
            if excel_data:
                st.session_state.processed_files.append({
                    "name": result_msg,
                    "data": excel_data
                })
            else:
                st.error(f"❌ '{file.name}' 실패: {result_msg}")
        
        progress_bar.progress(100, text="완료!")

# --- 결과 화면 ---
if st.session_state.processed_files:
    st.success(f"총 {len(st.session_state.processed_files)}개 변환 성공!")
    st.markdown("---")
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in st.session_state.processed_files:
            zf.writestr(f['name'], f['data'])
            
    st.download_button(
        label="📦 전체 압축 다운로드 (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Converted_Files.zip",
        mime="application/zip",
        use_container_width=True
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    for i, f in enumerate(st.session_state.processed_files):
        with st.container():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**📊 {f['name']}**")
            with col2:
                st.download_button(
                    label="다운로드",
                    data=f['data'],
                    file_name=f['name'],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"down_{i}",
                    use_container_width=True
                )