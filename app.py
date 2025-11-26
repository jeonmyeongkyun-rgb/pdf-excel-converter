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

# [수정 완료] 사용자가 확인한 최신 모델 적용 (gemini-2.5-flash)
try:
    model = genai.GenerativeModel('gemini-2.5-flash')
except:
    st.error("모델 로딩 실패. API 키가 정확한지 확인해주세요.")

# 화면 설정
st.set_page_config(page_title="Gemini 2.5 PDF 변환기", page_icon="⚡️", layout="wide")
st.title("⚡️ Gemini 2.5 PDF ➡️ 엑셀 변환기 (범용)")
st.write("2025년 최신 모델(Gemini 2.5)을 사용하여 모든 종류의 표를 엑셀로 변환합니다.")

# 키 입력 실수 방지
if "여기에" in GOOGLE_API_KEY:
    st.error("🚨 코드 16번째 줄에 '새로 발급받은 API 키'를 입력해주세요!")
    st.stop()

# 세션 상태 초기화 (새로고침 방지)
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []

# 파일 업로드
uploaded_files = st.file_uploader(
    "변환할 PDF 파일들을 여기에 드래그하세요 (여러 개 가능)", 
    type="pdf", 
    accept_multiple_files=True
)

# --- 변환 함수 ---
def process_pdf_universal(file_bytes, original_name):
    input_pdf = f"temp_{original_name}"
    # 확장자 변경 (.pdf -> .xlsx)
    output_xls = os.path.splitext(original_name)[0] + ".xlsx"
    
    with open(input_pdf, "wb") as f:
        f.write(file_bytes)

    try:
        doc = fitz.open(input_pdf)
        all_dfs = []
        
        for i, page in enumerate(doc):
            # 이미지 변환 (2배 확대)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            image_parts = [{"mime_type": "image/png", "data": img_data}]
            
            # [Gemini 2.5에게 내리는 범용 프롬프트]
            prompt = """
            이 이미지에서 '표(Table)' 데이터를 찾아서 CSV 형식으로 변환해줘.
            
            [규칙]
            1. 문서의 종류(금융, 견적서, 명단 등)에 상관없이 표 구조를 보이는 그대로 유지해.
            2. 배경의 워터마크나 표 바깥의 불필요한 글자는 무시해.
            3. 헤더(제목 줄)가 있다면 포함해.
            4. 금액이나 숫자에 있는 쉼표(,)는 제거하지 말고 그대로 둬.
            5. 오직 CSV 데이터만 출력해. (설명이나 마크다운 태그 ```csv 넣지 마)
            """
            
            response = model.generate_content([prompt, image_parts[0]])
            csv_text = response.text.strip().replace("```csv", "").replace("```", "")
            
            try:
                # CSV를 데이터프레임으로 변환 (칸 개수 자동 감지)
                df = pd.read_csv(io.StringIO(csv_text))
                if not df.empty:
                    all_dfs.append(df)
            except Exception as e:
                print(f"페이지 {i+1} 변환 건너뜀: {e}")

        if all_dfs:
            # 모든 페이지 합치기
            final_df = pd.concat(all_dfs, ignore_index=True)
            final_df.to_excel(output_xls, index=False)

            # 엑셀 디자인 (선 그리기 + 자동 너비)
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
                        # 숫자인 경우 엑셀이 숫자로 인식하도록 처리 시도
                        try:
                            if isinstance(cell.value, str) and cell.value.replace(',', '').replace('.', '').isdigit():
                                pass # 텍스트로 유지하되(지수표현 방지) 정렬은 가운데로
                        except:
                            pass

            # A4 용지 설정
            ws.page_setup.paperSize = 9
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = False

            # 칸 너비 자동 조절
            for column_cells in ws.columns:
                try:
                    length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                    final_width = max(10, min(length + 4, 60)) # 최소 10, 최대 60
                    ws.column_dimensions[get_column_letter(column_cells[0].column)].width = final_width
                except:
                    pass

            wb.save(output_xls)
            
            with open(output_xls, "rb") as f:
                data = f.read()
            
            # 임시 파일 삭제
            if os.path.exists(input_pdf): os.remove(input_pdf)
            if os.path.exists(output_xls): os.remove(output_xls)
            
            return data, output_xls
            
    except Exception as e:
        return None, str(e)
    
    return None, "표를 찾지 못함"


# --- 메인 실행 로직 ---
if uploaded_files:
    if st.button("🚀 일괄 변환 시작 (클릭)"):
        st.session_state.processed_files = [] # 초기화
        
        progress_bar = st.progress(0, text="작업 시작...")
        total = len(uploaded_files)
        
        for idx, file in enumerate(uploaded_files):
            # 진행률 바 업데이트
            progress_bar.progress(int((idx / total) * 100), text=f"🔄 Gemini 2.5가 변환 중... ({idx+1}/{total}): {file.name}")
            
            excel_data, result_name = process_pdf_universal(file.getbuffer(), file.name)
            
            if excel_data:
                st.session_state.processed_files.append({
                    "name": result_name,
                    "data": excel_data
                })
        
        progress_bar.progress(100, text="✅ 모든 변환 완료!")

# 결과 표시 화면
if st.session_state.processed_files:
    st.success(f"🎉 총 {len(st.session_state.processed_files)}개의 문서 변환 완료!")
    
    # 1. 전체 ZIP 다운로드
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in st.session_state.processed_files:
            zf.writestr(f['name'], f['data'])
            
    st.download_button(
        label="📦 전체 압축 다운로드 (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="변환결과_모음.zip",
        mime="application/zip"
    )
    
    st.divider()
    
    # 2. 개별 다운로드
    st.write("📂 개별 파일 다운로드")
    cols = st.columns(3)
    for i, f in enumerate(st.session_state.processed_files):
        with cols[i % 3]:
            st.download_button(
                label=f"📥 {f['name']}",
                data=f['data'],
                file_name=f['name'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"down_{i}"
            )