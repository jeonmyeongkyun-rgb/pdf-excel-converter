import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import pandas as pd
import os
import io
import zipfile  # 압축 파일 만들기용
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# ==============================================================================
# [필수] AIzaSyBQjCBOwYNjiy5Z-Ej_OQR8XSUHsbfvKPk
GOOGLE_API_KEY = "AIzaSyBQjCBOwYNjiy5Z-Ej_OQR8XSUHsbfvKPk"
# ==============================================================================

genai.configure(api_key=GOOGLE_API_KEY)
try:
    model = genai.GenerativeModel('gemini-2.5-flash')
except:
    st.error("모델 오류: gemini-1.5-flash 모델을 찾을 수 없습니다.")

st.set_page_config(page_title="Gemini PDF 변환기 Pro", page_icon="💳", layout="wide")
st.title("💳 대량 PDF 엑셀 변환기 (사라짐 방지 + ZIP 다운)")

if "여기에" in GOOGLE_API_KEY:
    st.error("🚨 코드 16번째 줄에 API 키를 입력해주세요!")
    st.stop()

# --- [핵심 1] 기억 저장소 초기화 (새로고침 되어도 데이터 유지) ---
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []

# 파일 업로더
uploaded_files = st.file_uploader(
    "변환할 PDF 파일들을 여기에 모두 드래그하세요", 
    type="pdf", 
    accept_multiple_files=True
)

# --- 변환 함수 (기존과 동일) ---
def process_pdf(file_bytes, original_name):
    input_pdf = f"temp_{original_name}"
    output_xls = os.path.splitext(original_name)[0] + ".xlsx"
    
    with open(input_pdf, "wb") as f:
        f.write(file_bytes)

    try:
        doc = fitz.open(input_pdf)
        all_dfs = []
        columns = ["거래일자", "거래시간", "상태", "거래구분", "거래금액", "표면잔액", "취급점", "적요", "은행명", "계좌번호"]

        # 페이지별 처리
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            image_parts = [{"mime_type": "image/png", "data": img_data}]
            
            prompt = """
            이 이미지의 은행 거래내역 표를 파이프(|) 기호로 구분된 텍스트로 추출해.
            [규칙] 10개 항목: 날짜|시간|상태|구분|거래금액|표면잔액|취급점|적요|은행명|계좌번호
            '표면잔액'과 '취급점' 구분선(|) 필수. 금액 콤마 유지. 계좌번호 숫자만. 헤더 무시.
            """
            
            response = model.generate_content([prompt, image_parts[0]])
            raw_text = response.text.strip().replace("```", "")
            
            data_rows = []
            for line in raw_text.split('\n'):
                if "|" in line:
                    parts = line.split('|')
                    if len(parts) < 10: parts += [""] * (10 - len(parts))
                    if len(parts) > 10: parts = parts[:10]
                    parts = [p.strip() for p in parts]
                    data_rows.append(parts)
            
            if data_rows:
                df = pd.DataFrame(data_rows, columns=columns)
                all_dfs.append(df)

        if all_dfs:
            final_df = pd.concat(all_dfs, ignore_index=True)
            for col in ["거래금액", "표면잔액"]:
                final_df[col] = final_df[col].astype(str).str.replace(',', '').str.replace('원', '')
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            final_df.to_excel(output_xls, index=False)

            # 디자인 적용
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
                        if cell.column in [5, 6]: 
                            cell.number_format = '#,##0'
                            cell.alignment = right_align
                        elif cell.column == 10:
                            cell.number_format = '@'
                            cell.value = str(cell.value)
                            cell.alignment = center_align
                        else:
                            cell.alignment = center_align
            
            ws.page_setup.paperSize = 9
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = False

            for column_cells in ws.columns:
                length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                ws.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 4

            wb.save(output_xls)
            
            with open(output_xls, "rb") as f:
                data = f.read()
            
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
        # 기존 기록 초기화
        st.session_state.processed_files = []
        
        # --- [핵심 2] 전체 진행률 바 생성 ---
        progress_text = "작업 시작..."
        my_bar = st.progress(0, text=progress_text)
        
        total_files = len(uploaded_files)
        
        for idx, file in enumerate(uploaded_files):
            # 진행률 업데이트 (0% ~ 100%)
            percent = int(((idx) / total_files) * 100)
            my_bar.progress(percent, text=f"🔄 처리 중 ({idx+1}/{total_files}): {file.name}")
            
            excel_data, result_name = process_pdf(file.getbuffer(), file.name)
            
            if excel_data:
                # 결과물을 기억 저장소(Session State)에 저장
                st.session_state.processed_files.append({
                    "name": result_name,
                    "data": excel_data
                })
        
        my_bar.progress(100, text="✅ 모든 변환이 완료되었습니다!")

# --- 결과 화면 표시 (저장소에 데이터가 있을 때만 표시) ---
if st.session_state.processed_files:
    st.success(f"🎉 총 {len(st.session_state.processed_files)}개의 파일 변환 완료!")
    
    # 1. 개별 다운로드 버튼 보여주기
    st.write("### 📂 개별 파일 다운로드")
    cols = st.columns(3) # 3열로 예쁘게 배치
    for i, file_info in enumerate(st.session_state.processed_files):
        with cols[i % 3]:
            st.download_button(
                label=f"📥 {file_info['name']}",
                data=file_info['data'],
                file_name=file_info['name'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"btn_{i}"
            )
    
    st.divider() # 구분선
    
    # 2. [핵심 3] 전체 ZIP 다운로드 버튼 생성
    st.write("### 📦 한 번에 다운로드 (ZIP)")
    
    # 메모리 상에서 ZIP 파일 만들기
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for file_info in st.session_state.processed_files:
            zf.writestr(file_info['name'], file_info['data'])
    
    st.download_button(
        label="📦 전체 파일 압축 다운로드 (.zip)",
        data=zip_buffer.getvalue(),
        file_name="변환결과_모음.zip",
        mime="application/zip"
    )