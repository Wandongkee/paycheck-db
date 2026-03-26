import os
import streamlit as st
import pandas as pd
import openpyxl
import io

# ==========================================
# ⚙️ 스크립트 실행 경로 자동 설정
# ==========================================
try:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_dir)
except NameError:
    pass 

# ==========================================
# ⚙️ 페이지 설정
# ==========================================
st.set_page_config(page_title="급여DB 자동 통합 툴", page_icon="💰", layout="centered")

st.title("💰 급여DB 자동 통합 툴")
st.markdown("""
급여 마스터 파일과 각 팀별 OT 파일을 업로드하면 데이터를 자동으로 매칭하여 최종 엑셀 파일을 생성합니다.
* **지원 형식**: `.xlsx`, `.xls` 파일 모두 업로드 가능합니다.
* **업로드 방식**: 팀별 OT 파일은 **파일명과 관계없이** 모두 한 번에 드래그 앤 드롭으로 올려주시면 됩니다.
""")

# ==========================================
# 🛠️ 기능 함수
# ==========================================
def convert_xls_to_xlsx_buffer(uploaded_file):
    """.xls 파일을 메모리 상에서 .xlsx 파일로 변환하는 함수"""
    df = pd.read_excel(uploaded_file, engine='xlrd')
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    return buffer

def load_ot_data_from_uploaded_file(uploaded_file):
    """업로드된 OT 엑셀 파일 내의 모든 시트에서 {이름: 금액} 데이터를 모아오는 함수"""
    combined_data = {}
    try:
        if uploaded_file.name.lower().endswith('.xls'):
            file_to_read = convert_xls_to_xlsx_buffer(uploaded_file)
        else:
            file_to_read = uploaded_file

        xls = pd.ExcelFile(file_to_read)
        
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(file_to_read, sheet_name=sheet, header=None)
                if len(df.columns) < 20:
                    continue

                df = df.iloc[7:] 
                valid_rows = df[df[4].notna()]
                
                for _, row in valid_rows.iterrows():
                    name = str(row[4]).strip()
                    amount = row[19]
                    
                    try:
                        amount = float(amount)
                        combined_data[name] = amount
                    except:
                        pass
            except Exception:
                continue
                
        return combined_data
    except Exception as e:
        st.error(f"'{uploaded_file.name}' 처리 중 오류 발생: {e}")
        return {}

def move_column(ws, col_from, col_to):
    """열 이동 함수"""
    max_row = ws.max_row
    col_data = [ws.cell(row=r, column=col_from).value for r in range(1, max_row + 1)]
    ws.insert_cols(col_to)
    for i, val in enumerate(col_data):
        ws.cell(row=i+1, column=col_to).value = val
    del_target = col_from + 1 if col_to <= col_from else col_from
    ws.delete_cols(del_target)

# ==========================================
# 🚀 메인 데이터 처리 로직
# ==========================================
def process_salary_master(db_file, ot_files):
    if db_file.name.lower().endswith('.xls'):
        db_to_process = convert_xls_to_xlsx_buffer(db_file)
    else:
        db_to_process = db_file

    wb = openpyxl.load_workbook(db_to_process)
    ws = wb.active

    # [Step 1] 행 삭제 & 이름 정리
    targets = ["부서별총계", "사업장별총계", "총계"]
    for row in range(ws.max_row, 1, -1):
        if ws.cell(row=row, column=1).value in targets:
            ws.delete_rows(row)
            
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=4)
        if cell.value:
            cell.value = str(cell.value).split("(")[0].strip()

    ws.freeze_panes = 'H2'

    # [Step 2] 열 이동
    move_column(ws, 38, 18) # AL -> R
    move_column(ws, 26, 18) # Z -> R
    move_column(ws, 30, 17) # AD -> Q

    # [Step 3] 열 삽입
    ws.insert_cols(21, 3) 

    # [Step 4] U열 합산 (양중 T/C 인원만)
    for row in range(2, ws.max_row + 1):
        k_val = ws.cell(row=row, column=11).value
        if k_val == "양중(T/C)":
            total = sum([
                float(ws.cell(row=row, column=c).value or 0) 
                for c in range(17, 21)
            ])
            ws.cell(row=row, column=21).value = total

    # [Step 5] OT 파일 데이터 통합 (수정된 부분)
    # 파일명 상관없이 모든 OT 데이터를 하나의 딕셔너리로 병합
    global_ot_db = {}
    for ot_file in ot_files:
        file_data = load_ot_data_from_uploaded_file(ot_file)
        global_ot_db.update(file_data) # 전체 명단에 추가

    # [Step 6] 데이터 매칭 및 수식 입력 (수정된 부분)
    count_match = 0
    ws.insert_cols(48, 1) # AV열 생성

    for row in range(2, ws.max_row + 1):
        name = ws.cell(row=row, column=4).value   # D열(이름)
        k_val = ws.cell(row=row, column=11).value # K열(직종)
        u_val = ws.cell(row=row, column=21).value # U열(합산금액)
        
        # 파일명/팀명 조건 제거하고 전체 통합 DB에서 이름만으로 검색
        if k_val == "양중(T/C)":
            found_ot = global_ot_db.get(name)
            if found_ot is not None:
                ws.cell(row=row, column=22).value = found_ot # V열에 입력
                count_match += 1

        if u_val not in [None, ""]:
            ws.cell(row=row, column=23).value = f"=U{row}=V{row}" # W열
            
        ws.cell(row=row, column=48).value = f"=AU{row}-X{row}-U{row}-AS{row}" # AV열

    # [Step 7] 결과물 저장
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, count_match

# ==========================================
# 🖥️ UI 및 파일 업로드 처리
# ==========================================
st.subheader("1. 파일 업로드")
uploaded_db = st.file_uploader("메인 급여DB 파일 (급여DB.xlsx 또는 .xls) 업로드", type=["xlsx", "xls"])
uploaded_ots = st.file_uploader("팀별 OT 파일들 업로드 (여러 개 선택 가능, 파일명 무관)", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_db and uploaded_ots:
    st.success(f"✅ DB 파일 1개와 OT 파일 {len(uploaded_ots)}개가 업로드 되었습니다.")
    
    if st.button("🚀 데이터 통합 실행하기"):
        with st.spinner("데이터를 처리하는 중입니다. 잠시만 기다려주세요..."):
            try:
                processed_file, match_count = process_salary_master(uploaded_db, uploaded_ots)
                
                st.success(f"🎉 작업 완료! 총 {match_count}명의 OT 금액을 성공적으로 매칭했습니다.")
                
                st.download_button(
                    label="📥 완료된 급여DB 파일 다운로드",
                    data=processed_file,
                    file_name="급여DB_최종완료.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ 처리 중 예상치 못한 오류가 발생했습니다: {e}")
