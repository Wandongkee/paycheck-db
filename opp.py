import os
import streamlit as st
import pandas as pd
import openpyxl
import io

# ==========================================
# ⚙️ 스크립트 실행 경로 자동 설정 (이전 수정사항 유지)
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
동명이인 방지를 위해 **각 본부/팀별로 분리해서 업로드** 해주세요.
""")

# ==========================================
# 🛠️ 기능 함수 (유지)
# ==========================================
def convert_xls_to_xlsx_buffer(uploaded_file):
    df = pd.read_excel(uploaded_file, engine='xlrd')
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    return buffer

def load_ot_data_from_uploaded_file(uploaded_file):
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
                        # 주의: 동일 파일(팀) 내 동명이인이 있으면 덮어씌워짐
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
    max_row = ws.max_row
    col_data = [ws.cell(row=r, column=col_from).value for r in range(1, max_row + 1)]
    ws.insert_cols(col_to)
    for i, val in enumerate(col_data):
        ws.cell(row=i+1, column=col_to).value = val
    del_target = col_from + 1 if col_to <= col_from else col_from
    ws.delete_cols(del_target)

# ==========================================
# 🚀 메인 데이터 처리 로직 (수정됨)
# ==========================================
def process_salary_master(db_file, ot_files_op1, ot_files_op2, ot_files_op):
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

    # [Step 2 & 3] 열 이동 및 삽입
    move_column(ws, 38, 18) # AL -> R
    move_column(ws, 26, 18) # Z -> R
    move_column(ws, 30, 17) # AD -> Q
    ws.insert_cols(21, 3) 

    # [Step 4] U열 합산
    for row in range(2, ws.max_row + 1):
        k_val = ws.cell(row=row, column=11).value
        if k_val == "양중(T/C)":
            total = sum([
                float(ws.cell(row=row, column=c).value or 0) 
                for c in range(17, 21)
            ])
            ws.cell(row=row, column=21).value = total

    # [Step 5] 그룹별 OT 데이터 각각 로드 (3개의 딕셔너리로 분리)
    db_op1, db_op2, db_op_gen = {}, {}, {}
    
    if ot_files_op1:
        for f in ot_files_op1: db_op1.update(load_ot_data_from_uploaded_file(f))
    if ot_files_op2:
        for f in ot_files_op2: db_op2.update(load_ot_data_from_uploaded_file(f))
    if ot_files_op:
        for f in ot_files_op: db_op_gen.update(load_ot_data_from_uploaded_file(f))

    # [Step 6] 데이터 매칭
    count_match = 0
    ws.insert_cols(48, 1) # AV열 생성

    for row in range(2, ws.max_row + 1):
        # ⚠️ 중요: 메인 DB에서 소속 부서를 나타내는 열 번호를 지정해야 합니다.
        # 아래는 B열(column=2)에 부서명이 있다고 가정한 예시입니다. 실제 엑셀에 맞게 수정하세요.
        dept_val = str(ws.cell(row=row, column=2).value or "").strip() 
        name = ws.cell(row=row, column=4).value   
        k_val = ws.cell(row=row, column=11).value 
        u_val = ws.cell(row=row, column=21).value 
        
        if k_val == "양중(T/C)":
            found_ot = None
            
            # 부서명에 따라 참조할 OT 딕셔너리 결정
            if "운영1" in dept_val:
                found_ot = db_op1.get(name)
            elif "운영2" in dept_val:
                found_ot = db_op2.get(name)
            elif "운영" in dept_val: # 운영1, 2가 아닌 일반 운영팀
                found_ot = db_op_gen.get(name)
            else:
                # 부서명이 명확하지 않을 경우 최후의 수단으로 전체 검색
                found_ot = db_op1.get(name) or db_op2.get(name) or db_op_gen.get(name)

            if found_ot is not None:
                ws.cell(row=row, column=22).value = found_ot # V열
                count_match += 1

        if u_val not in [None, ""]:
            ws.cell(row=row, column=23).value = f"=U{row}=V{row}" 
            
        ws.cell(row=row, column=48).value = f"=AU{row}-X{row}-U{row}-AS{row}" 

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, count_match

# ==========================================
# 🖥️ UI 및 파일 업로드 처리 (3분할)
# ==========================================
st.subheader("1. 메인 급여 DB 업로드")
uploaded_db = st.file_uploader("메인 급여DB 파일 (급여DB.xlsx 또는 .xls)", type=["xlsx", "xls"])

st.subheader("2. 본부/팀별 OT 파일 업로드")
col1, col2, col3 = st.columns(3)
with col1:
    up_op1 = st.file_uploader("🏢 운영1본부/운영1팀", type=["xlsx", "xls"], accept_multiple_files=True)
with col2:
    up_op2 = st.file_uploader("🏢 운영2본부/운영2팀", type=["xlsx", "xls"], accept_multiple_files=True)
with col3:
    up_op = st.file_uploader("🏢 운영팀", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_db and (up_op1 or up_op2 or up_op):
    st.success("✅ 파일 업로드 완료 (OT 파일이 최소 1개 이상 업로드되었습니다.)")
    
    if st.button("🚀 데이터 통합 실행하기"):
        with st.spinner("데이터를 처리하는 중입니다. 잠시만 기다려주세요..."):
            try:
                processed_file, match_count = process_salary_master(uploaded_db, up_op1, up_op2, up_op)
                
                st.success(f"🎉 작업 완료! 총 {match_count}명의 OT 금액을 성공적으로 매칭했습니다.")
                
                st.download_button(
                    label="📥 완료된 급여DB 파일 다운로드",
                    data=processed_file,
                    file_name="급여DB_최종완료.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ 처리 중 예상치 못한 오류가 발생했습니다: {e}")
