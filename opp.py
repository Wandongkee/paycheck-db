import os
import streamlit as st
import pandas as pd
import openpyxl
import io
import datetime
import re

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
* 동명이인 방지를 위해 **각 본부/팀별로 분리해서 업로드** 해주세요.
* 이름과 **입사일자**를 동시 비교하여 동명이인을 정확히 구분합니다.
* 메인 급여DB 파일은 **.xlsx**형식을 권장합니다.
""")

# ==========================================
# 🛠️ 기능 함수
# ==========================================
def convert_xls_to_xlsx_buffer(uploaded_file):
    df = pd.read_excel(uploaded_file, engine='xlrd')
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    return buffer

def clean_date_string(date_val):
    """엑셀의 다양한 날짜 형식을 순수 숫자(YYYYMMDD) 문자열로 변환"""
    if date_val is None or pd.isna(date_val):
        return ""
    
    # 1. 이미 날짜 객체(datetime)로 읽힌 경우
    if isinstance(date_val, (pd.Timestamp, datetime.datetime, datetime.date)):
        return date_val.strftime("%Y%m%d")
    
    # 2. 문자열인 경우 (예: "2020-01-01", "2020/01/01")
    date_str = str(date_val).split(" ")[0] # 시간 데이터가 붙어있을 경우 날짜만 분리
    clean_str = re.sub(r'[^0-9]', '', date_str) # 숫자 이외의 문자(-, / 등) 모두 제거
    return clean_str

def load_ot_data_from_uploaded_file(uploaded_file):
    """OT 파일에서 {(이름, 입사일자): 금액} 형태로 데이터를 가져옴"""
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
                    hire_date_raw = row[6] # G열 (인덱스 6)
                    amount = row[19]       # T열 (인덱스 19)
                    
                    try:
                        amount = float(amount)
                        hire_date = clean_date_string(hire_date_raw)
                        
                        # 이름과 입사일자를 조합한 고유 키 생성
                        unique_key = f"{name}_{hire_date}"
                        combined_data[unique_key] = amount
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
# 🚀 메인 데이터 처리 로직
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

    # [Step 5] 그룹별 OT 데이터 각각 로드 
    db_op1, db_op2, db_op_gen = {}, {}, {}
    
    if ot_files_op1:
        for f in ot_files_op1: db_op1.update(load_ot_data_from_uploaded_file(f))
    if ot_files_op2:
        for f in ot_files_op2: db_op2.update(load_ot_data_from_uploaded_file(f))
    if ot_files_op:
        for f in ot_files_op: db_op_gen.update(load_ot_data_from_uploaded_file(f))

    # [Step 6] 데이터 매칭 (이름 + 입사일자)
    count_match = 0
    ws.insert_cols(48, 1) # AV열 생성

    for row in range(2, ws.max_row + 1):
        dept_val = str(ws.cell(row=row, column=2).value or "").strip() # B열(부서)
        name = ws.cell(row=row, column=4).value                        # D열(이름)
        hire_date_raw = ws.cell(row=row, column=6).value               # F열(입사일자)
        k_val = ws.cell(row=row, column=11).value                      # K열(직종)
        u_val = ws.cell(row=row, column=21).value                      # U열(합산금액)
        
        if k_val == "양중(T/C)":
            hire_date = clean_date_string(hire_date_raw)
            search_key = f"{name}_{hire_date}"
            found_ot = None
            
            # 부서명에 따라 참조할 OT 딕셔너리 결정
            if "운영1" in dept_val:
                found_ot = db_op1.get(search_key)
            elif "운영2" in dept_val:
                found_ot = db_op2.get(search_key)
            elif "운영" in dept_val:
                found_ot = db_op_gen.get(search_key)
            else:
                found_ot = db_op1.get(search_key) or db_op2.get(search_key) or db_op_gen.get(search_key)

            if found_ot is not None:
                ws.cell(row=row, column=22).value = found_ot # V열에 입력
                count_match += 1

        if u_val not in [None, ""]:
            ws.cell(row=row, column=23).value = f"=U{row}=V{row}" 
            
        ws.cell(row=row, column=48).value = f"=AU{row}-X{row}-U{row}-AS{row}" 

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, count_match

# ==========================================
# 🖥️ UI 및 파일 업로드 처리 
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

# ==========================================
# 🛠️ 추가 기능: OT 데이터 요약 리스트 생성 함수
# ==========================================
def generate_ot_summary_excel(op1_files, op2_files, op_files):
    data = []
    
    def process_files(files, dept_name):
        if not files: return
        for f in files:
            if f.name.lower().endswith('.xls'):
                file_to_read = convert_xls_to_xlsx_buffer(f)
            else:
                file_to_read = f
                
            xls = pd.ExcelFile(file_to_read)
            for sheet in xls.sheet_names:
                try:
                    df = pd.read_excel(file_to_read, sheet_name=sheet, header=None)
                    if len(df.columns) < 20: continue
                    df = df.iloc[7:]
                    valid_rows = df[df[4].notna()]
                    
                    for _, row in valid_rows.iterrows():
                        name = str(row[4]).strip()
                        hire_date = clean_date_string(row[6]) # G열 입사일자 (동명이인 구분용)
                        
                        # 안전한 숫자 변환 함수 (소수점 떨어지면 정수로)
                        def safe_val(val):
                            try:
                                v = float(val)
                                return int(v) if v.is_integer() else v
                            except:
                                return 0
                                
                        val_j = safe_val(row[9])  # J열: 조출점심저녁
                        val_l = safe_val(row[11]) # L열: 연장OT
                        val_n = safe_val(row[13]) # N열: 야간OT
                        val_p = safe_val(row[15]) # P열: 휴일근무
                        val_r = safe_val(row[17]) # R열: 휴일OT
                        
                        # 0시간인 항목도 출력하도록 포맷팅 (제안하신 양식 반영)
                        data.append({
                            "부서": dept_name,
                            "성명": name,
                            "입사일자": hire_date,
                            "연장OT": f"연장OT:{val_l}H",
                            "야간OT": f"야간OT:{val_n}H",
                            "휴일근무": f"휴일근무:{val_p}D",
                            "휴일OT": f"휴일OT:{val_r}H",
                            "조출점심저녁": f"조출점심저녁:{val_j}H"
                        })
                except Exception:
                    continue
                    
    # 업로드 창별로 부서명을 다르게 매핑하여 처리
    process_files(op1_files, "운영1본부")
    process_files(op2_files, "운영2본부")
    process_files(op_files, "운영팀")
    
    # 리스트를 데이터프레임으로 변환 후 엑셀로 저장
    df_result = pd.DataFrame(data)
    output = io.BytesIO()
    df_result.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    
    return output, len(df_result)

# ------------------------------------------
# UI 적용 부분 (기존 2번 섹션 파일 업로드 코드 아래에 배치)
# ------------------------------------------
if up_op1 or up_op2 or up_op:
    st.info("💡 팁: 마스터 DB에 통합하기 전에, 업로드한 OT 파일들의 내역만 요약해서 엑셀로 뽑아볼 수 있습니다.")
    if st.button("📋 팀별 OT 요약 리스트 추출하기"):
        with st.spinner("OT 데이터를 요약하고 있습니다..."):
            try:
                summary_file, total_count = generate_ot_summary_excel(up_op1, up_op2, up_op)
                st.success(f"✅ 총 {total_count}명의 OT 요약 리스트가 생성되었습니다.")
                st.download_button(
                    label="📥 OT 요약 리스트 엑셀 다운로드",
                    data=summary_file,
                    file_name="팀별_OT요약_리스트.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_summary"
                )
            except Exception as e:
                st.error(f"❌ 요약 리스트 생성 중 오류 발생: {e}")

# ==========================================
# 🛠️ 3. 특정 수식(VLOOKUP)만 선택적 값 복사 기능
# ==========================================
def convert_only_vlookup_to_values(uploaded_file):
    """
    엑셀 파일에서 VLOOKUP 수식만 찾아 결과값으로 덮어쓰고,
    나머지 수식(SUM, IF 등)은 그대로 유지합니다.
    """
    # 1. 수식을 확인하기 위해 원본 그대로 워크북 로드
    wb_formula = openpyxl.load_workbook(uploaded_file, data_only=False)
    
    # 2. 결과값을 추출하기 위해 data_only=True 속성으로 워크북 한 번 더 로드
    # (Streamlit 파일 버퍼의 읽기 위치를 처음으로 초기화)
    uploaded_file.seek(0) 
    wb_value = openpyxl.load_workbook(uploaded_file, data_only=True)
    
    # 모든 시트를 순회하며 VLOOKUP 찾기
    for sheet_name in wb_formula.sheetnames:
        ws_f = wb_formula[sheet_name]
        ws_v = wb_value[sheet_name]
        
        for row in ws_f.iter_rows():
            for cell_f in row:
                # 셀에 값이 있고, 문자열이며, '='로 시작하면 수식으로 간주
                if isinstance(cell_f.value, str) and cell_f.value.startswith('='):
                    # 해당 수식 안에 'VLOOKUP'이 포함되어 있는지 확인 (대소문자 구분 없이)
                    if 'VLOOKUP' in cell_f.value.upper():
                        # 값 전용 워크북에서 동일한 좌표의 결과값을 가져와서 덮어쓰기
                        cell_v = ws_v[cell_f.coordinate]
                        cell_f.value = cell_v.value
                        
    output = io.BytesIO()
    wb_formula.save(output)
    output.seek(0)
    return output

st.divider()

st.subheader("3. VLOOKUP 수식만 선택적 값 변환")
st.markdown("""
파일 내의 다른 수식(SUM, IF 등)은 그대로 살려두고, **VLOOKUP 수식이 있는 셀만 찾아 화면에 보이는 '결과값'으로 덮어씌웁니다.**
* ⚠️ **주의사항**: 파이썬은 엑셀의 캐시를 읽어오므로, 업로드 전 엑셀 프로그램에서 파일을 **반드시 한 번 저장(Ctrl+S)**한 후 올려주세요.
""")

uploaded_formula_file = st.file_uploader("VLOOKUP을 제거할 엑셀 파일 (.xlsx) 업로드", type=["xlsx"], key="formula_uploader")

if uploaded_formula_file:
    if st.button("🪄 VLOOKUP만 값으로 변환하기"):
        with st.spinner("파일을 스캔하여 VLOOKUP 수식을 처리하는 중입니다..."):
            try:
                value_only_file = convert_only_vlookup_to_values(uploaded_formula_file)
                
                st.success("✅ VLOOKUP 수식 변환 완료! 나머지 수식은 안전하게 유지되었습니다.")
                
                st.download_button(
                    label="📥 VLOOKUP 처리 완료 파일 다운로드",
                    data=value_only_file,
                    file_name=f"VLOOKUP제거_{uploaded_formula_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="formula_download"
                )
            except Exception as e:
                st.error(f"❌ 처리 중 오류 발생: {e}")
