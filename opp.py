import os
import streamlit as st
import pandas as pd
import openpyxl
import io
import datetime
import re
import zipfile
from openpyxl.utils import get_column_letter

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
