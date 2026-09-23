import hashlib
import io
import zipfile
from pathlib import Path
import streamlit as st
from excel_engine import (SALARY_ROLES, detect_header, candidates, headers,
                          integrate, read_book, to_xlsx, statement, vlookup_values)
from mapping_ui import mapping_panel

st.set_page_config(page_title='급여DB 자동 통합 툴', page_icon='💰', layout='centered')
st.title('💰 급여DB 자동 통합 툴')
st.markdown('ERP에서 받은 **.xls 파일을 그대로 업로드**할 수 있습니다. 열 제목을 확인한 뒤 데이터를 통합합니다.')

def prepared(upload, key):
    raw = upload.getvalue()
    digest = hashlib.sha256(raw).hexdigest()
    cached = st.session_state.get(key)
    if cached is None or cached[0] != digest:
        with st.spinner(f'{upload.name} 파일을 준비하는 중입니다...'):
            data = to_xlsx(raw, upload.name)
        st.session_state[key] = (digest, data)
    return st.session_state[key][1], digest

def selected_sources(upload, group, key, roles):
    data, digest = prepared(upload, key + ':converted')
    wb = read_book(data, True)
    eligible = [ws.title for ws in wb if candidates(headers(ws, detect_header(ws)), 'name')]
    defaults = [wb.active.title] if wb.active.title in eligible else eligible if len(eligible) == 1 else []
    st.caption('파일에 저장된 활성 시트를 기본 선택합니다. 회사·정산월을 확인하고, 과거 자료나 복사본을 함께 선택하지 마세요.')
    sheets = st.multiselect(f'{upload.name} — 처리할 시트', wb.sheetnames,
                            default=defaults, key=key + digest + ':sheets')
    sources, ready = [], bool(sheets)
    for sheet in sheets:
        with st.expander(f'{upload.name} / {sheet} 항목 확인', expanded=True):
            header, mapping, valid = mapping_panel(wb[sheet], roles, key + digest + ':' + sheet)
            sources.append((data, sheet, header, mapping, group, upload.name))
            ready = ready and valid
    if not sheets:
        st.info('처리할 시트를 선택해주세요. 표지·집계 시트는 제외하세요.')
    return sources, ready

st.subheader('1. 메인 급여DB 업로드 및 자동 변환')
db = st.file_uploader('메인 급여DB (.xlsx 또는 .xls)', type=['xlsx', 'xls'])
db_data, db_config, db_ready = None, None, False
if db:
    try:
        db_data, digest = prepared(db, 'master:converted')
        st.download_button('📥 XLSX 원본 변환본 다운로드', db_data,
            file_name=Path(db.name).stem + '_변환.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        st.caption('통합 처리 전 변환본입니다. XLS의 모든 시트를 변환하며 특수 서식은 Excel과 차이가 있을 수 있습니다.')
        wb = read_book(db_data, True)
        sheet = st.selectbox('급여 데이터 시트', wb.sheetnames, key='master:sheet:' + digest)
        with st.expander('급여DB 항목 확인', expanded=True):
            st.write('이름·입사일자·수당은 제목으로 자동 인식합니다. 인식하지 못한 항목은 실제 제목을 보고 선택해주세요.')
            st.caption('차액 = 기준금액 − 차감금액 1 − 수당 4개 합산 − 차감금액 2')
            header, mapping, db_ready = mapping_panel(wb[sheet], SALARY_ROLES, 'master:' + digest + ':' + sheet)
            db_config = (sheet, header, mapping)
    except Exception as exc:
        st.error(str(exc))

st.subheader('2. 본부·팀별 OT 통합')
sources, ot_ready, uploads_present = [], True, False
for i, group in enumerate(('운영1', '운영2', '운영')):
    uploads = st.file_uploader(group + ' OT 파일', type=['xlsx', 'xls'],
                               accept_multiple_files=True, key='ot:' + group)
    for j, upload in enumerate(uploads):
        uploads_present = True
        try:
            entries, valid = selected_sources(upload, group, f'ot:{i}:{j}:', ('name', 'hire', 'amount'))
            sources.extend(entries)
            ot_ready = ot_ready and valid
        except Exception as exc:
            ot_ready = False
            st.error(f'{upload.name}: {exc}')
st.info('통합 결과는 기존 U·V·W·AV열 위치를 유지합니다. 원본 시트도 함께 보관하며 「대조내역」에서 미매칭·금액 불일치를 확인할 수 있습니다.')
request = None
if db_data is not None and db_config and sources and db_ready and ot_ready:
    request = hashlib.sha256(repr((hashlib.sha256(db_data).hexdigest(), db_config,
        [(hashlib.sha256(x[0]).hexdigest(), x[1:]) for x in sources])).encode()).hexdigest()
if st.button('🚀 데이터 통합 실행하기', disabled=not(db_ready and ot_ready and uploads_present and sources)):
    st.session_state.pop('salary_result', None)
    try:
        with st.spinner('급여DB와 OT를 비교하는 중입니다...'):
            result = integrate(db_data, *db_config, sources)
        st.session_state['salary_result'] = (request, result)
    except Exception as exc:
        st.error(str(exc))
saved = st.session_state.get('salary_result')
if saved and saved[0] == request:
    output, count, rows = saved[1]
    st.success(f'OT 금액 매칭 {count}건. 확인사항을 검토한 뒤 다운로드하세요.')
    st.dataframe([dict(zip(rows[0], row)) for row in rows[1:]], hide_index=True)
    st.download_button('📥 급여DB 통합 결과 다운로드', output, file_name='급여DB_최종완료.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.divider()
st.subheader('3. 급여명세서 작업')
st.write('OT 항목을 제목으로 선택하면 기존 BA~BE열에 수당 문구를 추가합니다. 해당 위치에 기존 값이 있으면 덮어쓰지 않고 안내합니다.')
text_uploads = st.file_uploader('문구를 추가할 OT 파일', type=['xlsx', 'xls'], accept_multiple_files=True, key='text:uploads')
text_entries, text_ready = [], bool(text_uploads)
for i, upload in enumerate(text_uploads):
    try:
        entries, ready = selected_sources(upload, '', f'text:{i}:',
            ('name', 'early', 'extension', 'night', 'holiday_days', 'holiday_hours'))
        text_entries.append((upload.name, entries))
        text_ready = text_ready and ready
    except Exception as exc:
        text_ready = False
        st.error(str(exc))
text_request = hashlib.sha256(repr([(name, [(hashlib.sha256(e[0]).hexdigest(), e[1:]) for e in entries]) for name, entries in text_entries]).encode()).hexdigest() if text_ready else None
if st.button('🪄 OT 급여명세서 작업실행', disabled=not text_ready):
    st.session_state.pop('text_result', None)
    try:
        outputs = [(f'{i + 1}_텍스트추가_{Path(name).stem}.xlsx', statement(entries))
                   for i, (name, entries) in enumerate(text_entries)]
        if len(outputs) == 1:
            filename, output = outputs[0]
        else:
            archive = io.BytesIO()
            with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as zipped:
                for name, data in outputs:
                    zipped.writestr(name, data)
            filename, output = 'OT_급여명세서_결과.zip', archive.getvalue()
        st.session_state['text_result'] = (text_request, filename, output)
    except Exception as exc:
        st.error(str(exc))
saved = st.session_state.get('text_result')
if saved and saved[0] == text_request:
    st.download_button('📥 급여명세서 작업 결과 다운로드', saved[2], file_name=saved[1])

st.divider()
st.subheader('4. VLOOKUP 수식만 선택적 값 변환')
st.write('Excel에서 재계산 후 저장한 파일을 올려주세요. 저장된 결과가 없는 VLOOKUP 셀이 있으면 변환을 중단합니다.')
formula_file = st.file_uploader('VLOOKUP을 제거할 파일', type=['xlsx'], key='formula:upload')
formula_key = hashlib.sha256(formula_file.getvalue()).hexdigest() if formula_file else None
if st.button('🪄 VLOOKUP만 값으로 변환하기', disabled=formula_file is None):
    st.session_state.pop('formula_result', None)
    try:
        output = vlookup_values(formula_file.getvalue())
        st.session_state['formula_result'] = (formula_key, output)
    except Exception as exc:
        st.error(str(exc))
saved = st.session_state.get('formula_result')
if formula_file and saved and saved[0] == formula_key:
    st.download_button('📥 VLOOKUP 처리 완료 파일 다운로드', saved[1], file_name='VLOOKUP제거_' + formula_file.name)
