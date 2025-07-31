import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(layout="wide")

st.title("📋 작업목록표 입력 시스템")

# 헬퍼 함수: 문자열에서 숫자 추출 (단위 제거)
def parse_value(value_str, default_val=0, val_type=float):
    if pd.isna(value_str) or str(value_str).strip() == "":
        return default_val
    try:
        cleaned_value = str(value_str).replace("시간", "").replace("분", "").replace("kg", "").replace("회", "").replace("일", "").replace("/", "").replace("초", "").strip()
        return val_type(cleaned_value)
    except ValueError:
        return default_val

# --- 세션 상태 초기화 및 기본값 설정 ---
def initialize_session_state():
    if 'task_units' not in st.session_state:
        st.session_state.task_units = []
    if 'unit_count' not in st.session_state:
        st.session_state.unit_count = 0
    if 'group_name' not in st.session_state:
        st.session_state.group_name = ""
    if '소속' not in st.session_state:
        st.session_state.소속 = ""
    if '반' not in st.session_state:
        st.session_state.반 = ""
    if 'file_processed' not in st.session_state:
        st.session_state.file_processed = False

def create_default_unit():
    return {
        "회사명": st.session_state.group_name, 
        "소속": st.session_state.소속, 
        "반": st.session_state.반,
        "단위작업명": "", 
        "작업내용(상세설명)": "",
        "작업자 수": 1, 
        "작업자 이름": "",
        "작업형태": "주간", 
        "1일 작업시간": 0,
        "자세": {}, 
        "중량물": [], 
        "도구": [],
        "유해요인_원인분석": [{"유형": "", "부담작업": "", "부담작업자세": ""}],
        "보호구": [], 
        "작성자": "", 
        "연락처": ""
    }

# 세션 상태 초기화
initialize_session_state()

# 앱 시작 시 또는 파일 로드 후, 최소 1개의 단위작업공정이 있도록 보장
if st.session_state.unit_count == 0 and not st.session_state.task_units:
    st.session_state.unit_count = 1
    st.session_state.task_units.append(create_default_unit())

# 엑셀 파일 업로드 섹션
st.sidebar.header("📊 데이터 불러오기/내보내기")
uploaded_file = st.sidebar.file_uploader("엑셀 파일 업로드 (재시작/수정)", type=["xlsx"], key="file_uploader")

# 파일 업로드 처리 (호환성 강화 + 데이터 보존)
if uploaded_file is not None and not st.session_state.file_processed:
    try:
        # 다양한 인코딩과 옵션으로 파일 읽기 시도
        df_uploaded = None
        try:
            df_uploaded = pd.read_excel(uploaded_file, sheet_name='작업목록', dtype=str, na_filter=False)
        except:
            try:
                df_uploaded = pd.read_excel(uploaded_file, sheet_name='작업목록', engine='openpyxl', dtype=str, na_filter=False)
            except:
                df_uploaded = pd.read_excel(uploaded_file, dtype=str, na_filter=False)
        
        if df_uploaded is None or df_uploaded.empty:
            raise ValueError("파일을 읽을 수 없거나 비어있습니다.")
        
        # 업로드된 파일의 컬럼들을 확인하고 안전하게 처리
        st.sidebar.info(f"📄 파일 정보: {len(df_uploaded)} 행, {len(df_uploaded.columns)} 컬럼")
        
        # 디버깅용: 주요 컬럼들 확인
        important_cols = ['회사명', '소속', '반', '단위작업명', '작업내용(상세설명)']
        available_cols = [col for col in important_cols if col in df_uploaded.columns]
        st.sidebar.info(f"🔍 감지된 주요 컬럼: {', '.join(available_cols) if available_cols else '없음'}")
        
        # 기존 데이터 초기화
        st.session_state.task_units = []
        st.session_state.unit_count = 0

        loaded_task_units = []
        successful_loads = 0
        
        for index, row in df_uploaded.iterrows():
            try:
                # 안전한 값 추출 함수
                def safe_get(column_name, default=""):
                    value = row.get(column_name, default)
                    if pd.isna(value) or value is None:
                        return default
                    return str(value).strip()
                
                def safe_get_int(column_name, default=0):
                    value = row.get(column_name, default)
                    if pd.isna(value) or value is None or str(value).strip() == "":
                        return default
                    try:
                        return int(float(str(value)))
                    except:
                        return default
                
                def safe_get_float(column_name, default=0.0):
                    value = row.get(column_name, default)
                    if pd.isna(value) or value is None or str(value).strip() == "":
                        return default
                    try:
                        return float(str(value))
                    except:
                        return default

                unit = {
                    "회사명": safe_get("회사명"),
                    "소속": safe_get("소속"),
                    "반": safe_get("반"),
                    "단위작업명": safe_get("단위작업명"),
                    "작업내용(상세설명)": safe_get("작업내용(상세설명)"),
                    "작업자 수": safe_get_int("작업자 수", 1),
                    "작업자 이름": safe_get("작업자 이름"),
                    "작업형태": safe_get("작업형태", "주간"),
                    "1일 작업시간": safe_get_int("1일 작업시간", 0),
                    "자세": {},
                    "중량물": [],
                    "도구": [],
                    "유해요인_원인분석": [],
                    "보호구": [],
                    "작성자": safe_get("작성자"),
                    "연락처": safe_get("연락처")
                }
                
                # 보호구 처리
                protection_gear_str = safe_get("보호구")
                if protection_gear_str:
                    unit["보호구"] = [item.strip() for item in protection_gear_str.split(",") if item.strip()]

                # 부담작업 데이터 로드 (안전하게)
                for k_crit in range(1, 13):
                    unit[f"부담작업_{k_crit}호"] = safe_get(f"부담작업_{k_crit}호", "X")

                # 유해요인 원인분석 데이터 로드 (호환성 강화)
                FIXED_MAX_HAZARD_ANALYTICS_FOR_PARSE = 5
                for j_hazard in range(FIXED_MAX_HAZARD_ANALYTICS_FOR_PARSE):
                    hazard_type = safe_get(f"유해요인_원인분석_유형_{j_hazard+1}")
                    if hazard_type and hazard_type != "":
                        hazard_entry = {"유형": hazard_type}
                        
                        if hazard_type == "반복동작":
                            hazard_entry["부담작업"] = safe_get(f"유해요인_원인분석_부담작업_{j_hazard+1}_반복")
                            hazard_entry["수공구 종류"] = safe_get(f"유해요인_원인분석_수공구_종류_{j_hazard+1}")
                            hazard_entry["수공구 용도"] = safe_get(f"유해요인_원인분석_수공구_용도_{j_hazard+1}")
                            hazard_entry["수공구 무게(kg)"] = safe_get_float(f"유해요인_원인분석_수공구_무게(kg)_{j_hazard+1}")
                            hazard_entry["수공구 사용시간(분)"] = safe_get(f"유해요인_원인분석_수공구_사용시간(분)_{j_hazard+1}")
                            hazard_entry["부담부위"] = safe_get(f"유해요인_원인분석_부담부위_{j_hazard+1}")
                            hazard_entry["회당 반복시간(초/회)"] = safe_get(f"유해요인_원인분석_반복_회당시간(초/회)_{j_hazard+1}")
                            hazard_entry["작업시간동안 반복횟수(회/일)"] = safe_get(f"유해요인_원인분석_반복_총횟수(회/일)_{j_hazard+1}")
                            hazard_entry["총 작업시간(분)"] = safe_get(f"유해요인_원인분석_반복_총시간(분)_{j_hazard+1}")
                            hazard_entry["물체 무게(kg)_10호"] = safe_get_float(f"유해요인_원인분석_반복_물체무게_10호(kg)_{j_hazard+1}")
                            hazard_entry["분당 반복횟수(회/분)_10호"] = safe_get(f"유해요인_원인분석_반복_분당반복횟수_10호(회/분)_{j_hazard+1}")
                            hazard_entry["작업내용_12호_정적"] = safe_get(f"유해요인_원인분석_반복_작업내용_12호_정적_{j_hazard+1}")
                            hazard_entry["작업시간(분)_12호_정적"] = safe_get(f"유해요인_원인분석_반복_작업시간_12호_정적_{j_hazard+1}")
                            hazard_entry["휴식시간(분)_12호_정적"] = safe_get(f"유해요인_원인분석_반복_휴식시간_12호_정적_{j_hazard+1}")
                            hazard_entry["인체부담부위_12호_정적"] = safe_get(f"유해요인_원인분석_반복_인체부담부위_12호_정적_{j_hazard+1}")

                        elif hazard_type == "부자연스러운 자세":
                            hazard_entry["부담작업자세"] = safe_get(f"유해요인_원인분석_부담작업자세_{j_hazard+1}")
                            hazard_entry["회당 반복시간(초/회)"] = safe_get(f"유해요인_원인분석_자세_회당시간(초/회)_{j_hazard+1}")
                            hazard_entry["작업시간동안 반복횟수(회/일)"] = safe_get(f"유해요인_원인분석_자세_총횟수(회/일)_{j_hazard+1}")
                            hazard_entry["총 작업시간(분)"] = safe_get(f"유해요인_원인분석_자세_총시간(분)_{j_hazard+1}")
                            
                        elif hazard_type == "과도한 힘":
                            hazard_entry["부담작업"] = safe_get(f"유해요인_원인분석_부담작업_{j_hazard+1}_힘")
                            hazard_entry["중량물 명칭"] = safe_get(f"유해요인_원인분석_힘_중량물_명칭_{j_hazard+1}")
                            hazard_entry["중량물 용도"] = safe_get(f"유해요인_원인분석_힘_중량물_용도_{j_hazard+1}")
                            
                            # 🔧 새 필드들 - 안전하게 로드 (없으면 기본값)
                            hazard_entry["중량물 무게(kg)"] = safe_get_float(f"유해요인_원인분석_중량물_무게(kg)_{j_hazard+1}")
                            hazard_entry["하루 8시간동안 중량물을 드는 횟수(회)"] = safe_get_int(f"유해요인_원인분석_하루8시간_중량물_횟수(회)_{j_hazard+1}")
                            
                            hazard_entry["취급방법"] = safe_get(f"유해요인_원인분석_힘_취급방법_{j_hazard+1}")
                            hazard_entry["중량물 이동방법"] = safe_get(f"유해요인_원인분석_힘_이동방법_{j_hazard+1}")
                            hazard_entry["작업자가 직접 밀고/당기기"] = safe_get(f"유해요인_원인분석_힘_직접_밀당_{j_hazard+1}")
                            hazard_entry["기타_밀당_설명"] = safe_get(f"유해요인_원인분석_힘_기타_밀당_설명_{j_hazard+1}")
                            hazard_entry["작업시간동안 작업횟수(회/일)"] = safe_get(f"유해요인_원인분석_힘_총횟수(회/일)_{j_hazard+1}")
                            
                        elif hazard_type == "접촉스트레스 또는 기타(진동, 밀고 당기기 등)":
                            hazard_entry["부담작업"] = safe_get(f"유해요인_원인분석_부담작업_{j_hazard+1}_기타")
                            
                            # 11호 관련 필드들
                            if "(11호)" in str(hazard_entry["부담작업"]):
                                hazard_entry["작업시간(분)"] = safe_get(f"유해요인_원인분석_기타_작업시간(분)_{j_hazard+1}")
                            
                            # 12호 진동 관련 필드들
                            elif "(12호)" in str(hazard_entry["부담작업"]):
                                hazard_entry["진동수공구명"] = safe_get(f"유해요인_원인분석_기타_진동수공구명_{j_hazard+1}")
                                hazard_entry["진동수공구 용도"] = safe_get(f"유해요인_원인분석_기타_진동수공구_용도_{j_hazard+1}")
                                hazard_entry["작업시간(분)_진동"] = safe_get(f"유해요인_원인분석_기타_작업시간_진동_{j_hazard+1}")
                                hazard_entry["작업빈도(초/회)_진동"] = safe_get(f"유해요인_원인분석_기타_작업빈도_진동_{j_hazard+1}")
                                hazard_entry["작업량(회/일)_진동"] = safe_get(f"유해요인_원인분석_기타_작업량_진동_{j_hazard+1}")
                                hazard_entry["수공구사용시 지지대가 있는가?"] = safe_get(f"유해요인_원인분석_기타_지지대_여부_{j_hazard+1}")
                        
                        unit["유해요인_원인분석"].append(hazard_entry)
                
                # 로드된 데이터에 원인분석 항목이 없으면 기본 1개 추가
                if not unit["유해요인_원인분석"]:
                    unit["유해요인_원인분석"].append({"유형": "", "부담작업": "", "부담작업자세": ""})

                loaded_task_units.append(unit)
                successful_loads += 1
                
            except Exception as row_error:
                st.sidebar.warning(f"⚠️ 행 {index+1} 처리 중 오류: {str(row_error)[:50]}...")
                continue
        
        if loaded_task_units:
            # 회사 정보 업데이트
            st.session_state.group_name = loaded_task_units[0].get("회사명", "")
            st.session_state.소속 = loaded_task_units[0].get("소속", "")
            st.session_state.반 = loaded_task_units[0].get("반", "")
            
            # 작업 단위 데이터 업데이트
            st.session_state.task_units = loaded_task_units
            st.session_state.unit_count = len(loaded_task_units)
            st.session_state.file_processed = True
            
            st.sidebar.success(f"✅ {successful_loads}개의 작업이 성공적으로 로드되었습니다!")
            
            # 🔧 데이터 확인 메시지
            sample_unit = loaded_task_units[0] if loaded_task_units else {}
            if sample_unit.get("단위작업명"):
                st.sidebar.info(f"📝 첫 번째 작업: '{sample_unit.get('단위작업명')}'")
            
            st.rerun()
        else:
            st.sidebar.error("❌ 유효한 작업 데이터를 찾을 수 없습니다.")
            st.sidebar.info("💡 파일 형식을 확인하거나 새로 작성해주세요.")
            # 기본 데이터로 초기화
            st.session_state.unit_count = 1
            st.session_state.task_units = [create_default_unit()]

    except Exception as e:
        st.sidebar.error(f"⚠️ 파일 로드 중 오류 발생: {str(e)}")
        st.sidebar.info("💡 다음을 확인해주세요:")
        st.sidebar.info("- 파일이 엑셀 형식(.xlsx)인지")
        st.sidebar.info("- '작업목록' 시트가 있는지")
        st.sidebar.info("- 파일이 손상되지 않았는지")
        # 오류 발생 시 기본 데이터로 초기화
        st.session_state.task_units = [create_default_unit()]
        st.session_state.unit_count = 1

# 파일 처리 상태 리셋 버튼 (새 파일 업로드를 위해)
if st.session_state.file_processed:
    if st.sidebar.button("새 파일 업로드 준비"):
        st.session_state.file_processed = False
        st.rerun()

# 회사 정보 입력
new_group_name = st.text_input("회사명을 입력하세요", value=st.session_state.group_name, key="input_group_name")
new_소속 = st.text_input("소속/팀/그룹", value=st.session_state.소속, key="input_affiliation")
new_반 = st.text_input("반", value=st.session_state.반, key="input_class")

# 회사 정보가 변경되면 모든 단위작업에 반영
if (new_group_name != st.session_state.group_name or 
    new_소속 != st.session_state.소속 or 
    new_반 != st.session_state.반):
    
    st.session_state.group_name = new_group_name
    st.session_state.소속 = new_소속
    st.session_state.반 = new_반
    
    # 모든 기존 단위작업에 새 회사 정보 적용
    for unit in st.session_state.task_units:
        unit["회사명"] = st.session_state.group_name
        unit["소속"] = st.session_state.소속
        unit["반"] = st.session_state.반

# 단위작업 추가 버튼
col_unit_add_btn, _ = st.columns([0.2, 0.8])
with col_unit_add_btn:
    if st.button("단위작업 추가", key="add_unit_button"):
        st.session_state.unit_count += 1
        new_unit = create_default_unit()
        new_unit["회사명"] = st.session_state.group_name
        new_unit["소속"] = st.session_state.소속
        new_unit["반"] = st.session_state.반
        st.session_state.task_units.append(new_unit)
        st.rerun()

# 단위작업 입력 폼
for i in range(st.session_state.unit_count):
    # 새로운 단위작업이 추가되었을 때 빈 데이터 구조로 초기화
    if i >= len(st.session_state.task_units):
        new_unit = create_default_unit()
        new_unit["회사명"] = st.session_state.group_name
        new_unit["소속"] = st.session_state.소속
        new_unit["반"] = st.session_state.반
        st.session_state.task_units.append(new_unit)

    unit_data = st.session_state.task_units[i]

    with st.expander(f"단위작업공정 {i+1} 입력", expanded=True):
        # 단위작업 삭제 버튼 (첫 번째 작업은 삭제 불가)
        if i > 0:
            col_delete, _ = st.columns([0.2, 0.8])
            with col_delete:
                if st.button(f"작업 {i+1} 삭제", key=f"delete_unit_{i}"):
                    st.session_state.task_units.pop(i)
                    st.session_state.unit_count -= 1
                    st.rerun()
        
        # 기본 정보 입력 (회사 정보 자동 업데이트)
        unit_data["회사명"] = st.session_state.group_name
        unit_data["소속"] = st.session_state.소속
        unit_data["반"] = st.session_state.반
        
        unit_data["단위작업명"] = st.text_input(f"[{i+1}] 단위작업명", value=unit_data.get("단위작업명", ""), key=f"작업명_{i}")
        unit_data["작업내용(상세설명)"] = st.text_area(f"[{i+1}] 작업내용(상세설명)", value=unit_data.get("작업내용(상세설명)", ""), key=f"작업내용_{i}")
        unit_data["작업자 수"] = st.number_input(f"[{i+1}] 단위작업별 작업근로자수", min_value=1, step=1, value=unit_data.get("작업자 수", 1), key=f"작업자수_{i}")
        unit_data["작업자 이름"] = st.text_input(f"[{i+1}] 작업근로자 이름", value=unit_data.get("작업자 이름", ""), key=f"작업자이름_{i}")
        
        작업형태_options = ["주간", "교대"]
        current_작업형태_index = 작업형태_options.index(unit_data.get("작업형태", "주간")) if unit_data.get("작업형태", "주간") in 작업형태_options else 0
        unit_data["작업형태"] = st.selectbox(f"[{i+1}] 작업형태", 작업형태_options, index=current_작업형태_index, key=f"작업형태_{i}")
        
        # 데이터 구조 유지
        unit_data["1일 작업시간"] = 0
        unit_data["자세"] = {}
        unit_data["중량물"] = []
        unit_data["도구"] = []

        st.markdown("---")
        
        # 유해요인 원인분석 섹션
        col_hazard_title, col_hazard_add_btn = st.columns([0.8, 0.2])
        with col_hazard_title:
            st.subheader("작업별 유해요인에 대한 원인분석")
        with col_hazard_add_btn:
            if st.button(f"[{i+1}] 항목 추가", key=f"add_hazard_analysis_{i}"):
                unit_data["유해요인_원인분석"].append({"유형": "", "부담작업": "", "부담작업자세": ""})
                st.rerun()
        
        current_hazard_analysis_data = unit_data.get("유해요인_원인분석", [])
        
        # 유해요인 원인분석 항목들 처리
        for k, hazard_entry in enumerate(current_hazard_analysis_data):
            st.markdown(f"**유해요인 원인분석 항목 {k+1}**")
            
            hazard_type_options = ["", "반복동작", "부자연스러운 자세", "과도한 힘", "접촉스트레스 또는 기타(진동, 밀고 당기기 등)"]
            selected_hazard_type_index = hazard_type_options.index(hazard_entry.get("유형", "")) if hazard_entry.get("유형", "") in hazard_type_options else 0
            
            hazard_entry["유형"] = st.selectbox(
                f"[{i+1}-{k+1}] 유해요인 유형 선택", 
                hazard_type_options, 
                index=selected_hazard_type_index, 
                key=f"hazard_type_{i}_{k}"
            )

            # 각 유해요인 유형별 세부 입력 필드들
            if hazard_entry["유형"] == "반복동작":
                burden_task_options = [
                    "",
                    "(1호)하루에 4시간 이상 집중적으로 자료입력 등을 위해 키보드 또는 마우스를 조작하는 작업",
                    "(2호)하루에 총 2시간 이상 목, 어깨, 팔꿈치, 손목 또는 손을 사용하여 같은 동작을 반복하는 작업",
                    "(6호)하루에 총 2시간 이상 지지되지 않은 상태에서 1kg 이상의 물건을 한손의 손가락으로 집어 옮기거나, 2kg 이상에 상응하는 힘을 가하여 한손의 손가락으로 물건을 쥐는 작업",
                    "(7호)하루에 총 2시간 이상 지지되지 않은 상태에서 4.5kg 이상의 물건을 한 손으로 들거나 동일한 힘으로 쥐는 작업",
                    "(10호)하루에 총 2시간 이상, 분당 2회 이상 4.5kg 이상의 물체를 드는 작업",
                    "(1호)하루에 4시간 이상 집중적으로 자료입력 등을 위해 키보드 또는 마우스를 조작하는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)",
                    "(2호)하루에 총 2시간 이상 목, 어깨, 팔꿈치, 손목 또는 손을 사용하여 같은 동작을 반복하는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)",
                    "(6호)하루에 총 2시간 이상 지지되지 않은 상태에서 1kg 이상의 물건을 한손의 손가락으로 집어 옮기거나, 2kg 이상에 상응하는 힘을 가하여 한손의 손가락으로 물건을 쥐는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)",
                    "(7호)하루에 총 2시간 이상 지지되지 않은 상태에서 4.5kg 이상의 물건을 한 손으로 들거나 동일한 힘으로 쥐는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)",
                    "(10호)하루에 총 2시간 이상, 분당 2회 이상 4.5kg 이상의 물체를 드는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)"
                ]
                selected_burden_task_index = burden_task_options.index(hazard_entry.get("부담작업", "")) if hazard_entry.get("부담작업", "") in burden_task_options else 0
                hazard_entry["부담작업"] = st.selectbox(f"[{i+1}-{k+1}] 부담작업", burden_task_options, index=selected_burden_task_index, key=f"burden_task_반복_{i}_{k}")
                
                hazard_entry["수공구 종류"] = st.text_input(f"[{i+1}-{k+1}] 수공구 종류", value=hazard_entry.get("수공구 종류", ""), key=f"수공구_종류_{i}_{k}")
                hazard_entry["수공구 용도"] = st.text_input(f"[{i+1}-{k+1}] 수공구 용도", value=hazard_entry.get("수공구 용도", ""), key=f"수공구_용도_{i}_{k}")
                hazard_entry["수공구 무게(kg)"] = st.number_input(f"[{i+1}-{k+1}] 수공구 무게(kg)", value=hazard_entry.get("수공구 무게(kg)", 0.0), key=f"수공구_무게_{i}_{k}")
                hazard_entry["수공구 사용시간(분)"] = st.text_input(f"[{i+1}-{k+1}] 수공구 사용시간(분)", value=hazard_entry.get("수공구 사용시간(분)", ""), key=f"수공구_사용시간_{i}_{k}")
                hazard_entry["부담부위"] = st.text_input(f"[{i+1}-{k+1}] 부담부위", value=hazard_entry.get("부담부위", ""), key=f"부담부위_{i}_{k}")
                
                # 총 작업시간 자동 계산을 위한 입력 필드
                회당_반복시간_초_회 = st.text_input(f"[{i+1}-{k+1}] 회당 반복시간(초/회)", value=hazard_entry.get("회당 반복시간(초/회)", ""), key=f"반복_회당시간_{i}_{k}")
                작업시간동안_반복횟수_회_일 = st.text_input(f"[{i+1}-{k+1}] 작업시간동안 반복횟수(회/일)", value=hazard_entry.get("작업시간동안 반복횟수(회/일)", ""), key=f"반복_총횟수_{i}_{k}")
                
                hazard_entry["회당 반복시간(초/회)"] = 회당_반복시간_초_회
                hazard_entry["작업시간동안 반복횟수(회/일)"] = 작업시간동안_반복횟수_회_일

                # 총 작업시간(분) 자동 계산
                calculated_total_work_time = 0.0
                try:
                    parsed_회당_반복시간 = parse_value(회당_반복시간_초_회, val_type=float)
                    parsed_작업시간동안_반복횟수 = parse_value(작업시간동안_반복횟수_회_일, val_type=float)
                    
                    if parsed_회당_반복시간 > 0 and parsed_작업시간동안_반복횟수 > 0:
                        calculated_total_work_time = (parsed_회당_반복시간 * parsed_작업시간동안_반복횟수) / 60
                except Exception:
                    pass

                hazard_entry["총 작업시간(분)"] = st.text_input(
                    f"[{i+1}-{k+1}] 총 작업시간(분) (자동계산)",
                    value=f"{calculated_total_work_time:.2f}" if calculated_total_work_time > 0 else "",
                    key=f"반복_총시간_{i}_{k}"
                )

                # 10호 추가 필드
                if "(10호)" in hazard_entry["부담작업"]:
                    hazard_entry["물체 무게(kg)_10호"] = st.number_input(f"[{i+1}-{k+1}] (10호)물체 무게(kg)", value=hazard_entry.get("물체 무게(kg)_10호", 0.0), key=f"물체_무게_10호_{i}_{k}")
                    hazard_entry["분당 반복횟수(회/분)_10호"] = st.text_input(f"[{i+1}-{k+1}] (10호)분당 반복횟수(회/분)", value=hazard_entry.get("분당 반복횟수(회/분)_10호", ""), key=f"분당_반복횟수_10호_{i}_{k}")
                else:
                    hazard_entry["물체 무게(kg)_10호"] = 0.0
                    hazard_entry["분당 반복횟수(회/분)_10호"] = ""

                # 12호 정적자세 관련 필드
                if "(12호)정적자세" in hazard_entry["부담작업"]:
                    hazard_entry["작업내용_12호_정적"] = st.text_input(f"[{i+1}-{k+1}] (정적자세)작업내용", value=hazard_entry.get("작업내용_12호_정적", ""), key=f"반복_작업내용_12호_정적_{i}_{k}")
                    hazard_entry["작업시간(분)_12호_정적"] = st.number_input(f"[{i+1}-{k+1}] (정적자세)작업시간(분)", value=hazard_entry.get("작업시간(분)_12호_정적", 0), key=f"반복_작업시간_12호_정적_{i}_{k}")
                    hazard_entry["휴식시간(분)_12호_정적"] = st.number_input(f"[{i+1}-{k+1}] (정적자세)휴식시간(분)", value=hazard_entry.get("휴식시간(분)_12호_정적", 0), key=f"반복_휴식시간_12호_정적_{i}_{k}")
                    hazard_entry["인체부담부위_12호_정적"] = st.text_input(f"[{i+1}-{k+1}] (정적자세)인체부담부위", value=hazard_entry.get("인체부담부위_12호_정적", ""), key=f"반복_인체부담부위_12호_정적_{i}_{k}")
                else:
                    hazard_entry["작업내용_12호_정적"] = ""
                    hazard_entry["작업시간(분)_12호_정적"] = 0
                    hazard_entry["휴식시간(분)_12호_정적"] = 0
                    hazard_entry["인체부담부위_12호_정적"] = ""

            elif hazard_entry["유형"] == "부자연스러운 자세":
                burden_pose_options = [
                    "",
                    "(3호)하루에 총 2시간 이상 머리 위에 손이 있거나, 팔꿈치가 어깨위에 있거나, 팔꿈치를 몸통으로부터 들거나, 팔꿈치를 몸통뒤쪽에 위치하도록 하는 상태에서 이루어지는 작업",
                    "(4호)지지되지 않은 상태이거나 임의로 자세를 바꿀 수 없는 조건에서, 하루에 총 2시간 이상 목이나 허리를 구부리거나 트는 상태에서 이루어지는 작업",
                    "(5호)하루에 총 2시간 이상 쪼그리고 앉거나 무릎을 굽힌 자세에서 이루어지는 작업"
                ]
                selected_burden_pose_index = burden_pose_options.index(hazard_entry.get("부담작업자세", "")) if hazard_entry.get("부담작업자세", "") in burden_pose_options else 0
                hazard_entry["부담작업자세"] = st.selectbox(f"[{i+1}-{k+1}] 부담작업자세", burden_pose_options, index=selected_burden_pose_index, key=f"burden_pose_{i}_{k}")
                
                hazard_entry["회당 반복시간(초/회)"] = st.text_input(f"[{i+1}-{k+1}] 회당 반복시간(초/회)", value=hazard_entry.get("회당 반복시간(초/회)", ""), key=f"자세_회당시간_{i}_{k}")
                hazard_entry["작업시간동안 반복횟수(회/일)"] = st.text_input(f"[{i+1}-{k+1}] 작업시간동안 반복횟수(회/일)", value=hazard_entry.get("작업시간동안 반복횟수(회/일)", ""), key=f"자세_총횟수_{i}_{k}")
                hazard_entry["총 작업시간(분)"] = st.text_input(f"[{i+1}-{k+1}] 총 작업시간(분)", value=hazard_entry.get("총 작업시간(분)", ""), key=f"자세_총시간_{i}_{k}")

            elif hazard_entry["유형"] == "과도한 힘":
                burden_force_options = [
                    "",
                    "(8호)하루에 10회 이상 25kg 이상의 물체를 드는 작업",
                    "(9호)하루에 25회 이상 10kg 이상의 물체를 무릎 아래에서 들거나, 어깨 위에서 들거나, 팔을 뻗은 상태에서 드는 작업",
                    "(12호)밀기/당기기 작업",
                    "(8호)하루에 10회 이상 25kg 이상의 물체를 드는 작업+(12호)밀기/당기기 작업",
                    "(9호)하루에 25회 이상 10kg 이상의 물체를 무릎 아래에서 들거나, 어깨 위에서 들거나, 팔을 뻗은 상태에서 드는 작업+(12호)밀기/당기기 작업"
                ]
                selected_burden_force_index = burden_force_options.index(hazard_entry.get("부담작업", "")) if hazard_entry.get("부담작업", "") in burden_force_options else 0
                hazard_entry["부담작업"] = st.selectbox(f"[{i+1}-{k+1}] 부담작업", burden_force_options, index=selected_burden_force_index, key=f"burden_force_{i}_{k}")
                
                hazard_entry["중량물 명칭"] = st.text_input(f"[{i+1}-{k+1}] 중량물 명칭", value=hazard_entry.get("중량물 명칭", ""), key=f"힘_중량물_명칭_{i}_{k}")
                hazard_entry["중량물 용도"] = st.text_input(f"[{i+1}-{k+1}] 중량물 용도", value=hazard_entry.get("중량물 용도", ""), key=f"힘_중량물_용도_{i}_{k}")
                hazard_entry["중량물 무게(kg)"] = st.number_input(f"[{i+1}-{k+1}] 중량물 무게(kg)", value=hazard_entry.get("중량물 무게(kg)", 0.0), key=f"중량물_무게_기본_{i}_{k}")
                hazard_entry["하루 8시간동안 중량물을 드는 횟수(회)"] = st.number_input(f"[{i+1}-{k+1}] 하루 8시간동안 중량물을 드는 횟수(회)", value=hazard_entry.get("하루 8시간동안 중량물을 드는 횟수(회)", 0), min_value=0, step=1, key=f"중량물_횟수_{i}_{k}")
                
                취급방법_options = ["", "직접 취급", "크레인 사용"]
                selected_취급방법_index = 취급방법_options.index(hazard_entry.get("취급방법", "")) if hazard_entry.get("취급방법", "") in 취급방법_options else 0
                hazard_entry["취급방법"] = st.selectbox(f"[{i+1}-{k+1}] 취급방법", 취급방법_options, index=selected_취급방법_index, key=f"힘_취급방법_{i}_{k}")

                if hazard_entry["취급방법"] == "직접 취급":
                    이동방법_options = ["", "1인 직접이동", "2인1조 직접이동", "여러명 직접이동", "이동대차(인력이동)", "이동대차(전력이동)", "지게차"]
                    selected_이동방법_index = 이동방법_options.index(hazard_entry.get("중량물 이동방법", "")) if hazard_entry.get("중량물 이동방법", "") in 이동방법_options else 0
                    hazard_entry["중량물 이동방법"] = st.selectbox(f"[{i+1}-{k+1}] 중량물 이동방법", 이동방법_options, index=selected_이동방법_index, key=f"힘_이동방법_{i}_{k}")
                    
                    if hazard_entry["중량물 이동방법"] == "이동대차(인력이동)":
                        직접_밀당_options = ["", "작업자가 직접 바퀴달린 이동대차를 밀고/당기기", "자동이동대차(AGV)", "기타"]
                        selected_직접_밀당_index = 직접_밀당_options.index(hazard_entry.get("작업자가 직접 밀고/당기기", "")) if hazard_entry.get("작업자가 직접 밀고/당기기", "") in 직접_밀당_options else 0
                        hazard_entry["작업자가 직접 밀고/당기기"] = st.selectbox(f"[{i+1}-{k+1}] 작업자가 직접 밀고/당기기", 직접_밀당_options, index=selected_직접_밀당_index, key=f"힘_직접_밀당_{i}_{k}")
                        
                        if hazard_entry["작업자가 직접 밀고/당기기"] == "기타":
                            hazard_entry["기타_밀당_설명"] = st.text_input(f"[{i+1}-{k+1}] 기타 밀기/당기기 설명", value=hazard_entry.get("기타_밀당_설명", ""), key=f"힘_기타_밀당_설명_{i}_{k}")
                        else:
                            hazard_entry["기타_밀당_설명"] = ""
                    else:
                        hazard_entry["작업자가 직접 밀고/당기기"] = ""
                        hazard_entry["기타_밀당_설명"] = ""
                else:
                    hazard_entry["중량물 이동방법"] = ""
                    hazard_entry["작업자가 직접 밀고/당기기"] = ""
                    hazard_entry["기타_밀당_설명"] = ""

                if "(12호)밀기/당기기 작업" not in hazard_entry["부담작업"]:
                    # 밀기/당기기 작업이 아닌 경우에만 기존 필드들 숨김 처리 (이미 위에서 입력받음)
                    pass
                else:
                    # 밀기/당기기 작업 선택 시 중량물 관련 필드들 초기화
                    hazard_entry["중량물 무게(kg)"] = 0.0
                    hazard_entry["하루 8시간동안 중량물을 드는 횟수(회)"] = 0

            elif hazard_entry["유형"] == "접촉스트레스 또는 기타(진동, 밀고 당기기 등)":
                burden_other_options = [
                    "",
                    "(11호)하루에 총 2시간 이상 시간당 10회 이상 손 또는 무릎을 사용하여 반복적으로 충격을 가하는 작업",
                    "(12호)진동작업(그라인더, 임팩터 등)"
                ]
                selected_burden_other_index = burden_other_options.index(hazard_entry.get("부담작업", "")) if hazard_entry.get("부담작업", "") in burden_other_options else 0
                hazard_entry["부담작업"] = st.selectbox(f"[{i+1}-{k+1}] 부담작업", burden_other_options, index=selected_burden_other_index, key=f"burden_other_{i}_{k}")

                if hazard_entry["부담작업"] == "(11호)하루에 총 2시간 이상 시간당 10회 이상 손 또는 무릎을 사용하여 반복적으로 충격을 가하는 작업":
                    hazard_entry["작업시간(분)"] = st.text_input(f"[{i+1}-{k+1}] 작업시간(분)", value=hazard_entry.get("작업시간(분)", ""), key=f"기타_작업시간_{i}_{k}")
                else:
                    hazard_entry["작업시간(분)"] = ""

                if hazard_entry["부담작업"] == "(12호)진동작업(그라인더, 임팩터 등)":
                    st.markdown("**(12호) 세부 유형에 대한 추가 정보 (선택적 입력)**")
                    hazard_entry["진동수공구명"] = st.text_input(f"[{i+1}-{k+1}] 진동수공구명", value=hazard_entry.get("진동수공구명", ""), key=f"기타_진동수공구명_{i}_{k}")
                    hazard_entry["진동수공구 용도"] = st.text_input(f"[{i+1}-{k+1}] 진동수공구 용도", value=hazard_entry.get("진동수공구 용도", ""), key=f"기타_진동수공구_용도_{i}_{k}")
                    hazard_entry["작업시간(분)_진동"] = st.text_input(f"[{i+1}-{k+1}] 작업시간(분)", value=hazard_entry.get("작업시간(분)_진동", ""), key=f"기타_작업시간_진동_{i}_{k}")
                    hazard_entry["작업빈도(초/회)_진동"] = st.text_input(f"[{i+1}-{k+1}] 작업빈도(초/회)", value=hazard_entry.get("작업빈도(초/회)_진동", ""), key=f"기타_작업빈도_진동_{i}_{k}")
                    hazard_entry["작업량(회/일)_진동"] = st.text_input(f"[{i+1}-{k+1}] 작업량(회/일)", value=hazard_entry.get("작업량(회/일)_진동", ""), key=f"기타_작업량_진동_{i}_{k}")
                    
                    지지대_options = ["", "예", "아니오"]
                    selected_지지대_index = 지지대_options.index(hazard_entry.get("수공구사용시 지지대가 있는가?", "")) if hazard_entry.get("수공구사용시 지지대가 있는가?", "") in 지지대_options else 0
                    hazard_entry["수공구사용시 지지대가 있는가?"] = st.selectbox(f"[{i+1}-{k+1}] 수공구사용시 지지대가 있는가?", 지지대_options, index=selected_지지대_index, key=f"기타_지지대_여부_{i}_{k}")
                else:
                    hazard_entry["작업시간(분)"] = ""
                    hazard_entry["진동수공구명"] = ""
                    hazard_entry["진동수공구 용도"] = ""
                    hazard_entry["작업시간(분)_진동"] = ""
                    hazard_entry["작업빈도(초/회)_진동"] = ""
                    hazard_entry["작업량(회/일)_진동"] = ""
                    hazard_entry["수공구사용시 지지대가 있는가?"] = ""

            # 현재 항목의 변경사항을 unit_data에 반영
            unit_data["유해요인_원인분석"][k] = hazard_entry

            # 삭제 버튼 (첫 번째 항목은 삭제 불가)
            if k > 0 or len(current_hazard_analysis_data) > 1:
                col_delete_btn, _ = st.columns([0.2, 0.8])
                with col_delete_btn:
                    if st.button(f"[{i+1}-{k+1}] 항목 삭제", key=f"delete_hazard_analysis_{i}_{k}"):
                        unit_data["유해요인_원인분석"].pop(k)
                        st.rerun()

        # 보호구 및 작성자 정보
        unit_data["보호구"] = st.multiselect(f"[{i+1}] 착용 보호구", ["무릎보호대", "손목보호대", "허리보호대", "각반", "기타"], default=unit_data.get("보호구", []), key=f"protection_gear_{i}")
        unit_data["작성자"] = st.text_input(f"[{i+1}] 작성자 이름", value=unit_data.get("작성자", ""), key=f"author_name_{i}")
        unit_data["연락처"] = st.text_input(f"[{i+1}] 작성자 연락처", value=unit_data.get("연락처", ""), key=f"author_contact_{i}")

        # 근골격계 부담작업 판단 기준 계산 및 업데이트
        burden_criteria = {f"부담작업_{k}호": "X" for k in range(1, 13)}

        for hazard_entry in unit_data.get("유해요인_원인분석", []):
            hazard_type = hazard_entry.get("유형")
            burden_detail_option = hazard_entry.get("부담작업") or hazard_entry.get("부담작업자세")

            if hazard_type == "반복동작":
                total_work_time_min = parse_value(hazard_entry.get("총 작업시간(분)"), val_type=float)
                
                if "(1호)" in burden_detail_option:
                    if burden_criteria["부담작업_1호"] != "O":
                        if total_work_time_min >= 240:
                            burden_criteria["부담작업_1호"] = "O"
                        else:
                            burden_criteria["부담작업_1호"] = "△"
                    if "(12호)정적자세" in burden_detail_option:
                        burden_criteria["부담작업_12호"] = "△"
                        
                elif "(2호)" in burden_detail_option:
                    if burden_criteria["부담작업_2호"] != "O":
                        if total_work_time_min >= 120:
                            burden_criteria["부담작업_2호"] = "O"
                        else:
                            burden_criteria["부담작업_2호"] = "△"
                    if "(12호)정적자세" in burden_detail_option:
                        burden_criteria["부담작업_12호"] = "△"
                        
                elif "(6호)" in burden_detail_option:
                    if burden_criteria["부담작업_6호"] != "O":
                        if total_work_time_min >= 120:
                            burden_criteria["부담작업_6호"] = "O"
                        else:
                            burden_criteria["부담작업_6호"] = "△"
                    if "(12호)정적자세" in burden_detail_option:
                        burden_criteria["부담작업_12호"] = "△"
                        
                elif "(7호)" in burden_detail_option:
                    if burden_criteria["부담작업_7호"] != "O":
                        if total_work_time_min >= 120:
                            burden_criteria["부담작업_7호"] = "O"
                        else:
                            burden_criteria["부담작업_7호"] = "△"
                    if "(12호)정적자세" in burden_detail_option:
                        burden_criteria["부담작업_12호"] = "△"
                        
                elif "(10호)" in burden_detail_option:
                    if burden_criteria["부담작업_10호"] != "O":
                        total_work_time_min_10 = parse_value(hazard_entry.get("총 작업시간(분)"), val_type=float)
                        min_repeat_count = parse_value(hazard_entry.get("분당 반복횟수(회/분)_10호"), val_type=float)
                        object_weight_10 = hazard_entry.get("물체 무게(kg)_10호", 0.0)

                        if total_work_time_min_10 >= 120 and min_repeat_count >= 2 and object_weight_10 >= 4.5:
                            burden_criteria["부담작업_10호"] = "O"
                        else:
                            burden_criteria["부담작업_10호"] = "△"
                    if "(12호)정적자세" in burden_detail_option:
                        burden_criteria["부담작업_12호"] = "△"

            elif hazard_type == "부자연스러운 자세":
                total_work_time_min = parse_value(hazard_entry.get("총 작업시간(분)"), val_type=float)

                if burden_detail_option == "(3호)하루에 총 2시간 이상 머리 위에 손이 있거나, 팔꿈치가 어깨위에 있거나, 팔꿈치를 몸통으로부터 들거나, 팔꿈치를 몸통뒤쪽에 위치하도록 하는 상태에서 이루어지는 작업":
                    if burden_criteria["부담작업_3호"] != "O":
                        if total_work_time_min >= 120:
                            burden_criteria["부담작업_3호"] = "O"
                        else:
                            burden_criteria["부담작업_3호"] = "△"
                elif burden_detail_option == "(4호)지지되지 않은 상태이거나 임의로 자세를 바꿀 수 없는 조건에서, 하루에 총 2시간 이상 목이나 허리를 구부리거나 트는 상태에서 이루어지는 작업":
                    if burden_criteria["부담작업_4호"] != "O":
                        if total_work_time_min >= 120:
                            burden_criteria["부담작업_4호"] = "O"
                        else:
                            burden_criteria["부담작업_4호"] = "△"
                elif burden_detail_option == "(5호)하루에 총 2시간 이상 쪼그리고 앉거나 무릎을 굽힌 자세에서 이루어지는 작업":
                    if burden_criteria["부담작업_5호"] != "O":
                        if total_work_time_min >= 120:
                            burden_criteria["부담작업_5호"] = "O"
                        else:
                            burden_criteria["부담작업_5호"] = "△"

            elif hazard_type == "과도한 힘":
                work_count_per_day = hazard_entry.get("하루 8시간동안 중량물을 드는 횟수(회)", 0)
                object_weight = hazard_entry.get("중량물 무게(kg)", 0.0)

                if burden_detail_option == "(8호)하루에 10회 이상 25kg 이상의 물체를 드는 작업":
                    if burden_criteria["부담작업_8호"] != "O":
                        if work_count_per_day >= 10 and object_weight >= 25:
                            burden_criteria["부담작업_8호"] = "O"
                        else:
                            burden_criteria["부담작업_8호"] = "△"
                elif burden_detail_option == "(9호)하루에 25회 이상 10kg 이상의 물체를 무릎 아래에서 들거나, 어깨 위에서 들거나, 팔을 뻗은 상태에서 드는 작업":
                    if burden_criteria["부담작업_9호"] != "O":
                        if work_count_per_day >= 25 and object_weight >= 10:
                            burden_criteria["부담작업_9호"] = "O"
                        else:
                            burden_criteria["부담작업_9호"] = "△"
                elif burden_detail_option == "(12호)밀기/당기기 작업":
                     burden_criteria["부담작업_12호"] = "△"
                elif "(8호)" in burden_detail_option and "(12호)밀기/당기기" in burden_detail_option:
                    if burden_criteria["부담작업_8호"] != "O":
                        if work_count_per_day >= 10 and object_weight >= 25:
                            burden_criteria["부담작업_8호"] = "O"
                        else:
                            burden_criteria["부담작업_8호"] = "△"
                    burden_criteria["부담작업_12호"] = "△"
                elif "(9호)" in burden_detail_option and "(12호)밀기/당기기" in burden_detail_option:
                    if burden_criteria["부담작업_9호"] != "O":
                        if work_count_per_day >= 25 and object_weight >= 10:
                            burden_criteria["부담작업_9호"] = "O"
                        else:
                            burden_criteria["부담작업_9호"] = "△"
                    burden_criteria["부담작업_12호"] = "△"

            elif hazard_type == "접촉스트레스 또는 기타(진동, 밀고 당기기 등)":
                if burden_detail_option == "(11호)하루에 총 2시간 이상 시간당 10회 이상 손 또는 무릎을 사용하여 반복적으로 충격을 가하는 작업":
                    if burden_criteria["부담작업_11호"] != "O":
                        work_time_min = parse_value(hazard_entry.get("작업시간(분)"), val_type=float)
                        if work_time_min >= 120:
                            burden_criteria["부담작업_11호"] = "O"
                        else:
                            burden_criteria["부담작업_11호"] = "△"
                elif burden_detail_option == "(12호)진동작업(그라인더, 임팩터 등)":
                    burden_criteria["부담작업_12호"] = "△"

        unit_data.update(burden_criteria)

# 엑셀 다운로드 섹션
if st.session_state.task_units:
    output = io.BytesIO()
    rows = []
    
    ordered_columns_prefix = [
        "회사명", "소속", "반", "단위작업명", "작업내용(상세설명)",
        "작업자 수", "작업자 이름", 
        "작업형태", "1일 작업시간"
    ]

    ordered_columns_burden = [f"부담작업_{k}호" for k in range(1, 13)]

    FIXED_MAX_HAZARD_ANALYTICS = 5

    ordered_columns_hazard_analysis = []
    for j in range(FIXED_MAX_HAZARD_ANALYTICS):
        ordered_columns_hazard_analysis.extend([
            f"유해요인_원인분석_유형_{j+1}", 
            f"유해요인_원인분석_부담작업_{j+1}_반복",
            f"유해요인_원인분석_수공구_종류_{j+1}", f"유해요인_원인분석_수공구_용도_{j+1}", 
            f"유해요인_원인분석_수공구_무게(kg)_{j+1}", f"유해요인_원인분석_수공구_사용시간(분)_{j+1}",
            f"유해요인_원인분석_부담부위_{j+1}", f"유해요인_원인분석_반복_회당시간(초/회)_{j+1}", 
            f"유해요인_원인분석_반복_총횟수(회/일)_{j+1}", f"유해요인_원인분석_반복_총시간(분)_{j+1}",
            f"유해요인_원인분석_반복_물체무게_10호(kg)_{j+1}", f"유해요인_원인분석_반복_분당반복횟수_10호(회/분)_{j+1}",
            f"유해요인_원인분석_반복_작업내용_12호_정적_{j+1}", f"유해요인_원인분석_반복_작업시간_12호_정적_{j+1}", 
            f"유해요인_원인분석_반복_휴식시간_12호_정적_{j+1}", f"유해요인_원인분석_반복_인체부담부위_12호_정적_{j+1}",
            f"유해요인_원인분석_부담작업자세_{j+1}",
            f"유해요인_원인분석_자세_회당시간(초/회)_{j+1}", f"유해요인_원인분석_자세_총횟수(회/일)_{j+1}", 
            f"유해요인_원인분석_자세_총시간(분)_{j+1}",
            f"유해요인_원인분석_부담작업_{j+1}_힘",
            f"유해요인_원인분석_힘_중량물_명칭_{j+1}", f"유해요인_원인분석_힘_중량물_용도_{j+1}", 
            f"유해요인_원인분석_중량물_무게(kg)_{j+1}", f"유해요인_원인분석_하루8시간_중량물_횟수(회)_{j+1}",
            f"유해요인_원인분석_힘_취급방법_{j+1}", f"유해요인_원인분석_힘_이동방법_{j+1}", 
            f"유해요인_원인분석_힘_직접_밀당_{j+1}", f"유해요인_원인분석_힘_기타_밀당_설명_{j+1}",
            f"유해요인_원인분석_힘_총횟수(회/일)_{j+1}",
            f"유해요인_원인분석_부담작업_{j+1}_기타",
            f"유해요인_원인분석_기타_작업시간(분)_{j+1}",
            f"유해요인_원인분석_기타_진동수공구명_{j+1}", f"유해요인_원인분석_기타_진동수공구_용도_{j+1}",
            f"유해요인_원인분석_기타_작업시간_진동_{j+1}", f"유해요인_원인분석_기타_작업빈도_진동_{j+1}",
            f"유해요인_원인분석_기타_작업량_진동_{j+1}", f"유해요인_원인분석_기타_지지대_여부_{j+1}"
        ])

    ordered_columns_suffix = ["보호구", "작성자", "연락처"]
    ordered_columns = ordered_columns_prefix + ordered_columns_burden + ordered_columns_hazard_analysis + ordered_columns_suffix

    for unit in st.session_state.task_units:
        base_row = {
            "회사명": unit["회사명"], "소속": unit["소속"], "반": unit["반"],
            "단위작업명": unit["단위작업명"], "작업내용(상세설명)": unit["작업내용(상세설명)"],
            "작업자 수": unit["작업자 수"], "작업자 이름": unit["작업자 이름"],
            "작업형태": unit["작업형태"], "1일 작업시간": unit["1일 작업시간"],
            "보호구": ", ".join(unit["보호구"]), "작성자": unit["작성자"], "연락처": unit["연락처"]
        }
        
        for k_crit in range(1, 13):
            base_row[f"부담작업_{k_crit}호"] = unit.get(f"부담작업_{k_crit}호", "X")

        # 유해요인 원인분석 데이터 평면화
        for j in range(FIXED_MAX_HAZARD_ANALYTICS):
            if j < len(unit["유해요인_원인분석"]):
                hazard_entry = unit["유해요인_원인분석"][j]
                base_row[f"유해요인_원인분석_유형_{j+1}"] = hazard_entry.get("유형", "")
                
                if hazard_entry.get("유형") == "반복동작":
                    base_row[f"유해요인_원인분석_부담작업_{j+1}_반복"] = hazard_entry.get("부담작업", "")
                    base_row[f"유해요인_원인분석_수공구_종류_{j+1}"] = hazard_entry.get("수공구 종류", "")
                    base_row[f"유해요인_원인분석_수공구_용도_{j+1}"] = hazard_entry.get("수공구 용도", "")
                    base_row[f"유해요인_원인분석_수공구_무게(kg)_{j+1}"] = hazard_entry.get("수공구 무게(kg)", 0.0)
                    base_row[f"유해요인_원인분석_수공구_사용시간(분)_{j+1}"] = hazard_entry.get("수공구 사용시간(분)", "")
                    base_row[f"유해요인_원인분석_부담부위_{j+1}"] = hazard_entry.get("부담부위", "")
                    base_row[f"유해요인_원인분석_반복_회당시간(초/회)_{j+1}"] = hazard_entry.get("회당 반복시간(초/회)", "")
                    base_row[f"유해요인_원인분석_반복_총횟수(회/일)_{j+1}"] = hazard_entry.get("작업시간동안 반복횟수(회/일)", "")
                    base_row[f"유해요인_원인분석_반복_총시간(분)_{j+1}"] = hazard_entry.get("총 작업시간(분)", "")
                    base_row[f"유해요인_원인분석_반복_물체무게_10호(kg)_{j+1}"] = hazard_entry.get("물체 무게(kg)_10호", 0.0)
                    base_row[f"유해요인_원인분석_반복_분당반복횟수_10호(회/분)_{j+1}"] = hazard_entry.get("분당 반복횟수(회/분)_10호", "")
                    base_row[f"유해요인_원인분석_반복_작업내용_12호_정적_{j+1}"] = hazard_entry.get("작업내용_12호_정적", "")
                    base_row[f"유해요인_원인분석_반복_작업시간_12호_정적_{j+1}"] = hazard_entry.get("작업시간(분)_12호_정적", "")
                    base_row[f"유해요인_원인분석_반복_휴식시간_12호_정적_{j+1}"] = hazard_entry.get("휴식시간(분)_12호_정적", "")
                    base_row[f"유해요인_원인분석_반복_인체부담부위_12호_정적_{j+1}"] = hazard_entry.get("인체부담부위_12호_정적", "")

                elif hazard_entry.get("유형") == "부자연스러운 자세":
                    base_row[f"유해요인_원인분석_부담작업자세_{j+1}"] = hazard_entry.get("부담작업자세", "")
                    base_row[f"유해요인_원인분석_자세_회당시간(초/회)_{j+1}"] = hazard_entry.get("회당 반복시간(초/회)", "")
                    base_row[f"유해요인_원인분석_자세_총횟수(회/일)_{j+1}"] = hazard_entry.get("작업시간동안 반복횟수(회/일)", "")
                    base_row[f"유해요인_원인분석_자세_총시간(분)_{j+1}"] = hazard_entry.get("총 작업시간(분)", "")
                    
                elif hazard_entry.get("유형") == "과도한 힘":
                    base_row[f"유해요인_원인분석_부담작업_{j+1}_힘"] = hazard_entry.get("부담작업", "")
                    base_row[f"유해요인_원인분석_힘_중량물_명칭_{j+1}"] = hazard_entry.get("중량물 명칭", "")
                    base_row[f"유해요인_원인분석_힘_중량물_용도_{j+1}"] = hazard_entry.get("중량물 용도", "")
                    base_row[f"유해요인_원인분석_중량물_무게(kg)_{j+1}"] = hazard_entry.get("중량물 무게(kg)", 0.0)
                    base_row[f"유해요인_원인분석_하루8시간_중량물_횟수(회)_{j+1}"] = hazard_entry.get("하루 8시간동안 중량물을 드는 횟수(회)", 0)
                    base_row[f"유해요인_원인분석_힘_취급방법_{j+1}"] = hazard_entry.get("취급방법", "")
                    base_row[f"유해요인_원인분석_힘_이동방법_{j+1}"] = hazard_entry.get("중량물 이동방법", "")
                    base_row[f"유해요인_원인분석_힘_직접_밀당_{j+1}"] = hazard_entry.get("작업자가 직접 밀고/당기기", "")
                    base_row[f"유해요인_원인분석_힘_기타_밀당_설명_{j+1}"] = hazard_entry.get("기타_밀당_설명", "")
                    base_row[f"유해요인_원인분석_힘_총횟수(회/일)_{j+1}"] = hazard_entry.get("작업시간동안 작업횟수(회/일)", "")
                    
                elif hazard_entry.get("유형") == "접촉스트레스 또는 기타(진동, 밀고 당기기 등)":
                    base_row[f"유해요인_원인분석_부담작업_{j+1}_기타"] = hazard_entry.get("부담작업", "")
                    if hazard_entry.get("부담작업") == "(11호)하루에 총 2시간 이상 시간당 10회 이상 손 또는 무릎을 사용하여 반복적으로 충격을 가하는 작업":
                        base_row[f"유해요인_원인분석_기타_작업시간(분)_{j+1}"] = hazard_entry.get("작업시간(분)", "")
                    elif hazard_entry.get("부담작업") == "(12호)진동작업(그라인더, 임팩터 등)":
                        base_row[f"유해요인_원인분석_기타_진동수공구명_{j+1}"] = hazard_entry.get("진동수공구명", "")
                        base_row[f"유해요인_원인분석_기타_진동수공구_용도_{j+1}"] = hazard_entry.get("진동수공구 용도", "")
                        base_row[f"유해요인_원인분석_기타_작업시간_진동_{j+1}"] = hazard_entry.get("작업시간(분)_진동", "")
                        base_row[f"유해요인_원인분석_기타_작업빈도_진동_{j+1}"] = hazard_entry.get("작업빈도(초/회)_진동", "")
                        base_row[f"유해요인_원인분석_기타_작업량_진동_{j+1}"] = hazard_entry.get("작업량(회/일)_진동", "")
                        base_row[f"유해요인_원인분석_기타_지지대_여부_{j+1}"] = hazard_entry.get("수공구사용시 지지대가 있는가?", "")
            else:
                # 해당 인덱스에 데이터가 없으면 None으로 채움
                start_idx = ordered_columns_hazard_analysis.index(f"유해요인_원인분석_유형_{j+1}")
                end_idx = start_idx + (len(ordered_columns_hazard_analysis) // FIXED_MAX_HAZARD_ANALYTICS) 
                
                if j < FIXED_MAX_HAZARD_ANALYTICS - 1: 
                    if f"유해요인_원인분석_유형_{j+2}" in ordered_columns_hazard_analysis:
                        end_idx = ordered_columns_hazard_analysis.index(f"유해요인_원인분석_유형_{j+2}")
                    else: 
                        end_idx = len(ordered_columns_hazard_analysis)

                for col_name in ordered_columns_hazard_analysis[start_idx:end_idx]:
                    base_row[col_name] = None

        rows.append(base_row)

    df = pd.DataFrame(rows)
    df = df.reindex(columns=ordered_columns, fill_value=None)

    # 파일명 생성
    if st.session_state.반:
        file_name_base = st.session_state.반
    else:
        file_name_base = "미정반" 
    
    current_date = datetime.now().strftime("%y%m%d")
    file_name = f"작업목록표_{file_name_base}_{current_date}.xlsx"

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='작업목록')

    st.download_button(
        label="📥 작업목록표 다운로드",
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
