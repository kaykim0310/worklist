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

# 앱 시작 시 또는 파일 로드 후, 최소 1개의 단위작업공정이 있도록 보장
if st.session_state.unit_count == 0 and not st.session_state.task_units:
    st.session_state.unit_count = 1
    st.session_state.task_units.append({
        "회사명": st.session_state.group_name, "소속": st.session_state.소속, "반": st.session_state.반,
        "단위작업명": "", "작업내용(상세설명)": "", # 새로운 필드 초기화
        "작업자 수": 1, "작업자 이름": "",
        "작업형태": "주간", "1일 작업시간": 0,
        "자세": {}, "중량물": [], "도구": [],
        "유해요인_원인분석": [{"유형": "", "부담작업": "", "부담작업자세": ""}], # 기본 유해요인 원인분석 항목 1개 추가
        "보호구": [], "작성자": "", "연락처": ""
    })


# 엑셀 파일 업로드 섹션
st.sidebar.header("📊 데이터 불러오기/내보내기")
uploaded_file = st.sidebar.file_uploader("엑셀 파일 업로드 (재시작/수정)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_uploaded = pd.read_excel(uploaded_file, sheet_name='작업목록')

        loaded_task_units = []
        for index, row in df_uploaded.iterrows():
            unit = {
                "회사명": row.get("회사명", ""),
                "소속": row.get("소속", ""),
                "반": row.get("반", ""),
                "단위작업명": row.get("단위작업명", ""),
                "작업내용(상세설명)": row.get("작업내용(상세설명)", ""), # 새로운 필드 파싱
                "작업자 수": row.get("작업자 수", 1),
                "작업자 이름": row.get("작업자 이름", ""),
                "작업형태": row.get("작업형태", "주간"),
                "1일 작업시간": row.get("1일 작업시간", 0),
                "자세": {},
                "중량물": [],
                "도구": [],
                "유해요인_원인분석": [], # 초기화 후 파싱된 데이터 추가
                "보호구": row.get("보호구", "").split(", ") if isinstance(row.get("보호구"), str) else [],
                "작성자": row.get("작성자", ""),
                "연락처": row.get("연락처", "")
            }

            for k_crit in range(1, 13):
                unit[f"부담작업_{k_crit}호"] = row.get(f"부담작업_{k_crit}호", "X")

            # NOTE: 파싱 시 FIXED_MAX_HAZARD_ANALYTICS와 일치해야 합니다.
            FIXED_MAX_HAZARD_ANALYTICS_FOR_PARSE = 5 # 유해요인 항목이 더 늘어날 수 있으므로 조정

            for j_hazard in range(FIXED_MAX_HAZARD_ANALYTICS_FOR_PARSE):
                hazard_type = row.get(f"유해요인_원인분석_유형_{j_hazard+1}")
                if pd.notna(hazard_type) and str(hazard_type).strip() != "":
                    hazard_entry = {"유형": hazard_type}
                    
                    if hazard_type == "반복동작":
                        hazard_entry["부담작업"] = row.get(f"유해요인_원인분석_부담작업_{j_hazard+1}_반복", "")
                        hazard_entry["수공구 종류"] = row.get(f"유해요인_원인분석_수공구_종류_{j_hazard+1}", "")
                        hazard_entry["수공구 용도"] = row.get(f"유해요인_원인분석_수공구_용도_{j_hazard+1}", "")
                        hazard_entry["수공구 무게(kg)"] = row.get(f"유해요인_원인분석_수공구_무게(kg)_{j_hazard+1}", 0.0)
                        hazard_entry["수공구 사용시간(분)"] = row.get(f"유해요인_원인분석_수공구_사용시간(분)_{j_hazard+1}", "")
                        hazard_entry["부담부위"] = row.get(f"유해요인_원인분석_부담부위_{j_hazard+1}", "")
                        hazard_entry["회당 반복시간(초/회)"] = row.get(f"유해요인_원인분석_반복_회당시간(초/회)_{j_hazard+1}", "")
                        hazard_entry["작업시간동안 반복횟수(회/일)"] = row.get(f"유해요인_원인분석_반복_총횟수(회/일)_{j_hazard+1}", "")
                        hazard_entry["총 작업시간(분)"] = row.get(f"유해요인_원인분석_반복_총시간(분)_{j_hazard+1}", "")
                        # 10호 관련 필드
                        hazard_entry["물체 무게(kg)_10호"] = row.get(f"유해요인_원인분석_반복_물체무게_10호(kg)_{j_hazard+1}", 0.0)
                        hazard_entry["분당 반복횟수(회/분)_10호"] = row.get(f"유해요인_원인분석_반복_분당반복횟수_10호(회/분)_{j_hazard+1}", "")
                        # 12호 정적자세 관련 필드 파싱
                        hazard_entry["작업내용_12호_정적"] = row.get(f"유해요인_원인분석_반복_작업내용_12호_정적_{j_hazard+1}", "")
                        hazard_entry["작업시간(분)_12호_정적"] = row.get(f"유해요인_원인분석_반복_작업시간_12호_정적_{j_hazard+1}", "")
                        hazard_entry["휴식시간(분)_12호_정적"] = row.get(f"유해요인_원인분석_반복_휴식시간_12호_정적_{j_hazard+1}", "")
                        hazard_entry["인체부담부위_12호_정적"] = row.get(f"유해요인_원인분석_반복_인체부담부위_12호_정적_{j_hazard+1}", "")

                    elif hazard_type == "부자연스러운 자세":
                        hazard_entry["부담작업자세"] = row.get(f"유해요인_원인분석_부담작업자세_{j_hazard+1}", "")
                        hazard_entry["회당 반복시간(초/회)"] = row.get(f"유해요인_원인분석_자세_회당시간(초/회)_{j_hazard+1}", "")
                        hazard_entry["작업시간동안 반복횟수(회/일)"] = row.get(f"유해요인_원인분석_자세_총횟수(회/일)_{j_hazard+1}", "")
                        hazard_entry["총 작업시간(분)"] = row.get(f"유해요인_원인분석_자세_총시간(분)_{j_hazard+1}", "")
                    elif hazard_type == "과도한 힘":
                        hazard_entry["부담작업"] = row.get(f"유해요인_원인분석_부담작업_{j_hazard+1}_힘", "")
                        hazard_entry["중량물 명칭"] = row.get(f"유해요인_원인분석_힘_중량물_명칭_{j_hazard+1}", "")
                        hazard_entry["중량물 용도"] = row.get(f"유해요인_원인분석_힘_중량물_용도_{j_hazard+1}", "")
                        hazard_entry["취급방법"] = row.get(f"유해요인_원인분석_힘_취급방법_{j_hazard+1}", "")
                        hazard_entry["중량물 이동방법"] = row.get(f"유해요인_원인분석_힘_이동방법_{j_hazard+1}", "")
                        hazard_entry["작업자가 직접 밀고/당기기"] = row.get(f"유해요인_원인분석_힘_직접_밀당_{j_hazard+1}", "")
                        hazard_entry["기타_밀당_설명"] = row.get(f"유해요인_원인분석_힘_기타_밀당_설명_{j_hazard+1}", "") # 기타 밀당 설명 필드 파싱
                        hazard_entry["중량물 무게(kg)"] = row.get(f"유해요인_원인분석_중량물_무게(kg)_{j_hazard+1}", 0.0)
                        hazard_entry["작업시간동안 작업횟수(회/일)"] = row.get(f"유해요인_원인분석_힘_총횟수(회/일)_{j_hazard+1}", "")
                    elif hazard_type == "접촉스트레스 또는 기타(진동, 밀고 당기기 등)":
                        hazard_entry["부담작업"] = row.get(f"유해요인_원인분석_부담작업_{j_hazard+1}_기타", "")
                        if hazard_entry["부담작업"] == "(11호)접촉스트레스":
                            hazard_entry["작업시간(분)"] = row.get(f"유해요인_원인분석_기타_작업시간(분)_{j_hazard+1}", "")
                        elif hazard_entry["부담작업"] == "(12호)진동작업(그라인더, 임팩터 등)":
                            hazard_entry["진동수공구명"] = row.get(f"유해요인_원인분석_기타_진동수공구명_{j_hazard+1}", "")
                            hazard_entry["진동수공구 용도"] = row.get(f"유해요인_원인분석_기타_진동수공구_용도_{j_hazard+1}", "")
                            hazard_entry["작업시간(분)_진동"] = row.get(f"유해요인_원인분석_기타_작업시간_진동_{j_hazard+1}", "")
                            hazard_entry["작업빈도(초/회)_진동"] = row.get(f"유해요인_원인분석_기타_작업빈도_진동_{j_hazard+1}", "")
                            hazard_entry["작업량(회/일)_진동"] = row.get(f"유해요인_원인분석_기타_작업량_진동_{j_hazard+1}", "")
                            hazard_entry["수공구사용시 지지대가 있는가?"] = row.get(f"유해요인_원인분석_기타_지지대_여부_{j_hazard+1}", "")
                    unit["유해요인_원인분석"].append(hazard_entry)
            
            # 로드된 데이터에 원인분석 항목이 없으면 기본 1개 추가 (가독성 목적)
            if not unit["유해요인_원인분석"]:
                unit["유해요인_원인분석"].append({"유형": "", "부담작업": "", "부담작업자세": ""})

            loaded_task_units.append(unit)
        
        if loaded_task_units:
            st.session_state.group_name = loaded_task_units[0].get("회사명", "")
            st.session_state.소속 = loaded_task_units[0].get("소속", "")
            st.session_state.반 = loaded_task_units[0].get("반", "")
            st.session_state.task_units = loaded_task_units
            st.session_state.unit_count = len(loaded_task_units)
            st.sidebar.success("✅ 파일이 성공적으로 로드되었습니다!")
        else:
            st.sidebar.warning("업로드된 파일에 유효한 작업 데이터가 없습니다.")
            # 데이터 없으면 기본 단위작업공정 1개와 원인분석 1개로 초기화 (기존 초기화 로직을 다시 호출)
            st.session_state.unit_count = 1
            st.session_state.task_units = [{
                "회사명": st.session_state.group_name, "소속": st.session_state.소속, "반": st.session_state.반,
                "단위작업명": "", "작업내용(상세설명)": "",
                "작업자 수": 1, "작업자 이름": "",
                "작업형태": "주간", "1일 작업시간": 0,
                "자세": {}, "중량물": [], "도구": [],
                "유해요인_원인분석": [{"유형": "", "부담작업": "", "부담작업자세": ""}],
                "보호구": [], "작성자": "", "연락처": ""
            }]


    except Exception as e:
        st.sidebar.error(f"파일 로드 중 오류 발생: {e}. 올바른 형식의 엑셀 파일인지 확인해주세요.")
        # 오류 발생 시 기본 단위작업공정 1개와 원인분석 1개로 초기화 (기존 초기화 로직을 다시 호출)
        st.session_state.task_units = [{
            "회사명": st.session_state.group_name, "소속": st.session_state.소속, "반": st.session_state.반,
            "단위작업명": "", "작업내용(상세설명)": "",
            "작업자 수": 1, "작업자 이름": "",
            "작업형태": "주간", "1일 작업시간": 0,
            "자세": {}, "중량물": [], "도구": [],
            "유해요인_원인분석": [{"유형": "", "부담작업": "", "부담작업자세": ""}],
            "보호구": [], "작성자": "", "연락처": ""
        }]
        st.session_state.unit_count = 1


st.session_state.group_name = st.text_input("회사명을 입력하세요", value=st.session_state.group_name, key="input_group_name")
st.session_state.소속 = st.text_input("소속/팀/그룹", value=st.session_state.소속, key="input_affiliation")
st.session_state.반 = st.text_input("반", value=st.session_state.반, key="input_class")

# 단위작업 추가 버튼을 입력 필드 옆에 배치
col_unit_add_btn, _ = st.columns([0.2, 0.8])
with col_unit_add_btn:
    add_unit = st.button("단위작업 추가")

if add_unit:
    st.session_state.unit_count += 1
    # 새로운 단위작업 추가 시 빈 데이터 구조 초기화 및 기본 유해요인 분석 항목 추가
    st.session_state.task_units.append({
        "회사명": st.session_state.group_name, "소속": st.session_state.소속, "반": st.session_state.반,
        "단위작업명": "", "작업내용(상세설명)": "", # 새로운 필드 초기화
        "작업자 수": 1, "작업자 이름": "",
        "작업형태": "주간", "1일 작업시간": 0,
        "자세": {}, "중량물": [], "도구": [],
        "유해요인_원인분석": [{"유형": "", "부담작업": "", "부담작업자세": ""}], # 기본 유해요인 원인분석 항목 1개 추가
        "보호구": [], "작성자": "", "연락처": ""
    })
    st.rerun() # UI 즉시 갱신

for i in range(st.session_state.unit_count):
    # 새로운 단위작업이 추가되었을 때 빈 데이터 구조로 초기화 (UI 업데이트를 위해)
    # (add_unit 버튼 로직에서 이미 추가되므로 이 조건문은 사실상 새로 추가된 항목에 대한 방어 로직)
    if i >= len(st.session_state.task_units):
        st.session_state.task_units.append({
            "회사명": st.session_state.group_name, "소속": st.session_state.소속, "반": st.session_state.반,
            "단위작업명": "", "작업내용(상세설명)": "",
            "작업자 수": 1, "작업자 이름": "",
            "작업형태": "주간", "1일 작업시간": 0,
            "자세": {},
            "중량물": [],
            "도구": [],
            "유해요인_원인분석": [{"유형": "", "부담작업": "", "부담작업자세": ""}], # 기본 유해요인 원인분석 항목 1개 추가
            "보호구": [], "작성자": "", "연락처": ""
        })

    unit_data = st.session_state.task_units[i]

    with st.expander(f"단위작업공정 {i+1} 입력", expanded=True):
        unit_data["단위작업명"] = st.text_input(f"[{i+1}] 단위작업명", value=unit_data.get("단위작업명", ""), key=f"작업명_{i}")
        # 새로운 필드: 작업내용(상세설명)
        unit_data["작업내용(상세설명)"] = st.text_area(f"[{i+1}] 작업내용(상세설명)", value=unit_data.get("작업내용(상세설명)", ""), key=f"작업내용_{i}")

        unit_data["작업자 수"] = st.number_input(f"[{i+1}] 단위작업별 작업근로자수", min_value=1, step=1, value=unit_data.get("작업자 수", 1), key=f"작업자수_{i}")
        unit_data["작업자 이름"] = st.text_input(f"[{i+1}] 작업근로자 이름", value=unit_data.get("작업자 이름", ""), key=f"작업자이름_{i}")
        
        작업형태_options = ["주간", "교대"]
        current_작업형태_index = 작업형태_options.index(unit_data.get("작업형태", "주간")) if unit_data.get("작업형태", "주간") in 작업형태_options else 0
        unit_data["작업형태"] = st.selectbox(f"[{i+1}] 작업형태", 작업형태_options, index=current_작업형태_index, key=f"작업형태_{i}")
        
        # "1일 작업시간"과 기존 "근골격계 유해위험요인 선택" (자세, 중량물)은 UI에서 제거합니다.
        # 데이터 구조는 session_state 내부에 유지 (빈 값으로).
        unit_data["1일 작업시간"] = 0
        unit_data["자세"] = {}
        unit_data["중량물"] = []
        unit_data["도구"] = []


        st.markdown("---")
        # "작업별 유해요인에 대한 원인분석" 제목과 추가 버튼을 한 줄에 배치
        col_hazard_title, col_hazard_add_btn = st.columns([0.8, 0.2])
        with col_hazard_title:
            st.subheader("작업별 유해요인에 대한 원인분석")
        with col_hazard_add_btn:
            add_hazard_analysis = st.button(f"[{i+1}] 항목 추가", key=f"add_hazard_analysis_{i}") # 버튼 텍스트 간소화
        
        current_hazard_analysis_data = unit_data.get("유해요인_원인분석", [])
        
        # '유해요인 원인분석 항목 추가' 버튼 클릭 시 빈 데이터 추가
        if add_hazard_analysis:
            current_hazard_analysis_data.append({"유형": "", "부담작업": "", "부담작업자세": ""})
            st.session_state.task_units[i]["유해요인_원인분석"] = current_hazard_analysis_data # session_state 업데이트
            st.rerun() # 항목 추가 후 UI를 즉시 갱신하기 위해 강제 재실행


        # 삭제 후 reruns를 위해 리스트 복사본 사용 (Streamlit 특성상 필요)
        hazard_entries_to_process = list(current_hazard_analysis_data)
        
        for k, hazard_entry in enumerate(hazard_entries_to_process):
            st.markdown(f"**유해요인 원인분석 항목 {k+1}**")
            
            hazard_type_options = ["", "반복동작", "부자연스러운 자세", "과도한 힘", "접촉스트레스 또는 기타(진동, 밀고 당기기 등)"]
            selected_hazard_type_index = hazard_type_options.index(hazard_entry.get("유형", "")) if hazard_entry.get("유형", "") in hazard_type_options else 0
            
            hazard_entry["유형"] = st.selectbox(
                f"[{i+1}-{k+1}] 유해요인 유형 선택", 
                hazard_type_options, 
                index=selected_hazard_type_index, 
                key=f"hazard_type_{i}_{k}"
            )

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
                hazard_entry["수공구 무게(kg)"] = st.number_input(f"[{i+1}-{k+1}] 수공구 무게(kg)", value=hazard_entry.get("수공구 무게(kg)", 0.0), key=f"수공구_무게_{i}_{k}") # 단위 명시
                hazard_entry["수공구 사용시간(분)"] = st.text_input(f"[{i+1}-{k+1}] 수공구 사용시간(분)", value=hazard_entry.get("수공구 사용시간(분)", ""), key=f"수공구_사용시간_{i}_{k}") # 단위 명시
                hazard_entry["부담부위"] = st.text_input(f"[{i+1}-{k+1}] 부담부위", value=hazard_entry.get("부담부위", ""), key=f"부담부위_{i}_{k}")
                
                # --- 총 작업시간(분) 자동 계산을 위한 입력 필드 ---
                # 회당 반복시간(초/회) 및 작업시간동안 반복횟수(회/일) 입력
                회당_반복시간_초_회 = st.text_input(f"[{i+1}-{k+1}] 회당 반복시간(초/회)", value=hazard_entry.get("회당 반복시간(초/회)", ""), key=f"반복_회당시간_{i}_{k}") # 단위 명시
                작업시간동안_반복횟수_회_일 = st.text_input(f"[{i+1}-{k+1}] 작업시간동안 반복횟수(회/일)", value=hazard_entry.get("작업시간동안 반복횟수(회/일)", ""), key=f"반복_총횟수_{i}_{k}") # 단위 명시
                
                # 값 저장
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
                    pass # 계산 오류 시 기본값 0.0 유지

                # 자동 계산된 총 작업시간(분) 표시
                hazard_entry["총 작업시간(분)"] = st.text_input(
                    f"[{i+1}-{k+1}] 총 작업시간(분) (자동계산)",
                    value=f"{calculated_total_work_time:.2f}" if calculated_total_work_time > 0 else "",
                    key=f"반복_총시간_{i}_{k}"
                )


                # 10호 추가 필드
                if "(10호)" in hazard_entry["부담작업"]: # 10호가 포함된 항목이 선택된 경우
                    hazard_entry["물체 무게(kg)_10호"] = st.number_input(f"[{i+1}-{k+1}] (10호)물체 무게(kg)", value=hazard_entry.get("물체 무게(kg)_10호", 0.0), key=f"물체_무게_10호_{i}_{k}")
                    hazard_entry["분당 반복횟수(회/분)_10호"] = st.text_input(f"[{i+1}-{k+1}] (10호)분당 반복횟수(회/분)", value=hazard_entry.get("분당 반복횟수(회/분)_10호", ""), key=f"분당_반복횟수_10호_{i}_{k}")
                else: # 10호 선택 해제 시 필드 초기화
                    hazard_entry["물체 무게(kg)_10호"] = 0.0
                    hazard_entry["분당 반복횟수(회/분)_10호"] = ""

                # 12호 정적자세 관련 필드
                if "(12호)정적자세" in hazard_entry["부담작업"]:
                    hazard_entry["작업내용_12호_정적"] = st.text_input(f"[{i+1}-{k+1}] (12호)작업내용", value=hazard_entry.get("작업내용_12호_정적", ""), key=f"반복_작업내용_12호_정적_{i}_{k}")
                    hazard_entry["작업시간(분)_12호_정적"] = st.number_input(f"[{i+1}-{k+1}] (12호)작업시간(분)", value=hazard_entry.get("작업시간(분)_12호_정적", 0), key=f"반복_작업시간_12호_정적_{i}_{k}")
                    hazard_entry["휴식시간(분)_12호_정적"] = st.number_input(f"[{i+1}-{k+1}] (12호)휴식시간(분)", value=hazard_entry.get("휴식시간(분)_12호_정적", 0), key=f"반복_휴식시간_12호_정적_{i}_{k}")
                    hazard_entry["인체부담부위_12호_정적"] = st.text_input(f"[{i+1}-{k+1}] (12호)인체부담부위", value=hazard_entry.get("인체부담부위_12호_정적", ""), key=f"반복_인체부담부위_12호_정적_{i}_{k}")
                else: # 12호 정적자세 선택 해제 시 필드 초기화
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
                
                hazard_entry["회당 반복시간(초/회)"] = st.text_input(f"[{i+1}-{k+1}] 회당 반복시간(초/회)", value=hazard_entry.get("회당 반복시간(초/회)", ""), key=f"자세_회당시간_{i}_{k}") # 단위 명시
                hazard_entry["작업시간동안 반복횟수(회/일)"] = st.text_input(f"[{i+1}-{k+1}] 작업시간동안 반복횟수(회/일)", value=hazard_entry.get("작업시간동안 반복횟수(회/일)", ""), key=f"자세_총횟수_{i}_{k}") # 단위 명시
                hazard_entry["총 작업시간(분)"] = st.text_input(f"[{i+1}-{k+1}] 총 작업시간(분)", value=hazard_entry.get("총 작업시간(분)", ""), key=f"자세_총시간_{i}_{k}") # 단위 명시

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
                
                # 과도한 힘 세부 정보 (12호 밀기/당기기 작업과 무관)
                hazard_entry["중량물 명칭"] = st.text_input(f"[{i+1}-{k+1}] 중량물 명칭", value=hazard_entry.get("중량물 명칭", ""), key=f"힘_중량물_명칭_{i}_{k}")
                hazard_entry["중량물 용도"] = st.text_input(f"[{i+1}-{k+1}] 중량물 용도", value=hazard_entry.get("중량물 용도", ""), key=f"힘_중량물_용도_{i}_{k}")
                
                # 취급방법
                취급방법_options = ["", "직접 취급", "크레인 사용"]
                selected_취급방법_index = 취급방법_options.index(hazard_entry.get("취급방법", "")) if hazard_entry.get("취급방법", "") in 취급방법_options else 0
                hazard_entry["취급방법"] = st.selectbox(f"[{i+1}-{k+1}] 취급방법", 취급방법_options, index=selected_취급방법_index, key=f"힘_취급방법_{i}_{k}")

                # 중량물 이동방법 (취급방법이 "직접 취급"인 경우만 해당)
                if hazard_entry["취급방법"] == "직접 취급":
                    이동방법_options = ["", "1인 직접이동", "2인1조 직접이동", "여러명 직접이동", "이동대차(인력이동)", "이동대차(전력이동)", "지게차"]
                    selected_이동방법_index = 이동방법_options.index(hazard_entry.get("중량물 이동방법", "")) if hazard_entry.get("중량물 이동방법", "") in 이동방법_options else 0
                    hazard_entry["중량물 이동방법"] = st.selectbox(f"[{i+1}-{k+1}] 중량물 이동방법", 이동방법_options, index=selected_이동방법_index, key=f"힘_이동방법_{i}_{k}")
                    
                    # 이동대차(인력이동) 선택 시 추가 드롭다운
                    if hazard_entry["중량물 이동방법"] == "이동대차(인력이동)":
                        직접_밀당_options = ["", "작업자가 직접 바퀴달린 이동대차를 밀고/당기기", "자동이동대차(AGV)", "기타"]
                        selected_직접_밀당_index = 직접_밀당_options.index(hazard_entry.get("작업자가 직접 밀고/당기기", "")) if hazard_entry.get("작업자가 직접 밀고/당기기", "") in 직접_밀당_options else 0
                        hazard_entry["작업자가 직접 밀고/당기기"] = st.selectbox(f"[{i+1}-{k+1}] 작업자가 직접 밀고/당기기", 직접_밀당_options, index=selected_직접_밀당_index, key=f"힘_직접_밀당_{i}_{k}")
                        # '기타' 선택 시 설명 적는 난 추가
                        if hazard_entry["작업자가 직접 밀고/당기기"] == "기타":
                            hazard_entry["기타_밀당_설명"] = st.text_input(f"[{i+1}-{k+1}] 기타 밀기/당기기 설명", value=hazard_entry.get("기타_밀당_설명", ""), key=f"힘_기타_밀당_설명_{i}_{k}")
                        else:
                            hazard_entry["기타_밀당_설명"] = ""
                    else:
                        hazard_entry["작업자가 직접 밀고/당기기"] = "" # 초기화
                        hazard_entry["기타_밀당_설명"] = "" # 기타 설명도 초기화
                else: # 취급방법이 "직접 취급"이 아닌 경우 이동방법 필드 초기화
                    hazard_entry["중량물 이동방법"] = ""
                    hazard_entry["작업자가 직접 밀고/당기기"] = ""
                    hazard_entry["기타_밀당_설명"] = "" # 기타 설명도 초기화

                # 중량물 무게와 작업 횟수 필드는 복합 항목이 아닌 경우에만 의미가 있으므로,
                # 밀기/당기기 관련 항목이 아닐 때만 표시되도록 조건 추가 (가독성 향상)
                if "(12호)밀기/당기기 작업" not in hazard_entry["부담작업"]:
                    hazard_entry["중량물 무게(kg)"] = st.number_input(f"[{i+1}-{k+1}] 중량물 무게(kg)", value=hazard_entry.get("중량물 무게(kg)", 0.0), key=f"중량물_무게_{i}_{k}") # 단위 명시
                    hazard_entry["작업시간동안 작업횟수(회/일)"] = st.text_input(f"[{i+1}-{k+1}] 작업시간동안 작업횟수(회/일)", value=hazard_entry.get("작업시간동안 작업횟수(회/일)", ""), key=f"힘_총횟수_{i}_{k}") # 단위 명시
                else: # 밀기/당기기 작업 선택 시 중량물 무게, 작업횟수 필드 초기화
                    hazard_entry["중량물 무게(kg)"] = 0.0
                    hazard_entry["작업시간동안 작업횟수(회/일)"] = ""

            elif hazard_entry["유형"] == "접촉스트레스 또는 기타(진동, 밀고 당기기 등)":
                burden_other_options = [
                    "",
                    "(11호)하루에 총 2시간 이상 시간당 10회 이상 손 또는 무릎을 사용하여 반복적으로 충격을 가하는 작업",
                    "(12호)진동작업(그라인더, 임팩터 등)" # 12호 수정
                ]
                selected_burden_other_index = burden_other_options.index(hazard_entry.get("부담작업", "")) if hazard_entry.get("부담작업", "") in burden_other_options else 0
                hazard_entry["부담작업"] = st.selectbox(f"[{i+1}-{k+1}] 부담작업", burden_other_options, index=selected_burden_other_index, key=f"burden_other_{i}_{k}")

                if hazard_entry["부담작업"] == "(11호)하루에 총 2시간 이상 시간당 10회 이상 손 또는 무릎을 사용하여 반복적으로 충격을 가하는 작업":
                    hazard_entry["작업시간(분)"] = st.text_input(f"[{i+1}-{k+1}] 작업시간(분)", value=hazard_entry.get("작업시간(분)", ""), key=f"기타_작업시간_{i}_{k}") # 단위 명시
                else: # 11호 선택 해제 시 필드 초기화
                    hazard_entry["작업시간(분)"] = ""

                if hazard_entry["부담작업"] == "(12호)진동작업(그라인더, 임팩터 등)": # 변경된 12호 옵션
                    st.markdown("**(12호) 세부 유형에 대한 추가 정보 (선택적 입력)**")
                    # 진동작업 필드만 해당 (밀기-당기기는 '과도한 힘'으로 이동)
                    hazard_entry["진동수공구명"] = st.text_input(f"[{i+1}-{k+1}] 진동수공구명", value=hazard_entry.get("진동수공구명", ""), key=f"기타_진동수공구명_{i}_{k}")
                    hazard_entry["진동수공구 용도"] = st.text_input(f"[{i+1}-{k+1}] 진동수공구 용도", value=hazard_entry.get("진동수공구 용도", ""), key=f"기타_진동수공구_용도_{i}_{k}")
                    hazard_entry["작업시간(분)_진동"] = st.text_input(f"[{i+1}-{k+1}] 작업시간(분)", value=hazard_entry.get("작업시간(분)_진동", ""), key=f"기타_작업시간_진동_{i}_{k}") # 단위 명시
                    hazard_entry["작업빈도(초/회)_진동"] = st.text_input(f"[{i+1}-{k+1}] 작업빈도(초/회)", value=hazard_entry.get("작업빈도(초/회)_진동", ""), key=f"기타_작업빈도_진동_{i}_{k}") # 단위 명시
                    hazard_entry["작업량(회/일)_진동"] = st.text_input(f"[{i+1}-{k+1}] 작업량(회/일)", value=hazard_entry.get("작업량(회/일)_진동", ""), key=f"기타_작업량_진동_{i}_{k}") # 단위 명시
                    
                    지지대_options = ["", "예", "아니오"]
                    selected_지지대_index = 지지대_options.index(hazard_entry.get("수공구사용시 지지대가 있는가?", "")) if hazard_entry.get("수공구사용시 지지대가 있는가?", "") in 지지대_options else 0
                    hazard_entry["수공구사용시 지지대가 있는가?"] = st.selectbox(f"[{i+1}-{k+1}] 수공구사용시 지지대가 있는가?", 지지대_options, index=selected_지지대_index, key=f"기타_지지대_여부_{i}_{k}") # 단위 명시
                else: # 12호 선택 해제 시 필드 초기화
                    hazard_entry["작업시간(분)"] = ""
                    hazard_entry["진동수공구명"] = ""
                    hazard_entry["진동수공구 용도"] = ""
                    hazard_entry["작업시간(분)_진동"] = ""
                    hazard_entry["작업빈도(초/회)_진동"] = ""
                    hazard_entry["작업량(회/일)_진동"] = ""
                    hazard_entry["수공구사용시 지지대가 있는가?"] = ""


            # 현재 항목의 모든 변경사항을 unit_data에 반영 (Streamlit의 상태 관리)
            unit_data["유해요인_원인분석"][k] = hazard_entry

            # 삭제 버튼
            col_delete_btn, _ = st.columns([0.2, 0.8])
            with col_delete_btn:
                if st.button(f"[{i+1}-{k+1}] 항목 삭제", key=f"delete_hazard_analysis_{i}_{k}"): # 버튼 텍스트 간소화
                    unit_data["유해요인_원인분석"].pop(k)
                    st.rerun()


        unit_data["보호구"] = st.multiselect(f"[{i+1}] 착용 보호구", ["무릎보호대", "손목보호대", "허리보호대", "각반", "기타"], default=unit_data.get("보호구", []), key=f"protection_gear_{i}")
        unit_data["작성자"] = st.text_input(f"[{i+1}] 작성자 이름", value=unit_data.get("작성자", ""), key=f"author_name_{i}")
        unit_data["연락처"] = st.text_input(f"[{i+1}] 작성자 연락처", value=unit_data.get("연락처", ""), key=f"author_contact_{i}")

        # --- 근골격계 부담작업 판단 기준 계산 및 업데이트 (원인분석 섹션 기반) ---
        # 모든 부담작업호를 "X"로 초기화
        burden_criteria = {f"부담작업_{k}호": "X" for k in range(1, 13)}

        for hazard_entry in unit_data.get("유해요인_원인분석", []):
            hazard_type = hazard_entry.get("유형")
            burden_detail_option = hazard_entry.get("부담작업") or hazard_entry.get("부담작업자세") # 두 가지 필드 모두 확인

            if hazard_type == "반복동작":
                total_work_time_min = parse_value(hazard_entry.get("총 작업시간(분)"), val_type=float)
                
                # 1호
                if "(1호)" in burden_detail_option:
                    if burden_criteria["부담작업_1호"] != "O": # 이미 O인 경우는 덮어쓰지 않음
                        if total_work_time_min >= 240: # 4시간 = 240분
                            burden_criteria["부담작업_1호"] = "O"
                        else:
                            burden_criteria["부담작업_1호"] = "△"
                    if "(12호)정적자세" in burden_detail_option:
                        burden_criteria["부담작업_12호"] = "△" # 12호는 무조건 △
                # 2호
                elif "(2호)" in burden_detail_option:
                    if burden_criteria["부담작업_2호"] != "O":
                        if total_work_time_min >= 120: # 2시간 = 120분
                            burden_criteria["부담작업_2호"] = "O"
                        else:
                            burden_criteria["부담작업_2호"] = "△"
                    if "(12호)정적자세" in burden_detail_option:
                        burden_criteria["부담작업_12호"] = "△"
                # 6호
                elif "(6호)" in burden_detail_option:
                    if burden_criteria["부담작업_6호"] != "O":
                        if total_work_time_min >= 120: # 2시간 = 120분
                            burden_criteria["부담작업_6호"] = "O"
                        else:
                            burden_criteria["부담작업_6호"] = "△"
                    if "(12호)정적자세" in burden_detail_option:
                        burden_criteria["부담작업_12호"] = "△"
                # 7호
                elif "(7호)" in burden_detail_option:
                    if burden_criteria["부담작업_7호"] != "O":
                        if total_work_time_min >= 120: # 2시간 = 120분
                            burden_criteria["부담작업_7호"] = "O"
                        else:
                            burden_criteria["부담작업_7호"] = "△"
                    if "(12호)정적자세" in burden_detail_option:
                        burden_criteria["부담작업_12호"] = "△"
                # 10호
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
                work_count_per_day = parse_value(hazard_entry.get("작업시간동안 작업횟수(회/일)"), val_type=int)
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
                elif burden_detail_option == "(12호)밀기/당기기 작업": # 과도한 힘 유형에서 단독으로 "밀기/당기기 작업" 선택 시
                     burden_criteria["부담작업_12호"] = "△" # 12호는 잠재위험
                elif "(8호)" in burden_detail_option and "(12호)밀기/당기기" in burden_detail_option: # 8호 + 12호 복합
                    if burden_criteria["부담작업_8호"] != "O":
                        if work_count_per_day >= 10 and object_weight >= 25:
                            burden_criteria["부담작업_8호"] = "O"
                        else:
                            burden_criteria["부담작업_8호"] = "△"
                    burden_criteria["부담작업_12호"] = "△" # 12호는 잠재위험
                elif "(9호)" in burden_detail_option and "(12호)밀기/당기기" in burden_detail_option: # 9호 + 12호 복합
                    if burden_criteria["부담작업_9호"] != "O":
                        if work_count_per_day >= 25 and object_weight >= 10:
                            burden_criteria["부담작업_9호"] = "O"
                        else:
                            burden_criteria["부담작업_9호"] = "△"
                    burden_criteria["부담작업_12호"] = "△" # 12호는 잠재위험


            elif hazard_type == "접촉스트레스 또는 기타(진동, 밀고 당기기 등)":
                if burden_detail_option == "(11호)하루에 총 2시간 이상 시간당 10회 이상 손 또는 무릎을 사용하여 반복적으로 충격을 가하는 작업":
                    if burden_criteria["부담작업_11호"] != "O":
                        work_time_min = parse_value(hazard_entry.get("작업시간(분)"), val_type=float)
                        if work_time_min >= 120:
                            burden_criteria["부담작업_11호"] = "O"
                        else:
                            burden_criteria["부담작업_11호"] = "△"
                elif burden_detail_option == "(12호)진동작업(그라인더, 임팩터 등)":
                    burden_criteria["부담작업_12호"] = "△" # 12호는 잠재위험 (진동작업)

        unit_data.update(burden_criteria)

# 엑셀 다운로드 섹션
if st.session_state.task_units:
    output = io.BytesIO()
    rows = []
    
    ordered_columns_prefix = [
        "회사명", "소속", "반", "단위작업명", "작업내용(상세설명)", # 새로운 필드 추가
        "작업자 수", "작업자 이름", 
        "작업형태", "1일 작업시간" # 이 값은 이제 유해요인 분석에서 가져오지 않으므로 값이 없을 수 있음
    ]

    ordered_columns_burden = [f"부담작업_{k}호" for k in range(1, 13)] # 12호까지 포함

    FIXED_MAX_HAZARD_ANALYTICS = 5 # 유해요인 항목이 더 늘어났으므로 5개로 조정

    ordered_columns_hazard_analysis = []
    for j in range(FIXED_MAX_HAZARD_ANALYTICS):
        ordered_columns_hazard_analysis.extend([
            f"유해요인_원인분석_유형_{j+1}", 
            f"유해요인_원인분석_부담작업_{j+1}_반복", # 반복동작
            f"유해요인_원인분석_수공구_종류_{j+1}", f"유해요인_원인분석_수공구_용도_{j+1}", 
            f"유해요인_원인분석_수공구_무게(kg)_{j+1}", f"유해요인_원인분석_수공구_사용시간(분)_{j+1}",
            f"유해요인_원인분석_부담부위_{j+1}", f"유해요인_원인분석_반복_회당시간(초/회)_{j+1}", 
            f"유해요인_원인분석_반복_총횟수(회/일)_{j+1}", f"유해요인_원인분석_반복_총시간(분)_{j+1}",
            f"유해요인_원인분석_반복_물체무게_10호(kg)_{j+1}", f"유해요인_원인분석_반복_분당반복횟수_10호(회/분)_{j+1}", # 10호 관련 필드
            f"유해요인_원인분석_반복_작업내용_12호_정적_{j+1}", f"유해요인_원인분석_반복_작업시간_12호_정적_{j+1}", 
            f"유해요인_원인분석_반복_휴식시간_12호_정적_{j+1}", f"유해요인_원인분석_반복_인체부담부위_12호_정적_{j+1}", # 12호 정적자세 관련 필드
            f"유해요인_원인분석_부담작업자세_{j+1}", # 부자연스러운 자세
            f"유해요인_원인분석_자세_회당시간(초/회)_{j+1}", f"유해요인_원인분석_자세_총횟수(회/일)_{j+1}", 
            f"유해요인_원인분석_자세_총시간(분)_{j+1}",
            f"유해요인_원인분석_부담작업_{j+1}_힘", # 과도한 힘
            f"유해요인_원인분석_힘_중량물_명칭_{j+1}", f"유해요인_원인분석_힘_중량물_용도_{j+1}", 
            f"유해요인_원인분석_힘_취급방법_{j+1}", f"유해요인_원인분석_힘_이동방법_{j+1}", 
            f"유해요인_원인분석_힘_직접_밀당_{j+1}", f"유해요인_원인분석_힘_기타_밀당_설명_{j+1}", # 기타 밀당 설명 필드 추가
            f"유해요인_원인분석_중량물_무게(kg)_{j+1}", f"유해요인_원인분석_힘_총횟수(회/일)_{j+1}",
            f"유해요인_원인분석_부담작업_{j+1}_기타", # 접촉스트레스 또는 기타
            f"유해요인_원인분석_기타_작업시간(분)_{j+1}",
            f"유해요인_원인분석_기타_진동수공구명_{j+1}", f"유해요인_원인분석_기타_진동수공구_용도_{j+1}", # 새로운 진동 필드
            f"유해요인_원인분석_기타_작업시간_진동_{j+1}", f"유해요인_원인분석_기타_작업빈도_진동_{j+1}", # 새로운 진동 필드
            f"유해요인_원인분석_기타_작업량_진동_{j+1}", f"유해요인_원인분석_기타_지지대_여부_{j+1}" # 새로운 진동 필드
        ])

    ordered_columns_suffix = ["보호구", "작성자", "연락처"]

    ordered_columns = ordered_columns_prefix + ordered_columns_burden + ordered_columns_hazard_analysis + ordered_columns_suffix


    for unit in st.session_state.task_units:
        base_row = {
            "회사명": unit["회사명"], "소속": unit["소속"], "반": unit["반"],
            "단위작업명": unit["단위작업명"], "작업내용(상세설명)": unit["작업내용(상세설명)"], # 새로운 필드 추가
            "작업자 수": unit["작업자 수"], "작업자 이름": unit["작업자 이름"],
            "작업형태": unit["작업형태"], "1일 작업시간": unit["1일 작업시간"],
            "보호구": ", ".join(unit["보호구"]), "작성자": unit["작성자"], "연락처": unit["연락처"]
        }
        
        for k_crit in range(1, 13):
            base_row[f"부담작업_{k_crit}호"] = unit.get(f"부담작업_{k_crit}호", "X")

        # 유해요인 원인분석 데이터 평면화 (컬럼명에 단위 추가)
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
                    base_row[f"유해요인_원인분석_힘_취급방법_{j+1}"] = hazard_entry.get("취급방법", "")
                    base_row[f"유해요인_원인분석_힘_이동방법_{j+1}"] = hazard_entry.get("중량물 이동방법", "")
                    base_row[f"유해요인_원인분석_힘_직접_밀당_{j+1}"] = hazard_entry.get("작업자가 직접 밀고/당기기", "")
                    base_row[f"유해요인_원인분석_힘_기타_밀당_설명_{j+1}"] = hazard_entry.get("기타_밀당_설명", "") # 기타 밀당 설명 필드 추가
                    base_row[f"유해요인_원인분석_중량물_무게(kg)_{j+1}"] = hazard_entry.get("중량물 무게(kg)", 0.0)
                    base_row[f"유해요인_원인분석_힘_총횟수(회/일)_{j+1}"] = hazard_entry.get("작업시간동안 반복횟수(회/일)", "")
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
            else: # 해당 인덱스에 데이터가 없으면 None으로 채움
                start_idx = ordered_columns_hazard_analysis.index(f"유해요인_원인분석_유형_{j+1}")
                end_idx = start_idx + (len(ordered_columns_hazard_analysis) // FIXED_MAX_HAZARD_ANALYTICS) 
                
                if j < FIXED_MAX_HAZARD_ANALYTICS -1 : 
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
