import streamlit as st
import pandas as pd
import io

st.title("📋 작업목록표 입력 시스템")

task_units = []

group_name = st.text_input("회사명을 입력하세요")
소속 = st.text_input("소속/팀/그룹")
반 = st.text_input("반")

add_unit = st.button("단위작업 추가")
if 'unit_count' not in st.session_state:
    st.session_state.unit_count = 0

if add_unit:
    st.session_state.unit_count += 1

for i in range(st.session_state.unit_count):
    with st.expander(f"단위작업공정 {i+1} 입력"):
        작업명 = st.text_input(f"[{i+1}] 단위작업명")
        작업자수 = st.number_input(f"[{i+1}] 단위작업별 작업근로자수", min_value=1, step=1)
        작업자이름 = st.text_input(f"[{i+1}] 작업근로자 이름")
        작업형태 = st.selectbox(f"[{i+1}] 작업형태", ["주간", "교대"])
        작업시간 = st.number_input(f"[{i+1}] 1일 작업시간 (시간 단위)", min_value=0, step=1)

        유해요인 = st.multiselect(f"[{i+1}] 근골격계 유해위험요인 선택", ["자세", "중량물"])

        자세 = {}
        중량물 = []
        도구 = []

        if "자세" in 유해요인:
            st.markdown("**자세 관련 정보**")
            자세["어깨"] = st.number_input(f"[{i+1}] 어깨 위로 팔이 올라가는 자세 (작업시간)", min_value=0.0, step=0.5)
            자세["몸통"] = st.number_input(f"[{i+1}] 몸통이 비트는 자세 (작업시간)", min_value=0.0, step=0.5)
            자세["쪼그림"] = st.number_input(f"[{i+1}] 쪼그려 앉는 자세 (작업시간)", min_value=0.0, step=0.5)
            자세["반복전체"] = st.number_input(f"[{i+1}] 반복작업 (1일 작업시간)", min_value=0.0, step=0.5)
            자세["반복무거운"] = st.number_input(f"[{i+1}] 반복작업 (4.5kg 이상, 분당 작업횟수)", min_value=0, step=1)

        if "중량물" in 유해요인:
            st.markdown("**중량물 관련 정보**")
            수공구_수 = st.number_input(f"[{i+1}] 수공구 종류 수", min_value=0, step=1)
            for j in range(수공구_수):
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    명칭 = st.text_input(f"[{i+1}-{j+1}] 수공구명")
                with col2:
                    용도 = st.text_input(f"[{i+1}-{j+1}] 수공구 용도")
                with col3:
                    무게 = st.number_input(f"[{i+1}-{j+1}] 수공구 무게(kg)", min_value=0.0)
                with col4:
                    시간 = st.text_input(f"[{i+1}-{j+1}] 작업 횟수/시간")
                도구.append((명칭, 용도, 무게, 시간))

            중량물_수 = st.number_input(f"[{i+1}] 중량물 종류 수", min_value=0, step=1)
            for j in range(중량물_수):
                col1, col2, col3 = st.columns(3)
                with col1:
                    명칭 = st.text_input(f"[{i+1}-{j+1}] 중량물명")
                with col2:
                    무게 = st.number_input(f"[{i+1}-{j+1}] 중량물 무게(kg)", min_value=0.0)
                with col3:
                    횟수 = st.number_input(f"[{i+1}-{j+1}] 1일 작업 횟수", min_value=0)
                중량물.append((명칭, 무게, 횟수))

        보호구 = st.multiselect(f"[{i+1}] 착용 보호구", ["무릎보호대", "손목보호대", "허리보호대", "각반", "기타"])
        작성자 = st.text_input(f"[{i+1}] 작성자 이름")
        연락처 = st.text_input(f"[{i+1}] 작성자 연락처")

        저장 = st.button(f"저장하기", key=f"save_{i}")
        if 저장:
            st.success("✅ 저장이 완료되었습니다!")

        task_units.append({
            "회사명": group_name,
            "소속": 소속,
            "반": 반,
            "단위작업명": 작업명,
            "작업자 수": 작업자수,
            "작업자 이름": 작업자이름,
            "작업형태": 작업형태,
            "1일 작업시간": 작업시간,
            "자세": 자세,
            "중량물": 중량물,
            "도구": 도구,
            "보호구": 보호구,
            "작성자": 작성자,
            "연락처": 연락처
        })

if task_units:
    output = io.BytesIO()
    rows = []
    for unit in task_units:
        base_row = {
            "회사명": unit["회사명"],
            "소속": unit["소속"],
            "반": unit["반"],
            "단위작업명": unit["단위작업명"],
            "작업자 수": unit["작업자 수"],
            "작업자 이름": unit["작업자 이름"],
            "작업형태": unit["작업형태"],
            "1일 작업시간": unit["1일 작업시간"],
            "자세_어깨": unit["자세"].get("어깨"),
            "자세_몸통": unit["자세"].get("몸통"),
            "자세_쪼그림": unit["자세"].get("쪼그림"),
            "자세_반복전체": unit["자세"].get("반복전체"),
            "자세_반복무거운": unit["자세"].get("반복무거운"),
            "보호구": ", ".join(unit["보호구"]),
            "작성자": unit["작성자"],
            "연락처": unit["연락처"]
        }
        for tool in unit["도구"]:
            rows.append({**base_row, "구분": "수공구", "명칭": tool[0], "용도": tool[1], "무게(kg)": tool[2], "작업횟수/시간": tool[3]})
        for mat in unit["중량물"]:
            rows.append({**base_row, "구분": "중량물", "명칭": mat[0], "용도": "-", "무게(kg)": mat[1], "작업횟수/시간": mat[2]})
        if not unit["도구"] and not unit["중량물"]:
            rows.append(base_row)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame(rows).to_excel(writer, index=False, sheet_name='작업목록')

    st.download_button(
        label="📥 작업목록표 다운로드",
        data=output.getvalue(),
        file_name=f"작업목록표_{반}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
