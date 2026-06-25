import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(layout="wide")

st.title("📊 성적 분석 대시보드")
st.subheader("Advanced, On, Developing 점수 구간 분석 (캠퍼스별/레벨별)")

# =====================================================================
# ⚙️ [내부용 설정 1] 월간 평가(MT) 기준 커트라인 (100점 환산 기준)
# =====================================================================
MT_CRITERIA = {
    'Penta': {'Phonics': (88, 73), 'Reading': (81, 60), 'Overall': (85, 67)},
    'Hexa': {'Listening': (91, 76), 'Reading': (86, 70), 'Vocabulary': (91, 71), 'Overall': (89, 72)},
    'Hepta': {'Listening': (95, 75), 'Reading': (95, 75), 'Vocabulary': (90, 70), 'Grammar': (90, 70), 'Overall': (92, 73)},
    'Octa': {'Listening': (90, 70), 'Reading': (90, 70), 'Vocabulary': (90, 70), 'Grammar': (90, 70), 'Overall': (90, 70)},
    'Nona': {'Listening': (90, 70), 'Reading': (90, 70), 'Vocabulary': (90, 70), 'Grammar': (90, 70), 'Overall': (90, 70)}
}
MT_MAX_SCORES = {
    'Penta': {'Phonics': 75, 'Reading': 25},
    'Hexa': {'Listening': 100, 'Reading': 100, 'Vocabulary': 30},
    'Hepta': {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30},
    'Octa': {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30},
    'Nona': {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30}
}

# =====================================================================
# ⚙️ [내부용 설정 2] 학기 평가(TT) 이미지 기준 커트라인 데이터 전면 반영
# =====================================================================
TT_CRITERIA = {
    'Hepta 1': {'Listening': (79, 60), 'Reading': (79, 60), 'Vocabulary': (74, 51), 'Grammar': (68, 44), 'Overall': (75, 54)},
    'Hepta 2': {'Listening': (92, 73), 'Reading': (92, 73), 'Vocabulary': (88, 68), 'Grammar': (84, 61), 'Overall': (89, 69)},
    'Octa 1':  {'Listening': (71, 40), 'Reading': (63, 37), 'Vocabulary': (68, 41), 'Grammar': (61, 44), 'Overall': (66, 41)},
    'Octa 2':  {'Listening': (82, 53), 'Reading': (78, 53), 'Vocabulary': (81, 54), 'Grammar': (71, 54), 'Overall': (78, 54)},
    'Octa 3':  {'Listening': (92, 65), 'Reading': (86, 62), 'Vocabulary': (88, 61), 'Grammar': (81, 61), 'Overall': (87, 62)},
    'Nona 1':  {'Listening': (69, 37), 'Reading': (60, 34), 'Vocabulary': (68, 41), 'Grammar': (64, 48), 'Overall': (65, 40)},
    'Nona 2':  {'Listening': (80, 50), 'Reading': (75, 51), 'Vocabulary': (84, 54), 'Grammar': (71, 54), 'Overall': (78, 52)},
    'Nona 3':  {'Listening': (90, 63), 'Reading': (89, 65), 'Vocabulary': (88, 61), 'Grammar': (81, 58), 'Overall': (87, 62)}
}
TT_MAX_SCORES = {
    'Hepta 1': {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30, 'Overall': 100},
    'Hepta 2': {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30, 'Overall': 100},
    'Octa 1':  {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30, 'Overall': 100},
    'Octa 2':  {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30, 'Overall': 100},
    'Octa 3':  {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30, 'Overall': 100},
    'Nona 1':  {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30, 'Overall': 100},
    'Nona 2':  {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30, 'Overall': 100},
    'Nona 3':  {'Listening': 100, 'Reading': 100, 'Vocabulary': 30, 'Grammar': 30, 'Overall': 100}
}

# 사이드바 설정
st.sidebar.header("⚙️ 분석 설정")
exam_type = st.sidebar.radio("분석할 시험 종류를 선택하세요:", ["월간 평가 (MT)", "학기 평가 (TT)"])

st.sidebar.divider()
view_type = st.sidebar.radio("📊 데이터 표기 기준 선택", ["학생 수 기준", "캠퍼스별 비율 기준 (100%)"])

if exam_type == "학기 평가 (TT)":
    current_criteria = TT_CRITERIA
    current_max_scores = TT_MAX_SCORES
else:
    current_criteria = MT_CRITERIA
    current_max_scores = MT_MAX_SCORES

st.info(f"💡 현재 [{exam_type}] + [{view_type}] 조건으로 대시보드가 구동 중입니다.")

# 파일 업로드 영역
st.divider()
st.write("📂 **레벨별 엑셀 파일(들)을 선택하세요 (다중 선택 가능)**")
uploaded_files = st.file_uploader("", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=5)
            if '응시여부' in df.columns:
                df = df[df['응시여부'] == 'Y']
            else:
                continue
            all_data.append(df)
        except Exception as e:
            st.error(f"{file.name} 파일을 읽는 중 오류가 발생했습니다: {e}")
            
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        combined_df['캠퍼스'] = combined_df['캠퍼스'].astype(str).str.strip().str.replace(" ", "")
        combined_df['레벨'] = combined_df['레벨'].astype(str).str.strip()
        
        # Hepta S1/S2 통합 매칭 처리
        combined_df['레벨'] = combined_df['레벨'].str.replace('Hepta S1', 'Hepta 1', case=False)
        combined_df['레벨'] = combined_df['레벨'].str.replace('HeptaS1', 'Hepta 1', case=False)
        combined_df['레벨'] = combined_df['레벨'].str.replace('Hepta S2', 'Hepta 2', case=False)
        combined_df['레벨'] = combined_df['레벨'].str.replace('HeptaS2', 'Hepta 2', case=False)
        
        available_subjects = [col for col in combined_df.columns if col in ['Overall', 'Phonics', 'Reading', 'Listening', 'Vocabulary', 'Grammar']]
        
        if not available_subjects:
            st.warning("분석할 수 있는 과목 열을 찾지 못했습니다.")
        else:
            for subj in available_subjects:
                combined_df[subj] = pd.to_numeric(combined_df[subj], errors='coerce').fillna(0)
                combined_df[f'{subj}_구간'] = 'Developing'
                combined_df['temp_converted'] = combined_df[subj]
                
                for lvl in combined_df['레벨'].unique():
                    main_lvl = None
                    for k in current_criteria.keys():
                        if k.lower() in str(lvl).lower():
                            main_lvl = k
                            break
                    
                    if main_lvl and subj in current_criteria[main_lvl]:
                        adv_score, on_score = current_criteria[main_lvl][subj]
                        mask_base = (combined_df['레벨'] == lvl)
                        
                        if main_lvl in current_max_scores and subj in current_max_scores[main_lvl]:
                            max_score = current_max_scores[main_lvl][subj]
                            combined_df.loc[mask_base, 'temp_converted'] = np.round((combined_df.loc[mask_base, subj] / max_score) * 100)
                        
                        mask_adv = mask_base & (combined_df['temp_converted'] >= adv_score)
                        mask_on = mask_base & (combined_df['temp_converted'] >= on_score) & (combined_df['temp_converted'] < adv_score)
                        
                        combined_df.loc[mask_adv, f'{subj}_구간'] = 'Advanced'
                        combined_df.loc[mask_on, f'{subj}_구간'] = 'On'
            
            tab1, tab2 = st.tabs(["📊 요약표", "📈 분포 그래프"])
            
            with tab1:
                selected_col = st.selectbox("결과를 확인할 과목을 선택하세요", available_subjects)
                target_segment_col = f'{selected_col}_구간'
                
                pivot_count = pd.crosstab(index=[combined_df['캠퍼스'], combined_df['레벨']], columns=combined_df[target_segment_col], margins=True, margins_name='합계')
                pivot_combined = pivot_count.copy().astype(str)
                segments = [c for c in ['Advanced', 'On', 'Developing'] if c in pivot_count.columns]
                
                if view_type == "학생 수 기준":
                    for col in segments:
                        pivot_combined[col] = pivot_count[col].astype(str) + "명"
                    if '합계' in pivot_count.columns:
                        pivot_combined['합계'] = pivot_count['합계'].astype(str) + "명"
                else:
                    for col in segments:
                        pct = (pivot_count[col] / pivot_count['합계'] * 100).round(1)
                        pivot_combined[col] = pct.astype(str) + "%"
                    if '합계' in pivot_count.columns:
                        pivot_combined['합계'] = "100.0%"
                
                pivot_combined = pivot_combined.reset_index()
                desired_order = ['캠퍼스', '레벨'] + [c for c in ['Advanced', 'On', 'Developing', '합계'] if c in pivot_combined.columns]
                st.dataframe(pivot_combined[desired_order], use_container_width=True, hide_index=True)
                
                csv = pivot_combined[desired_order].to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 현재 요약표 엑셀(CSV) 다운로드", data=csv, file_name=f"{exam_type}_{selected_col}_요약_{view_type}.csv", mime="text/csv")
                
            with tab2:
                selected_col_graph = st.selectbox("그래프로 볼 과목을 선택하세요", available_subjects, key="graph_col")
                graph_df = combined_df.groupby(['캠퍼스', f'{selected_col_graph}_구간']).size().reset_index(name='학생수')
                graph_df['비율'] = (graph_df['학생수'] / graph_df.groupby('캠퍼스')['학생수'].transform('sum') * 100).round(1)
                graph_df['비율(%)'] = graph_df['비율'].astype(str) + "%"
                
                if view_type == "학생 수 기준":
                    y_axis = '학생수'
                    title_suffix = "학생 수 기준 분포"
                    label_dict = {'학생수': '학생 수(명)'}
                else:
                    y_axis = '비율'
                    title_suffix = "100% 비율 기준 분포"
                    label_dict = {'비율': '비율(%)'}
                
                fig = px.bar(
                    graph_df, x="캠퍼스", y=y_axis, color=f'{selected_col_graph}_구간',
                    title=f"캠퍼스별 {selected_col_graph} {title_suffix}", barmode="stack",
                    color_discrete_map={'Advanced': '#93c5fd', 'On': '#d1d5db', 'Developing': '#fca5a5'},
                    category_orders={f'{selected_col_graph}_구간': ['Advanced', 'On', 'Developing']},
                    labels=label_dict,
                    hover_data={'캠퍼스': True, f'{selected_col_graph}_구간': True, '학생수': True, '비율(%)': True}
                )
                if view_type != "학생 수 기준":
                    fig.update_layout(yaxis_range=[0, 100])
                st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("'응시여부'가 'Y'인 데이터가 없습니다.")
