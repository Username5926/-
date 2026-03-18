import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.util import Inches
import io
import os

# [매핑] 이미지 및 요청 사항 기반 최종 확정
PHASE_MAPPING = {
    "Position": [6, 7, 14, 15, 19, 28], "Personality": [2, 10, 11, 18, 21, 25],
    "Relationship": [3, 4, 20, 22, 27, 29], "Results": [5, 13, 26, 30],
    "Development": [1, 9, 17, 24], "Principles": [8, 12, 16, 23]
}
STRATEGY_MAPPING = {
    "우호성": [1, 9, 17, 24], "동기유발": [2, 10, 18, 25], "자문": [3, 11],
    "협력제휴": [4, 12, 19, 26], "협상거래": [5, 13, 20, 27], "합리적설득": [6, 14, 21, 28],
    "합법화": [7, 15, 22, 29], "강요": [8, 16, 23, 30]
}

def process_automation(df, template_path):
    if not os.path.exists(template_path):
        st.error(f"템플릿 파일({template_path})을 찾을 수 없습니다.")
        return None, None

    prs = Presentation(template_path)
    # 템플릿의 첫 번째 슬라이드 레이아웃 저장
    template_layout = prs.slides[0].slide_layout
    
    refined_results = []

    for idx, row in df.iterrows():
        name = str(row['성함을 작성해주세요.'])
        # 1번 문항부터 30번 문항까지 점수 추출 (인덱스 2번부터 시작)
        scores = row.iloc[2:32].values
        
        # 1. 점수 계산
        p_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in PHASE_MAPPING.items()}
        s_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in STRATEGY_MAPPING.items()}
        
        res_entry = {"성함": name}
        res_entry.update(p_scores)
        res_entry.update(s_scores)
        refined_results.append(res_entry)

        # 2. 슬라이드 처리 (첫 번째는 재사용, 이후는 추가)
        slide = prs.slides[0] if idx == 0 else prs.slides.add_slide(template_layout)

        # 3. 개체 업데이트
        for shape in slide.shapes:
            # 텍스트 치환
            if shape.has_text_frame and "{{NAME}}" in shape.text:
                shape.text = shape.text.replace("{{NAME}}", name)

            # 표(Table) 업데이트
            if shape.name == 'table_phase' and shape.has_table:
                for i, val in enumerate(p_scores.values()):
                    shape.table.cell(i, 1).text = str(val)
            
            if shape.name == 'table_strategy' and shape.has_table:
                for i, val in enumerate(s_scores.values()):
                    shape.table.cell(i+1, 1).text = str(val) # 헤더 제외 row index 조정 가능

            # 차트(Chart) 업데이트
            if shape.name == 'chart_phase' and shape.has_chart:
                chart_data = CategoryChartData()
                chart_data.categories = list(p_scores.keys())
                chart_data.add_series('Score', list(p_scores.values()))
                shape.chart.replace_data(chart_data)

            if shape.name == 'chart_strategy' and shape.has_chart:
                chart_data = CategoryChartData()
                chart_data.categories = list(s_scores.keys())
                chart_data.add_series('Score', list(s_scores.values()))
                shape.chart.replace_data(chart_data)

            # 동그라미(Circle) 이동 - 최댓값/최솟값 강조 예시
            # if shape.name == 'circle1': # 최댓값 강조용
            #     max_key = max(p_scores, key=p_scores.get)
            #     # 여기에 특정 좌표값(left, top)을 계산해서 넣으면 이동합니다.
            #     pass

    return pd.DataFrame(refined_results), prs

# Streamlit UI
st.set_page_config(page_title="HRD 진단 자동화", layout="centered")
st.title("🚀 리더십 영향력 진단 리포트 생성기")

uploaded_file = st.file_uploader("구글 폼 응답 엑셀(XLSX)을 업로드하세요", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # 깃허브의 파일명과 일치해야 합니다 (확장자 주의!)
    template_name = "template.pptx.pptx" 
    
    refined_df, final_ppt = process_automation(df, template_name)

    if final_ppt:
        st.success(f"✅ {len(df)}명의 진단 데이터가 성공적으로 분석되었습니다.")

        col1, col2 = st.columns(2)
        with col1:
            # 엑셀 다운로드
            output_xlsx = io.BytesIO()
            refined_df.to_excel(output_xlsx, index=False)
            st.download_button("📂 정제된 엑셀 다운로드", output_xlsx.getvalue(), "진단결과_요약.xlsx")
        
        with col2:
            # PPT 다운로드
            output_ppt = io.BytesIO()
            final_ppt.save(output_ppt)
            st.download_button("📊 최종 PPT 보고서 다운로드", output_ppt.getvalue(), "리더십_영향력_보고서.pptx")
