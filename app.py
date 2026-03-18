import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
import io

# [매핑] 이미지의 빨간색 체크 및 엑셀 수식 기준 (Q = C - 3)
PHASE_MAPPING = {
    "Position": [6, 7, 14, 15, 19, 28],
    "Personality": [2, 10, 11, 18, 21, 25],
    "Relationship": [3, 4, 20, 22, 27, 29],
    "Results": [5, 13, 26, 30],
    "Development": [1, 9, 17, 24],
    "Principles": [8, 12, 16, 23]
}

STRATEGY_MAPPING = {
    "우호성": [1, 9, 17, 24], "동기유발": [2, 10, 18, 25], "자문": [3, 11],
    "협력제휴": [4, 12, 19, 26], "협상거래": [5, 13, 20, 27], "합리적설득": [6, 14, 21, 28],
    "합법화": [7, 15, 22, 29], "강요": [8, 16, 23, 30]
}

def process_ppt(input_df, template_path):
    prs = Presentation(template_path)
    base_slide = prs.slides[0] # 첫 번째 슬라이드를 템플릿으로 사용

    for _, row in input_df.iterrows():
        name = row['성함을 작성해주세요.']
        # 문항 응답 데이터 추출 (데이터 시작 위치에 따라 인덱스 조정 필요)
        scores = row.iloc[2:32].values 
        
        # 새 슬라이드 생성 (템플릿 복제 로직은 라이브러리 특성상 수동 구현이 필요하여 기존 슬라이드 활용 권장)
        slide = prs.slides.add_slide(base_slide.slide_layout)
        
        # 1. 이름 치환 {{NAME}}
        for shape in slide.shapes:
            if shape.has_text_frame and "{{NAME}}" in shape.text:
                shape.text = shape.text.replace("{{NAME}}", name)

        # 2. 표(Table) 데이터 입력
        for shape in slide.shapes:
            if shape.name == 'table_phase':
                for i, (key, q_list) in enumerate(PHASE_MAPPING.items()):
                    val = sum([float(scores[q-1]) for q in q_list]) / len(q_list)
                    shape.table.cell(i, 1).text = str(round(val, 2))
            
            if shape.name == 'table_strategy':
                for i, (key, q_list) in enumerate(STRATEGY_MAPPING.items()):
                    val = sum([float(scores[q-1]) for q in q_list]) / len(q_list)
                    shape.table.cell(i, 1).text = str(round(val, 2))

        # 3. 차트(Chart) 데이터 업데이트
        for shape in slide.shapes:
            if shape.name == 'chart_phase':
                chart_data = CategoryChartData()
                chart_data.categories = list(PHASE_MAPPING.keys())
                vals = [sum([float(scores[q-1]) for q in q_list])/len(q_list) for q_list in PHASE_MAPPING.values()]
                chart_data.add_series('Score', vals)
                shape.chart.replace_data(chart_data)

        # 4. 동그라미(Circle) 제어 (예시: 최고점 항목에 circle1 이동)
        # 실제 구현 시 각 점수 위치의 좌표(left, top)를 파악하여 shape.left/top을 수정합니다.
    
    return prs

# 스트림릿 UI 구현
st.title("📊 리더십 진단 보고서 자동 생성기")
file = st.file_uploader("구글 폼 엑셀 업로드", type=['xlsx'])
if file:
    df = pd.read_excel(file)
    final_prs = process_ppt(df, "template.pptx.pptx") # 템플릿 파일명 확인
    
    output = io.BytesIO()
    final_prs.save(output)
    st.download_button("PPT 리포트 다운로드", output.getvalue(), "진단결과보고서.pptx")
