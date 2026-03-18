import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
import io
import os
import openpyxl
from copy import deepcopy

# [매핑] 이미지 및 원본 엑셀 기준 (Q1 ~ Q30)
PHASE_MAPPING = {
    "Position": [6, 7, 14, 15, 19, 28], "Personality": [2, 10, 11, 18, 21, 25],
    "Relationship": [3, 4, 20, 22, 27, 29], "Results": [5, 13, 26, 30],
    "Development": [1, 9, 17, 24], "Principles": [8, 12, 16, 23]
}
STRATEGY_LIST = ["우호성", "동기유발", "자문", "협력제휴", "협상거래", "합리적설득", "합법화", "강요"]
STRATEGY_MAPPING = {
    "우호성": [1, 9, 17, 24], "동기유발": [2, 10, 18, 25], "자문": [3, 11],
    "협력제휴": [4, 12, 19, 26], "협상거래": [5, 13, 20, 27], "합리적설득": [6, 14, 21, 28],
    "합법화": [7, 15, 22, 29], "강요": [8, 16, 23, 30]
}

def duplicate_slide(pres, index):
    """첫 번째 슬라이드의 레이아웃을 사용하여 완벽하게 복제하는 함수"""
    template = pres.slides[index]
    # 인덱스 6 대신 템플릿 슬라이드가 사용 중인 레이아웃을 그대로 가져옴
    slide_layout = template.slide_layout 
    copied_slide = pres.slides.add_slide(slide_layout)
    
    for shape in template.shapes:
        el = shape.element
        newel = deepcopy(el)
        copied_slide.shapes._spctree.insert_element_before(newel, 'p:extLst')
    return copied_slide

def process_all(df, ppt_path, excel_path):
    if not os.path.exists(ppt_path) or not os.path.exists(excel_path):
        st.error("템플릿 파일이 서버에 없습니다. 파일명을 확인하세요.")
        return None, None

    # 1. 엑셀 로드
    wb = openpyxl.load_workbook(excel_path)
    template_sheet = wb.active
    
    # 2. PPT 로드
    prs = Presentation(ppt_path)
    
    for idx, row in df.iterrows():
        # 열 이름 대신 인덱스 활용하여 에러 방지
        name = str(row.iloc[1])
        scores = row.iloc[2:32].values
        
        # 점수 계산
        p_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in PHASE_MAPPING.items()}
        s_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in STRATEGY_MAPPING.items()}
        
        # --- 엑셀: 시트 복제 및 데이터 기입 ---
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = name[:30].replace('/', '_') # 시트명 금칙어 처리
        new_sheet['C1'] = name # 상단 이름1
        new_sheet['C2'] = name # 상단 이름2 (원본 양식 위치에 맞춰 수정)
        
        for i, score in enumerate(scores):
            new_sheet.cell(row=4+i, column=3).value = score

        # --- PPT: 슬라이드 복제 및 데이터 기입 ---
        slide = prs.slides[0] if idx == 0 else duplicate_slide(prs, 0)

        for shape in slide.shapes:
            if shape.has_text_frame and "{{NAME}}" in shape.text:
                shape.text = shape.text.replace("{{NAME}}", name)
            
            # table_phase 처리 (헤더 포함)
            if shape.name == 'table_phase' and shape.has_table:
                table = shape.table
                table.cell(0, 0).text, table.cell(0, 1).text = "항목명", "평균점수"
                for i, (k, v) in enumerate(p_scores.items()):
                    table.cell(i+1, 0).text = k
                    table.cell(i+1, 1).text = str(v)

            # table_strategy 처리 (헤더 포함 + 모든 항목 기입)
            if shape.name == 'table_strategy' and shape.has_table:
                table = shape.table
                table.cell(0, 0).text, table.cell(0, 1).text = "항목명", "평균점수"
                for i, k in enumerate(STRATEGY_LIST):
                    if i + 1 < len(table.rows): # 표의 행 개수 확인
                        table.cell(i+1, 0).text = k
                        table.cell(i+1, 1).text = str(s_scores[k])

            # 차트 업데이트
            if shape.name == 'chart_phase' and shape.has_chart:
                chart_data = CategoryChartData()
                chart_data.categories = list(p_scores.keys())
                chart_data.add_series('점수', list(p_scores.values()))
                shape.chart.replace_data(chart_data)

    # 원본 가이드 시트 삭제
    wb.remove(template_sheet)
    
    ex_buffer = io.BytesIO()
    wb.save(ex_buffer)
    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    
    return ex_buffer.getvalue(), ppt_buffer.getvalue()

# Streamlit UI
st.title("🚗 넥센타이어 진단 결과 자동 생성기")
uploaded_file = st.file_uploader("구글 폼 결과 엑셀 업로드", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # 알려주신 파일명으로 최종 세팅
    ppt_temp = "template.pptx.pptx"
    excel_temp = "excel_template.xlsx"
    
    xlsx_out, ppt_out = process_all(df, ppt_temp, excel_temp)
    
    if xlsx_out:
        st.success("🎉 모든 데이터 분석 및 리포트 생성이 완료되었습니다!")
        st.download_button("📂 1. 응답자별 시트 엑셀 다운로드", xlsx_out, "진단결과_전체시트.xlsx")
        st.download_button("📊 2. 통합 PPT 보고서 다운로드", ppt_out, "최종진단보고서.pptx")
