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
# table_strategy용 정밀 매핑 (순서 중요)
STRATEGY_LIST = ["우호성", "동기유발", "자문", "협력제휴", "협상거래", "합리적설득", "합법화", "강요"]
STRATEGY_MAPPING = {
    "우호성": [1, 9, 17, 24], "동기유발": [2, 10, 18, 25], "자문": [3, 11],
    "협력제휴": [4, 12, 19, 26], "협상거래": [5, 13, 20, 27], "합리적설득": [6, 14, 21, 28],
    "합법화": [7, 15, 22, 29], "강요": [8, 16, 23, 30]
}

def duplicate_slide(pres, index):
    """템플릿 슬라이드를 복제하는 함수"""
    template = pres.slides[index]
    blank_slide_layout = pres.slide_layouts[6] # 빈 레이아웃
    copied_slide = pres.slides.add_slide(blank_slide_layout)
    for shape in template.shapes:
        el = shape.element
        newel = deepcopy(el)
        copied_slide.shapes._spctree.insert_element_before(newel, 'p:extLst')
    return copied_slide

def process_all(df, ppt_path, excel_path):
    # 1. 엑셀 처리 (원본 양식 복제)
    wb = openpyxl.load_workbook(excel_path)
    template_sheet = wb.active
    
    # 2. PPT 처리
    prs = Presentation(ppt_path)
    
    refined_data = []

    for idx, row in df.iterrows():
        name = str(row.iloc[1])
        scores = row.iloc[2:32].values
        
        # 점수 계산
        p_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in PHASE_MAPPING.items()}
        s_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in STRATEGY_MAPPING.items()}
        
        # --- 엑셀 시트 복제 및 데이터 기입 ---
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = name[:30]
        new_sheet['C1'] = name # 상단 성함1 (위치에 맞춰 수정 가능) 
        # 문항 점수 기입 (C4부터 C33까지 순서대로) 
        for i, score in enumerate(scores):
            new_sheet.cell(row=4+i, column=3).value = score

        # --- PPT 슬라이드 복제 및 데이터 기입 ---
        if idx == 0:
            slide = prs.slides[0] # 첫 사람은 기존 슬라이드 사용
        else:
            slide = duplicate_slide(prs, 0) # 두 번째부터 무한 복제

        for shape in slide.shapes:
            if shape.has_text_frame and "{{NAME}}" in shape.text:
                shape.text = shape.text.replace("{{NAME}}", name)
            
            # table_phase 처리
            if shape.name == 'table_phase' and shape.has_table:
                table = shape.table
                table.cell(0, 0).text, table.cell(0, 1).text = "항목명", "평균점수"
                for i, (k, v) in enumerate(p_scores.items()):
                    table.cell(i+1, 0).text = k
                    table.cell(i+1, 1).text = str(v)

            # table_strategy 처리 (우호성~강요 + 평균)
            if shape.name == 'table_strategy' and shape.has_table:
                table = shape.table
                table.cell(0, 0).text, table.cell(0, 1).text = "항목명", "평균점수"
                # 요청하신 순서대로 데이터 주입 (중간 평균 계산 포함 가능)
                for i, k in enumerate(STRATEGY_LIST):
                    table.cell(i+1, 0).text = k
                    table.cell(i+1, 1).text = str(s_scores[k])

            # 차트 업데이트 로직 생략(위와 동일)

    # 원본 템플릿 시트 삭제 후 저장
    wb.remove(template_sheet)
    ex_buffer = io.BytesIO()
    wb.save(ex_buffer)
    
    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    
    return ex_buffer.getvalue(), ppt_buffer.getvalue()

# Streamlit UI
st.title("🚗 넥센타이어 진단 결과 통합 생성기")
uploaded_file = st.file_uploader("구글 폼 결과 엑셀 업로드", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # 깃허브에 올릴 파일명들
    ppt_temp = "template.pptx.pptx"
    excel_temp = "excel_template.xlsx"
    
    xlsx_out, ppt_out = process_all(df, ppt_temp, excel_temp)
    
    st.success("분석 완료!")
    st.download_button("📂 1. 응답자별 시트 엑셀 다운로드", xlsx_out, "진단결과_원본양식통합.xlsx")
    st.download_button("📊 2. 통합 PPT 보고서 다운로드", ppt_out, "최종진단보고서.pptx")
