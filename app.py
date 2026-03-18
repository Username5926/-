import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
import io
import os
import openpyxl
from copy import deepcopy

# [매핑] 이미지 및 원본 엑셀 수식 기준 (Q1 ~ Q30)
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

def duplicate_slide_safe(pres, index):
    """라이브러리 버전에 상관없이 안전하게 슬라이드를 복제하는 함수"""
    source_slide = pres.slides[index]
    slide_layout = source_slide.slide_layout
    new_slide = pres.slides.add_slide(slide_layout)
    
    for shape in source_slide.shapes:
        # 텍스트 박스, 표, 차트 등 모든 개체를 새 슬라이드로 복제
        new_shape_el = deepcopy(shape.element)
        new_slide.shapes._spctree.append(new_shape_el)
    return new_slide

def process_all(df, ppt_path, excel_path):
    if not os.path.exists(ppt_path) or not os.path.exists(excel_path):
        st.error(f"템플릿 파일이 없습니다. 현재 경로: {os.listdir('.')}")
        return None, None

    # 1. 엑셀 로드 (원본 양식 유지) 
    wb = openpyxl.load_workbook(excel_path)
    template_sheet = wb.active
    
    # 2. PPT 로드 
    prs = Presentation(ppt_path)
    
    for idx, row in df.iterrows():
        # 두 번째 열(Index 1)에서 성함 추출
        name = str(row.iloc[1]) 
        # 세 번째 열(Index 2)부터 30개 문항 점수 추출 [cite: 25, 26, 27, 28, 29, 30]
        scores = row.iloc[2:32].values 
        
        # 역량 및 기술별 점수 계산
        p_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in PHASE_MAPPING.items()}
        s_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in STRATEGY_MAPPING.items()}
        
        # --- 엑셀: 시트 복제 및 데이터 기입 ---
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = name[:30].replace('/', '_') # 시트명 금칙어 처리
        new_sheet['C1'] = name # 성함 입력 (C1 셀 예시) [cite: 1, 5, 11, 17, 23]
        for i, score in enumerate(scores):
            new_sheet.cell(row=4+i, column=3).value = score # C4부터 점수 입력

        # --- PPT: 슬라이드 복제 및 데이터 주입 ---
        slide = prs.slides[0] if idx == 0 else duplicate_slide_safe(prs, 0)

        for shape in slide.shapes:
            # 이름 치환 
            if shape.has_text_frame and "{{NAME}}" in shape.text:
                shape.text = shape.text.replace("{{NAME}}", name)
            
            # table_phase 처리 (헤더 및 데이터 주입)
            if shape.name == 'table_phase' and shape.has_table:
                table = shape.table
                table.cell(0, 0).text, table.cell(0, 1).text = "항목명", "평균점수"
                for i, (k, v) in enumerate(p_scores.items()):
                    if i + 1 < len(table.rows):
                        table.cell(i+1, 0).text, table.cell(i+1, 1).text = k, str(v)

            # table_strategy 처리 (헤더 및 모든 전략 데이터 주입)
            if shape.name == 'table_strategy' and shape.has_table:
                table = shape.table
                table.cell(0, 0).text, table.cell(0, 1).text = "항목명", "평균점수"
                for i, k in enumerate(STRATEGY_LIST):
                    if i + 1 < len(table.rows):
                        table.cell(i+1, 0).text, table.cell(i+1, 1).text = k, str(s_scores[k])

            # 차트 데이터 업데이트
            if shape.name == 'chart_phase' and shape.has_chart:
                chart_data = CategoryChartData()
                chart_data.categories = list(p_scores.keys())
                chart_data.add_series('점수', list(p_scores.values()))
                shape.chart.replace_data(chart_data)

    # 원본 가이드 시트 삭제 후 저장
    wb.remove(template_sheet)
    ex_buf, ppt_buf = io.BytesIO(), io.BytesIO()
    wb.save(ex_buf)
    prs.save(ppt_buf)
    return ex_buf.getvalue(), ppt_buf.getvalue()

# Streamlit UI
st.set_page_config(page_title="진단 자동화 시스템", layout="centered")
st.title("📊 리더십 진단 통합 리포트 생성기")

uploaded_file = st.file_uploader("구글 폼 결과 엑셀(XLSX)을 업로드하세요", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # GitHub에 업로드된 정확한 파일명 지정
    ppt_temp, excel_temp = "template.pptx.pptx", "excel_template.xlsx"
    
    xlsx_out, ppt_out = process_all(df, ppt_temp, excel_temp)
    
    if xlsx_out:
        st.success("🎉 모든 분석이 완료되었습니다. 결과물을 확인하세요!")
        st.download_button("📂 1. 응답자별 시트 엑셀 다운로드", xlsx_out, "진단결과_전체시트.xlsx")
        st.download_button("📊 2. 통합 PPT 보고서 다운로드", ppt_out, "최종진단보고서.pptx")
