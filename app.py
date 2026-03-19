import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import copy
import io
import zipfile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import lxml.etree as etree

# ─────────────────────────────────────────────
# 1. 매핑 정의
# ─────────────────────────────────────────────
COMPETENCY_MAP = {
    "Position":     [6, 7, 14, 15, 19, 28],
    "Personality":  [2, 10, 11, 18, 21, 25],
    "Relationship": [3, 4, 20, 22, 27, 29],
    "Results":      [5, 13, 26, 30],
    "Development":  [1, 9, 17, 24],
    "Principles":   [8, 12, 16, 23],
}

SKILL_MAP = {
    "우호성":      [1, 9, 17, 24],
    "동기유발":    [2, 10, 18, 25],
    "자문":        [3, 11],
    "협력제휴":    [4, 12, 19, 26],
    "협상거래":    [5, 13, 20, 27],
    "합리적설득":  [6, 14, 21, 28],
    "합법화":      [7, 15, 22, 29],
    "강요":        [8, 16, 23, 30],
}

# 소프트 스킬 (우호성, 동기유발, 자문) → 별도 평균
SOFT_SKILLS = ["우호성", "동기유발", "자문"]
# 나머지 5가지
HARD_SKILLS = ["협력제휴", "협상거래", "합리적설득", "합법화", "강요"]


# ─────────────────────────────────────────────
# 2. 점수 계산 헬퍼
# ─────────────────────────────────────────────
def calc_avg(scores: dict, q_list: list) -> float:
    vals = [scores[q] for q in q_list if q in scores]
    return round(sum(vals) / len(vals), 2) if vals else 0.0


def compute_person(scores: dict) -> dict:
    """scores: {1: val, 2: val, ..., 30: val}"""
    competency = {k: calc_avg(scores, v) for k, v in COMPETENCY_MAP.items()}
    skill_raw   = {k: calc_avg(scores, v) for k, v in SKILL_MAP.items()}

    soft_avg = round(
        sum(skill_raw[k] for k in SOFT_SKILLS) / len(SOFT_SKILLS), 2
    )
    hard_avg = round(
        sum(skill_raw[k] for k in HARD_SKILLS) / len(HARD_SKILLS), 2
    )

    return {
        "competency": competency,
        "skill_raw":  skill_raw,
        "soft_avg":   soft_avg,
        "hard_avg":   hard_avg,
    }


# ─────────────────────────────────────────────
# 3. 입력 데이터 파싱
# ─────────────────────────────────────────────
def parse_response_excel(file) -> list[dict]:
    """구글 폼 응답 엑셀 파싱. 반환: [{name, scores:{1..30}}, ...]"""
    df = pd.read_excel(file, header=0)
    people = []
    for _, row in df.iterrows():
        name = str(row.iloc[1]).strip()
        scores = {}
        for q in range(1, 31):
            col_idx = q + 2          # C열 = index 2 → Q1=index 2, Q2=index 3 ...
            # Q = C열 인덱스 - 3 역변환: col_idx = Q + 2
            try:
                scores[q] = float(row.iloc[col_idx])
            except Exception:
                scores[q] = 0.0
        people.append({"name": name, "scores": scores})
    return people


# ─────────────────────────────────────────────
# 4. Excel 출력 생성
# ─────────────────────────────────────────────
def build_excel(people: list[dict], template_file) -> bytes:
    """
    excel_template.xlsx 의 첫 번째 시트를 응답자 수만큼 복제.
    C1 = 성함, C4:C33 = Q1..Q30 점수
    """
    wb_tpl = load_workbook(template_file)
    tpl_sheet = wb_tpl.worksheets[0]
    tpl_name  = tpl_sheet.title

    # 새 워크북
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)   # 기본 시트 제거

    for person in people:
        # 템플릿 시트를 deep copy
        wb_tpl_copy = load_workbook(template_file)
        src = wb_tpl_copy.worksheets[0]

        safe_name = person["name"][:31]   # 시트명 최대 31자
        new_ws = wb_out.copy_worksheet(src) if False else _copy_sheet(wb_out, src, safe_name)

        # 데이터 주입
        new_ws["C1"] = person["name"]
        for q in range(1, 31):
            row = q + 3          # C4 = Q1, C5 = Q2, ...
            new_ws.cell(row=row, column=3).value = person["scores"].get(q, 0)

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf.read()


def _copy_sheet(wb_dest, src_ws, new_title: str):
    """openpyxl 워크시트를 다른 워크북으로 안전하게 복사."""
    ws = wb_dest.create_sheet(title=new_title)

    # 열 너비
    for col_letter, cd in src_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = cd.width

    # 행 높이
    for row_num, rd in src_ws.row_dimensions.items():
        ws.row_dimensions[row_num].height = rd.height

    # 셀 값 + 스타일
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = ws.cell(row=cell.row, column=cell.column)
            new_cell.value = cell.value
            if cell.has_style:
                new_cell.font      = copy.copy(cell.font)
                new_cell.border    = copy.copy(cell.border)
                new_cell.fill      = copy.copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection    = copy.copy(cell.protection)
                new_cell.alignment     = copy.copy(cell.alignment)

    # 병합 셀
    for merge in src_ws.merged_cells.ranges:
        ws.merge_cells(str(merge))

    return ws


# ─────────────────────────────────────────────
# 5. PowerPoint 출력 생성
# ─────────────────────────────────────────────
def _clone_slide(prs: Presentation, src_slide_idx: int = 0):
    """
    템플릿의 첫 번째 슬라이드를 안전하게 복제하여 새 슬라이드 반환.
    _spctree를 직접 건드리지 않고 lxml deep-copy를 이용.
    """
    template_slide = prs.slides[src_slide_idx]
    layout = template_slide.slide_layout

    new_slide = prs.slides.add_slide(layout)

    # 레이아웃에서 자동 추가된 placeholder 제거
    for ph in new_slide.placeholders:
        sp = ph._element
        sp.getparent().remove(sp)

    # 원본 슬라이드 XML 트리에서 spTree 내부 모든 요소를 deep-copy
    src_sp_tree = template_slide.shapes._spTree
    dst_sp_tree = new_slide.shapes._spTree

    for child in list(src_sp_tree):
        # nvGrpSpPr, grpSpPr 는 이미 존재하므로 건너뜀
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        dst_sp_tree.append(copy.deepcopy(child))

    return new_slide


def _set_table_data(shape, headers: list[str], rows: list[tuple]):
    """
    표 shape에 0행: 헤더, 1행~: 데이터 주입.
    rows = [(항목명, 평균점수), ...]
    """
    table = shape.table
    needed_rows = 1 + len(rows)   # header + data

    # 필요한 행 수 확인 (템플릿이 충분한 행을 가져야 함)
    for col_idx, header in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = header

    for r_idx, (label, val) in enumerate(rows, start=1):
        if r_idx < len(table.rows):
            table.cell(r_idx, 0).text = str(label)
            table.cell(r_idx, 1).text = f"{val:.2f}"


def _replace_text_in_shape(shape, old: str, new: str):
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)


def build_ppt(people: list[dict], template_file) -> bytes:
    prs = Presentation(template_file)

    # 첫 번째 슬라이드를 템플릿으로 사용
    # 이미 슬라이드 1개(idx=0)가 있으므로, 나머지를 추가 후 맨 앞 제거
    first_slide_idx = 0

    slides_to_fill = []

    # 첫 번째 사람은 기존 슬라이드 사용, 나머지는 복제
    for i, person in enumerate(people):
        if i == 0:
            slide = prs.slides[0]
        else:
            slide = _clone_slide(prs, src_slide_idx=0 if i == 1 else len(prs.slides) - 1)
            # 주의: 복제 기준을 항상 원본 첫 슬라이드로 고정
        slides_to_fill.append((slide, person))

    # 각 슬라이드에 데이터 주입
    # 복제 이후에는 인덱스가 달라지므로 slides_to_fill 활용
    # 단, 첫 번째 사람 이후 복제 시 원본이 변해버리는 문제 → 원본 XML을 미리 저장
    prs2 = Presentation(template_file)
    # 원본 XML snapshot
    src_xml = copy.deepcopy(prs2.slides[0]._element)

    prs3 = Presentation(template_file)

    for i, person in enumerate(people):
        result = compute_person(person["scores"])

        if i == 0:
            slide = prs3.slides[0]
        else:
            # 원본 첫 슬라이드 기준으로 항상 복제
            slide = _clone_slide_from_xml(prs3, src_xml)

        _fill_slide(slide, person["name"], result)

    buf = io.BytesIO()
    prs3.save(buf)
    buf.seek(0)
    return buf.read()


def _clone_slide_from_xml(prs: Presentation, src_slide_xml_element):
    """
    저장된 원본 슬라이드 XML element를 기반으로 새 슬라이드를 복제.
    """
    # 첫 번째 슬라이드 레이아웃 재사용
    layout = prs.slides[0].slide_layout
    new_slide = prs.slides.add_slide(layout)

    # 자동 추가된 placeholder 제거
    for ph in list(new_slide.placeholders):
        ph._element.getparent().remove(ph._element)

    src_sp_tree = src_slide_xml_element.find(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}spTree"
    )
    # pptx namespace
    pptx_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    draw_ns  = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

    # spTree는 p:spTree
    ns_map = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    src_sp_tree = src_slide_xml_element.find(
        ".//{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"
        "/{http://schemas.openxmlformats.org/presentationml/2006/main}spTree"
    )

    if src_sp_tree is None:
        # fallback: python-pptx namespace
        cSld = src_slide_xml_element.find(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"
        )
        if cSld is not None:
            src_sp_tree = cSld.find(
                "{http://schemas.openxmlformats.org/presentationml/2006/main}spTree"
            )

    if src_sp_tree is None:
        # 마지막 fallback: lxml xpath
        results = src_slide_xml_element.xpath(
            "//*[local-name()='spTree']"
        )
        src_sp_tree = results[0] if results else None

    dst_sp_tree = new_slide.shapes._spTree

    if src_sp_tree is not None:
        for child in list(src_sp_tree):
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag in ("nvGrpSpPr", "grpSpPr"):
                continue
            dst_sp_tree.append(copy.deepcopy(child))

    return new_slide


def _fill_slide(slide, name: str, result: dict):
    """슬라이드 내 개체에 데이터 주입."""
    competency = result["competency"]
    skill_raw  = result["skill_raw"]
    soft_avg   = result["soft_avg"]
    hard_avg   = result["hard_avg"]

    for shape in slide.shapes:
        shape_name = shape.name.lower()

        # {{NAME}} 치환
        _replace_text_in_shape(shape, "{{NAME}}", name)
        _replace_text_in_shape(shape, "{{name}}", name)

        # table_phase: 6대 역량 표
        if "table_phase" in shape_name and shape.has_table:
            rows_data = [(k, v) for k, v in competency.items()]
            _set_table_data(shape, ["항목명", "평균점수"], rows_data)

        # table_strategy: 8대 기술 표
        elif "table_strategy" in shape_name and shape.has_table:
            rows_data = []
            for k in SOFT_SKILLS:
                rows_data.append((k, skill_raw[k]))
            rows_data.append(("소프트스킬 평균", soft_avg))
            for k in HARD_SKILLS:
                rows_data.append((k, skill_raw[k]))
            rows_data.append(("하드스킬 평균", hard_avg))
            _set_table_data(shape, ["항목명", "평균점수"], rows_data)

        # chart_phase: 역량 차트 (텍스트 프레임에 값 삽입 또는 차트 업데이트)
        elif "chart_phase" in shape_name:
            if shape.has_chart:
                _update_chart(shape.chart, competency)
            elif shape.has_text_frame:
                summary = "\n".join(
                    [f"{k}: {v:.2f}" for k, v in competency.items()]
                )
                shape.text_frame.text = summary


def _update_chart(chart, competency: dict):
    """python-pptx 차트 데이터 업데이트 (레이더/바 공통 처리)."""
    from pptx.chart.data import ChartData
    try:
        chart_data = ChartData()
        chart_data.categories = list(competency.keys())
        chart_data.add_series("역량 점수", list(competency.values()))
        chart.replace_data(chart_data)
    except Exception:
        pass   # 차트 구조가 맞지 않을 경우 무시


# ─────────────────────────────────────────────
# 6. Streamlit UI
# ─────────────────────────────────────────────
st.set_page_config(page_title="리더십 진단 보고서 자동화", layout="centered")
st.title("📊 리더십 진단 보고서 자동화 툴")
st.markdown("구글 폼 응답 엑셀을 업로드하면 **개인별 엑셀**과 **통합 PPT**를 자동 생성합니다.")

with st.expander("📌 파일 업로드 안내"):
    st.markdown("""
- **구글 폼 응답 엑셀**: A열=타임스탬프, B열=성함, C~AF열=1~30번 문항 점수
- **엑셀 템플릿** (`excel_template.xlsx`): 디자인 완료된 원본 양식
- **PPT 템플릿** (`template.pptx`): `{{NAME}}`, `table_phase`, `table_strategy`, `chart_phase` 포함
""")

col1, col2, col3 = st.columns(3)
with col1:
    response_file  = st.file_uploader("구글 폼 응답 엑셀", type=["xlsx", "xls"], key="response")
with col2:
    excel_template = st.file_uploader("엑셀 템플릿", type=["xlsx"], key="excel_tpl")
with col3:
    ppt_template   = st.file_uploader("PPT 템플릿", type=["pptx"], key="ppt_tpl")

if st.button("🚀 보고서 생성", type="primary", use_container_width=True):
    if not response_file:
        st.error("구글 폼 응답 엑셀을 업로드해주세요.")
        st.stop()
    if not excel_template:
        st.error("엑셀 템플릿을 업로드해주세요.")
        st.stop()
    if not ppt_template:
        st.error("PPT 템플릿을 업로드해주세요.")
        st.stop()

    with st.spinner("데이터 파싱 중..."):
        people = parse_response_excel(response_file)

    st.success(f"✅ {len(people)}명의 응답 데이터를 읽었습니다.")

    # 미리보기
    with st.expander("응답자 목록 미리보기"):
        preview = []
        for p in people:
            result = compute_person(p["scores"])
            row = {"성함": p["name"]}
            row.update({k: f"{v:.2f}" for k, v in result["competency"].items()})
            row["소프트스킬 평균"] = result["soft_avg"]
            row["하드스킬 평균"]   = result["hard_avg"]
            preview.append(row)
        st.dataframe(pd.DataFrame(preview), use_container_width=True)

    with st.spinner("엑셀 생성 중..."):
        excel_bytes = build_excel(people, excel_template)

    with st.spinner("PPT 생성 중..."):
        ppt_bytes = build_ppt(people, ppt_template)

    # ZIP으로 묶어서 제공
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("리더십진단_개인별.xlsx", excel_bytes)
        zf.writestr("리더십진단_통합.pptx",  ppt_bytes)
    zip_buf.seek(0)

    st.balloons()
    st.download_button(
        label="⬇️ 결과 파일 다운로드 (ZIP)",
        data=zip_buf,
        file_name="리더십진단_결과.zip",
        mime="application/zip",
        use_container_width=True,
    )

    col_a, col_b = st.columns(2)
    with col_a:
        st.download_button(
            "⬇️ 엑셀만 다운로드",
            data=excel_bytes,
            file_name="리더십진단_개인별.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_b:
        st.download_button(
            "⬇️ PPT만 다운로드",
            data=ppt_bytes,
            file_name="리더십진단_통합.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
