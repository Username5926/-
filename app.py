"""
리더십 진단 보고서 자동화 툴
- 버그 수정: _copy_sheet가 return ws를 올바르게 반환하도록 수정
- 신규 기능: 엑셀/PPT 템플릿 미업로드 시 코드로 기본 템플릿 자동 생성
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import copy
import io
import zipfile

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE

# ══════════════════════════════════════════════════════════════════
# 1. 매핑 정의
# ══════════════════════════════════════════════════════════════════

COMPETENCY_MAP = {
    "Position":     [6, 7, 14, 15, 19, 28],
    "Personality":  [2, 10, 11, 18, 21, 25],
    "Relationship": [3, 4, 20, 22, 27, 29],
    "Results":      [5, 13, 26, 30],
    "Development":  [1, 9, 17, 24],
    "Principles":   [8, 12, 16, 23],
}

SKILL_MAP = {
    "우호성":     [1, 9, 17, 24],
    "동기유발":   [2, 10, 18, 25],
    "자문":       [3, 11],
    "협력제휴":   [4, 12, 19, 26],
    "협상거래":   [5, 13, 20, 27],
    "합리적설득": [6, 14, 21, 28],
    "합법화":     [7, 15, 22, 29],
    "강요":       [8, 16, 23, 30],
}

SOFT_SKILLS = ["우호성", "동기유발", "자문"]
HARD_SKILLS = ["협력제휴", "협상거래", "합리적설득", "합법화", "강요"]

# ══════════════════════════════════════════════════════════════════
# 2. 점수 계산
# ══════════════════════════════════════════════════════════════════

def calc_avg(scores: dict, q_list: list) -> float:
    vals = [scores[q] for q in q_list if q in scores]
    return round(sum(vals) / len(vals), 2) if vals else 0.0


def compute_person(scores: dict) -> dict:
    competency = {k: calc_avg(scores, v) for k, v in COMPETENCY_MAP.items()}
    skill_raw  = {k: calc_avg(scores, v) for k, v in SKILL_MAP.items()}
    soft_avg   = round(sum(skill_raw[k] for k in SOFT_SKILLS) / len(SOFT_SKILLS), 2)
    hard_avg   = round(sum(skill_raw[k] for k in HARD_SKILLS) / len(HARD_SKILLS), 2)
    return {"competency": competency, "skill_raw": skill_raw,
            "soft_avg": soft_avg, "hard_avg": hard_avg}

# ══════════════════════════════════════════════════════════════════
# 3. 입력 파싱
# ══════════════════════════════════════════════════════════════════

def parse_response_excel(file) -> list:
    """
    A열(idx 0): 타임스탬프, B열(idx 1): 성함, C열(idx 2)~: Q1~Q30
    Q = C열 인덱스 - 3 규칙: C열=엑셀 3번째 열(1-based) → Q1=C열(0-based idx 2)
    즉 col_idx(0-based) = Q + 1
    """
    df = pd.read_excel(file, header=0)
    people = []
    for _, row in df.iterrows():
        name = str(row.iloc[1]).strip()
        if not name or name.lower() in ("nan", ""):
            continue
        scores = {}
        for q in range(1, 31):
            col_idx = q + 1        # Q1 → iloc[2] = C열 (0-based)
            try:
                scores[q] = float(row.iloc[col_idx])
            except Exception:
                scores[q] = 0.0
        people.append({"name": name, "scores": scores})
    return people

# ══════════════════════════════════════════════════════════════════
# 4-A. 엑셀 템플릿 자동 생성
# ══════════════════════════════════════════════════════════════════

def make_default_excel_template() -> io.BytesIO:
    """디자인이 적용된 기본 엑셀 템플릿을 코드로 생성."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Template"

    DARK_BLUE  = "1F3864"
    MID_BLUE   = "2E75B6"
    LIGHT_BLUE = "BDD7EE"
    ACCENT     = "ED7D31"
    WHITE      = "FFFFFF"
    GRAY_BG    = "F2F2F2"
    YELLOW_BG  = "FFF2CC"

    thin = Side(border_style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr(ws, coord, val, bg=DARK_BLUE, fg=WHITE, bold=True, size=11, align="center"):
        c = ws[coord]
        c.value = val
        c.font = Font(bold=bold, color=fg, size=size, name="맑은 고딕")
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        c.border = border
        return c

    def dat(ws, coord, val, bg=WHITE, fg="000000", bold=False, align="center"):
        c = ws[coord]
        c.value = val
        c.font = Font(bold=bold, color=fg, size=10, name="맑은 고딕")
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border = border
        return c

    # 열 너비
    for col, w in {"A":12,"B":8,"C":16,"D":2,"E":18,"F":12,"G":2,"H":18,"I":12}.items():
        ws.column_dimensions[col].width = w

    # 행 높이
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 8
    ws.row_dimensions[3].height = 22
    for r in range(4, 35):
        ws.row_dimensions[r].height = 18

    # 타이틀
    hdr(ws, "A1", "리더십 진단 보고서", bg=DARK_BLUE, size=14)
    hdr(ws, "B1", "응답자", bg=MID_BLUE, size=11)
    c = ws["C1"]
    c.value = ""
    c.font = Font(bold=True, size=12, color=DARK_BLUE, name="맑은 고딕")
    c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = border

    # 문항 헤더
    hdr(ws, "A3", "문항번호", bg=MID_BLUE)
    hdr(ws, "B3", "Q#",      bg=MID_BLUE)
    hdr(ws, "C3", "점수",    bg=MID_BLUE)

    # Q1~Q30
    for q in range(1, 31):
        row = q + 3
        bg = GRAY_BG if q % 2 == 0 else WHITE
        dat(ws, f"A{row}", f"문항 {q}", bg=bg)
        dat(ws, f"B{row}", q,            bg=bg)
        sc = ws[f"C{row}"]
        sc.value = None
        sc.font = Font(size=10, name="맑은 고딕")
        sc.fill = PatternFill("solid", fgColor=YELLOW_BG)
        sc.alignment = Alignment(horizontal="center", vertical="center")
        sc.border = border

    # 6대 역량
    hdr(ws, "E3", "6대 역량",  bg=DARK_BLUE)
    hdr(ws, "F3", "평균점수",  bg=DARK_BLUE)
    for i, (key, qlist) in enumerate(COMPETENCY_MAP.items()):
        row = i + 4
        bg = GRAY_BG if i % 2 == 0 else WHITE
        dat(ws, f"E{row}", key, bg=bg, align="left")
        refs = ",".join([f"C{q+3}" for q in qlist])
        fc = ws[f"F{row}"]
        fc.value = f"=IFERROR(AVERAGE({refs}),0)"
        fc.font = Font(size=10, name="맑은 고딕")
        fc.fill = PatternFill("solid", fgColor=bg)
        fc.alignment = Alignment(horizontal="center", vertical="center")
        fc.number_format = "0.00"
        fc.border = border

    # 8대 기술
    hdr(ws, "H3", "8대 기술",  bg=ACCENT)
    hdr(ws, "I3", "평균점수",  bg=ACCENT)
    skill_keys = list(SKILL_MAP.keys())
    for i, (key, qlist) in enumerate(SKILL_MAP.items()):
        row = i + 4
        is_soft = key in SOFT_SKILLS
        bg = "E2EFDA" if is_soft else (GRAY_BG if i % 2 == 0 else WHITE)
        dat(ws, f"H{row}", key, bg=bg, align="left")
        refs = ",".join([f"C{q+3}" for q in qlist])
        fc = ws[f"I{row}"]
        fc.value = f"=IFERROR(AVERAGE({refs}),0)"
        fc.font = Font(size=10, name="맑은 고딕")
        fc.fill = PatternFill("solid", fgColor=bg)
        fc.alignment = Alignment(horizontal="center", vertical="center")
        fc.number_format = "0.00"
        fc.border = border

    # 소프트스킬 평균
    sr = len(skill_keys) + 4
    hdr(ws, f"H{sr}", "소프트스킬 평균", bg="70AD47", size=10)
    soft_refs = ",".join([f"I{skill_keys.index(k)+4}" for k in SOFT_SKILLS])
    fc = ws[f"I{sr}"]
    fc.value = f"=IFERROR(AVERAGE({soft_refs}),0)"
    fc.font = Font(bold=True, size=10, name="맑은 고딕")
    fc.fill = PatternFill("solid", fgColor="A9D18E")
    fc.alignment = Alignment(horizontal="center", vertical="center")
    fc.number_format = "0.00"
    fc.border = border

    # 하드스킬 평균
    hr_ = sr + 1
    hdr(ws, f"H{hr_}", "하드스킬 평균", bg=ACCENT, size=10)
    hard_refs = ",".join([f"I{skill_keys.index(k)+4}" for k in HARD_SKILLS])
    fc = ws[f"I{hr_}"]
    fc.value = f"=IFERROR(AVERAGE({hard_refs}),0)"
    fc.font = Font(bold=True, size=10, name="맑은 고딕")
    fc.fill = PatternFill("solid", fgColor="F4B183")
    fc.alignment = Alignment(horizontal="center", vertical="center")
    fc.number_format = "0.00"
    fc.border = border

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ══════════════════════════════════════════════════════════════════
# 4-B. 엑셀 출력 생성
# ══════════════════════════════════════════════════════════════════

def _copy_sheet(wb_dest, src_ws, new_title: str):
    """워크시트를 다른 워크북으로 안전하게 복사. ws 반환 필수."""
    ws = wb_dest.create_sheet(title=new_title)

    for col_letter, cd in src_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = cd.width
    for row_num, rd in src_ws.row_dimensions.items():
        ws.row_dimensions[row_num].height = rd.height

    for row in src_ws.iter_rows():
        for cell in row:
            nc = ws.cell(row=cell.row, column=cell.column)
            nc.value = cell.value
            if cell.has_style:
                nc.font          = copy.copy(cell.font)
                nc.border        = copy.copy(cell.border)
                nc.fill          = copy.copy(cell.fill)
                nc.number_format = cell.number_format
                nc.protection    = copy.copy(cell.protection)
                nc.alignment     = copy.copy(cell.alignment)

    for merge in src_ws.merged_cells.ranges:
        ws.merge_cells(str(merge))

    return ws    # ← 핵심: 반드시 return


def build_excel(people: list, template_src) -> bytes:
    """C1=성함, C4:C33=Q1~Q30 점수를 각 응답자 시트에 주입."""
    if hasattr(template_src, "read"):
        raw = template_src.read()
        template_src = io.BytesIO(raw)

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    for person in people:
        template_src.seek(0)
        wb_tpl = load_workbook(template_src)
        src_ws = wb_tpl.worksheets[0]

        safe_name = person["name"][:31]
        new_ws = _copy_sheet(wb_out, src_ws, safe_name)   # 반환값 받음

        # 성함 주입 (C1)
        new_ws.cell(row=1, column=3).value = person["name"]

        # 점수 주입 (C4:C33 = Q1~Q30)
        for q in range(1, 31):
            new_ws.cell(row=q + 3, column=3).value = person["scores"].get(q, 0)

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf.read()

# ══════════════════════════════════════════════════════════════════
# 5-A. PPT 템플릿 자동 생성
# ══════════════════════════════════════════════════════════════════

def make_default_ppt_template() -> io.BytesIO:
    """기본 PPT 템플릿(개체명 포함)을 코드로 생성."""
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 빈 레이아웃

    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(0xF5, 0xF7, 0xFA)

    def add_txt(slide, l, t, w, h, text, size=12, bold=False,
                color=RGBColor(0,0,0), align=PP_ALIGN.LEFT, name=None):
        box = slide.shapes.add_textbox(l, t, w, h)
        if name:
            box.name = name
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        return box

    # 상단 타이틀 바
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Inches(0.9))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(0x1F, 0x38, 0x64)
    bar.line.fill.background()
    tf = bar.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    run = tf.paragraphs[0].add_run()
    run.text = "  리더십 진단 보고서"
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # {{NAME}} 텍스트박스
    add_txt(slide, Inches(9.5), Inches(0.12), Inches(3.5), Inches(0.65),
            "{{NAME}}", size=18, bold=True,
            color=RGBColor(0xFF, 0xD7, 0x00),
            align=PP_ALIGN.RIGHT, name="name_label")

    # 섹션 레이블
    add_txt(slide, Inches(0.2), Inches(1.0), Inches(4.4), Inches(0.4),
            "6대 역량 (Phase)", size=13, bold=True, color=RGBColor(0x1F,0x38,0x64))
    add_txt(slide, Inches(4.8), Inches(1.0), Inches(4.4), Inches(0.4),
            "8대 영향력 기술 (Strategy)", size=13, bold=True, color=RGBColor(0xED,0x7D,0x31))
    add_txt(slide, Inches(9.5), Inches(1.0), Inches(3.6), Inches(0.4),
            "역량 레이더 차트", size=13, bold=True, color=RGBColor(0x1F,0x38,0x64))

    # ── table_phase (6대 역량: 헤더 1행 + 데이터 6행 = 7행) ──
    tbl_phase_shape = slide.shapes.add_table(
        7, 2, Inches(0.2), Inches(1.45), Inches(4.4), Inches(5.8)
    )
    tbl_phase_shape.name = "table_phase"
    tbl = tbl_phase_shape.table
    for ci, hdr in enumerate(["항목명", "평균점수"]):
        cell = tbl.cell(0, ci)
        cell.text = hdr
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x1F, 0x38, 0x64)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(11)
    for ri, key in enumerate(COMPETENCY_MAP.keys(), start=1):
        c0 = tbl.cell(ri, 0); c0.text = key
        c1 = tbl.cell(ri, 1); c1.text = ""
        bg = RGBColor(0xBD,0xD7,0xEE) if ri % 2 == 0 else RGBColor(0xFF,0xFF,0xFF)
        for ci, c in enumerate([c0, c1]):
            c.fill.solid(); c.fill.fore_color.rgb = bg
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if ci == 1 else PP_ALIGN.LEFT
            if p.runs: p.runs[0].font.size = Pt(10)

    # ── table_strategy (8대 기술: 헤더 1행 + 소프트3 + 소프트평균 + 하드5 + 하드평균 = 11행) ──
    tbl_strat_shape = slide.shapes.add_table(
        11, 2, Inches(4.8), Inches(1.45), Inches(4.4), Inches(5.8)
    )
    tbl_strat_shape.name = "table_strategy"
    tbl2 = tbl_strat_shape.table
    for ci, hdr in enumerate(["항목명", "평균점수"]):
        cell = tbl2.cell(0, ci)
        cell.text = hdr
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0xED, 0x7D, 0x31)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0]
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(11)
    strat_rows = (SOFT_SKILLS + ["소프트스킬 평균"] + HARD_SKILLS + ["하드스킬 평균"])
    for ri, label in enumerate(strat_rows, start=1):
        is_avg  = "평균" in label
        is_soft = label in SOFT_SKILLS or label == "소프트스킬 평균"
        c0 = tbl2.cell(ri, 0); c0.text = label
        c1 = tbl2.cell(ri, 1); c1.text = ""
        if is_avg:
            bg = RGBColor(0xA9,0xD1,0x8E) if is_soft else RGBColor(0xF4,0xB1,0x83)
        elif is_soft:
            bg = RGBColor(0xE2,0xEF,0xDA)
        else:
            bg = RGBColor(0xF2,0xF2,0xF2) if ri % 2 == 0 else RGBColor(0xFF,0xFF,0xFF)
        for ci, c in enumerate([c0, c1]):
            c.fill.solid(); c.fill.fore_color.rgb = bg
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if ci == 1 else PP_ALIGN.LEFT
            if p.runs:
                p.runs[0].font.size = Pt(10)
                if is_avg: p.runs[0].font.bold = True

    # ── chart_phase (레이더 차트) ──
    chart_data = ChartData()
    chart_data.categories = list(COMPETENCY_MAP.keys())
    chart_data.add_series("역량 점수", [3.0] * 6)
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.RADAR_FILLED,
        Inches(9.5), Inches(1.45), Inches(3.6), Inches(5.8),
        chart_data
    )
    chart_shape.name = "chart_phase"

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf

# ══════════════════════════════════════════════════════════════════
# 5-B. PPT 출력 생성
# ══════════════════════════════════════════════════════════════════

def _replace_text_in_shape(shape, old: str, new_val: str):
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new_val)


def _set_table_data(shape, rows_data: list):
    """1행부터 데이터 주입 (0행=헤더는 템플릿 그대로 유지)."""
    tbl = shape.table
    for r_idx, (label, val) in enumerate(rows_data, start=1):
        if r_idx >= len(tbl.rows):
            break
        tbl.cell(r_idx, 0).text = str(label)
        tbl.cell(r_idx, 1).text = f"{val:.2f}" if isinstance(val, float) else str(val)


def _update_chart(shape, competency: dict):
    try:
        cd = ChartData()
        cd.categories = list(competency.keys())
        cd.add_series("역량 점수", list(competency.values()))
        shape.chart.replace_data(cd)
    except Exception:
        pass


def _clone_slide_safe(prs: Presentation, src_xml_snapshot):
    """원본 슬라이드 XML snapshot으로 새 슬라이드 복제 (_spTree 직접 참조 없음)."""
    layout    = prs.slides[0].slide_layout
    new_slide = prs.slides.add_slide(layout)

    for ph in list(new_slide.placeholders):
        ph._element.getparent().remove(ph._element)

    sp_trees = src_xml_snapshot.xpath("//*[local-name()='spTree']")
    if not sp_trees:
        return new_slide

    src_sp_tree = sp_trees[0]
    dst_sp_tree = new_slide.shapes._spTree

    for child in list(src_sp_tree):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        dst_sp_tree.append(copy.deepcopy(child))

    return new_slide


def _fill_slide(slide, name: str, result: dict):
    competency = result["competency"]
    skill_raw  = result["skill_raw"]
    soft_avg   = result["soft_avg"]
    hard_avg   = result["hard_avg"]

    for shape in slide.shapes:
        sname = shape.name.lower()

        _replace_text_in_shape(shape, "{{NAME}}", name)
        _replace_text_in_shape(shape, "{{name}}", name)

        if "table_phase" in sname and shape.has_table:
            _set_table_data(shape, [(k, v) for k, v in competency.items()])

        elif "table_strategy" in sname and shape.has_table:
            rows = [(k, skill_raw[k]) for k in SOFT_SKILLS]
            rows.append(("소프트스킬 평균", soft_avg))
            rows += [(k, skill_raw[k]) for k in HARD_SKILLS]
            rows.append(("하드스킬 평균", hard_avg))
            _set_table_data(shape, rows)

        elif "chart_phase" in sname and shape.has_chart:
            _update_chart(shape, competency)


def build_ppt(people: list, template_src) -> bytes:
    if hasattr(template_src, "read"):
        raw = template_src.read()
        template_src = io.BytesIO(raw)

    # 원본 슬라이드 XML snapshot
    template_src.seek(0)
    prs_ref = Presentation(template_src)
    src_xml_snapshot = copy.deepcopy(prs_ref.slides[0]._element)

    template_src.seek(0)
    prs = Presentation(template_src)

    for i, person in enumerate(people):
        result = compute_person(person["scores"])
        slide  = prs.slides[0] if i == 0 else _clone_slide_safe(prs, src_xml_snapshot)
        _fill_slide(slide, person["name"], result)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()

# ══════════════════════════════════════════════════════════════════
# 6. Streamlit UI
# ══════════════════════════════════════════════════════════════════

st.set_page_config(page_title="리더십 진단 보고서 자동화", layout="wide")

st.markdown("""
<style>
.main-title{font-size:2rem;font-weight:800;color:#1F3864;}
.sub{font-size:1rem;color:#555;margin-bottom:.5rem;}
.sec{font-size:1.05rem;font-weight:700;color:#2E75B6;margin-top:1rem;margin-bottom:.3rem;}
</style>""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 리더십 진단 보고서 자동화 툴</div>', unsafe_allow_html=True)
st.markdown('<div class="sub">구글 폼 응답 엑셀 → 개인별 엑셀 + 통합 PPT 자동 생성</div>', unsafe_allow_html=True)
st.markdown("---")

# ── 파일 업로드 ──
st.markdown('<div class="sec">① 구글 폼 응답 엑셀 (필수)</div>', unsafe_allow_html=True)
response_file = st.file_uploader(
    "A열: 타임스탬프 / B열: 성함 / C~AF열: Q1~Q30 점수",
    type=["xlsx", "xls"], key="response"
)

st.markdown('<div class="sec">② 엑셀 템플릿 (선택 – 미업로드 시 자동 생성)</div>', unsafe_allow_html=True)
excel_tpl_file = st.file_uploader(
    "C1=성함 위치, C4:C33=점수 위치인 디자인 완료 양식",
    type=["xlsx"], key="excel_tpl"
)

st.markdown('<div class="sec">③ PPT 템플릿 (선택 – 미업로드 시 자동 생성)</div>', unsafe_allow_html=True)
ppt_tpl_file = st.file_uploader(
    "{{NAME}}, table_phase, table_strategy, chart_phase 개체 포함 1슬라이드 파일",
    type=["pptx"], key="ppt_tpl"
)

# ── 기본 템플릿 다운로드 ──
with st.expander("📥 기본 템플릿 다운로드 (편집 후 업로드 가능)"):
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "⬇️ 기본 엑셀 템플릿",
            data=make_default_excel_template(),
            file_name="excel_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with c2:
        st.download_button(
            "⬇️ 기본 PPT 템플릿",
            data=make_default_ppt_template(),
            file_name="template.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

st.markdown("---")

# ── 생성 버튼 ──
if st.button("🚀 보고서 생성", type="primary", use_container_width=True):
    if not response_file:
        st.error("❌ 구글 폼 응답 엑셀을 업로드해주세요.")
        st.stop()

    with st.spinner("📂 응답 데이터 파싱 중..."):
        try:
            people = parse_response_excel(response_file)
        except Exception as e:
            st.error(f"❌ 응답 엑셀 파싱 실패: {e}")
            st.stop()

    if not people:
        st.error("❌ 응답자 데이터가 비어있습니다.")
        st.stop()

    st.success(f"✅ {len(people)}명 응답 데이터 파싱 완료")

    with st.expander(f"📋 응답자 미리보기 ({len(people)}명)"):
        preview = []
        for p in people:
            r = compute_person(p["scores"])
            row = {"성함": p["name"]}
            row.update({k: f"{v:.2f}" for k, v in r["competency"].items()})
            row["소프트스킬 평균"] = f"{r['soft_avg']:.2f}"
            row["하드스킬 평균"]   = f"{r['hard_avg']:.2f}"
            preview.append(row)
        st.dataframe(pd.DataFrame(preview), use_container_width=True)

    # 템플릿 소스 결정
    excel_src = excel_tpl_file if excel_tpl_file else make_default_excel_template()
    ppt_src   = ppt_tpl_file   if ppt_tpl_file   else make_default_ppt_template()

    if not excel_tpl_file:
        st.info("ℹ️ 엑셀 템플릿 미업로드 → 기본 템플릿 자동 사용")
    if not ppt_tpl_file:
        st.info("ℹ️ PPT 템플릿 미업로드 → 기본 템플릿 자동 사용")

    with st.spinner("📊 개인별 엑셀 생성 중..."):
        try:
            excel_bytes = build_excel(people, excel_src)
        except Exception as e:
            st.error(f"❌ 엑셀 생성 실패: {e}")
            st.exception(e)
            st.stop()

    with st.spinner("📑 통합 PPT 생성 중..."):
        try:
            ppt_bytes = build_ppt(people, ppt_src)
        except Exception as e:
            st.error(f"❌ PPT 생성 실패: {e}")
            st.exception(e)
            st.stop()

    # ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("리더십진단_개인별.xlsx", excel_bytes)
        zf.writestr("리더십진단_통합.pptx",  ppt_bytes)
    zip_buf.seek(0)

    st.balloons()
    st.success("🎉 보고서 생성 완료!")

    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.download_button("⬇️ ZIP (전체)", data=zip_buf,
            file_name="리더십진단_결과.zip", mime="application/zip",
            use_container_width=True)
    with dl2:
        st.download_button("⬇️ 엑셀 (개인별)", data=excel_bytes,
            file_name="리더십진단_개인별.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with dl3:
        st.download_button("⬇️ PPT (통합)", data=ppt_bytes,
            file_name="리더십진단_통합.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True)
