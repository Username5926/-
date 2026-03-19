"""
리더십 진단 보고서 자동화 툴 v6
핵심 수정:
  1. st.session_state로 결과 캐싱 → download_button 클릭 시 rerun 문제 해결
  2. pathlib.Path(__file__).parent 로 템플릿 자동 로드 → 경로 문제 해결
  3. PPT: ChartPart 개별 deep-copy로 슬라이드 완전 복제
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import copy, io, zipfile
from pathlib import Path

from pptx import Presentation
from pptx.opc.package import _Relationship
from pptx.opc.packuri import PackURI
from pptx.parts.chart import ChartPart
from pptx.chart.data import ChartData

# ══════════════════════════════════════════════════════════════════
# 0. 경로 설정 (GitHub 루트 = app.py 위치)
# ══════════════════════════════════════════════════════════════════

BASE_DIR = Path(__file__).parent

# ══════════════════════════════════════════════════════════════════
# 1. 매핑 정의 (C열 행번호 기준: row = Q번호 + 3)
# ══════════════════════════════════════════════════════════════════

COMPETENCY_MAP = {
    "Position":     [9, 10, 17, 18, 22, 31],
    "Personality":  [5, 13, 14, 21, 24, 28],
    "Relationship": [6,  7, 23, 25, 30, 32],
    "Results":      [8, 16, 29, 33],
    "Development":  [4, 12, 20, 27],
    "Principles":   [11, 15, 19, 26],
}

SKILL_MAP = {
    "우호성":     [4, 12, 20, 27],
    "동기유발":   [5, 13, 21, 28],
    "자문":       [6, 14],
    "협력제휴":   [7, 15, 22, 29],
    "협상거래":   [8, 16, 23, 30],
    "합리적설득": [9, 17, 24, 31],
    "합법화":     [10, 18, 25, 32],
    "강요":       [11, 19, 26, 33],
}

SOFT_SKILLS = ["우호성", "동기유발", "자문"]
HARD_SKILLS = ["협력제휴", "협상거래", "합리적설득", "합법화", "강요"]

COMP_ROW  = {"Position":4,"Personality":5,"Relationship":6,
              "Results":7,"Development":8,"Principles":9}
SKILL_ROW = {"우호성":12,"동기유발":13,"자문":14,"협력제휴":15,
              "협상거래":16,"합리적설득":17,"합법화":18,"강요":19}

# ══════════════════════════════════════════════════════════════════
# 2. 점수 계산
# ══════════════════════════════════════════════════════════════════

def avg_by_rows(scores: dict, row_list: list) -> float:
    vals = [scores[r - 3] for r in row_list if (r - 3) in scores]
    return round(sum(vals) / len(vals), 2) if vals else 0.0


def compute_person(scores: dict) -> dict:
    competency = {k: avg_by_rows(scores, v) for k, v in COMPETENCY_MAP.items()}
    skill_raw  = {k: avg_by_rows(scores, v) for k, v in SKILL_MAP.items()}
    soft_avg   = round(sum(skill_raw[k] for k in SOFT_SKILLS) / len(SOFT_SKILLS), 2)
    hard_avg   = round(sum(skill_raw[k] for k in HARD_SKILLS) / len(HARD_SKILLS), 2)
    return {"competency": competency, "skill_raw": skill_raw,
            "soft_avg": soft_avg, "hard_avg": hard_avg}

# ══════════════════════════════════════════════════════════════════
# 3. 입력 파싱
# ══════════════════════════════════════════════════════════════════

def parse_response_excel(file) -> list:
    """A열:타임스탬프, B열:성함, C~AF열:Q1~Q30"""
    df = pd.read_excel(file, header=0)
    people = []
    for _, row in df.iterrows():
        name = str(row.iloc[1]).strip()
        if not name or name.lower() in ("nan", ""):
            continue
        scores = {}
        for q in range(1, 31):
            try:
                scores[q] = float(row.iloc[q + 1])
            except Exception:
                scores[q] = 0.0
        people.append({"name": name, "scores": scores})
    return people

# ══════════════════════════════════════════════════════════════════
# 4. 엑셀 출력
# ══════════════════════════════════════════════════════════════════

def _copy_sheet(wb_dest, src_ws, new_title: str):
    ws = wb_dest.create_sheet(title=new_title)
    for cl, cd in src_ws.column_dimensions.items():
        ws.column_dimensions[cl].width = cd.width
    for rn, rd in src_ws.row_dimensions.items():
        ws.row_dimensions[rn].height = rd.height
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
    return ws


def build_excel(people: list, template_bytes: bytes) -> bytes:
    """template_bytes: 미리 읽어둔 bytes"""
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    for person in people:
        src_ws = load_workbook(io.BytesIO(template_bytes)).worksheets[0]
        ws     = _copy_sheet(wb_out, src_ws, person["name"][:31])
        result = compute_person(person["scores"])

        # 성함 (A1 — 병합 A1:C1)
        ws.cell(row=1, column=1).value = person["name"]

        # Q1~Q30 점수 (C4:C33)
        for q in range(1, 31):
            ws.cell(row=q + 3, column=3).value = person["scores"].get(q, 0)

        # 6대 역량 소계(G) / 평균(H)
        for key, row_list in COMPETENCY_MAP.items():
            er  = COMP_ROW[key]
            avg = result["competency"][key]
            g = ws.cell(row=er, column=7); h = ws.cell(row=er, column=8)
            g.value = round(avg * len(row_list), 2); g.number_format = "0.00"
            h.value = avg;                           h.number_format = "0.00"

        # 8대 기술 소계(G) / 평균(H)
        for key, row_list in SKILL_MAP.items():
            er  = SKILL_ROW[key]
            avg = result["skill_raw"][key]
            g = ws.cell(row=er, column=7); h = ws.cell(row=er, column=8)
            g.value = round(avg * len(row_list), 2); g.number_format = "0.00"
            h.value = avg;                           h.number_format = "0.00"

    buf = io.BytesIO()
    wb_out.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════
# 5. PPT 출력
# ══════════════════════════════════════════════════════════════════

def _clone_chart_part(pkg, orig_cp, new_partname: PackURI) -> ChartPart:
    """ChartPart 독립 복제 — 슬라이드 간 공유 방지"""
    new_element = copy.deepcopy(orig_cp._element)
    new_cp = ChartPart(new_partname, orig_cp.content_type, pkg, new_element)
    for rId2, rel2 in orig_cp.rels.items():
        new_cp.rels._rels[rId2] = _Relationship(
            new_cp.partname.baseURI,
            rel2._rId, rel2._reltype, rel2._target_mode, rel2._target
        )
    return new_cp


def _clone_slide(prs: Presentation, src_slide_index: int = 0):
    """슬라이드 완전 복제 (spTree + ChartPart 개별 복제 + rels)"""
    src_slide = prs.slides[src_slide_index]
    src_part  = src_slide.part
    pkg       = prs.part.package

    new_slide = prs.slides.add_slide(src_slide.slide_layout)
    new_part  = new_slide.part

    # 자동 placeholder 제거
    for ph in list(new_slide.placeholders):
        ph._element.getparent().remove(ph._element)

    # shapes 복사
    for child in list(src_slide.shapes._spTree):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        new_slide.shapes._spTree.append(copy.deepcopy(child))

    # 현재 패키지 내 chart 파트 수 파악 (새 이름 충돌 방지)
    chart_count = sum(
        1 for p in pkg.iter_parts()
        if str(p.partname).startswith("/ppt/charts/chart")
    )

    # rels 복사: chart는 개별 복제, 나머지는 공유
    for rId, rel in src_part.rels.items():
        if rId in new_part.rels:
            continue
        if "chart" in rel._reltype:
            chart_count += 1
            new_cp = _clone_chart_part(
                pkg, rel._target,
                PackURI(f"/ppt/charts/chart{chart_count}.xml")
            )
            new_part.rels._rels[rId] = _Relationship(
                new_part.partname.baseURI,
                rel._rId, rel._reltype, rel._target_mode, new_cp
            )
        else:
            new_part.rels._rels[rId] = _Relationship(
                new_part.partname.baseURI,
                rel._rId, rel._reltype, rel._target_mode, rel._target
            )
    return new_slide


def _replace_text(shape, old: str, new_val: str):
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if old in run.text:
                run.text = run.text.replace(old, new_val)


def _set_table_data(shape, rows_data: list):
    tbl = shape.table
    for r_idx, (label, val) in enumerate(rows_data, start=1):
        if r_idx >= len(tbl.rows):
            break
        tbl.cell(r_idx, 0).text = str(label)
        tbl.cell(r_idx, 1).text = f"{val:.2f}" if isinstance(val, float) else str(val)


def _fill_slide(slide, name: str, result: dict):
    competency = result["competency"]
    skill_raw  = result["skill_raw"]
    soft_avg   = result["soft_avg"]
    hard_avg   = result["hard_avg"]

    for shape in slide.shapes:
        _replace_text(shape, "{{NAME}}", name)

        if shape.name == "table_phase" and shape.has_table:
            _set_table_data(shape, list(competency.items()))

        elif shape.name == "table_strategy" and shape.has_table:
            rows = [(k, skill_raw[k]) for k in SOFT_SKILLS]
            rows.append(("소프트스킬 평균", soft_avg))
            rows += [(k, skill_raw[k]) for k in HARD_SKILLS]
            rows.append(("하드스킬 평균", hard_avg))
            _set_table_data(shape, rows)

        elif shape.name == "chart_phase" and shape.has_chart:
            try:
                cd = ChartData()
                cd.categories = list(competency.keys())
                cd.add_series("역량 점수", list(competency.values()))
                shape.chart.replace_data(cd)
            except Exception:
                pass

        elif shape.name == "chart_strategy" and shape.has_chart:
            try:
                # 카테고리: 소프트3 + 소프트평균 + 하드5 + 하드평균
                cats = SOFT_SKILLS + ["소프트 평균"] + HARD_SKILLS + ["하드 평균"]
                vals = ([skill_raw[k] for k in SOFT_SKILLS] + [soft_avg] +
                        [skill_raw[k] for k in HARD_SKILLS] + [hard_avg])
                cd = ChartData()
                cd.categories = cats
                cd.add_series("계열 1", vals)
                shape.chart.replace_data(cd)
            except Exception:
                pass


def build_ppt(people: list, template_bytes: bytes) -> bytes:
    """template_bytes: 미리 읽어둔 bytes"""
    prs = Presentation(io.BytesIO(template_bytes))

    # 원본 슬라이드[0] 기준으로 나머지 복제
    slides = [prs.slides[0]]
    for _ in range(len(people) - 1):
        slides.append(_clone_slide(prs, src_slide_index=0))

    # 데이터 주입
    for slide, person in zip(slides, people):
        _fill_slide(slide, person["name"], compute_person(person["scores"]))

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════
# 6. 템플릿 로드 헬퍼
# ══════════════════════════════════════════════════════════════════

def load_template_bytes(uploaded_file, filename: str) -> bytes:
    """
    업로드 파일 우선, 없으면 app.py와 같은 폴더(GitHub 루트)에서 자동 로드.
    항상 bytes로 반환 → rerun 시 재사용 가능.
    """
    if uploaded_file is not None:
        return uploaded_file.read()
    path = BASE_DIR / filename
    if path.exists():
        return path.read_bytes()
    return None

# ══════════════════════════════════════════════════════════════════
# 7. Streamlit UI
# ══════════════════════════════════════════════════════════════════

st.set_page_config(page_title="리더십 진단 보고서 자동화", layout="wide")
st.markdown("""
<style>
.main-title{font-size:2rem;font-weight:800;color:#1F3864;}
.sub{font-size:1rem;color:#666;margin-bottom:.5rem;}
.sec{font-size:1.05rem;font-weight:700;color:#2E75B6;margin-top:1.2rem;margin-bottom:.3rem;}
</style>""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 리더십 진단 보고서 자동화 툴</div>', unsafe_allow_html=True)
st.markdown('<div class="sub">구글 폼 응답 엑셀 → 개인별 엑셀(평균 포함) + 응답자별 슬라이드 PPT 자동 생성</div>',
            unsafe_allow_html=True)
st.markdown("---")

st.markdown('<div class="sec">① 구글 폼 응답 엑셀 (필수)</div>', unsafe_allow_html=True)
response_file = st.file_uploader(
    "A열: 타임스탬프 / B열: 성함 / C~AF열: Q1~Q30 점수",
    type=["xlsx", "xls"], key="response"
)

st.markdown('<div class="sec">② 엑셀 템플릿 (선택 — GitHub 루트의 excel_template.xlsx 자동 사용)</div>',
            unsafe_allow_html=True)
excel_tpl_file = st.file_uploader(
    "A1:C1 병합=성함, C4:C33=점수, G/H열=소계/평균 위치인 양식",
    type=["xlsx"], key="excel_tpl"
)

st.markdown('<div class="sec">③ PPT 템플릿 (선택 — GitHub 루트의 template_pptx.pptx 자동 사용)</div>',
            unsafe_allow_html=True)
ppt_tpl_file = st.file_uploader(
    "{{NAME}}, table_phase, table_strategy, chart_phase, chart_strategy 개체 포함",
    type=["pptx"], key="ppt_tpl"
)

st.markdown("---")

# ── 생성 버튼 ──
if st.button("🚀 보고서 생성", type="primary", use_container_width=True):

    # 1. 응답 엑셀 파싱
    if not response_file:
        st.error("❌ 구글 폼 응답 엑셀을 업로드해주세요.")
        st.stop()

    with st.spinner("📂 응답 데이터 파싱 중..."):
        try:
            people = parse_response_excel(response_file)
        except Exception as e:
            st.error(f"❌ 파싱 실패: {e}"); st.stop()

    if not people:
        st.error("❌ 응답자 데이터가 없습니다."); st.stop()

    # 2. 템플릿 bytes 로드 (rerun 안전하게 bytes로 읽어둠)
    excel_bytes_tpl = load_template_bytes(excel_tpl_file, "excel_template.xlsx")
    ppt_bytes_tpl   = load_template_bytes(ppt_tpl_file,   "template_pptx.pptx")

    if excel_bytes_tpl is None:
        st.error("❌ 엑셀 템플릿을 업로드하거나, GitHub 루트에 `excel_template.xlsx`를 추가해주세요.")
        st.stop()
    if ppt_bytes_tpl is None:
        st.error("❌ PPT 템플릿을 업로드하거나, GitHub 루트에 `template_pptx.pptx`를 추가해주세요.")
        st.stop()

    if not excel_tpl_file: st.info("ℹ️ 엑셀 템플릿: GitHub 루트의 excel_template.xlsx 자동 사용")
    if not ppt_tpl_file:   st.info("ℹ️ PPT 템플릿: GitHub 루트의 template_pptx.pptx 자동 사용")

    st.success(f"✅ {len(people)}명 응답 데이터 파싱 완료")

    # 3. 미리보기
    with st.expander(f"📋 응답자 미리보기 ({len(people)}명)"):
        preview = []
        for p in people:
            r = compute_person(p["scores"])
            row = {"성함": p["name"]}
            row.update({k: f"{v:.2f}" for k, v in r["competency"].items()})
            for k in SOFT_SKILLS + HARD_SKILLS:
                row[k] = f"{r['skill_raw'][k]:.2f}"
            row["소프트스킬 평균"] = f"{r['soft_avg']:.2f}"
            row["하드스킬 평균"]   = f"{r['hard_avg']:.2f}"
            preview.append(row)
        st.dataframe(pd.DataFrame(preview), use_container_width=True)

    # 4. 엑셀 생성
    with st.spinner("📊 개인별 엑셀 생성 중..."):
        try:
            excel_out = build_excel(people, excel_bytes_tpl)
        except Exception as e:
            st.error(f"❌ 엑셀 생성 실패: {e}"); st.exception(e); st.stop()

    # 5. PPT 생성
    with st.spinner(f"📑 PPT 생성 중 ({len(people)}슬라이드)..."):
        try:
            ppt_out = build_ppt(people, ppt_bytes_tpl)
        except Exception as e:
            st.error(f"❌ PPT 생성 실패: {e}"); st.exception(e); st.stop()

    # 6. session_state에 저장 (download_button 클릭 후 rerun 시에도 유지)
    st.session_state["excel_out"] = excel_out
    st.session_state["ppt_out"]   = ppt_out
    st.session_state["n_people"]  = len(people)
    st.balloons()

# ── 다운로드 버튼 (session_state에서 읽음 → rerun 안전) ──
if "excel_out" in st.session_state and "ppt_out" in st.session_state:
    excel_out = st.session_state["excel_out"]
    ppt_out   = st.session_state["ppt_out"]
    n         = st.session_state["n_people"]

    st.success(f"🎉 완료! 엑셀 {n}시트 + PPT {n}슬라이드 생성")

    # ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("리더십진단_개인별.xlsx", excel_out)
        zf.writestr("리더십진단_통합.pptx",  ppt_out)

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "⬇️ ZIP (전체)",
            data=zip_buf.getvalue(),
            file_name="리더십진단_결과.zip",
            mime="application/zip",
            use_container_width=True
        )
    with d2:
        st.download_button(
            "⬇️ 엑셀 (개인별)",
            data=excel_out,
            file_name="리더십진단_개인별.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with d3:
        st.download_button(
            "⬇️ PPT (통합)",
            data=ppt_out,
            file_name="리더십진단_통합.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
