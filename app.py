import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import copy, io, zipfile, os, sys
from pathlib import Path

from pptx import Presentation
from pptx.opc.package import _Relationship
from pptx.opc.packuri import PackURI
from pptx.parts.chart import ChartPart
from pptx.chart.data import ChartData

# ══════════════════════════════════════════════════════════════════
# 1. 매핑 (C열 행번호 기준: row = Q번호 + 3)
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

def avg_by_rows(scores, row_list):
    vals = [scores[r - 3] for r in row_list if (r - 3) in scores]
    return round(sum(vals) / len(vals), 2) if vals else 0.0

def compute_person(scores):
    competency = {k: avg_by_rows(scores, v) for k, v in COMPETENCY_MAP.items()}
    skill_raw  = {k: avg_by_rows(scores, v) for k, v in SKILL_MAP.items()}
    soft_avg   = round(sum(skill_raw[k] for k in SOFT_SKILLS) / 3, 2)
    hard_avg   = round(sum(skill_raw[k] for k in HARD_SKILLS) / 5, 2)
    return {"competency": competency, "skill_raw": skill_raw,
            "soft_avg": soft_avg, "hard_avg": hard_avg}

# ══════════════════════════════════════════════════════════════════
# 3. 응답 엑셀 파싱
# ══════════════════════════════════════════════════════════════════

def parse_response_excel(raw_bytes: bytes) -> list:
    """bytes로 받아서 파싱 → rerun 시 재사용 가능"""
    df = pd.read_excel(io.BytesIO(raw_bytes), header=0)
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
# 4. 엑셀 생성
# ══════════════════════════════════════════════════════════════════

def _copy_sheet(wb_dest, src_ws, title):
    ws = wb_dest.create_sheet(title=title)
    for cl, cd in src_ws.column_dimensions.items():
        ws.column_dimensions[cl].width = cd.width
    for rn, rd in src_ws.row_dimensions.items():
        ws.row_dimensions[rn].height = rd.height
    for row in src_ws.iter_rows():
        for cell in row:
            nc = ws.cell(row=cell.row, column=cell.column)
            nc.value = cell.value
            if cell.has_style:
                nc.font=copy.copy(cell.font); nc.border=copy.copy(cell.border)
                nc.fill=copy.copy(cell.fill); nc.number_format=cell.number_format
                nc.protection=copy.copy(cell.protection); nc.alignment=copy.copy(cell.alignment)
    for m in src_ws.merged_cells.ranges:
        ws.merge_cells(str(m))
    return ws

def build_excel(people, excel_tpl_bytes):
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    for person in people:
        src_ws = load_workbook(io.BytesIO(excel_tpl_bytes)).worksheets[0]
        ws = _copy_sheet(wb_out, src_ws, person["name"][:31])
        result = compute_person(person["scores"])
        ws.cell(1, 1).value = person["name"]
        for q in range(1, 31):
            ws.cell(q + 3, 3).value = person["scores"].get(q, 0)
        for key, row_list in COMPETENCY_MAP.items():
            er = COMP_ROW[key]; avg = result["competency"][key]
            g = ws.cell(er, 7); h = ws.cell(er, 8)
            g.value = round(avg * len(row_list), 2); g.number_format = "0.00"
            h.value = avg; h.number_format = "0.00"
        for key, row_list in SKILL_MAP.items():
            er = SKILL_ROW[key]; avg = result["skill_raw"][key]
            g = ws.cell(er, 7); h = ws.cell(er, 8)
            g.value = round(avg * len(row_list), 2); g.number_format = "0.00"
            h.value = avg; h.number_format = "0.00"
    buf = io.BytesIO(); wb_out.save(buf); return buf.getvalue()

# ══════════════════════════════════════════════════════════════════
# 5. PPT 생성
# ══════════════════════════════════════════════════════════════════

def _clone_chart_part(pkg, orig_cp, new_partname):
    new_cp = ChartPart(new_partname, orig_cp.content_type, pkg,
                       copy.deepcopy(orig_cp._element))
    for rId2, rel2 in orig_cp.rels.items():
        new_cp.rels._rels[rId2] = _Relationship(
            new_cp.partname.baseURI, rel2._rId, rel2._reltype,
            rel2._target_mode, rel2._target)
    return new_cp

def _clone_slide(prs, src_idx=0):
    src = prs.slides[src_idx]; sp = src.part; pkg = prs.part.package
    ns = prs.slides.add_slide(src.slide_layout); np_ = ns.part
    for ph in list(ns.placeholders):
        ph._element.getparent().remove(ph._element)
    for child in list(src.shapes._spTree):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("nvGrpSpPr", "grpSpPr"): continue
        ns.shapes._spTree.append(copy.deepcopy(child))
    cc = sum(1 for p in pkg.iter_parts()
             if str(p.partname).startswith("/ppt/charts/chart"))
    for rId, rel in sp.rels.items():
        if rId in np_.rels: continue
        if "chart" in rel._reltype:
            cc += 1
            new_cp = _clone_chart_part(
                pkg, rel._target, PackURI(f"/ppt/charts/chart{cc}.xml"))
            np_.rels._rels[rId] = _Relationship(
                np_.partname.baseURI, rel._rId, rel._reltype,
                rel._target_mode, new_cp)
        else:
            np_.rels._rels[rId] = _Relationship(
                np_.partname.baseURI, rel._rId, rel._reltype,
                rel._target_mode, rel._target)
    return ns

def _replace_text(shape, old, new_val):
    if not shape.has_text_frame: return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if old in run.text:
                run.text = run.text.replace(old, new_val)

def _set_table(shape, rows_data):
    tbl = shape.table
    for ri, (label, val) in enumerate(rows_data, start=1):
        if ri >= len(tbl.rows): break
        tbl.cell(ri, 0).text = str(label)
        tbl.cell(ri, 1).text = f"{val:.2f}" if isinstance(val, float) else str(val)

def _fill_slide(slide, name, result):
    c = result["competency"]; sr = result["skill_raw"]
    sa = result["soft_avg"];  ha = result["hard_avg"]
    for shape in slide.shapes:
        _replace_text(shape, "{{NAME}}", name)
        if shape.name == "table_phase" and shape.has_table:
            _set_table(shape, list(c.items()))
        elif shape.name == "table_strategy" and shape.has_table:
            rows = [(k, sr[k]) for k in SOFT_SKILLS]
            rows.append(("소프트스킬 평균", sa))
            rows += [(k, sr[k]) for k in HARD_SKILLS]
            rows.append(("하드스킬 평균", ha))
            _set_table(shape, rows)
        elif shape.name == "chart_phase" and shape.has_chart:
            try:
                cd = ChartData()
                cd.categories = list(c.keys())
                cd.add_series("역량 점수", list(c.values()))
                shape.chart.replace_data(cd)
            except Exception: pass
        elif shape.name == "chart_strategy" and shape.has_chart:
            try:
                cats = SOFT_SKILLS + ["소프트 평균"] + HARD_SKILLS + ["하드 평균"]
                vals = ([sr[k] for k in SOFT_SKILLS] + [sa] +
                        [sr[k] for k in HARD_SKILLS] + [ha])
                cd = ChartData()
                cd.categories = cats
                cd.add_series("계열 1", vals)
                shape.chart.replace_data(cd)
            except Exception: pass

def build_ppt(people, ppt_tpl_bytes):
    """
    매번 bytes에서 새로 Presentation을 열어 복제.
    Streamlit rerun과 무관하게 독립적으로 동작.
    """
    prs = Presentation(io.BytesIO(ppt_tpl_bytes))
    slides = [prs.slides[0]]
    for _ in range(len(people) - 1):
        slides.append(_clone_slide(prs, src_idx=0))
    for slide, person in zip(slides, people):
        _fill_slide(slide, person["name"], compute_person(person["scores"]))
    buf = io.BytesIO(); prs.save(buf); return buf.getvalue()

# ══════════════════════════════════════════════════════════════════
# 6. 템플릿 파일 탐색 (Streamlit Cloud 포함 다중 경로)
# ══════════════════════════════════════════════════════════════════

def find_template_bytes(filename: str) -> bytes | None:
    """
    GitHub 루트에 있는 파일을 여러 경로에서 탐색.
    Streamlit Cloud: cwd = /mount/src/{repo}/
    """
    candidates = []

    # 1. __file__ 기준 (일반 Python 실행)
    try:
        candidates.append(Path(__file__).parent / filename)
    except Exception:
        pass

    # 2. 현재 작업 디렉토리 (Streamlit Cloud 기준)
    candidates.append(Path(os.getcwd()) / filename)

    # 3. sys.argv[0] 기준
    try:
        if sys.argv and sys.argv[0]:
            candidates.append(
                Path(os.path.dirname(os.path.abspath(sys.argv[0]))) / filename)
    except Exception:
        pass

    # 4. /mount/src 하위 전체 탐색 (Streamlit Cloud fallback)
    try:
        mount = Path("/mount/src")
        if mount.exists():
            for p in mount.rglob(filename):
                candidates.append(p)
    except Exception:
        pass

    for path in candidates:
        try:
            if Path(path).exists():
                return Path(path).read_bytes()
        except Exception:
            continue
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

st.markdown('<div class="main-title">📊 리더십 진단 보고서 자동화 툴</div>',
            unsafe_allow_html=True)
st.markdown('<div class="sub">구글 폼 응답 엑셀 → 개인별 엑셀(평균 포함) + 응답자별 PPT 슬라이드 자동 생성</div>',
            unsafe_allow_html=True)
st.markdown("---")

st.markdown('<div class="sec">① 구글 폼 응답 엑셀 (필수)</div>', unsafe_allow_html=True)
response_file = st.file_uploader(
    "A열: 타임스탬프 / B열: 성함 / C~AF열: Q1~Q30 점수",
    type=["xlsx", "xls"], key="response")

st.markdown('<div class="sec">② 엑셀 템플릿 (선택 — GitHub 루트의 excel_template.xlsx 자동 사용)</div>',
            unsafe_allow_html=True)
excel_tpl_file = st.file_uploader(
    "A1:C1 병합=성함, C4:C33=점수, G/H열=소계/평균 양식",
    type=["xlsx"], key="excel_tpl")

st.markdown('<div class="sec">③ PPT 템플릿 (선택 — GitHub 루트의 template_pptx.pptx 자동 사용)</div>',
            unsafe_allow_html=True)
ppt_tpl_file = st.file_uploader(
    "{{NAME}}, table_phase, table_strategy, chart_phase, chart_strategy 개체 포함",
    type=["pptx"], key="ppt_tpl")

st.markdown("---")

# ── 생성 버튼 ──
if st.button("🚀 보고서 생성", type="primary", use_container_width=True):

    # ① 응답 엑셀 → bytes로 즉시 읽기 (rerun 후 사라지기 전에)
    if response_file is None:
        st.error("❌ 구글 폼 응답 엑셀을 업로드해주세요."); st.stop()
    response_bytes = response_file.read()

    # ② 템플릿 bytes 확보 (업로드 우선, 없으면 GitHub 루트 자동 탐색)
    if excel_tpl_file is not None:
        excel_tpl_bytes = excel_tpl_file.read()
        st.info("ℹ️ 엑셀 템플릿: 업로드 파일 사용")
    else:
        excel_tpl_bytes = find_template_bytes("excel_template.xlsx")
        if excel_tpl_bytes:
            st.info("ℹ️ 엑셀 템플릿: GitHub 루트의 excel_template.xlsx 자동 사용")
        else:
            st.error("❌ 엑셀 템플릿 없음. ② 업로더에서 직접 업로드하거나 "
                     "GitHub 루트에 excel_template.xlsx 를 추가해주세요."); st.stop()

    if ppt_tpl_file is not None:
        ppt_tpl_bytes = ppt_tpl_file.read()
        st.info("ℹ️ PPT 템플릿: 업로드 파일 사용")
    else:
        ppt_tpl_bytes = find_template_bytes("template_pptx.pptx")
        if ppt_tpl_bytes:
            st.info("ℹ️ PPT 템플릿: GitHub 루트의 template_pptx.pptx 자동 사용")
        else:
            st.error("❌ PPT 템플릿 없음. ③ 업로더에서 직접 업로드하거나 "
                     "GitHub 루트에 template_pptx.pptx 를 추가해주세요."); st.stop()

    # ③ 파싱
    with st.spinner("📂 응답 데이터 파싱 중..."):
        try:
            people = parse_response_excel(response_bytes)
        except Exception as e:
            st.error(f"❌ 파싱 실패: {e}"); st.stop()

    if not people:
        st.error("❌ 응답자 데이터가 없습니다."); st.stop()

    st.success(f"✅ {len(people)}명 응답 데이터 파싱 완료")

    # ④ 미리보기
    with st.expander(f"📋 응답자 미리보기 ({len(people)}명)"):
        preview = []
        for p in people:
            r = compute_person(p["scores"])
            row = {"성함": p["name"]}
            row.update({k: f"{v:.2f}" for k, v in r["competency"].items()})
            for k in SOFT_SKILLS + HARD_SKILLS:
                row[k] = f"{r['skill_raw'][k]:.2f}"
            row["소프트평균"] = f"{r['soft_avg']:.2f}"
            row["하드평균"]   = f"{r['hard_avg']:.2f}"
            preview.append(row)
        st.dataframe(pd.DataFrame(preview), use_container_width=True)

    # ⑤ 엑셀 생성
    with st.spinner("📊 개인별 엑셀 생성 중..."):
        try:
            excel_out = build_excel(people, excel_tpl_bytes)
        except Exception as e:
            st.error(f"❌ 엑셀 생성 실패: {e}"); st.exception(e); st.stop()

    # ⑥ PPT 생성
    with st.spinner(f"📑 PPT 생성 중 ({len(people)}슬라이드)..."):
        try:
            ppt_out = build_ppt(people, ppt_tpl_bytes)
        except Exception as e:
            st.error(f"❌ PPT 생성 실패: {e}"); st.exception(e); st.stop()

    # ⑦ session_state에 저장 (download_button rerun 대비)
    st.session_state["excel_out"]  = excel_out
    st.session_state["ppt_out"]    = ppt_out
    st.session_state["n_people"]   = len(people)
    st.session_state["generated"]  = True

# ── 다운로드 영역 (session_state에서 읽음 → rerun 안전) ──
if st.session_state.get("generated"):
    excel_out = st.session_state["excel_out"]
    ppt_out   = st.session_state["ppt_out"]
    n         = st.session_state["n_people"]

    st.success(f"🎉 완료! 엑셀 {n}시트 + PPT {n}슬라이드 생성")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("리더십진단_개인별.xlsx", excel_out)
        zf.writestr("리더십진단_통합.pptx",   ppt_out)

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button("⬇️ ZIP (전체)", data=zip_buf.getvalue(),
            file_name="리더십진단_결과.zip", mime="application/zip",
            use_container_width=True)
    with d2:
        st.download_button("⬇️ 엑셀 (개인별)", data=excel_out,
            file_name="리더십진단_개인별.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with d3:
        st.download_button("⬇️ PPT (통합)", data=ppt_out,
            file_name="리더십진단_통합.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True)
