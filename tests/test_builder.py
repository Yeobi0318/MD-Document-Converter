"""변환 엔진 테스트. 실행: .venv\\Scripts\\python -m pytest -q tests"""
import glob
import os
import sys
import zipfile

import pytest
from lxml import etree

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import docx  # noqa: E402

import ooxml_helpers as ox  # noqa: E402
import styles as S  # noqa: E402
from docx_builder import (NUM_RE, BuildOptions, DocxBuilder, build_docx, cell_groups,  # noqa: E402
                          find_mml2omml, split_feedback)
from md_parser import MdParseError, Span, parse, parse_file  # noqa: E402
from pdf_export import word_available  # noqa: E402
from worker import ConvertError, Converter, Job  # noqa: E402

SAMPLES = sorted(glob.glob(os.path.join(ROOT, "tests", "samples", "*.md")))
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


def q(t):
    return f"{{{W}}}{t}"


def build(tmp_path, md_text=None, path=None, **kw):
    if path:
        md = parse_file(path)
        name = os.path.basename(path)[:-3]
    else:
        md = parse(md_text)
        path = str(tmp_path / "t.md")
        name = "t"
    out = str(tmp_path / f"{name}.docx")
    opts = BuildOptions(mml2omml=find_mml2omml(), update_fields_on_open=False, **kw)
    res = build_docx(md, path, out, opts)
    return out, res


def body_xml(path):
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read("word/document.xml"))


# ── 샘플 전체 ──────────────────────────────────────────
@pytest.mark.parametrize("path", SAMPLES, ids=[os.path.basename(p) for p in SAMPLES])
def test_sample_converts_cleanly(tmp_path, path):
    out, res = build(tmp_path, path=path)
    # python-docx로 다시 열 수 있고 단락이 있다
    d = docx.Document(out)
    assert len(d.paragraphs) > 0
    root = body_xml(out)
    # 책갈피 id는 문서 전체에서 고유
    ids = [b.get(q("id")) for b in root.iter(q("bookmarkStart"))]
    assert len(ids) == len(set(ids))
    names = {b.get(q("name")) for b in root.iter(q("bookmarkStart"))}
    # 차례의 모든 링크가 실제 책갈피를 가리킨다
    for hl in root.iter(q("hyperlink")):
        a = hl.get(q("anchor"))
        if a:
            assert a in names
    # 마크다운 잔재 없음 (코드 제외)
    assert not [w for w in res.warnings if "잔재" in w], res.warnings
    texts = [t.text or "" for t in root.iter(q("t"))
             if not _in_code(t)]
    for pat in ("|", "---", "**", "##", "`"):
        assert not any(pat in t for t in texts), pat


def _in_code(t):
    r = t.getparent()
    rs = r.find(q("rPr") + "/" + q("rStyle"))
    if rs is not None and rs.get(q("val")) == "MDCodeChar":
        return True
    p = r
    while p is not None and p.tag != q("p"):
        p = p.getparent()
    ps = p.find(q("pPr") + "/" + q("pStyle")) if p is not None else None
    return ps is not None and ps.get(q("val")) == "MDCode"


@pytest.mark.parametrize("path", SAMPLES, ids=[os.path.basename(p) for p in SAMPLES])
def test_schema_child_order(tmp_path, path):
    """pPr/rPr/tblPr/tcPr/sectPr 자식이 스키마 순서대로인지 (Word 「복구」 방지)."""
    out, _ = build(tmp_path, path=path)
    with zipfile.ZipFile(out) as z:
        for part in ("word/document.xml", "word/styles.xml", "word/settings.xml"):
            root = etree.fromstring(z.read(part))
            # headerReference/footerReference는 스키마상 섞여도 되는 선택 그룹
            same = {q("footerReference"): q("headerReference")}
            for tag, order in ox._ORDER_Q.items():
                for e in root.iter(tag):
                    idx = [order.index(same.get(c.tag, c.tag)) for c in e if c.tag in order]
                    assert idx == sorted(idx), f"{part} {tag}: {[c.tag.split('}')[1] for c in e]}"


@pytest.mark.parametrize("path", SAMPLES, ids=[os.path.basename(p) for p in SAMPLES])
def test_tables_fill_body_width(tmp_path, path):
    out, _ = build(tmp_path, path=path)
    root = body_xml(out)
    body = root.find(q("body"))
    for tbl in body.findall(q("tbl")):
        grid = [int(g.get(q("w"))) for g in tbl.find(q("tblGrid"))]
        ind = tbl.find(q("tblPr") + "/" + q("tblInd"))
        indent = int(ind.get(q("w"))) if ind is not None else 0
        assert sum(grid) + indent == S.BODY_W
        assert tbl.find(q("tblPr") + "/" + q("tblLayout")).get(q("type")) == "fixed"
        assert tbl.find(q("tblPr") + "/" + q("tblW")).get(q("type")) == "dxa"
        for tr in tbl.findall(q("tr")):
            ws = [int(tc.find(q("tcPr") + "/" + q("tcW")).get(q("w"))) for tc in tr.findall(q("tc"))]
            assert sum(ws) == sum(grid)


def test_no_cell_borders_on_normal_tables(tmp_path):
    out, _ = build(tmp_path, "# T\n\n## A\n\n| a | b |\n|---|---|\n| 1 | 2 |\n")
    root = body_xml(out)
    for tbl in root.iter(q("tbl")):
        borders = tbl.find(q("tblPr") + "/" + q("tblBorders"))
        assert borders is not None
        assert [c.tag.split("}")[1] for c in borders] == ["top", "left", "bottom", "right", "insideH", "insideV"]
        assert not list(tbl.iter(q("tcBorders")))
        hdr = tbl.find(q("tr") + "/" + q("trPr") + "/" + q("tblHeader"))
        assert hdr is not None
        shd = tbl.find(q("tr") + "/" + q("tc") + "/" + q("tcPr") + "/" + q("shd"))
        assert shd.get(q("val")) == "clear" and shd.get(q("fill")) == S.HEADER_ROW_BG


# ── 요소별 ─────────────────────────────────────────────
def test_lists_restart_and_use_numbering(tmp_path):
    out, _ = build(tmp_path, "## A\n\n1. a\n2. b\n\n문단\n\n1. c\n2. d\n\n- x\n  - y\n")
    root = body_xml(out)
    nums = []
    for p in root.iter(q("p")):
        numId = p.find(q("pPr") + "/" + q("numPr") + "/" + q("numId"))
        if numId is not None:
            ilvl = p.find(q("pPr") + "/" + q("numPr") + "/" + q("ilvl")).get(q("val"))
            text = "".join(t.text or "" for t in p.iter(q("t")))
            nums.append((text, numId.get(q("val")), ilvl))
    by = {t: (n, l) for t, n, l in nums}
    assert by["a"][0] == by["b"][0]
    assert by["a"][0] != by["c"][0]            # 목록마다 새 번호
    assert by["y"][1] == "1"                   # 2단계
    texts = [t.text or "" for t in root.iter(q("t"))]
    assert "•" not in "".join(texts)           # 글머리 문자를 직접 넣지 않음


def test_task_list_checkboxes(tmp_path):
    out, _ = build(tmp_path, "## A\n\n- [ ] 하나\n- [x] 둘\n")
    root = body_xml(out)
    boxes = list(root.iter(f"{{{W14}}}checkbox"))
    assert len(boxes) == 2
    states = [b.find(f"{{{W14}}}checked").get(f"{{{W14}}}val") for b in boxes]
    assert states == ["0", "1"]
    sdt_ids = [s.find(q("sdtPr") + "/" + q("id")).get(q("val")) for s in root.iter(q("sdt"))]
    assert len(sdt_ids) == len(set(sdt_ids))
    glyphs = [t.text for s in root.iter(q("sdt")) for t in s.iter(q("t"))]
    assert glyphs == ["☐", "☑"]


def test_alerts_become_boxes(tmp_path):
    md = "## A\n\n> [!TIP]\n> **제목**\n> 본문\n\n> [!WARNING]\n> 경고\n"
    out, _ = build(tmp_path, md)
    root = body_xml(out)
    text = "".join(t.text or "" for t in root.iter(q("t")))
    assert "[💡 TIP] 제목" in text and "[⚠️ 주의]" in text and "[!TIP]" not in text
    bars = [b.find(q("left")) for b in root.iter(q("tcBorders"))]
    assert {b.get(q("color")) for b in bars} == {S.BLUE, S.ORANGE}


def test_auto_box_sections(tmp_path):
    md = "## 3. 강의자가 강조·반복한 내용\n\n**강조**\n1. 하나\n\n## 4. 원본 확인이 필요한 부분\n\n- 확인\n"
    out, _ = build(tmp_path, md)
    root = body_xml(out)
    assert len(list(root.iter(q("tcBorders")))) == 2
    out2, _ = build(tmp_path, md, auto_box=False)
    assert len(list(body_xml(out2).iter(q("tcBorders")))) == 0


def test_feedback_and_cell_split():
    spans = [Span("답변입니다. ▶ 피드백 문장")]
    groups = split_feedback(spans)
    assert [k for k, _ in groups] == ["normal", "feedback"]
    long = "첫 문장은 충분히 깁니다. " * 5 + "마지막 문장입니다!"
    cg = cell_groups([Span(long), Span(kind="br"), Span("둘째 줄")])
    kinds = [k for k, _ in cg]
    assert kinds.count("bullet") >= 2 and kinds[-1] == "normal"


def test_number_emphasis_regex():
    def hits(s):
        return [m.group(0) for m in NUM_RE.finditer(s)]
    assert hits("답변은 40~50초, 예산 100만 원, 10명, 3년, 85%, A+") == \
        ["40~50초", "100만 원", "10명", "3년", "85%", "A+"]
    assert hits("시각 00:12:30, M1, Q3") == []


def test_text_cleanup(tmp_path):
    out, _ = build(tmp_path, "## A\n\n앞 — 뒤, 범위 1 — 5, A | B\n\n| x — y | z |\n|---|---|\n| a — b | c |\n")
    root = body_xml(out)
    paras = ["".join(t.text or "" for t in p.iter(q("t"))) for p in root.iter(q("p"))]
    assert any(p.startswith("앞: 뒤, 범위 1 — 5, A B") for p in paras)
    cells = ["".join(t.text or "" for t in tc.iter(q("t"))) for tc in root.iter(q("tc"))]
    assert "a — b" in cells                    # 표 안은 그대로


def test_unparsed_bold_is_fixed(tmp_path):
    md = parse("**R(깨달음)**을 강조")
    sp = md.blocks[0].spans
    assert sp[0].bold and sp[0].text == "R(깨달음)" and not sp[1].bold


def test_title_from_heading_without_front_matter(tmp_path):
    out, res = build(tmp_path, "# 내 문서\n\n## 장 하나\n\n본문\n")
    assert res.title == "내 문서"
    assert [h.text for h in res.headings] == ["장 하나"]


def test_plain_md_without_headings(tmp_path):
    out, res = build(tmp_path, "그냥 문단 하나.\n\n- 목록\n")
    assert len(docx.Document(out).paragraphs) > 0
    assert res.title == "t"


def test_pageref_fields_and_page_footer(tmp_path):
    out, res = build(tmp_path, "# T\n\n## 하나\n\n### 둘\n\n본문\n")
    root = body_xml(out)
    instr = [i.text for i in root.iter(q("instrText"))]
    assert sum("PAGEREF" in i for i in instr) == 2
    sect = root.find(q("body") + "/" + q("sectPr"))
    assert sect.find(q("docGrid")).get(q("type")) == "default"
    assert sect.find(q("pgSz")).get(q("w")) == str(S.PAGE_W)
    assert sect.find(q("titlePg")) is not None


def test_all_paragraph_line_rules_fixed(tmp_path):
    out, _ = build(tmp_path, path=SAMPLES[0])
    with zipfile.ZipFile(out) as z:
        st = etree.fromstring(z.read("word/styles.xml"))
    for s in st.iter(q("style")):
        if s.get(q("type")) != "paragraph":
            continue
        sp = s.find(q("pPr") + "/" + q("spacing"))
        sid = s.get(q("styleId"))
        if sid in ("Normal", "Heading1", "Heading2", "Heading3", "MDTable", "MDBox", "MDSpacer"):
            assert sp.get(q("lineRule")) == "exact", sid


# ── 잘못된 입력 ────────────────────────────────────────
def test_empty_md_fails_cleanly(tmp_path):
    with pytest.raises(MdParseError):
        parse("   \n\n")
    p = tmp_path / "empty.md"
    p.write_text("", encoding="utf-8")
    conv = Converter("docx", "none")
    with pytest.raises(ConvertError):
        conv.convert(Job(0, str(p), str(tmp_path / "empty.docx")))


def test_binary_file_fails_cleanly(tmp_path):
    p = tmp_path / "bin.md"
    p.write_bytes(b"\x00\x01\x02PK\x03\x04" * 50)
    with pytest.raises(ConvertError):
        Converter("docx", "none").convert(Job(0, str(p), str(tmp_path / "bin.docx")))


def test_broken_md_still_converts(tmp_path):
    p = tmp_path / "broken.md"
    p.write_text("---\ntitle: [깨진 yaml\n---\n\n## 제목\n\n| a | b\n|---\n| 1 |\n\n**안 닫힌 굵게\n\n$$\n\\frac{\n$$\n"
                 "![](없는.png)\n\n> [!UNKNOWN]\n> 내용\n", encoding="utf-8")
    conv = Converter("docx", "none")
    conv.convert(Job(0, str(p), str(tmp_path / "broken.docx")))
    assert os.path.exists(tmp_path / "broken.docx")


def test_cp949_input(tmp_path):
    p = tmp_path / "cp.md"
    p.write_bytes("## 한글 제목\n\n본문입니다.\n".encode("cp949"))
    Converter("docx", "none").convert(Job(0, str(p), str(tmp_path / "cp.docx")))
    d = docx.Document(str(tmp_path / "cp.docx"))
    assert any("본문입니다" in x.text for x in d.paragraphs)


def test_pdf_without_engine_fails(tmp_path):
    p = tmp_path / "a.md"
    p.write_text("## A\n\n본문\n", encoding="utf-8")
    with pytest.raises(ConvertError):
        Converter("pdf", "none").convert(Job(0, str(p), str(tmp_path / "a.pdf")))


def test_no_engine_sets_update_fields(tmp_path):
    p = tmp_path / "a.md"
    p.write_text("## A\n\n본문\n", encoding="utf-8")
    Converter("docx", "none").convert(Job(0, str(p), str(tmp_path / "a.docx")))
    with zipfile.ZipFile(tmp_path / "a.docx") as z:
        st = etree.fromstring(z.read("word/settings.xml"))
    assert st.find(q("updateFields")) is not None


# ── Word 통합 (Word가 있을 때만) ───────────────────────
word = pytest.mark.skipif(not word_available(), reason="Word 없음")


@word
def test_word_opens_without_repair_and_checkbox_toggles(tmp_path):
    import pythoncom
    import win32com.client
    out, _ = build(tmp_path, path=os.path.join(ROOT, "tests", "samples", "요소_예제.md"))
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        # OpenAndRepair=False: 손상된 문서면 예외가 난다
        d = app.Documents.Open(out, False, False, False)
        try:
            assert d.Paragraphs.Count > 0
            cbs = [c for c in d.ContentControls if c.Type == 8]   # wdContentControlCheckBox
            assert len(cbs) == 3
            cbs[0].Checked = True
            assert cbs[0].Range.Text == "☑"
        finally:
            d.Close(0)
    finally:
        app.Quit(0)
        pythoncom.CoUninitialize()


@word
def test_word_pipeline_pdf_and_page_numbers(tmp_path):
    src = os.path.join(ROOT, "tests", "samples", "요소_예제.md")
    conv = Converter("pdf", "word")
    try:
        conv.convert(Job(0, src, str(tmp_path / "out.pdf")))
    finally:
        conv.close()
    from pypdf import PdfReader
    r = PdfReader(str(tmp_path / "out.pdf"))
    assert len(r.pages) >= 3
    assert len(r.outline) > 0                       # 제목 기반 PDF 책갈피
    links = [a for p in r.pages[:2] for a in (p.get("/Annots") or [])]
    assert links                                    # 차례 링크
