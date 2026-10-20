"""OOXML을 직접 다루는 도우미: 스키마 순서 삽입, 간격, 테두리, 책갈피, 필드, 체크박스.

Word는 pPr/rPr/tblPr 등의 자식 순서가 스키마와 다르면 「복구」 메시지를 띄운다.
그래서 모든 속성 요소는 set_child()로 스키마 순서에 맞춰 넣는다.
"""
from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

# 스키마(transitional) 자식 순서
ORDER = {
    "w:pPr": ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
              "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
              "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
              "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
              "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
              "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
              "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
              "w:sectPr", "w:pPrChange"],
    "w:rPr": ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
              "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint",
              "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing",
              "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
              "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
              "w:eastAsianLayout", "w:specVanish", "w:oMath"],
    "w:tblPr": ["w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize",
                "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd",
                "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
                "w:tblCaption", "w:tblDescription"],
    "w:tcPr": ["w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders",
               "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
               "w:hideMark"],
    "w:sectPr": ["w:headerReference", "w:footerReference", "w:footnotePr", "w:endnotePr",
                 "w:type", "w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType",
                 "w:pgNumType", "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
                 "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings"],
    "w:settings": ["w:writeProtection", "w:view", "w:zoom", "w:removePersonalInformation",
                   "w:removeDateAndTime", "w:doNotDisplayPageBoundaries",
                   "w:displayBackgroundShape", "w:printPostScriptOverText",
                   "w:printFractionalCharacterWidth", "w:printFormsData",
                   "w:embedTrueTypeFonts", "w:embedSystemFonts", "w:saveSubsetFonts",
                   "w:saveFormsData", "w:mirrorMargins", "w:alignBordersAndEdges",
                   "w:bordersDoNotSurroundHeader", "w:bordersDoNotSurroundFooter",
                   "w:gutterAtTop", "w:hideSpellingErrors", "w:hideGrammaticalErrors",
                   "w:activeWritingStyle", "w:proofState", "w:formsDesign",
                   "w:attachedTemplate", "w:linkStyles", "w:stylePaneFormatFilter",
                   "w:stylePaneSortMethod", "w:documentType", "w:mailMerge",
                   "w:revisionView", "w:trackRevisions", "w:doNotTrackMoves",
                   "w:doNotTrackFormatting", "w:documentProtection", "w:autoFormatOverride",
                   "w:styleLockTheme", "w:styleLockQFSet", "w:defaultTabStop",
                   "w:autoHyphenation", "w:consecutiveHyphenLimit", "w:hyphenationZone",
                   "w:doNotHyphenateCaps", "w:showEnvelope", "w:summaryLength",
                   "w:clickAndTypeStyle", "w:defaultTableStyle", "w:evenAndOddHeaders",
                   "w:bookFoldRevPrinting", "w:bookFoldPrinting", "w:bookFoldPrintingSheets",
                   "w:drawingGridHorizontalSpacing", "w:drawingGridVerticalSpacing",
                   "w:displayHorizontalDrawingGridEvery", "w:displayVerticalDrawingGridEvery",
                   "w:doNotUseMarginsForDrawingGridOrigin", "w:drawingGridHorizontalOrigin",
                   "w:drawingGridVerticalOrigin", "w:doNotShadeFormData",
                   "w:noPunctuationKerning", "w:characterSpacingControl", "w:printTwoOnOne",
                   "w:strictFirstAndLastChars", "w:noLineBreaksAfter", "w:noLineBreaksBefore",
                   "w:savePreviewPicture", "w:doNotValidateAgainstSchema", "w:saveInvalidXml",
                   "w:ignoreMixedContent", "w:alwaysShowPlaceholderText",
                   "w:doNotDemarcateInvalidXml", "w:saveXmlDataOnly", "w:useXSLTWhenSaving",
                   "w:saveThroughXslt", "w:showXMLTags", "w:alwaysMergeEmptyNamespace",
                   "w:updateFields", "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr",
                   "w:compat", "w:docVars", "w:rsids", "m:mathPr", "w:attachedSchema",
                   "w:themeFontLang", "w:clrSchemeMapping"],
}
_ORDER_Q = {qn(k): [qn(x) for x in v] for k, v in ORDER.items()}


def el(tag: str, **attrs) -> OxmlElement:
    """el('w:spacing', after=120) → <w:spacing w:after="120"/>; 속성 이름의 접두어는 생략 가능."""
    e = OxmlElement(tag)
    for k, v in attrs.items():
        if v is None:
            continue
        name = k.replace("__", ":") if "__" in k else ("w:" + k)
        e.set(qn(name), str(v))
    return e


def set_child(parent, child, replace_existing=True):
    """스키마 순서에 맞춰 child를 넣는다. 같은 태그가 있으면 바꾼다."""
    order = _ORDER_Q.get(parent.tag)
    if replace_existing:
        old = parent.find(child.tag)
        if old is not None:
            old.addprevious(child)
            parent.remove(old)
            return child
    if order and child.tag in order:
        idx = order.index(child.tag)
        later = set(order[idx + 1:])
        for existing in parent:
            if existing.tag in later:
                existing.addprevious(child)
                return child
    parent.append(child)
    return child


def get_pPr(p):
    """w:p 요소의 pPr (없으면 만든다)."""
    return p.get_or_add_pPr()


def get_rPr(r):
    return r.get_or_add_rPr()


def pt2tw(v: float) -> int:
    return round(v * 20)


# ── 단락 속성 ──────────────────────────────────────────
def spacing(p, before=None, after=None, line=None, rule="exact"):
    """줄 간격·앞뒤 간격(pt). line이 None이면 줄 간격은 건드리지 않는다."""
    pPr = get_pPr(p)
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = set_child(pPr, el("w:spacing"))
    if before is not None:
        sp.set(qn("w:before"), str(pt2tw(before)))
        sp.set(qn("w:beforeAutospacing"), "0")
    if after is not None:
        sp.set(qn("w:after"), str(pt2tw(after)))
        sp.set(qn("w:afterAutospacing"), "0")
    if line is not None:
        if rule == "auto":
            sp.set(qn("w:line"), str(round(line * 240)))
        else:
            sp.set(qn("w:line"), str(pt2tw(line)))
        sp.set(qn("w:lineRule"), rule)
    return sp


def get_after(p) -> float | None:
    sp = p.find(qn("w:pPr") + "/" + qn("w:spacing"))
    if sp is not None and sp.get(qn("w:after")) is not None:
        return int(sp.get(qn("w:after"))) / 20
    return None


def flag(p, tag, on=True):
    """keepNext/keepLines/pageBreakBefore/widowControl 같은 on/off 속성."""
    pPr = get_pPr(p)
    if on:
        set_child(pPr, el(tag))
    else:
        set_child(pPr, el(tag, val=0))


def indent(p, left=None, hanging=None, first=None, right=None):
    pPr = get_pPr(p)
    ind = el("w:ind", left=left, hanging=hanging, firstLine=first, right=right)
    set_child(pPr, ind)


def align(p, val):
    set_child(get_pPr(p), el("w:jc", val=val))


def style(p, style_id):
    set_child(get_pPr(p), el("w:pStyle", val=style_id))


def numbering(p, num_id, ilvl=0):
    numPr = el("w:numPr")
    numPr.append(el("w:ilvl", val=ilvl))
    numPr.append(el("w:numId", val=num_id))
    set_child(get_pPr(p), numPr)


def tabs(p, stops):
    """stops: [(val, pos, leader)]"""
    t = el("w:tabs")
    for val, pos, leader in stops:
        t.append(el("w:tab", val=val, pos=pos, leader=leader))
    set_child(get_pPr(p), t)


def shade(parent_pr, fill):
    """pPr/tcPr/rPr에 배경색. CLEAR 패턴(SOLID는 검게 나온다)."""
    set_child(parent_pr, el("w:shd", val="clear", color="auto", fill=fill))


def p_border(p, side, sz, color, space=1, val="single"):
    pPr = get_pPr(p)
    bdr = pPr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = set_child(pPr, el("w:pBdr"))
    order = ["top", "left", "bottom", "right", "between", "bar"]
    new = el("w:" + side, val=val, sz=sz, space=space, color=color)
    old = bdr.find(qn("w:" + side))
    if old is not None:
        bdr.remove(old)
    later = {qn("w:" + s) for s in order[order.index(side) + 1:]}
    for e in bdr:
        if e.tag in later:
            e.addprevious(new)
            break
    else:
        bdr.append(new)


def outline_level(p, lvl):
    set_child(get_pPr(p), el("w:outlineLvl", val=lvl))


# ── 런 속성 ────────────────────────────────────────────
def run_fonts(r, font, hint=None):
    rPr = get_rPr(r)
    set_child(rPr, el("w:rFonts", ascii=font, hAnsi=font, eastAsia=font, cs=font, hint=hint))


def run_color(r, color):
    set_child(get_rPr(r), el("w:color", val=color))


def run_size(r, size_pt):
    rPr = get_rPr(r)
    set_child(rPr, el("w:sz", val=round(size_pt * 2)))
    set_child(rPr, el("w:szCs", val=round(size_pt * 2)))


def run_bold(r, on=True):
    rPr = get_rPr(r)
    set_child(rPr, el("w:b") if on else el("w:b", val=0))
    set_child(rPr, el("w:bCs") if on else el("w:bCs", val=0))


def run_italic(r, on=True):
    rPr = get_rPr(r)
    set_child(rPr, el("w:i") if on else el("w:i", val=0))


def run_flag(r, tag):
    set_child(get_rPr(r), el(tag))


def run_style(r, style_id):
    set_child(get_rPr(r), el("w:rStyle", val=style_id))


# ── 표 ─────────────────────────────────────────────────
def tbl_borders(tblPr, sz=4, color="BFBFBF", none=False):
    """표 단위 테두리. top, left, bottom, right, insideH, insideV 순서(스키마)."""
    b = el("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if none:
            b.append(el("w:" + side, val="nil"))
        else:
            b.append(el("w:" + side, val="single", sz=sz, space=0, color=color))
    set_child(tblPr, b)


def tc_borders(tcPr, **sides):
    """셀 테두리. sides: left=(val, sz, color) 형태. 나머지 변은 nil."""
    b = el("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        spec = sides.get(side)
        if spec:
            val, sz, color = spec
            b.append(el("w:" + side, val=val, sz=sz, space=0, color=color))
        else:
            b.append(el("w:" + side, val="nil"))
    set_child(tcPr, b)


def cell_margins(pr, top, left, bottom, right, tag="w:tblCellMar"):
    m = el(tag)
    for side, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        m.append(el("w:" + side, w=v, type="dxa"))
    set_child(pr, m)


# ── 책갈피·필드·하이퍼링크 ─────────────────────────────
class IdCounter:
    """문서 전체에서 고유한 id (책갈피, 콘텐츠 컨트롤)."""

    def __init__(self, start=1):
        self.n = start

    def next(self) -> int:
        v = self.n
        self.n += 1
        return v


def wrap_bookmark(p, name, bm_id):
    """단락 p의 내용 전체를 책갈피로 감싼다."""
    start = el("w:bookmarkStart", id=bm_id, name=name)
    end = el("w:bookmarkEnd", id=bm_id)
    pPr = p.find(qn("w:pPr"))
    if pPr is not None:
        pPr.addnext(start)
    else:
        p.insert(0, start)
    p.append(end)


def make_run(text=None, rpr_children=()):
    r = OxmlElement("w:r")
    if rpr_children:
        rPr = OxmlElement("w:rPr")
        for c in rpr_children:
            set_child(rPr, c)
        r.append(rPr)
    if text is not None:
        t = OxmlElement("w:t")
        t.text = text
        t.set(qn("xml:space"), "preserve")
        r.append(t)
    return r


def field_runs(instr, result="", rpr_children=()):
    """복합 필드(begin/instr/separate/result/end)를 이루는 w:r 목록."""
    def r_with(child):
        r = make_run(rpr_children=[c.__copy__() for c in rpr_children])
        r.append(child)
        return r
    runs = [r_with(el("w:fldChar", fldCharType="begin"))]
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = f" {instr} "
    runs.append(r_with(instr_el))
    runs.append(r_with(el("w:fldChar", fldCharType="separate")))
    runs.append(make_run(result, [c.__copy__() for c in rpr_children]))
    runs.append(r_with(el("w:fldChar", fldCharType="end")))
    return runs


def internal_hyperlink(anchor):
    return el("w:hyperlink", anchor=anchor, history=1)


def external_hyperlink(part, url):
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    rid = part.relate_to(url, RT.HYPERLINK, is_external=True)
    return el("w:hyperlink", r__id=rid, history=1)


# ── 체크박스 콘텐츠 컨트롤 ─────────────────────────────
def checkbox_sdt(checked: bool, sdt_id: int, font: str, size_pt: float):
    """클릭 가능한 Word 체크박스(w14:checkbox). 미체크 ☐ U+2610 / 체크 ☑ U+2611."""
    sdt = OxmlElement("w:sdt")
    pr = OxmlElement("w:sdtPr")
    rpr = OxmlElement("w:rPr")
    set_child(rpr, el("w:rFonts", ascii=font, hAnsi=font, eastAsia=font, cs=font))
    set_child(rpr, el("w:sz", val=round(size_pt * 2)))
    pr.append(rpr)
    pr.append(el("w:id", val=sdt_id))
    cb = OxmlElement("w14:checkbox")
    c = OxmlElement("w14:checked")
    c.set(f"{{{W14}}}val", "1" if checked else "0")
    cs = OxmlElement("w14:checkedState")
    cs.set(f"{{{W14}}}val", "2611")
    cs.set(f"{{{W14}}}font", font)
    us = OxmlElement("w14:uncheckedState")
    us.set(f"{{{W14}}}val", "2610")
    us.set(f"{{{W14}}}font", font)
    cb.extend([c, cs, us])
    pr.append(cb)
    sdt.append(pr)
    content = OxmlElement("w:sdtContent")
    content.append(make_run("☑" if checked else "☐",
                            [el("w:rFonts", ascii=font, hAnsi=font, eastAsia=font, cs=font),
                             el("w:sz", val=round(size_pt * 2))]))
    sdt.append(content)
    return sdt
