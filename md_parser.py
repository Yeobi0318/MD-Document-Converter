"""markdown-it 토큰 → 중간 모델(블록 목록).

중간 모델은 docx와 무관한 단순 자료 구조다. docx_builder가 이 모델만 보고 문서를 만든다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, replace

import yaml
from markdown_it import MarkdownIt
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.footnote import footnote_plugin
from mdit_py_plugins.front_matter import front_matter_plugin
from mdit_py_plugins.tasklists import tasklists_plugin


# ── 인라인 ─────────────────────────────────────────────
@dataclass
class Span:
    text: str = ""
    kind: str = "text"      # text | br | math | image | checkbox
    bold: bool = False
    italic: bool = False
    strike: bool = False
    code: bool = False
    sup: bool = False
    href: str | None = None
    src: str = ""           # image 경로 또는 math의 LaTeX


# ── 블록 ───────────────────────────────────────────────
@dataclass
class Heading:
    level: int
    spans: list[Span]


@dataclass
class Paragraph:
    spans: list[Span]


@dataclass
class ListItem:
    blocks: list
    checked: bool | None = None     # None: 일반 항목, True/False: 작업 목록


@dataclass
class ListBlock:
    ordered: bool
    start: int
    items: list[ListItem]
    tight: bool = True


@dataclass
class Table:
    header: list[list[Span]]
    rows: list[list[list[Span]]]
    aligns: list[str | None]
    accent: str | None = None       # 자동 상자화 절의 표: 머리글 행을 상자 색으로


@dataclass
class Quote:
    blocks: list
    kind: str = "quote"             # styles.BOX_KINDS 키
    title: list[Span] | None = None


@dataclass
class MathBlock:
    latex: str


@dataclass
class CodeBlock:
    text: str
    lang: str = ""


@dataclass
class Rule:
    pass


@dataclass
class Footnote:
    label: str
    blocks: list


@dataclass
class MdDocument:
    meta: dict = field(default_factory=dict)
    blocks: list = field(default_factory=list)
    footnotes: list[Footnote] = field(default_factory=list)


class MdParseError(Exception):
    pass


# ── 파서 ───────────────────────────────────────────────
def _make_md() -> MarkdownIt:
    md = MarkdownIt("commonmark", {"html": True})
    md.enable(["table", "strikethrough"])
    md.use(front_matter_plugin)
    md.use(tasklists_plugin)
    md.use(dollarmath_plugin, allow_digits=False, double_inline=True)
    md.use(footnote_plugin)
    return md


_MD = None
_TAG_RE = re.compile(r"<[^>]+>")
_BR_RE = re.compile(r"<br\s*/?>", re.I)
_ALERT_RE = re.compile(r"^\s*\[!(\w+)\]\s*")


def read_text(path: str) -> str:
    """UTF-8(BOM 포함) → CP949 순으로 읽는다. 둘 다 실패하면 깨진 파일로 본다."""
    with open(path, "rb") as f:
        raw = f.read()
    if b"\x00" in raw[:4096]:
        raise MdParseError("텍스트 파일이 아닙니다 (이진 데이터 포함)")
    for enc in ("utf-8-sig", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise MdParseError("문자 인코딩을 읽을 수 없습니다 (UTF-8/CP949 아님)")


def parse_file(path: str) -> MdDocument:
    return parse(read_text(path))


def parse(text: str) -> MdDocument:
    global _MD
    if _MD is None:
        _MD = _make_md()
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if not text.strip():
        raise MdParseError("내용이 없는 파일입니다")
    tokens = _MD.parse(text)
    doc = MdDocument()
    p = _Parser(tokens, doc)
    doc.blocks = p.blocks(None)
    if not doc.blocks and not doc.meta:
        raise MdParseError("변환할 본문이 없습니다")
    return doc


class _Parser:
    def __init__(self, tokens, doc: MdDocument):
        self.t = tokens
        self.i = 0
        self.doc = doc

    def blocks(self, close_type: str | None) -> list:
        out = []
        while self.i < len(self.t):
            tok = self.t[self.i]
            if close_type and tok.type == close_type:
                self.i += 1
                return out
            blk = self.block()
            if blk is not None:
                out.append(blk)
        return out

    def block(self):
        tok = self.t[self.i]
        tt = tok.type
        if tt == "front_matter":
            self.i += 1
            try:
                meta = yaml.safe_load(tok.content) or {}
            except yaml.YAMLError:
                meta = {}
            if isinstance(meta, dict):
                self.doc.meta = {str(k): v for k, v in meta.items()}
            return None
        if tt == "heading_open":
            level = int(tok.tag[1])
            spans = self.inline(self.t[self.i + 1])
            self.i += 3
            return Heading(level, spans)
        if tt == "paragraph_open":
            spans = self.inline(self.t[self.i + 1])
            self.i += 3
            return Paragraph(spans) if spans else None
        if tt in ("bullet_list_open", "ordered_list_open"):
            return self.list_block()
        if tt == "blockquote_open":
            self.i += 1
            return self.quote(self.blocks("blockquote_close"))
        if tt == "table_open":
            return self.table()
        if tt in ("fence", "code_block"):
            self.i += 1
            return CodeBlock(tok.content.rstrip("\n"), (tok.info or "").strip())
        if tt in ("math_block", "math_block_label"):
            self.i += 1
            return MathBlock(tok.content.strip())
        if tt == "hr":
            self.i += 1
            return Rule()
        if tt == "html_block":
            self.i += 1
            txt = _TAG_RE.sub("", _BR_RE.sub("\n", tok.content)).strip()
            if not txt:
                return None
            spans = []
            for n, line in enumerate(txt.split("\n")):
                if n:
                    spans.append(Span(kind="br"))
                spans.append(Span(line))
            return Paragraph(spans)
        if tt == "footnote_block_open":
            self.i += 1
            self.footnotes()
            return None
        # 알 수 없는 토큰은 건너뛴다
        self.i += 1
        return None

    def list_block(self) -> ListBlock:
        tok = self.t[self.i]
        ordered = tok.type == "ordered_list_open"
        close = "ordered_list_close" if ordered else "bullet_list_close"
        start = int(tok.attrGet("start") or 1) if ordered else 1
        self.i += 1
        items, tight = [], True
        while self.i < len(self.t) and self.t[self.i].type != close:
            it = self.t[self.i]
            if it.type != "list_item_open":
                self.i += 1
                continue
            self.i += 1
            # 느슨한 목록: 항목 안 단락 토큰이 숨겨지지 않는다
            if self.i < len(self.t) and self.t[self.i].type == "paragraph_open" and not self.t[self.i].hidden:
                tight = False
            blocks = self.blocks("list_item_close")
            checked = None
            if blocks and isinstance(blocks[0], Paragraph) and blocks[0].spans and blocks[0].spans[0].kind == "checkbox":
                checked = blocks[0].spans[0].text == "x"
                rest = blocks[0].spans[1:]
                if rest and rest[0].kind == "text":
                    rest[0] = replace(rest[0], text=rest[0].text.lstrip())
                blocks[0] = Paragraph(rest)
            items.append(ListItem(blocks, checked))
        self.i += 1
        return ListBlock(ordered, start, items, tight)

    def quote(self, blocks: list) -> Quote:
        q = Quote(blocks)
        if blocks and isinstance(blocks[0], Paragraph) and blocks[0].spans:
            first = blocks[0].spans[0]
            m = _ALERT_RE.match(first.text) if first.kind == "text" else None
            if m and m.group(1).upper() in _alert_map():
                q.kind = _alert_map()[m.group(1).upper()]
                spans = list(blocks[0].spans)
                spans[0] = replace(first, text=first.text[m.end():])
                q.title, rest = _take_title(spans)
                if rest:
                    q.blocks = [Paragraph(rest)] + blocks[1:]
                else:
                    q.blocks = blocks[1:]
        return q

    def table(self) -> Table:
        self.i += 1
        header, rows, aligns, cur = [], [], [], None
        in_head = False
        while self.i < len(self.t):
            tok = self.t[self.i]
            tt = tok.type
            if tt == "table_close":
                self.i += 1
                break
            if tt == "thead_open":
                in_head = True
            elif tt == "thead_close":
                in_head = False
            elif tt == "tr_open":
                cur = []
            elif tt == "tr_close":
                (header.extend if in_head else rows.append)(cur if not in_head else [cur])
                cur = None
            elif tt in ("th_open", "td_open"):
                style = tok.attrGet("style") or ""
                if in_head:
                    m = re.search(r"text-align:(\w+)", style)
                    aligns.append(m.group(1) if m else None)
                inl = self.t[self.i + 1]
                cur.append(self.inline(inl) if inl.type == "inline" else [])
            self.i += 1
        hdr = header[0] if header else []
        ncol = max([len(hdr)] + [len(r) for r in rows]) if (hdr or rows) else 0
        hdr = hdr + [[] for _ in range(ncol - len(hdr))] if hdr else []
        rows = [r + [[] for _ in range(ncol - len(r))] for r in rows]
        aligns = aligns + [None] * (ncol - len(aligns))
        return Table(hdr, rows, aligns)

    def footnotes(self):
        while self.i < len(self.t) and self.t[self.i].type != "footnote_block_close":
            tok = self.t[self.i]
            if tok.type == "footnote_open":
                label = str(tok.meta.get("label", tok.meta.get("id", "")))
                self.i += 1
                blocks = self.blocks("footnote_close")
                self.doc.footnotes.append(Footnote(label, blocks))
                continue
            self.i += 1
        self.i += 1

    # ── 인라인 토큰 → Span 목록 ──
    def inline(self, tok) -> list[Span]:
        spans: list[Span] = []
        st = dict(bold=0, italic=0, strike=0, href=None)

        def add(text, **kw):
            if not text:
                return
            spans.append(Span(text, bold=st["bold"] > 0, italic=st["italic"] > 0,
                              strike=st["strike"] > 0, href=st["href"], **kw))

        for c in tok.children or []:
            ct = c.type
            if ct == "text":
                add(c.content)
            elif ct == "softbreak" or ct == "hardbreak":
                spans.append(Span(kind="br"))
            elif ct == "strong_open":
                st["bold"] += 1
            elif ct == "strong_close":
                st["bold"] -= 1
            elif ct == "em_open":
                st["italic"] += 1
            elif ct == "em_close":
                st["italic"] -= 1
            elif ct == "s_open":
                st["strike"] += 1
            elif ct == "s_close":
                st["strike"] -= 1
            elif ct == "code_inline":
                add(c.content, code=True)
            elif ct == "link_open":
                st["href"] = c.attrGet("href")
            elif ct == "link_close":
                st["href"] = None
            elif ct == "image":
                spans.append(Span(c.content or "", kind="image", src=c.attrGet("src") or ""))
            elif ct in ("math_inline", "math_inline_double"):
                spans.append(Span(kind="math", src=c.content.strip()))
            elif ct == "footnote_ref":
                label = str(c.meta.get("label", c.meta.get("id", 0) + 1))
                spans.append(Span(f"[{label}]", sup=True))
            elif ct == "html_inline":
                h = c.content
                if 'type="checkbox"' in h or "type='checkbox'" in h:
                    spans.append(Span("x" if "checked" in h else " ", kind="checkbox"))
                elif _BR_RE.fullmatch(h.strip()):
                    spans.append(Span(kind="br"))
                elif h.lower().startswith(("<b>", "<strong>")):
                    st["bold"] += 1
                elif h.lower().startswith(("</b>", "</strong>")):
                    st["bold"] = max(0, st["bold"] - 1)
                # 그 밖의 태그는 버린다
            else:
                if c.content:
                    add(c.content)
        return _fix_unparsed_bold(_merge(spans))


def _alert_map():
    from styles import ALERT_MAP
    return ALERT_MAP


def _take_title(spans: list[Span]) -> tuple[list[Span] | None, list[Span]]:
    """Alert 뒤 첫 줄이 굵은 글씨로 시작하면 그 굵은 부분을 상자 제목으로 떼어 낸다."""
    i = 0
    while i < len(spans) and spans[i].kind == "text" and not spans[i].text.strip():
        i += 1
    if i < len(spans) and spans[i].kind == "br":
        i += 1
    if i < len(spans) and spans[i].bold and spans[i].kind == "text":
        title = []
        while i < len(spans) and spans[i].bold and spans[i].kind == "text":
            title.append(replace(spans[i], bold=False))
            i += 1
        rest = spans[i:]
        while rest and rest[0].kind == "text" and not rest[0].text.strip():
            rest = rest[1:]
        if rest and rest[0].kind == "text":
            rest[0] = replace(rest[0], text=re.sub(r"^\s*[:：]?\s*", "", rest[0].text))
            if not rest[0].text:
                rest = rest[1:]
        if rest and rest[0].kind == "br":
            rest = rest[1:]
        return title, rest
    rest = spans[i:]
    if rest and rest[0].kind == "br":
        rest = rest[1:]
    return None, rest


def _merge(spans: list[Span]) -> list[Span]:
    out: list[Span] = []
    for s in spans:
        if out and s.kind == "text" and out[-1].kind == "text" and \
                replace(out[-1], text="") == replace(s, text=""):
            out[-1] = replace(out[-1], text=out[-1].text + s.text)
        else:
            out.append(s)
    return out


def _fix_unparsed_bold(spans: list[Span]) -> list[Span]:
    """CommonMark 규칙상 풀리지 않은 `**`(예: `**R(깨달음)**을`)를 굵게로 보정한다.

    코드가 아닌 텍스트 스팬에 남은 `**`를 짝지어 그 사이를 굵게 바꾼다.
    짝이 없는 마지막 `**`는 그대로 둔다(잔재 검사에서 경고).
    """
    marks = sum(s.text.count("**") for s in spans if s.kind == "text" and not s.code)
    if marks < 2:
        return spans
    usable = marks - (marks % 2)
    out, toggled, seen = [], False, 0
    for s in spans:
        if s.kind != "text" or s.code or "**" not in s.text:
            out.append(replace(s, bold=s.bold or toggled) if s.kind == "text" else s)
            continue
        parts = s.text.split("**")
        for n, part in enumerate(parts):
            if n:
                if seen < usable:
                    toggled = not toggled
                    seen += 1
                else:
                    part = "**" + part
            if part:
                out.append(replace(s, text=part, bold=s.bold or toggled))
    return _merge(out)


def spans_text(spans: list[Span]) -> str:
    """스팬 목록의 순수 텍스트 (br은 줄바꿈)."""
    parts = []
    for s in spans:
        if s.kind == "text":
            parts.append(s.text)
        elif s.kind == "br":
            parts.append("\n")
        elif s.kind == "math":
            parts.append(s.src)
        elif s.kind == "image":
            parts.append(s.text)
    return "".join(parts)
