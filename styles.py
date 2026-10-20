"""색·글꼴·크기·간격 상수. 문서 모양은 이 파일 한 곳에서 조정한다.

단위: 크기와 간격은 pt, 길이는 twip(1pt = 20twip, 1cm = 567twip).
"""

# ── 페이지 ─────────────────────────────────────────────
PAGE_W = 11906          # A4 폭 (twip)
PAGE_H = 16838          # A4 높이
MARGIN = 1134           # 상하좌우 여백 2cm
HEADER_DIST = 567       # 머리글·바닥글 거리 1cm
BODY_W = PAGE_W - 2 * MARGIN          # 9638
BODY_H = PAGE_H - 2 * MARGIN          # 14570


def cm(v: float) -> int:
    """cm → twip"""
    return round(v * 567)


# ── 글꼴 ───────────────────────────────────────────────
FONT = "맑은 고딕"
EMOJI_FONT = "Segoe UI Emoji"
# PDF용: Word는 PDF에 이모지를 흑백으로 그리면서 Segoe UI Emoji(8MB)를 통째로 넣는다.
# 흑백 이모지 글꼴인 Segoe UI Symbol을 쓰면 모양은 같고 PDF가 수 MB 작아진다.
PDF_EMOJI_FONT = "Segoe UI Symbol"
SYMBOL_FONT = "Segoe UI Symbol"
MATH_FONT = "Cambria Math"
CODE_FONT = "Consolas"

# ── 색 ─────────────────────────────────────────────────
NAVY = "1F3864"
BLUE = "2E75B6"
ORANGE = "C55A11"
GREEN = "2E7D32"
GRAY_TEXT = "7F7F7F"
DARK_TEXT = "262626"
BORDER = "BFBFBF"
LABEL_BG = "F2F2F2"         # 표지 정보 표 레이블 열
FIRST_COL_BG = "EAF1FB"     # 레이블형 표의 첫 열
H3_BG = "F2F4F8"
CODE_BG = "F3F3F3"
BOLD_COLOR = NAVY           # 본문 굵은 글자
UNCERTAIN_COLOR = "E36C09"  # [?] 표시
LINK_COLOR = "2E75B6"

# ── 글자 크기와 줄 높이 배수 ───────────────────────────
BODY_SIZE = 10.5
BODY_LINE = 1.6
TABLE_SIZE = 9.5
TABLE_LINE = 1.55
COVER_TITLE_SIZE = 24
COVER_TITLE_LINE = 1.3
COVER_SUB_SIZE = 13
TOC_TITLE_SIZE = 20
TOC_CH_SIZE = 12
TOC_SUB_SIZE = 10.5
TOC_NOTE_SIZE = 9
HEADER_SIZE = 8
FOOTER_SIZE = 9
APPENDIX_SIZE = 8.5
CODE_SIZE = 9
FOOTNOTE_SIZE = 9

# 제목: (크기, 줄 배수, 앞 간격, 뒤 간격, 글자색)
HEADINGS = {
    1: dict(size=17, line=1.3, before=26, after=11, color=NAVY),
    2: dict(size=13, line=1.4, before=15, after=7, color=BLUE),
    3: dict(size=11, line=1.5, before=12, after=5, color=NAVY),
}
H1_RULE_COLOR = BLUE
H1_RULE_SIZE = 8            # 1/8pt 단위 (8 = 1pt)

# ── 단락 간격 위계 ─────────────────────────────────────
PARA_AFTER = 6              # 본문 일반 단락
LIST_AFTER = 2              # 같은 목록 안 항목 사이
LIST_LOOSE_AFTER = 4
LIST_END_AFTER = 6          # 목록 마지막 항목
CELL_PARA_AFTER = 2         # 같은 셀 안 단락
BOX_PARA_AFTER = 5.5        # 상자 안 줄글 단락 사이
LABEL_BEFORE = 12           # 상자·셀 안 소제목성 레이블
LABEL_AFTER = 3
BODY_LABEL_BEFORE = 10      # 본문의 굵은 한 줄 레이블
BODY_LABEL_AFTER = 4
SPACER_LINE = 5             # 빈 간격용 단락: 고정 5pt 줄 높이
SPACER_AFTER = 4            # 표·상자 뒤 간격(줄 5pt + 4pt = 9pt)
HR_AFTER = 7                # --- 를 대신하는 간격

# ── 목록 ───────────────────────────────────────────────
LIST_INDENT = 400           # 단계별 왼쪽 들여쓰기 (twip)
LIST_HANGING = 300
CELL_LIST_INDENT = 260
CELL_LIST_HANGING = 200
BULLETS = ["•", "-", "▪"]   # 1/2/3단계 글머리 기호
CHECKBOX_HANGING = cm(0.78)

# ── 표 ─────────────────────────────────────────────────
TABLE_BORDER_SZ = 4         # 0.5pt
HEADER_ROW_BG = NAVY
HEADER_ROW_TEXT = "FFFFFF"
CELL_MAR_TB = cm(0.16)
CELL_MAR_LR = cm(0.23)
LONG_CELL_CHARS = 80        # 이보다 긴 여러 문장 셀은 글머리 목록으로 쪼갬
LABEL_COL_MIN = 0.15
LABEL_COL_MAX = 0.20
ROW_SPLIT_RATIO = 1 / 3     # 행 높이가 쪽 높이의 이 비율을 넘으면 분리 허용

# ── 콜아웃 상자 ────────────────────────────────────────
BOX_BAR_SZ = 24             # 3pt
BOX_MAR = dict(top=cm(0.32), bottom=cm(0.28), left=cm(0.42), right=cm(0.39))
BOX_KEEP_LINES = 6          # 이하이면 쪽 경계에서 나누지 않음

# 종류: (제목 접두어, 막대색, 배경색)
BOX_KINDS = {
    "tip": ("[💡 TIP]", BLUE, "EAF1FB"),
    "warning": ("[⚠️ 주의]", ORANGE, "FFF4E5"),
    "note": ("📌", NAVY, "EEF0F7"),
    "example": ("✅", GREEN, "E9F5EC"),
    "quote": ("", BORDER, "F7F7F7"),
}
ALERT_MAP = {
    "TIP": "tip",
    "WARNING": "warning",
    "CAUTION": "warning",
    "IMPORTANT": "note",
    "NOTE": "note",
    "EXAMPLE": "example",
}

# ── 표지 ───────────────────────────────────────────────
COVER_TOP_SPACE = 150       # 표지 제목 위 여백 (pt)
COVER_RULE_SZ = 18          # 제목 아래 굵은 선 (2.25pt)
COVER_LABEL_COL = 0.22

# front matter 키 → 표지 레이블 (순서대로 표시)
COVER_KEYS = [
    ("provider", "제공 기관"),
    ("instructor", "강사"),
    ("speakers", "발표자"),
    ("author", "작성자"),
    ("date", "날짜"),
    ("duration", "강의 시간"),
    ("parts", "구성"),
    ("source_video", "원본 영상"),
]
# 표지에 넣지 않고 부록에 작게 넣는 파이프라인용 키
PIPELINE_KEYS = {
    "inputs": "입력 자료",
    "numeric_policy": "수치 처리 정책",
    "slip_policy": "실수·불일치 처리 정책",
}
# 표지에도 부록에도 넣지 않는 키
HIDDEN_KEYS = {"title", "subtitle"}

# ── 자동 상자화 대상 절 (제목에 모든 단어가 들어 있으면) ─
AUTO_BOX_SECTIONS = [
    (("강조", "반복"), "note"),
    (("원본", "확인"), "warning"),
]

FEEDBACK_LABEL = "💡 피드백: "
TOC_TITLE = "차례"
TOC_NOTE = "※ 항목을 Ctrl + 클릭하면 해당 위치로 이동합니다."
APPENDIX_TITLE = "부록. 문서 정보"
FOOTNOTE_TITLE = "각주"
