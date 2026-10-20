"""Word로 생성 문서를 열어 「9. 완료 기준」의 배치 항목을 측정한다 (개발용).

사용: python tests/verify_word.py 출력폴더 [--png]
  - 차례가 2쪽 안에 들어가는지
  - PAGEREF 쪽 번호가 실제 제목 위치와 같은지
  - 제목이 쪽 끝에 홀로 남는지
  - 본문 쪽 아래쪽에 큰 빈 공간이 있는지 (PDF 기준)
"""
import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import styles as S  # noqa: E402

WD_PAGE = 3  # wdActiveEndPageNumber


def check_docx(app, path):
    doc = app.Documents.Open(os.path.abspath(path), False, True, False)
    issues = []
    try:
        doc.Repaginate()
        pages = doc.ComputeStatistics(2)
        # 차례 안내 문구가 2쪽에 있어야 한다
        toc_page = None
        heading_info = []
        n = doc.Paragraphs.Count
        paras = doc.Paragraphs
        for i in range(1, n + 1):
            p = paras(i)
            txt = p.Range.Text.strip()
            if txt.startswith("※ 항목을 Ctrl"):
                toc_page = p.Range.Information(WD_PAGE)
            lvl = p.OutlineLevel
            if lvl <= 3 and txt:
                heading_info.append((i, txt[:30], p.Range.Information(WD_PAGE)))
        if toc_page != 2:
            issues.append(f"차례 끝이 {toc_page}쪽 (2쪽이어야 함)")
        # 제목 고립: 제목과 다음 단락의 쪽이 다르면 문제
        for i, txt, pg in heading_info:
            if i < n:
                nxt = paras(i + 1).Range
                if nxt.Information(WD_PAGE) != pg:
                    issues.append(f"제목 고립: '{txt}' {pg}쪽")
        # PAGEREF 결과 = 책갈피 쪽
        bad = 0
        for f in doc.Fields:
            code = f.Code.Text.strip()
            if code.startswith("PAGEREF"):
                name = code.split()[1]
                real = doc.Bookmarks(name).Range.Information(WD_PAGE)
                shown = f.Result.Text.strip()
                if str(real) != shown:
                    bad += 1
                    issues.append(f"쪽 번호 불일치 {name}: 표시 {shown!r} 실제 {real}")
        return pages, issues
    finally:
        doc.Close(0)


def gap_report(pdf_path, png_dir=None):
    import fitz
    d = fitz.open(pdf_path)
    out = []
    for i, page in enumerate(d):
        if png_dir:
            page.get_pixmap(dpi=60).save(os.path.join(png_dir, f"p{i + 1:02d}.png"))
        if i < 2 or i == len(d) - 1:
            continue
        blocks = [b for b in page.get_text("blocks") if b[4].strip()]
        # 바닥글(쪽 번호) 제외
        body = [b for b in blocks if b[3] < page.rect.height - 50]
        if not body:
            continue
        bottom = max(b[3] for b in body)
        ratio = (bottom - S.MARGIN / 20) / (S.BODY_H / 20)
        if ratio < 0.70:
            out.append(f"{i + 1}쪽 본문이 {ratio:.0%}에서 끝남")
    return out, len(d)


def main():
    folder = sys.argv[1]
    png = "--png" in sys.argv
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for f in sorted(glob.glob(os.path.join(folder, "*.docx"))):
            pages, issues = check_docx(app, f)
            pdf = f[:-5] + ".pdf"
            gaps = []
            if os.path.exists(pdf):
                pd = None
                if png:
                    pd = f[:-5] + "_png"
                    os.makedirs(pd, exist_ok=True)
                gaps, _ = gap_report(pdf, pd)
            name = os.path.basename(f)
            print(f"== {name}: {pages}쪽, 문제 {len(issues)}개, 여백 경고 {len(gaps)}개")
            for x in issues + gaps:
                print("   -", x)
    finally:
        app.Quit(0)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
