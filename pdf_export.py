"""쪽 번호(PAGEREF) 갱신과 PDF 변환: Word COM → LibreOffice → 없음."""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WD_EXPORT_PDF = 17
WD_EXPORT_HEADING_BOOKMARKS = 1


# ── 엔진 탐지 ──────────────────────────────────────────
def word_available() -> bool:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            return True
    except OSError:
        return False
    except ImportError:
        return False


def find_soffice() -> str | None:
    cands = []
    try:
        import winreg
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for key in (r"SOFTWARE\LibreOffice\UNO\InstallPath",
                        r"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath"):
                try:
                    with winreg.OpenKey(hive, key) as k:
                        cands.append(os.path.join(winreg.QueryValueEx(k, "")[0], "soffice.exe"))
                except OSError:
                    pass
    except ImportError:
        pass
    for root in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        cands.append(os.path.join(root, "LibreOffice", "program", "soffice.exe"))
    w = shutil.which("soffice")
    if w:
        cands.append(w)
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


def detect_engine() -> str:
    """'word' | 'libreoffice' | 'none'"""
    if word_available():
        return "word"
    if find_soffice():
        return "libreoffice"
    return "none"


ENGINE_LABEL = {"word": "Word", "libreoffice": "LibreOffice", "none": "없음"}


class ExportError(Exception):
    pass


# ── Word COM ───────────────────────────────────────────
class WordSession:
    """이 프로그램이 띄운 Word 인스턴스(DispatchEx)만 다룬다. 작업 스레드 안에서 만든다."""

    def __init__(self):
        self.app = None
        self._co = False

    def __enter__(self):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        self._co = True
        try:
            self.app = win32com.client.DispatchEx("Word.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = 0
        except Exception as e:  # noqa: BLE001
            self.close()
            raise ExportError(f"Word를 시작할 수 없습니다: {e}") from e
        return self

    def __exit__(self, *exc):
        self.close()

    def close(self):
        if self.app is not None:
            try:
                self.app.Quit(0)
            except Exception:  # noqa: BLE001
                pass
            self.app = None
        if self._co:
            import pythoncom
            pythoncom.CoUninitialize()
            self._co = False

    def process(self, docx_path: str, pdf_path: str | None = None):
        """필드(쪽 번호) 갱신 → 저장 → (선택) PDF 내보내기."""
        doc = None
        try:
            # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
            doc = self.app.Documents.Open(os.path.abspath(docx_path), False, False, False)
            doc.Repaginate()
            doc.Fields.Update()
            doc.Save()
            if pdf_path:
                # ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor,
                #   Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks,
                #   DocStructureTags, BitmapMissingFonts, UseISO19005_1)
                doc.ExportAsFixedFormat(os.path.abspath(pdf_path), WD_EXPORT_PDF, False, 0,
                                        0, 1, 1, 0, True, True, WD_EXPORT_HEADING_BOOKMARKS,
                                        True, True, False)
        except Exception as e:  # noqa: BLE001
            raise ExportError(f"Word 처리 실패: {_com_msg(e)}") from e
        finally:
            if doc is not None:
                try:
                    doc.Close(0)
                except Exception:  # noqa: BLE001
                    pass


def _com_msg(e) -> str:
    try:
        args = e.args
        if len(args) >= 3 and args[2] and len(args[2]) >= 3 and args[2][2]:
            return str(args[2][2]).strip()
    except Exception:  # noqa: BLE001
        pass
    return str(e)


# ── LibreOffice ────────────────────────────────────────
def lo_convert(soffice: str, docx_path: str, out_dir: str) -> str:
    profile = tempfile.mkdtemp(prefix="md2doc_lo_")
    try:
        cmd = [soffice, f"-env:UserInstallation=file:///{profile.replace(os.sep, '/')}",
               "--headless", "--norestore", "--convert-to", "pdf", "--outdir", out_dir, docx_path]
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        res = subprocess.run(cmd, capture_output=True, timeout=300, creationflags=flags)
        pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        if res.returncode != 0 or not os.path.exists(pdf):
            raise ExportError("LibreOffice 변환 실패: " +
                              (res.stderr or res.stdout).decode("utf-8", "replace")[-300:])
        return pdf
    except subprocess.TimeoutExpired as e:
        raise ExportError("LibreOffice 변환 시간 초과") from e
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", s)


def find_heading_pages(pdf_path: str, headings, first_body_page: int = 3) -> dict[str, int]:
    """PDF에서 각 제목이 처음 나오는 쪽(1부터)을 순서대로 찾는다."""
    from pypdf import PdfReader
    reader = PdfReader(pdf_path)
    pages = [_norm(p.extract_text() or "") for p in reader.pages]
    result, cur = {}, first_body_page - 1
    for h in headings:
        key = _norm(h.text)[:40]
        for i in range(cur, len(pages)):
            if key and key in pages[i]:
                result[h.bookmark] = i + 1
                cur = i
                break
    return result


def fill_pageref(docx_path: str, pages: dict[str, int]):
    """PAGEREF 필드의 결과값(캐시)에 쪽 번호를 기록한다."""
    tmp = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = _fill_xml(data, pages)
            zout.writestr(item, data)
    os.replace(tmp, docx_path)


def _fill_xml(data: bytes, pages: dict[str, int]) -> bytes:
    root = etree.fromstring(data)
    q = lambda t: f"{{{W}}}{t}"  # noqa: E731
    for instr in root.iter(q("instrText")):
        m = re.search(r"PAGEREF\s+(\S+)", instr.text or "")
        if not m or m.group(1) not in pages:
            continue
        r = instr.getparent()
        state, nxt = "instr", r.getnext()
        while nxt is not None:
            fc = nxt.find(q("fldChar"))
            if fc is not None and fc.get(q("fldCharType")) == "separate":
                state = "result"
            elif fc is not None and fc.get(q("fldCharType")) == "end":
                break
            elif state == "result":
                t = nxt.find(q("t"))
                if t is not None:
                    t.text = str(pages[m.group(1)])
                    break
            nxt = nxt.getnext()
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def lo_process(soffice: str, docx_path: str, headings, pdf_path: str | None = None):
    """2패스: PDF 렌더링 → 제목 쪽 탐색 → PAGEREF 캐시 기록 → (PDF면) 재변환."""
    work = tempfile.mkdtemp(prefix="md2doc_pdf_")
    try:
        first = lo_convert(soffice, docx_path, work)
        pages = find_heading_pages(first, headings)
        fill_pageref(docx_path, pages)
        if pdf_path:
            os.remove(first)
            second = lo_convert(soffice, docx_path, work)
            shutil.move(second, pdf_path)
        return pages
    finally:
        shutil.rmtree(work, ignore_errors=True)
