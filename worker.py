"""변환 작업: Qt와 무관한 Converter + GUI용 QThread 래퍼."""
from __future__ import annotations

import os
import shutil
import tempfile
import traceback
from dataclasses import dataclass

from docx_builder import BuildOptions, build_docx, find_mml2omml
from md_parser import MdParseError, parse_file
from pdf_export import ExportError, WordSession, find_soffice, lo_process


@dataclass
class Job:
    row: int
    md_path: str
    out_path: str       # 최종 .docx 또는 .pdf


class ConvertError(Exception):
    pass


class Converter:
    """파일 하나씩 변환한다. Word 인스턴스는 첫 사용 때 띄우고 close()에서 닫는다."""

    def __init__(self, fmt: str, engine: str, auto_number_bold=True, auto_box=True, log=print):
        self.fmt = fmt
        self.engine = engine
        self.auto_number_bold = auto_number_bold
        self.auto_box = auto_box
        self.log = log
        self.word: WordSession | None = None
        self.soffice = find_soffice() if engine == "libreoffice" else None
        self.mml2omml = find_mml2omml()

    def close(self):
        if self.word is not None:
            self.word.close()
            self.word = None

    def _word(self) -> WordSession:
        if self.word is None:
            self.word = WordSession().__enter__()
        return self.word

    def _opts(self, update_fields: bool) -> BuildOptions:
        return BuildOptions(auto_number_bold=self.auto_number_bold, auto_box=self.auto_box,
                            update_fields_on_open=update_fields, mml2omml=self.mml2omml,
                            for_pdf=self.fmt == "pdf")

    def convert(self, job: Job) -> list[str]:
        """성공하면 경고 목록을 돌려주고, 실패하면 ConvertError."""
        if self.fmt == "pdf" and self.engine == "none":
            raise ConvertError("PDF 변환 엔진(Word 또는 LibreOffice)이 없습니다")
        try:
            md = parse_file(job.md_path)
        except MdParseError as e:
            raise ConvertError(str(e)) from e
        except OSError as e:
            raise ConvertError(f"파일을 읽을 수 없습니다: {e.strerror or e}") from e
        except Exception as e:  # noqa: BLE001
            raise ConvertError(f"마크다운 해석 실패: {e}") from e

        stem = os.path.splitext(os.path.basename(job.md_path))[0]
        work = tempfile.mkdtemp(prefix="md2doc_")
        try:
            tmp_docx = os.path.join(work, stem + ".docx")
            tmp_pdf = os.path.join(work, stem + ".pdf") if self.fmt == "pdf" else None
            try:
                result = build_docx(md, job.md_path, tmp_docx, self._opts(self.engine == "none"))
            except Exception as e:  # noqa: BLE001
                self.log("    " + traceback.format_exc().strip().splitlines()[-1])
                raise ConvertError(f"문서 생성 실패: {e}") from e
            warnings = list(result.warnings)

            if self.engine == "word":
                try:
                    self._word().process(tmp_docx, tmp_pdf)
                except ExportError as e:
                    self.close()      # Word가 죽었을 수 있으니 다음 파일에서 새로 띄운다
                    if tmp_pdf:
                        raise ConvertError(str(e)) from e
                    # docx는 Word 없이도 쓸 수 있게 「열 때 필드 갱신」으로 다시 만든다
                    build_docx(md, job.md_path, tmp_docx, self._opts(True))
                    warnings.append(f"쪽 번호 갱신 실패, Word에서 열 때 갱신됩니다 ({e})")
            elif self.engine == "libreoffice" and self.soffice:
                try:
                    pages = lo_process(self.soffice, tmp_docx, result.headings, tmp_pdf)
                    missing = len(result.headings) - len(pages)
                    if missing > 0:
                        warnings.append(f"쪽 번호를 찾지 못한 제목 {missing}개")
                except ExportError as e:
                    if tmp_pdf:
                        raise ConvertError(str(e)) from e
                    warnings.append(f"쪽 번호 계산 실패: {e}")

            src = tmp_pdf if tmp_pdf else tmp_docx
            self._place(src, job.out_path)
            return warnings
        finally:
            shutil.rmtree(work, ignore_errors=True)

    @staticmethod
    def _place(src: str, dst: str):
        os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
        try:
            if os.path.exists(dst):
                os.remove(dst)
            shutil.move(src, dst)
        except PermissionError as e:
            raise ConvertError("출력 파일이 다른 프로그램에서 열려 있어 저장할 수 없습니다") from e
        except OSError as e:
            raise ConvertError(f"저장 실패: {e.strerror or e}") from e


# ── Qt 작업 스레드 ─────────────────────────────────────
try:
    from PySide6.QtCore import QThread, Signal
except ImportError:  # CLI·테스트에서는 Qt 없이도 쓸 수 있게
    QThread = object
    Signal = None


if Signal is not None:
    class ConvertWorker(QThread):
        file_started = Signal(int)
        file_done = Signal(int, str, str)       # row, 상태(완료/실패), 메시지
        log = Signal(str)
        progress = Signal(int, int)
        all_done = Signal(int, int, bool)       # 성공, 실패, 취소됨

        def __init__(self, jobs: list[Job], fmt: str, engine: str, auto_number_bold: bool,
                     auto_box: bool):
            super().__init__()
            self.jobs = jobs
            self.fmt = fmt
            self.engine = engine
            self.auto_number_bold = auto_number_bold
            self.auto_box = auto_box
            self._cancel = False

        def cancel(self):
            self._cancel = True

        def run(self):
            ok = fail = 0
            conv = Converter(self.fmt, self.engine, self.auto_number_bold, self.auto_box,
                             log=self.log.emit)
            try:
                for n, job in enumerate(self.jobs):
                    if self._cancel:
                        break
                    self.file_started.emit(job.row)
                    name = os.path.basename(job.md_path)
                    self.log.emit(f"변환 시작: {name}")
                    try:
                        warnings = conv.convert(job)
                        ok += 1
                        for w in warnings:
                            self.log.emit(f"    ⚠ {w}")
                        self.log.emit(f"    완료 → {job.out_path}")
                        self.file_done.emit(job.row, "완료", "\n".join(warnings))
                    except ConvertError as e:
                        fail += 1
                        self.log.emit(f"    실패: {e}")
                        self.file_done.emit(job.row, "실패", str(e))
                    except Exception as e:  # noqa: BLE001 - 한 파일 실패가 전체를 멈추면 안 된다
                        fail += 1
                        self.log.emit(f"    실패(예상하지 못한 오류): {e}")
                        self.file_done.emit(job.row, "실패", f"예상하지 못한 오류: {e}")
                    self.progress.emit(n + 1, len(self.jobs))
            finally:
                conv.close()
            self.all_done.emit(ok, fail, self._cancel)
