"""MD→문서 변환기: PySide6 GUI 진입점.

명령줄 변환(창 없이): app.py --convert a.md b.md [--format docx|pdf] [--out 폴더] [--log 로그.txt]
"""
from __future__ import annotations

import os
import sys
import time

APP_TITLE = "MD→문서 변환기"
APP_VERSION = "1.0.0"


def resource_path(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def human_size(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1024 * 1024:
        return f"{n / 1024:.1f} KB"
    return f"{n / 1024 / 1024:.1f} MB"


# ── 명령줄 모드 ────────────────────────────────────────
def run_cli(argv: list[str]) -> int:
    from pdf_export import detect_engine
    from worker import ConvertError, Converter, Job

    files, fmt, out, log_path = [], "docx", None, None
    it = iter(argv)
    for a in it:
        if a == "--convert":
            continue
        if a == "--format":
            fmt = next(it, "docx").lower()
        elif a == "--out":
            out = next(it, None)
        elif a == "--log":
            log_path = next(it, None)
        else:
            files.append(a)
    lines = []

    def log(msg):
        lines.append(msg)
        try:
            print(msg)
        except (OSError, UnicodeEncodeError, AttributeError):
            pass

    engine = detect_engine()
    log(f"PDF 변환 엔진: {engine}")
    conv = Converter(fmt, engine, log=log)
    fails = 0
    try:
        for i, f in enumerate(files):
            d = out or os.path.dirname(os.path.abspath(f))
            dst = os.path.join(d, os.path.splitext(os.path.basename(f))[0] + "." + fmt)
            try:
                for w in conv.convert(Job(i, f, dst)):
                    log(f"  경고: {w}")
                log(f"완료: {dst}")
            except ConvertError as e:
                fails += 1
                log(f"실패: {f}: {e}")
    finally:
        conv.close()
        if log_path:
            with open(log_path, "w", encoding="utf-8") as fh:
                fh.write("\n".join(lines) + "\n")
    return 1 if fails else 0


# ── GUI ────────────────────────────────────────────────
def run_gui(argv: list[str]) -> int:
    from PySide6.QtCore import Qt, QUrl
    from PySide6.QtGui import QAction, QColor, QDesktopServices, QIcon  # noqa: F401
    from PySide6.QtWidgets import (QAbstractItemView, QApplication, QButtonGroup, QCheckBox,
                                   QFileDialog, QGroupBox, QHBoxLayout, QHeaderView, QLabel,
                                   QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit,
                                   QProgressBar, QPushButton, QRadioButton, QTableWidget,
                                   QTableWidgetItem, QVBoxLayout, QWidget)

    import settings
    from pdf_export import ENGINE_LABEL, detect_engine
    from worker import ConvertWorker, Job

    STATUS_COLOR = {"대기": "#595959", "변환 중": "#2E75B6", "완료": "#2E7D32",
                    "실패": "#C00000", "건너뜀": "#7F7F7F"}

    class MainWindow(QMainWindow):
        COLS = ["파일명", "폴더 경로", "크기", "상태"]

        def __init__(self):
            super().__init__()
            self.setWindowTitle(APP_TITLE)
            icon = resource_path("app.ico")
            if os.path.exists(icon):
                self.setWindowIcon(QIcon(icon))
            self.setAcceptDrops(True)
            self.cfg = settings.load()
            self.engine = detect_engine()
            self.worker = None
            self.last_out_dirs: list[str] = []
            self._build_ui()
            self._restore()

        # ── 화면 구성 ──
        def _build_ui(self):
            central = QWidget()
            root = QVBoxLayout(central)

            box_files = QGroupBox("변환할 파일")
            lf = QVBoxLayout(box_files)
            self.table = QTableWidget(0, len(self.COLS))
            self.table.setHorizontalHeaderLabels(self.COLS)
            self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
            self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.table.verticalHeader().setVisible(False)
            self.table.setAlternatingRowColors(True)
            hh = self.table.horizontalHeader()
            hh.setSectionResizeMode(0, QHeaderView.Interactive)
            hh.setSectionResizeMode(1, QHeaderView.Stretch)
            hh.setSectionResizeMode(2, QHeaderView.ResizeToContents)
            hh.setSectionResizeMode(3, QHeaderView.ResizeToContents)
            self.table.setColumnWidth(0, 260)
            lf.addWidget(self.table)
            row = QHBoxLayout()
            self.btn_add = QPushButton("파일 추가")
            self.btn_del = QPushButton("선택 삭제")
            self.btn_clear = QPushButton("모두 지우기")
            for b in (self.btn_add, self.btn_del, self.btn_clear):
                row.addWidget(b)
            row.addStretch(1)
            hint = QLabel("md 파일이나 폴더를 이 창에 끌어다 놓아도 추가됩니다.")
            hint.setStyleSheet("color:#7F7F7F")
            row.addWidget(hint)
            lf.addLayout(row)
            root.addWidget(box_files, 3)

            box_out = QGroupBox("출력 설정")
            lo = QVBoxLayout(box_out)
            r1 = QHBoxLayout()
            r1.addWidget(QLabel("출력 폴더"))
            self.ed_out = QLineEdit()
            self.ed_out.setPlaceholderText("변환한 문서를 저장할 폴더")
            self.btn_browse = QPushButton("찾아보기")
            r1.addWidget(self.ed_out, 1)
            r1.addWidget(self.btn_browse)
            lo.addLayout(r1)
            self.chk_same = QCheckBox("원본 md와 같은 폴더에 저장")
            lo.addWidget(self.chk_same)
            r2 = QHBoxLayout()
            r2.addWidget(QLabel("출력 형식"))
            self.rb_docx = QRadioButton("DOCX (Word 문서)")
            self.rb_pdf = QRadioButton("PDF")
            self.fmt_group = QButtonGroup(self)
            self.fmt_group.addButton(self.rb_docx)
            self.fmt_group.addButton(self.rb_pdf)
            r2.addWidget(self.rb_docx)
            r2.addWidget(self.rb_pdf)
            r2.addStretch(1)
            lo.addLayout(r2)
            r3 = QHBoxLayout()
            r3.addWidget(QLabel("변환 옵션"))
            self.chk_num = QCheckBox("수치 자동 강조")
            self.chk_num.setToolTip("40~50초, 100만 원, 85%처럼 숫자와 단위가 붙은 표현을 굵게 표시합니다.")
            self.chk_box = QCheckBox("특정 절 자동 상자화")
            self.chk_box.setToolTip("「강의자가 강조·반복한 내용」은 📌 상자로, "
                                    "「원본 확인이 필요한 부분」은 ⚠️ 주의 상자로 만듭니다.")
            r3.addWidget(self.chk_num)
            r3.addWidget(self.chk_box)
            r3.addStretch(1)
            lo.addLayout(r3)
            root.addWidget(box_out)

            r4 = QHBoxLayout()
            self.btn_start = QPushButton("변환 시작")
            self.btn_start.setMinimumHeight(34)
            self.btn_start.setStyleSheet("QPushButton{font-weight:bold;}")
            self.btn_cancel = QPushButton("취소")
            self.btn_cancel.setEnabled(False)
            self.btn_open = QPushButton("출력 폴더 열기")
            self.btn_open.setVisible(False)
            self.progress = QProgressBar()
            self.progress.setValue(0)
            r4.addWidget(self.btn_start)
            r4.addWidget(self.btn_cancel)
            r4.addWidget(self.progress, 1)
            r4.addWidget(self.btn_open)
            root.addLayout(r4)

            box_log = QGroupBox("로그")
            ll = QVBoxLayout(box_log)
            self.log_view = QPlainTextEdit()
            self.log_view.setReadOnly(True)
            self.log_view.setMaximumBlockCount(5000)
            ll.addWidget(self.log_view)
            root.addWidget(box_log, 2)

            self.setCentralWidget(central)
            self.engine_label = QLabel(f"PDF 변환 엔진: {ENGINE_LABEL[self.engine]}")
            self.statusBar().addPermanentWidget(self.engine_label)
            self.statusBar().showMessage(f"버전 {APP_VERSION}")

            self.btn_add.clicked.connect(self.on_add)
            self.btn_del.clicked.connect(self.on_delete)
            self.btn_clear.clicked.connect(self.on_clear)
            self.btn_browse.clicked.connect(self.on_browse)
            self.chk_same.toggled.connect(self._sync_out_enabled)
            self.btn_start.clicked.connect(self.on_start)
            self.btn_cancel.clicked.connect(self.on_cancel)
            self.btn_open.clicked.connect(self.on_open_folder)

            if self.engine == "none":
                self.rb_pdf.setEnabled(False)
                self.rb_pdf.setToolTip("Word 또는 LibreOffice가 설치되어 있어야 PDF로 변환할 수 있습니다.\n"
                                       "DOCX로 변환한 뒤 Word에서 열면 쪽 번호가 갱신됩니다.")

        def _restore(self):
            c = self.cfg
            self.ed_out.setText(c.get("out_dir") or "")
            self.chk_same.setChecked(bool(c.get("same_folder")))
            if c.get("format") == "pdf" and self.rb_pdf.isEnabled():
                self.rb_pdf.setChecked(True)
            else:
                self.rb_docx.setChecked(True)
            self.chk_num.setChecked(bool(c.get("auto_number_bold", True)))
            self.chk_box.setChecked(bool(c.get("auto_box", True)))
            self._sync_out_enabled()
            geo = c.get("geometry")
            if isinstance(geo, list) and len(geo) == 4:
                self.setGeometry(*geo)
            else:
                self.resize(920, 720)

        def _save(self):
            g = self.geometry()
            self.cfg.update(out_dir=self.ed_out.text().strip(), same_folder=self.chk_same.isChecked(),
                            format="pdf" if self.rb_pdf.isChecked() else "docx",
                            auto_number_bold=self.chk_num.isChecked(),
                            auto_box=self.chk_box.isChecked(),
                            geometry=[g.x(), g.y(), g.width(), g.height()])
            settings.save(self.cfg)

        def _sync_out_enabled(self):
            on = not self.chk_same.isChecked()
            self.ed_out.setEnabled(on)
            self.btn_browse.setEnabled(on)

        def log(self, msg: str):
            self.log_view.appendPlainText(f"[{time.strftime('%H:%M:%S')}] {msg}")

        # ── 파일 목록 ──
        def paths(self) -> list[str]:
            return [self.table.item(r, 0).data(Qt.UserRole) for r in range(self.table.rowCount())]

        def add_files(self, files):
            have = {os.path.normcase(os.path.abspath(p)) for p in self.paths()}
            added = 0
            for f in files:
                f = os.path.abspath(f)
                key = os.path.normcase(f)
                if key in have or not f.lower().endswith((".md", ".markdown")) or not os.path.isfile(f):
                    continue
                have.add(key)
                r = self.table.rowCount()
                self.table.insertRow(r)
                it = QTableWidgetItem(os.path.basename(f))
                it.setData(Qt.UserRole, f)
                it.setToolTip(f)
                self.table.setItem(r, 0, it)
                self.table.setItem(r, 1, QTableWidgetItem(os.path.dirname(f)))
                try:
                    size = human_size(os.path.getsize(f))
                except OSError:
                    size = "?"
                si = QTableWidgetItem(size)
                si.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(r, 2, si)
                self.set_status(r, "대기")
                added += 1
            if added:
                self.log(f"파일 {added}개 추가")
            self.btn_open.setVisible(False)

        def set_status(self, row, status, tip=""):
            it = QTableWidgetItem(status)
            it.setForeground(QColor(STATUS_COLOR.get(status, "#000000")))
            it.setToolTip(tip or status)
            it.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 3, it)

        def on_add(self):
            files, _ = QFileDialog.getOpenFileNames(self, "md 파일 선택", self._start_dir(),
                                                    "마크다운 (*.md *.markdown);;모든 파일 (*.*)")
            self.add_files(files)

        def _start_dir(self):
            ps = self.paths()
            return os.path.dirname(ps[-1]) if ps else (self.ed_out.text() or os.path.expanduser("~"))

        def on_delete(self):
            rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
            for r in rows:
                self.table.removeRow(r)

        def on_clear(self):
            self.table.setRowCount(0)
            self.progress.setValue(0)
            self.btn_open.setVisible(False)

        def on_browse(self):
            d = QFileDialog.getExistingDirectory(self, "출력 폴더 선택", self.ed_out.text() or self._start_dir())
            if d:
                self.ed_out.setText(os.path.normpath(d))

        # ── 끌어다 놓기 ──
        def dragEnterEvent(self, e):
            if e.mimeData().hasUrls() and self.worker is None:
                e.acceptProposedAction()

        def dropEvent(self, e):
            files = []
            for url in e.mimeData().urls():
                p = url.toLocalFile()
                if os.path.isdir(p):
                    for dp, _, fns in os.walk(p):
                        files += [os.path.join(dp, f) for f in sorted(fns)
                                  if f.lower().endswith((".md", ".markdown"))]
                else:
                    files.append(p)
            self.add_files(files)

        # ── 변환 ──
        def on_start(self):
            if self.worker is not None:
                return
            n = self.table.rowCount()
            if not n:
                QMessageBox.information(self, APP_TITLE, "변환할 md 파일을 먼저 추가하세요.")
                return
            same = self.chk_same.isChecked()
            out_dir = self.ed_out.text().strip()
            if not same:
                if not out_dir:
                    QMessageBox.warning(self, APP_TITLE, "출력 폴더를 지정하거나 「원본 md와 같은 폴더에 저장」을 선택하세요.")
                    return
                try:
                    os.makedirs(out_dir, exist_ok=True)
                except OSError as e:
                    QMessageBox.warning(self, APP_TITLE, f"출력 폴더를 만들 수 없습니다.\n{e}")
                    return
            fmt = "pdf" if self.rb_pdf.isChecked() else "docx"
            self._save()

            jobs, seen, overwrite_all = [], set(), False
            for r in range(n):
                md = self.table.item(r, 0).data(Qt.UserRole)
                d = os.path.dirname(md) if same else out_dir
                dst = os.path.join(d, os.path.splitext(os.path.basename(md))[0] + "." + fmt)
                key = os.path.normcase(os.path.abspath(dst))
                if key in seen:
                    self.set_status(r, "실패", "같은 이름의 출력 파일이 이번 작업에 이미 있습니다")
                    continue
                seen.add(key)
                if os.path.exists(dst) and not overwrite_all:
                    ans = self._ask_overwrite(dst)
                    if ans is None:
                        self.log("변환을 시작하지 않았습니다.")
                        return
                    if ans == "skip":
                        self.set_status(r, "건너뜀", "이미 있는 파일을 건너뛰었습니다")
                        continue
                    if ans == "all":
                        overwrite_all = True
                self.set_status(r, "대기")
                jobs.append(Job(r, md, dst))
            if not jobs:
                self.log("변환할 파일이 없습니다.")
                return

            self.last_out_dirs = []
            self.progress.setRange(0, len(jobs))
            self.progress.setValue(0)
            self._set_running(True)
            self.log(f"── 변환 시작: {len(jobs)}개, 형식 {fmt.upper()}, 엔진 {ENGINE_LABEL[self.engine]} ──")
            self.worker = ConvertWorker(jobs, fmt, self.engine, self.chk_num.isChecked(),
                                        self.chk_box.isChecked())
            self.worker.file_started.connect(lambda r: self.set_status(r, "변환 중"))
            self.worker.file_done.connect(self._on_file_done)
            self.worker.log.connect(self.log)
            self.worker.progress.connect(lambda d, t: self.progress.setValue(d))
            self.worker.all_done.connect(self._on_all_done)
            self._jobs = {j.row: j for j in jobs}
            self.worker.start()

        def _ask_overwrite(self, path):
            box = QMessageBox(self)
            box.setWindowTitle(APP_TITLE)
            box.setIcon(QMessageBox.Question)
            box.setText(f"같은 이름의 파일이 이미 있습니다.\n\n{path}")
            b_over = box.addButton("덮어쓰기", QMessageBox.AcceptRole)
            b_skip = box.addButton("건너뛰기", QMessageBox.RejectRole)
            b_all = box.addButton("모두 덮어쓰기", QMessageBox.AcceptRole)
            b_cancel = box.addButton("작업 취소", QMessageBox.DestructiveRole)
            box.setEscapeButton(b_skip)
            box.exec()
            c = box.clickedButton()
            if c is b_over:
                return "over"
            if c is b_all:
                return "all"
            if c is b_cancel:
                return None
            return "skip"

        def _on_file_done(self, row, status, msg):
            self.set_status(row, status, msg if msg else status)
            if status == "완료":
                d = os.path.dirname(self._jobs[row].out_path)
                if d not in self.last_out_dirs:
                    self.last_out_dirs.append(d)

        def _on_all_done(self, ok, fail, canceled):
            self.worker.wait()
            self.worker = None
            self._set_running(False)
            head = "변환을 취소했습니다." if canceled else "변환을 마쳤습니다."
            self.log(f"── {head} 성공 {ok}, 실패 {fail} ──")
            self.btn_open.setVisible(bool(self.last_out_dirs))
            self.statusBar().showMessage(f"성공 {ok}, 실패 {fail}", 0)
            QMessageBox.information(self, APP_TITLE, f"{head}\n\n성공 {ok}개, 실패 {fail}개")

        def on_cancel(self):
            if self.worker is not None:
                self.worker.cancel()
                self.btn_cancel.setEnabled(False)
                self.log("취소 요청: 지금 변환 중인 파일까지 마치고 멈춥니다.")

        def on_open_folder(self):
            for d in self.last_out_dirs[:5]:
                QDesktopServices.openUrl(QUrl.fromLocalFile(d))

        def _set_running(self, running):
            for w in (self.btn_add, self.btn_del, self.btn_clear, self.btn_start, self.chk_same,
                      self.rb_docx, self.rb_pdf, self.chk_num, self.chk_box):
                w.setEnabled(not running)
            if not running and self.engine == "none":
                self.rb_pdf.setEnabled(False)
            self.ed_out.setEnabled(not running and not self.chk_same.isChecked())
            self.btn_browse.setEnabled(not running and not self.chk_same.isChecked())
            self.btn_cancel.setEnabled(running)
            if running:
                self.btn_open.setVisible(False)

        def closeEvent(self, e):
            if self.worker is not None:
                if QMessageBox.question(self, APP_TITLE, "변환 중입니다. 취소하고 종료할까요?") != QMessageBox.Yes:
                    e.ignore()
                    return
                self.worker.cancel()
                self.worker.wait(120000)
            self._save()
            e.accept()

    app = QApplication(sys.argv[:1])
    app.setApplicationName(APP_TITLE)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    win.add_files([a for a in argv if not a.startswith("--")])
    win.log(f"{APP_TITLE} {APP_VERSION} 시작. PDF 변환 엔진: {ENGINE_LABEL[win.engine]}")
    shot = os.environ.get("MD2DOC_SCREENSHOT")   # 개발용: 창을 그림 파일로 저장하고 종료
    if shot:
        from PySide6.QtCore import QTimer
        QTimer.singleShot(800, lambda: (win.grab().save(shot), app.quit()))
    return app.exec()


def main():
    argv = sys.argv[1:]
    if "--convert" in argv:
        sys.exit(run_cli(argv))
    sys.exit(run_gui(argv))


if __name__ == "__main__":
    main()
