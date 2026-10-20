# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 설정: 단일 exe(--onefile --windowed). 실행: pyinstaller --noconfirm --clean md2doc.spec
from PyInstaller.utils.hooks import collect_data_files

EXE_NAME = "MD→문서 변환기"

datas = [("app.ico", ".")]
datas += collect_data_files("latex2mathml")
datas += collect_data_files("docx")

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=["win32com", "win32com.client", "pythoncom", "pywintypes"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        "tkinter", "pytest", "fitz", "pymupdf", "PIL", "numpy",
        "PySide6.QtNetwork", "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtQuickWidgets",
        "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtWebChannel",
        "PySide6.QtMultimedia", "PySide6.Qt3DCore", "PySide6.QtCharts", "PySide6.QtDataVisualization",
        "PySide6.QtPdf", "PySide6.QtSql", "PySide6.QtTest", "PySide6.QtOpenGL",
        "PySide6.QtSvg", "PySide6.QtXml",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=EXE_NAME,
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon="app.ico",
)
