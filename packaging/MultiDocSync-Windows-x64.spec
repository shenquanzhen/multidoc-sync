# -*- mode: python ; coding: utf-8 -*-

import os

project_root = os.path.abspath(os.path.join(SPECPATH, ".."))

a = Analysis(
    [os.path.join(project_root, "multi_document_viewer.py")],
    pathex=[project_root],
    binaries=[],
    datas=[(os.path.join(project_root, "assets", "MultiDocSync.png"), "assets")],
    hiddenimports=[
        "pythoncom",
        "pywintypes",
        "win32com",
        "win32com.client",
        "win32gui",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "numpy", "PIL", "pandas", "lxml", "pytz", "dateutil",
        "pymupdf4llm", "bs4", "tkinter", "matplotlib", "scipy",
        "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtNetwork",
        "PySide6.QtPdf", "PySide6.QtPdfWidgets", "PySide6.QtOpenGL",
        "PySide6.QtOpenGLWidgets", "PySide6.QtWebEngineCore",
        "PySide6.QtWebEngineWidgets", "Pythonwin", "win32ui", "win32uiole",
    ],
    noarchive=False,
    optimize=0,
)

# PySide6 wheels contain optional Qt engines that PyInstaller discovers through
# binary dependency scanning even though this widget-only application never
# imports them. Removing them keeps the portable EXE materially smaller.
unused_qt_parts = (
    "qt6network.dll",
    "qt6opengl.dll",
    "qt6openglwidgets.dll",
    "qt6pdf.dll",
    "qt6pdfwidgets.dll",
    "qt6qml",
    "qt6quick",
    "opengl32sw.dll",
    "qdirect2d.dll",
    "qminimal.dll",
    "qoffscreen.dll",
    "qtplugins\\networkinformation\\",
    "qtplugins\\tls\\",
)

def keep_entry(entry):
    normalized = entry[0].replace("/", "\\").lower()
    return not any(part in normalized for part in unused_qt_parts)

a.binaries = [entry for entry in a.binaries if keep_entry(entry)]
a.datas = [entry for entry in a.datas if keep_entry(entry)]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="MultiDocSync-Windows-x64",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(project_root, "assets", "MultiDocSync.ico"),
)
