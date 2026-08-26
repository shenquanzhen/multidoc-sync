@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

py -3 -c "import fitz, pythoncom, win32com.client, win32gui; from PySide6.QtWidgets import QApplication" >nul 2>&1
if errorlevel 1 (
    echo 无法启动：缺少 Python 3、PySide6、PyMuPDF 或 pywin32。
    echo 请先运行：py -3 -m pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

start "" /b pyw -3 "%~dp0multi_document_viewer.py" %*
exit /b 0
