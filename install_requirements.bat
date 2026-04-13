@echo off
echo =========================================
echo  安裝 BurnIn Monitor 所需套件 v1.6
echo =========================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.10+
    echo 下載網址：https://www.python.org/downloads/
    pause
    exit /b 1
)

echo 正在安裝 Python 套件...
pip install mss Pillow pytesseract openpyxl numpy pywinauto pywin32 pyserial

echo.
echo =========================================
echo  重要：請另外安裝 Tesseract OCR 引擎
echo  （使用視窗直讀模式時可不安裝）
echo =========================================
echo  下載網址：
echo  https://github.com/UB-Mannheim/tesseract/wiki
echo  (下載 tesseract-ocr-w64-setup-*.exe 並安裝)
echo =========================================
echo.
echo =========================================
echo  重要：SMART 溫度功能需要管理員權限
echo  程式啟動時會自動彈出 UAC 提示視窗
echo  請點「是」以允許管理員權限
echo =========================================
echo.
echo 安裝完成！可執行 run.bat 啟動程式。
pause
