@echo off
chcp 65001 >nul
title Gemini for Excel - HTTPS Server

echo ============================================================
echo         KHOI DONG MAYS CHU LOCAL CHO GEMINI EXCEL ADD-IN
echo ============================================================
echo.

:: Kiem tra Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [LOI] Khong tim thay Python tren may tinh cua ban!
    echo Vui long cai dat Python de khoi dong may chu dev.
    pause
    exit /b 1
)

echo Dang khoi dong may chu HTTPS tren cong 3000...
echo.

python "%~dp0server.py"

pause
