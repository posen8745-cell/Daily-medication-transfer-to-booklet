@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

REM === 切到 BAT 所在資料夾 ===
cd /d "%~dp0"

REM === 使用系統 Python（自動找）===
set PYTHON=python

echo ===== 開始執行 管制藥轉入 =====
echo 目前資料夾：
cd
echo.

REM === 檢查 Python 是否存在 ===
where %PYTHON% >nul 2>&1
if errorlevel 1 (
    echo ❌ 找不到 Python
    echo 請確認已安裝 Python 3.11，且可在 CMD 執行 python
    pause
    exit /b 1
)

echo 使用的 Python：
%PYTHON% --version
echo.

REM === 執行主程式 ===
%PYTHON% main_win.py --prefer stream

echo.
echo ===== 程式已結束 =====
pause
