@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo   Offline install: pandas + openpyxl  (no internet needed)
echo ============================================================
echo.

if not exist "%~dp0offline_wheels" (
    echo [ERROR] "offline_wheels" folder not found next to this .bat
    pause
    exit /b 1
)

python -m pip install --no-index --find-links="%~dp0offline_wheels" pandas==3.0.5 openpyxl==3.1.5

echo.
echo ------------------------------------------------------------------
echo  Done.
echo  If you see "Successfully installed ..." above, dependencies ready.
echo  Then edit the 3 file paths inside requirement_trace.py, and run:
echo      python requirement_trace.py
echo ------------------------------------------------------------------
pause
