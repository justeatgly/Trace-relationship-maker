@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo   Requirement-trace tools dependency installer
echo   Scripts: requirement_trace.py, make_detection_file.py
echo   Usage:
echo     install_offline_deps.bat          Offline mode (default, needs Python 3.14)
echo     install_offline_deps.bat online   Online mode (needs internet, Python 3.11+)
echo ============================================================
echo.

set "MODE=%~1"
if /I "%MODE%"=="online" goto :online

rem ============ Offline mode (default) ============
if not exist "%~dp0offline_wheels" (
    echo [ERROR] Folder "offline_wheels" not found next to this script.
    goto :fail
)

set "PY="
python -c "import sys; raise SystemExit(0 if sys.version_info[:2]==(3,14) else 1)" >nul 2>nul
if not errorlevel 1 (
    set "PY=python"
) else (
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3.14 -c "import sys" >nul 2>nul
        if not errorlevel 1 set "PY=py -3.14"
    )
)

if not defined PY (
    echo [ERROR] Python 3.14 not found. offline_wheels contains cp314 wheels.
    echo         Install Python 3.14 64-bit and retry, or run:  %~nx0 online
    goto :fail
)

echo [Step 1/3] Using interpreter:
%PY% -V
echo [Step 2/3] Installing pandas==3.0.5 and openpyxl==3.1.5 (both .py tools need them) ...
%PY% -m pip install --no-index --find-links="%~dp0offline_wheels" pandas==3.0.5 openpyxl==3.1.5
if errorlevel 1 (
    echo [ERROR] Offline install failed. Common causes: not Python 3.14, or wheels missing/corrupt.
    goto :fail
)
goto :verify

rem ============ Online mode ============
:online
echo [Online mode] Needs internet. Installing from requirements.txt via PyPI ...
set "PY=python"
python --version >nul 2>nul
if errorlevel 1 set "PY=py"
%PY% -V
%PY% -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo [ERROR] Online install failed. Check your network connection.
    goto :fail
)
goto :verify

rem ============ Verify & usage notes ============
:verify
echo.
echo [Step 3/3] Verifying imports ...
%PY% -c "import sys, pandas, openpyxl; print('  Python   :', sys.version.split()[0]); print('  pandas   :', pandas.__version__); print('  openpyxl :', openpyxl.__version__); print('  json/urllib are stdlib, no extra package needed')"
if errorlevel 1 (
    echo [ERROR] Verification failed. Review the output above.
    goto :fail
)
%PY% -c "from openpyxl import Workbook; from openpyxl.styles import Alignment, Border, Font, PatternFill, Side; print('  openpyxl Workbook/styles OK (used by make_detection_file.py)')"
if errorlevel 1 (
    echo [ERROR] openpyxl Workbook/styles check failed. It is required by make_detection_file.py.
    goto :fail
)

echo.
echo ------------------------------------------------------------
echo  Dependencies installed successfully.
echo.
echo  Next steps:
echo   1. Edit the file paths and DeepSeek API key in requirement_trace.py
echo      (set LLM_API_KEY or the DEEPSEEK_API_KEY env var);
echo   2. Online machine - run keyword filter + LLM review:
echo        %PY% requirement_trace.py
echo   3. Offline machine - set ENABLE_LLM_REVIEW=False in requirement_trace.py
echo      to generate only the intermediate result Excel;
echo   4. Build the Sheet2-like "to-be-detected" Excel from the trace result
echo      (edit config paths first):
echo        %PY% make_detection_file.py
echo ------------------------------------------------------------
pause
exit /b 0

:fail
echo.
echo [FAILED] Dependencies were not installed. Fix the error above and retry.
pause
exit /b 1
