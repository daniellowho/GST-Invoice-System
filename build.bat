@echo off
REM ============================================================
REM  GSTR 2B Reconciliation - Windows build script
REM  Produces a single-file GSTR2BRecon.exe in the current folder.
REM ============================================================
setlocal

echo.
echo ============================================================
echo   GSTR 2B Reconciliation - Build
echo ============================================================
echo.

REM --- 1. Check Python is available ---------------------------
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python was not found on PATH.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during install, then re-run this file.
    pause
    exit /b 1
)

echo [1/4] Preparing the virtual environment ...
if exist build_env\Scripts\activate.bat (
    echo   Existing environment found - reusing it.
    goto env_ready
)
echo   No environment found - creating build_env ...
python -m venv build_env
if errorlevel 1 (
    echo ERROR: Could not create the virtual environment.
    pause
    exit /b 1
)
:env_ready
call build_env\Scripts\activate.bat

echo.
echo [2/4] Checking dependencies ...
REM Probe whether every required package already imports. If so, skip pip.
python -c "import pandas, openpyxl, rapidfuzz, webview, PyInstaller" >nul 2>&1
if not errorlevel 1 goto deps_ok
echo   Some dependencies are missing - installing from requirements.txt ...
python -m pip install --upgrade pip >nul
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Dependency installation failed. See messages above.
    pause
    exit /b 1
)
goto deps_done
:deps_ok
echo   All dependencies already installed - skipping install.
:deps_done

echo.
echo [3/4] Building single-file executable with PyInstaller ...
REM Source files live in the python\ subfolder.
REM --paths python      lets "import engine" resolve.
REM --add-data ui.html  bundles the HTML (it sits next to build.bat, in the root).
REM --collect-all ...    pulls in native bits for these libraries.
REM --distpath .         outputs the .exe directly here, no dist\ folder created.
pyinstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name "GSTR2BRecon" ^
  --distpath . ^
  --paths python ^
  --add-data "ui.html;." ^
  --collect-all rapidfuzz ^
  --collect-all pandas ^
  --collect-all openpyxl ^
  --collect-all webview ^
  python\app.py

if errorlevel 1 (
    echo ERROR: PyInstaller build failed. See messages above.
    pause
    exit /b 1
)

echo.
echo [4/4] Cleaning up build intermediates ...
if exist build rmdir /s /q build
if exist GSTR2BRecon.spec del GSTR2BRecon.spec

echo.
echo ============================================================
echo   DONE
echo ============================================================
echo   Your application:  GSTR2BRecon.exe
echo.
echo   Double-click that file to run the tool. You can copy the
echo   single .exe anywhere - it carries all its dependencies.
echo ============================================================
echo.
pause
endlocal
