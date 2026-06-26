@echo off
REM ============================================================
REM  GSTR 2B Reconciliation - Windows build script
REM  Produces a single-file GSTR2BRecon.exe in the current folder.
REM  Self-contained: auto-installs Python 3.12 if not present.
REM ============================================================
setlocal EnableDelayedExpansion

echo.
echo ============================================================
echo   GSTR 2B Reconciliation - Build
echo ============================================================
echo.

REM --- 1. Ensure Python is available --------------------------
python --version >nul 2>&1
if not errorlevel 1 goto python_ready

echo   Python not found. Attempting automatic installation...
echo.

REM Try winget (available on Windows 10 1709+ / Windows 11)
winget --version >nul 2>&1
if not errorlevel 1 (
    echo   Installing Python 3.12 via winget...
    winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
    REM Refresh PATH so the new python.exe is visible in this session
    for /f "tokens=*" %%i in ('where python 2^>nul') do set "_PY=%%i"
    if not "!_PY!"=="" goto python_ready
    REM winget may have put it somewhere not yet on PATH; search common locations
    for %%d in (
        "%LOCALAPPDATA%\Programs\Python\Python312"
        "%LOCALAPPDATA%\Programs\Python\Python311"
        "%LOCALAPPDATA%\Programs\Python\Python310"
        "C:\Python312"
        "C:\Python311"
        "C:\Python310"
    ) do (
        if exist "%%~d\python.exe" (
            set "PATH=%%~d;%%~d\Scripts;!PATH!"
            goto python_ready
        )
    )
)

REM Fallback: download the official Python 3.12 installer and run silently
echo   winget not available. Downloading Python 3.12 installer...
set "_PY_INSTALLER=%TEMP%\python-3.12-installer.exe"
set "_PY_URL=https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"

REM Use PowerShell to download (available on all modern Windows)
powershell -NoProfile -Command "Invoke-WebRequest -Uri '%_PY_URL%' -OutFile '%_PY_INSTALLER%' -UseBasicParsing" 2>nul
if not exist "%_PY_INSTALLER%" (
    REM Try older Invoke-WebRequest syntax as fallback
    powershell -NoProfile -Command "(New-Object Net.WebClient).DownloadFile('%_PY_URL%', '%_PY_INSTALLER%')" 2>nul
)
if not exist "%_PY_INSTALLER%" (
    echo ERROR: Could not download Python installer. Please install Python 3.10+
    echo manually from https://www.python.org/downloads/ and tick
    echo "Add python.exe to PATH", then re-run this script.
    pause
    exit /b 1
)

echo   Running Python installer silently (this may take a minute)...
"%_PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1
del "%_PY_INSTALLER%" >nul 2>&1

REM Refresh PATH from registry for this session
for /f "tokens=*" %%i in ('powershell -NoProfile -Command "[Environment]::GetEnvironmentVariable(\"PATH\",\"User\")"') do set "PATH=%%i;%PATH%"

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python installation appears to have failed.
    echo Please install Python 3.10+ manually from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH", then re-run this script.
    pause
    exit /b 1
)

:python_ready
echo [OK] Python is available.
python --version

echo.
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
REM --paths python      lets "import engine" resolve.
REM --add-data ui.html  bundles the HTML (it sits next to build.bat, in the root).
REM --distpath .        outputs the .exe directly here, no dist\ folder created.
python -m PyInstaller ^
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
