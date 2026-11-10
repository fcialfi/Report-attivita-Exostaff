@echo off
:: ============================================================
::  Build + Deploy — Exostaff Activity Report Tool
:: ============================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"

set APP_NAME=ActivityReportTool
set SCRIPT=Exostaff_Activity_Report_Tool.py
set BUILD_ROOT=C:\Temp\ExostaffBuild
set DIST_DIR=%BUILD_ROOT%\dist
set BUILD_DIR=%BUILD_ROOT%\build
set LOG_FILE=build_log.txt

echo. > %LOG_FILE%

:: ── Read config via Python helper ───────────────────────────────────────
python read_config.py build_config.txt > _cfg_tmp.txt 2>nul

set /p EXE_SHARE=<_cfg_tmp.txt
for /f "skip=1 delims=" %%l in (_cfg_tmp.txt) do (
    if not defined DATA_SHARE set "DATA_SHARE=%%l"
)
del _cfg_tmp.txt 2>nul

:: ── Ask only if not configured ──────────────────────────────────────────
set ASKED=
if "%EXE_SHARE%"=="" (
    echo.
    echo  Where should the exe be copied?
    echo  e.g. N:\temp\Fabio\Exostaff Report Generator
    set /p EXE_SHARE=  Exe destination: 
    set ASKED=1
)
if "%DATA_SHARE%"=="" (
    echo.
    echo  Where is the working data folder?
    set /p DATA_SHARE=  Data folder: 
    set ASKED=1
)

:: Save config for next time — only when something was asked, so that a
:: normal build leaves the repository files untouched (git pull keeps working)
if defined ASKED (
    (
    echo EXE_SHARE=%EXE_SHARE%
    echo DATA_SHARE=%DATA_SHARE%
    ) > build_config.txt
)

echo.
echo  Exe  ^> %EXE_SHARE%
echo  Data ^> %DATA_SHARE%
echo.

:: ── Python check ────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 ( echo [ERROR] Python not found. & pause & exit /b 1 )

:: ── Version ─────────────────────────────────────────────────────────────
for /f "delims=" %%v in ('python update_version.py') do set CURRENT_VERSION=%%v
echo  Current version: %CURRENT_VERSION%
set /p NEW_VERSION=New version (blank = keep %CURRENT_VERSION%): 
if "%NEW_VERSION%"=="" set NEW_VERSION=%CURRENT_VERSION%
python update_version.py %NEW_VERSION% "%EXE_SHARE%"
if errorlevel 1 ( echo [ERROR] Version update failed. & pause & exit /b 1 )
echo  [OK] Version: %NEW_VERSION%

:: ── Dependencies ────────────────────────────────────────────────────────
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 ( pip install pyinstaller )
pip install python-docx lxml openpyxl pywin32 --quiet >> %LOG_FILE% 2>&1
echo  [OK] Dependencies ready.

:: ── Clean ───────────────────────────────────────────────────────────────
:: An exe started from the build folder (to try it) locks it: PyInstaller then
:: fails at the end ("_append_data_to_exe failed"). Close only those copies,
:: not the tool someone is using from another folder.
powershell -NoProfile -Command "Get-Process -Name '%APP_NAME%' -ErrorAction SilentlyContinue | Where-Object { $_.Path -like '%BUILD_ROOT%\*' } | ForEach-Object { Write-Host ('  [INFO] Closing ' + $_.Path); Stop-Process -Id $_.Id -Force }"
timeout /t 1 /nobreak >nul
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%"  rmdir /s /q "%DIST_DIR%"
if exist "%DIST_DIR%\%APP_NAME%.exe" (
    echo [ERROR] %DIST_DIR%\%APP_NAME%.exe is in use and cannot be replaced.
    echo         Close it ^(also in Task Manager^) or wait for the antivirus, then run build.bat again.
    pause & exit /b 1
)
if not exist "%BUILD_ROOT%" mkdir "%BUILD_ROOT%"
if exist %APP_NAME%.spec del %APP_NAME%.spec

:: ── Build ───────────────────────────────────────────────────────────────
echo.
echo  [INFO] Building exe...
python -m PyInstaller --onefile --windowed --name "%APP_NAME%" --icon "assets\exostaff_icon.ico" --add-data "assets;assets" --add-data "build_info.json;." --distpath "%DIST_DIR%" --workpath "%BUILD_DIR%" --clean %SCRIPT% >> %LOG_FILE% 2>&1
if errorlevel 1 (
    echo [ERROR] Build FAILED. See %LOG_FILE%.
    findstr /c:"_append_data_to_exe" %LOG_FILE% >nul && echo         The new exe could not be written: the antivirus was probably scanning it. Run build.bat again.
    pause & exit /b 1
)
echo  [OK] Build successful.

:: ── Write VERSION.txt ───────────────────────────────────────────────────
echo %NEW_VERSION%> "%DIST_DIR%\VERSION.txt"

:: ── Deploy ──────────────────────────────────────────────────────────────
echo  [INFO] Deploying...
if not exist "%EXE_SHARE%" mkdir "%EXE_SHARE%"

:: One exe, without version in the name. EXE_SHA256.txt holds its fingerprint:
:: a PC where OneDrive has synced VERSION.txt but not yet the new exe sees that
:: the exe there is still the old one and waits, instead of "updating" to it.
del /q "%EXE_SHARE%\%APP_NAME%_*.exe" 2>nul
copy /y "%DIST_DIR%\%APP_NAME%.exe" "%EXE_SHARE%\%APP_NAME%.exe" >nul
if errorlevel 1 ( echo [ERROR] Copy exe failed: %APP_NAME%.exe in use from the shared folder, or network. & pause & exit /b 1 )
python -c "import hashlib,sys; print(hashlib.sha256(open(sys.argv[1],'rb').read()).hexdigest())" "%DIST_DIR%\%APP_NAME%.exe" > "%DATA_SHARE%\EXE_SHA256.txt"
if errorlevel 1 ( echo [ERROR] EXE_SHA256.txt not written. & pause & exit /b 1 )
echo %EXE_SHARE%\%APP_NAME%.exe> "%DATA_SHARE%\EXE_LOCATION.txt"
echo  [OK] Exe deployed: %APP_NAME%.exe, EXE_SHA256.txt, EXE_LOCATION.txt

:: VERSION.txt last: it is what tells the users a new version is out
copy /y "%DIST_DIR%\VERSION.txt" "%DATA_SHARE%\VERSION.txt" >nul
if errorlevel 1 ( echo [ERROR] VERSION.txt copy failed - users will NOT be notified. & pause & exit /b 1 )
echo  [OK] VERSION.txt updated.

:: The tool looks for VERSION.txt in the exe folder first (RELEASE_DIR):
:: publish it there too when the exe and data folders differ
if /i not "%EXE_SHARE%"=="%DATA_SHARE%" (
    copy /y "%DATA_SHARE%\EXE_SHA256.txt" "%EXE_SHARE%\EXE_SHA256.txt" >nul
    copy /y "%DATA_SHARE%\EXE_LOCATION.txt" "%EXE_SHARE%\EXE_LOCATION.txt" >nul
    copy /y "%DIST_DIR%\VERSION.txt" "%EXE_SHARE%\VERSION.txt" >nul
    echo  [OK] VERSION.txt, EXE_LOCATION.txt and EXE_SHA256.txt also in the exe folder.
)

set TEMPLATE=EPS-Sterna and EPS-Aeolus System Engineering Service Template.docx
set TEMPLATE_SRC=
if exist "%TEMPLATE%" set "TEMPLATE_SRC=%TEMPLATE%"
if exist "config\%TEMPLATE%" set "TEMPLATE_SRC=config\%TEMPLATE%"
if defined TEMPLATE_SRC (
    copy /y "%TEMPLATE_SRC%" "%EXE_SHARE%\%TEMPLATE%" >nul
    echo  [OK] Template copied.
) else (
    echo  [WARN] Template not found - not copied.
)

echo.
echo  ============================================================
echo   DONE  —  v%NEW_VERSION% deployed to:
echo   %EXE_SHARE%
echo  ============================================================
echo.
pause
