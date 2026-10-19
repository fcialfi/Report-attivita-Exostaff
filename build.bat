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
if "%EXE_SHARE%"=="" (
    echo.
    echo  Where should the exe be copied?
    echo  e.g. N:\temp\Fabio\Exostaff Report Generator
    set /p EXE_SHARE=  Exe destination: 
)
if "%DATA_SHARE%"=="" (
    echo.
    echo  Where is the working data folder?
    set /p DATA_SHARE=  Data folder: 
)

:: Save config for next time
(
echo EXE_SHARE=%EXE_SHARE%
echo DATA_SHARE=%DATA_SHARE%
) > build_config.txt

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
python update_version.py %NEW_VERSION%
if errorlevel 1 ( echo [ERROR] Version update failed. & pause & exit /b 1 )
echo  [OK] Version: %NEW_VERSION%

:: ── Dependencies ────────────────────────────────────────────────────────
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 ( pip install pyinstaller )
pip install python-docx lxml openpyxl --quiet >> %LOG_FILE% 2>&1
echo  [OK] Dependencies ready.

:: ── Clean ───────────────────────────────────────────────────────────────
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%"  rmdir /s /q "%DIST_DIR%"
if not exist "%BUILD_ROOT%" mkdir "%BUILD_ROOT%"
if exist %APP_NAME%.spec del %APP_NAME%.spec

:: ── Build ───────────────────────────────────────────────────────────────
echo.
echo  [INFO] Building exe...
python -m PyInstaller --onefile --windowed --name "%APP_NAME%" --distpath "%DIST_DIR%" --workpath "%BUILD_DIR%" --clean %SCRIPT% >> %LOG_FILE% 2>&1
if errorlevel 1 ( echo [ERROR] Build FAILED. See %LOG_FILE%. & pause & exit /b 1 )
echo  [OK] Build successful.

:: ── Write VERSION.txt ───────────────────────────────────────────────────
echo %NEW_VERSION%> "%DIST_DIR%\VERSION.txt"

:: ── Deploy ──────────────────────────────────────────────────────────────
echo  [INFO] Deploying...
if not exist "%EXE_SHARE%" mkdir "%EXE_SHARE%"

copy /y "%DIST_DIR%\%APP_NAME%.exe" "%EXE_SHARE%\%APP_NAME%.exe" >nul
if errorlevel 1 ( echo [ERROR] Copy exe failed. Check network. & pause & exit /b 1 )
echo  [OK] Exe deployed.

copy /y "%DIST_DIR%\VERSION.txt" "%DATA_SHARE%\VERSION.txt" >nul
if errorlevel 1 ( echo [ERROR] VERSION.txt copy failed - users will NOT be notified. & pause & exit /b 1 )
echo  [OK] VERSION.txt updated.

echo %EXE_SHARE%\%APP_NAME%.exe> "%DATA_SHARE%\EXE_LOCATION.txt"
echo  [OK] EXE_LOCATION.txt updated.

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
