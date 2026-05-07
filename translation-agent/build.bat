@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

:: ═══════════════════════════════════════════════════════════
::  Translation Agent - Windows Build Script
:: ═══════════════════════════════════════════════════════════
::  Usage:  build.bat              (build both)
::          build.bat onefile      (build onefile only)
::          build.bat onedir       (build onedir only)
::          build.bat clean        (clean build artifacts)
:: ═══════════════════════════════════════════════════════════

set "PROJECT_DIR=%~dp0"
set "VENV_DIR=%PROJECT_DIR%.venv-build"
set "DIST_DIR=%PROJECT_DIR%dist"

echo.
echo  ══════════════════════════════════════════════════
echo   Translation Agent - Build Script (Windows)
echo  ══════════════════════════════════════════════════
echo.

:: ── Parse arguments ──
set "BUILD_ONEFILE=1"
set "BUILD_ONEDIR=1"
set "DO_CLEAN=0"

if "%~1"=="onefile" (
    set "BUILD_ONEDIR=0"
)
if "%~1"=="onedir" (
    set "BUILD_ONEFILE=0"
)
if "%~1"=="clean" (
    set "DO_CLEAN=1"
    set "BUILD_ONEFILE=0"
    set "BUILD_ONEDIR=0"
)

:: ── Clean mode ──
if "%DO_CLEAN%"=="1" (
    echo  [CLEAN] Removing build artifacts...
    if exist "%PROJECT_DIR%build" rd /s /q "%PROJECT_DIR%build"
    if exist "%DIST_DIR%" rd /s /q "%DIST_DIR%"
    echo  [CLEAN] Done.
    echo.
    goto :eof
)

:: ── Check Python ──
where python >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found in PATH.
    echo  Please install Python 3.9+ from https://www.python.org/
    goto :error
)

for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set "PYVER=%%v"
echo  [INFO] Python version: %PYVER%

:: ── Create virtual environment ──
if not exist "%VENV_DIR%" (
    echo  [SETUP] Creating virtual environment...
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo  [ERROR] Failed to create virtual environment.
        goto :error
    )
    echo  [SETUP] Virtual environment created.
) else (
    echo  [SETUP] Using existing virtual environment.
)

:: ── Activate venv ──
call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
    echo  [ERROR] Failed to activate virtual environment.
    goto :error
)

:: ── Upgrade pip ──
echo  [SETUP] Upgrading pip...
python -m pip install --upgrade pip --quiet
if errorlevel 1 (
    echo  [WARN] pip upgrade failed, continuing anyway...
)

:: ── Install dependencies ──
echo  [SETUP] Installing dependencies...
pip install -r "%PROJECT_DIR%requirements-desktop.txt" --quiet
if errorlevel 1 (
    echo  [ERROR] Failed to install dependencies.
    echo  Try running manually: pip install -r requirements-desktop.txt
    goto :error
)
echo  [SETUP] Dependencies installed.

:: ── Build onefile ──
if "%BUILD_ONEFILE%"=="1" (
    echo.
    echo  ──────────────────────────────────────
    echo   Building ONEFILE (single EXE)...
    echo  ──────────────────────────────────────
    pyinstaller "%PROJECT_DIR%TranslationAgent-onefile.spec" --clean --noconfirm
    if errorlevel 1 (
        echo  [ERROR] onefile build failed!
        goto :error
    )
    echo  [OK] onefile build succeeded.
)

:: ── Build onedir ──
if "%BUILD_ONEDIR%"=="1" (
    echo.
    echo  ──────────────────────────────────────
    echo   Building ONEDIR (directory mode)...
    echo  ──────────────────────────────────────
    pyinstaller "%PROJECT_DIR%TranslationAgent-onedir.spec" --clean --noconfirm
    if errorlevel 1 (
        echo  [ERROR] onedir build failed!
        goto :error
    )
    echo  [OK] onedir build succeeded.
)

:: ── Copy .env.example ──
if exist "%PROJECT_DIR%.env.example" (
    if "%BUILD_ONEFILE%"=="1" (
        copy /y "%PROJECT_DIR%.env.example" "%DIST_DIR%.env.example" >nul
    )
    if "%BUILD_ONEDIR%"=="1" (
        copy /y "%PROJECT_DIR%.env.example" "%DIST_DIR%TranslationAgent\.env.example" >nul
    )
    echo  [INFO] .env.example copied to output directory.
)

:: ── Summary ──
echo.
echo  ══════════════════════════════════════════════════
echo   BUILD SUCCESSFUL
echo  ══════════════════════════════════════════════════
echo.

if "%BUILD_ONEFILE%"=="1" (
    echo  [onefile]  %DIST_DIR%\TranslationAgent.exe
    echo             (single executable, ~150-200 MB)
)
if "%BUILD_ONEDIR%"=="1" (
    echo  [onedir]   %DIST_DIR%\TranslationAgent\
    echo             (directory mode, exe + _internal\)
)

echo.
echo  To run:  Copy .env to the same directory as the EXE,
::            then double-click TranslationAgent.exe
echo.
echo  .env.template:
echo    OPENAI_API_KEY=sk-your-key-here
echo    OPENAI_BASE_URL=https://api.openai.com/v1
echo    OPENAI_MODEL=gpt-4o
echo.

goto :eof

:error
echo.
echo  ══════════════════════════════════════════════════
echo   BUILD FAILED
echo  ══════════════════════════════════════════════════
echo.
exit /b 1
