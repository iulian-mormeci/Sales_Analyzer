@echo off
REM ============================================================
REM build.bat — Build SalesAnalyzer.exe (Windows)
REM Usage: build.bat
REM Requirements: pip install pyinstaller (and all requirements.txt)
REM ============================================================

echo.
echo ============================================================
echo  Sales Analyzer — Build per Windows
echo ============================================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Python non trovato. Installalo da https://python.org
    pause
    exit /b 1
)

REM Install/upgrade dependencies
echo [1/4] Installazione dipendenze...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERRORE] Installazione dipendenze fallita.
    pause
    exit /b 1
)

REM Install PyInstaller
echo [2/4] Installazione PyInstaller...
pip install pyinstaller --quiet
if errorlevel 1 (
    echo [ERRORE] Installazione PyInstaller fallita.
    pause
    exit /b 1
)

REM Create icon if it doesn't exist
if not exist icon.ico (
    echo [3/4] Creazione icona placeholder...
    python create_icon.py 2>nul || echo        Icona non creata, continuo senza.
) else (
    echo [3/4] Icona trovata: icon.ico
)

REM Build
echo [4/4] Build con PyInstaller...
pyinstaller build.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERRORE] Build fallita.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build completata!
echo  Eseguibile: dist\SalesAnalyzer.exe
echo ============================================================
echo.
pause
