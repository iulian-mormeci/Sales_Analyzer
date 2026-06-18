@echo off
REM ============================================================
REM build.bat — Build SalesAnalyzer.exe (Windows)
REM Usage: doppio click su build.bat, oppure da terminale: build.bat
REM ============================================================

cd /d "%~dp0"

echo.
echo ============================================================
echo  Sales Analyzer - Build per Windows
echo ============================================================
echo.

REM ── 1. Python check ─────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Python non trovato. Installalo da https://python.org
    pause & exit /b 1
)
echo [1/5] Python trovato:
python --version

REM ── 2. Ambiente virtuale ────────────────────────────────────
set VENV_DIR=.venv_build
echo [2/5] Creazione ambiente virtuale (%VENV_DIR%)...
python -m venv %VENV_DIR%
if errorlevel 1 (
    echo [ERRORE] Creazione venv fallita.
    pause & exit /b 1
)
call %VENV_DIR%\Scripts\activate.bat
python -m pip install --upgrade pip --quiet

REM ── 3. Dipendenze ───────────────────────────────────────────
echo [3/5] Installazione dipendenze...
pip install -r requirements.txt --quiet
pip install pyinstaller --quiet
if errorlevel 1 (
    echo [ERRORE] Installazione dipendenze fallita.
    pause & exit /b 1
)

REM ── 4. Icona ────────────────────────────────────────────────
if not exist icon.ico (
    echo [4/5] Creazione icona placeholder...
    python create_icon.py 2>nul && echo        icon.ico creata || echo        Icona non creata, continuo senza.
) else (
    echo [4/5] Icona trovata: icon.ico
)

REM ── 5. Build ────────────────────────────────────────────────
echo [5/5] Build con PyInstaller...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

python -m PyInstaller build.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERRORE] Build fallita. Controlla i messaggi sopra.
    pause & exit /b 1
)

call %VENV_DIR%\Scripts\deactivate.bat 2>nul

echo.
echo ============================================================
echo  Build completata!
echo  Eseguibile: dist\SalesAnalyzer.exe
echo ============================================================
echo.
pause
