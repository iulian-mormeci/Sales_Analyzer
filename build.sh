#!/usr/bin/env bash
# ============================================================
# build.sh — Build SalesAnalyzer (macOS / Linux)
# Usage: chmod +x build.sh && ./build.sh
# ============================================================

set -e

echo ""
echo "============================================================"
echo " Sales Analyzer — Build per macOS / Linux"
echo "============================================================"
echo ""

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "[ERRORE] Python 3 non trovato."
    exit 1
fi
python3 --version

# Install dependencies
echo ""
echo "[1/4] Installazione dipendenze..."
pip3 install -r requirements.txt --quiet

# Install PyInstaller
echo "[2/4] Installazione PyInstaller..."
pip3 install pyinstaller --quiet

# Create icon placeholder if missing
if [ ! -f "icon.ico" ]; then
    echo "[3/4] Creazione icona placeholder..."
    python3 create_icon.py 2>/dev/null || echo "       Icona non creata, continuo senza."
else
    echo "[3/4] Icona trovata: icon.ico"
fi

# Clean previous build
rm -rf build dist

# Build
echo "[4/4] Build con PyInstaller..."
python3 -m PyInstaller build.spec --clean --noconfirm

echo ""
echo "============================================================"
echo " Build completata!"
echo " Eseguibile: dist/SalesAnalyzer  (o dist/SalesAnalyzer.app su macOS)"
echo "============================================================"
echo ""
