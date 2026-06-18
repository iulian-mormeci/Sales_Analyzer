#!/usr/bin/env bash
# ============================================================
# build.sh — Build SalesAnalyzer (macOS / Linux)
# Usage:
#   chmod +x build.sh   (only once)
#   ./build.sh
# ============================================================

set -e
cd "$(dirname "$0")"

echo ""
echo "============================================================"
echo " Sales Analyzer — Build per macOS / Linux"
echo "============================================================"
echo ""

# ── 1. Python check ──────────────────────────────────────────
PYTHON=""
for candidate in python3.13 python3.12 python3.11 python3.10 python3; do
    if command -v "$candidate" &>/dev/null; then
        PYTHON="$candidate"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "[ERRORE] Python 3 non trovato."
    exit 1
fi
echo "[1/5] Python trovato: $($PYTHON --version)"

# ── 2. Virtual environment ───────────────────────────────────
VENV_DIR=".venv_build"
echo "[2/5] Creazione ambiente virtuale ($VENV_DIR)..."
"$PYTHON" -m venv "$VENV_DIR"
source "$VENV_DIR/bin/activate"
pip install --upgrade pip --quiet

# ── 3. Dipendenze ────────────────────────────────────────────
echo "[3/5] Installazione dipendenze..."
pip install -r requirements.txt --quiet
pip install pyinstaller --quiet

# ── 4. Icona ─────────────────────────────────────────────────
if [ ! -f "icon.ico" ]; then
    echo "[4/5] Creazione icona placeholder..."
    python create_icon.py 2>/dev/null && echo "       icon.ico creata" \
        || echo "       Icona non creata (Pillow mancante?), continuo senza."
else
    echo "[4/5] Icona trovata: icon.ico"
fi

# ── 5. Build ─────────────────────────────────────────────────
echo "[5/5] Build con PyInstaller..."
rm -rf build dist
python -m PyInstaller build.spec --clean --noconfirm

deactivate

echo ""
echo "============================================================"
echo " Build completata!"
if [[ "$OSTYPE" == "darwin"* ]]; then
    echo " macOS: dist/SalesAnalyzer  (binary standalone)"
    echo " NOTA: per creare il .exe Windows esegui build.bat su Windows."
else
    echo " Linux: dist/SalesAnalyzer"
fi
echo "============================================================"
echo ""
