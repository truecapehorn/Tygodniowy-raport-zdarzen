#!/usr/bin/env bash
# Skrypt startowy dla Linux Ubuntu

set -e

echo ""
echo " ============================================="
echo "  Tygodniowy raport napraw DUR"
echo " ============================================="
echo ""

# Przejdź do katalogu skryptu (działa też przy wywołaniu z innego miejsca)
cd "$(dirname "$0")"

# ── Sprawdź Python ──────────────────────────────────────────────────
if ! command -v python3 &>/dev/null; then
    echo "[BŁĄD] Python3 nie został znaleziony."
    echo "Zainstaluj: sudo apt-get install python3 python3-pip python3-venv"
    exit 1
fi

PYTHON_VER=$(python3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
echo "[INFO] Python $PYTHON_VER"

# ── Sprawdź python3-venv ────────────────────────────────────────────
if ! python3 -m venv --help &>/dev/null; then
    echo "[BŁĄD] Moduł venv nie jest dostępny."
    echo "Zainstaluj: sudo apt-get install python3-venv"
    exit 1
fi

# ── Środowisko wirtualne ────────────────────────────────────────────
if [ ! -d "venv" ]; then
    echo "[INFO] Tworzenie środowiska wirtualnego..."
    python3 -m venv venv
fi

# ── Aktywacja ───────────────────────────────────────────────────────
echo "[INFO] Aktywacja środowiska wirtualnego..."
source venv/bin/activate

# ── Zależności ──────────────────────────────────────────────────────
echo "[INFO] Instalacja/aktualizacja zależności..."
pip install -r requirements.txt -q

# ── Uprawnienia do katalogu instance ───────────────────────────────
mkdir -p instance/uploads

# ── Uruchomienie serwera ────────────────────────────────────────────
echo ""
echo " Serwer uruchomiony:"
echo "   http://localhost:5001"

# Otwórz przeglądarkę jeśli dostępna (działa na Ubuntu Desktop)
if [ -n "${DISPLAY:-}" ] && command -v xdg-open &>/dev/null; then
    (sleep 2 && xdg-open http://localhost:5001) &
fi

export FLASK_DEBUG=0
python app.py
