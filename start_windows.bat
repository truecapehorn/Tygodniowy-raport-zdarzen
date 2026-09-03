@echo off
chcp 65001 >nul
title Tygodniowy raport napraw GORT

echo.
echo  =============================================
echo   Tygodniowy raport napraw GORT
echo  =============================================
echo.

:: ── Sprawdz Python ────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [BLAD] Python nie zostal znaleziony.
    echo Pobierz Python 3.11+ ze strony https://www.python.org/downloads/
    echo Podczas instalacji zaznacz opcje "Add Python to PATH".
    pause
    exit /b 1
)

:: ── Srodowisko wirtualne ──────────────────────────────────────────
if not exist "venv\Scripts\activate.bat" (
    echo [INFO] Tworzenie srodowiska wirtualnego...
    python -m venv venv
    if errorlevel 1 (
        echo [BLAD] Nie mozna utworzyc srodowiska wirtualnego.
        pause
        exit /b 1
    )
)

:: ── Aktywacja ─────────────────────────────────────────────────────
echo [INFO] Aktywacja srodowiska wirtualnego...
call venv\Scripts\activate.bat

:: ── Zaleznosci ───────────────────────────────────────────────────
echo [INFO] Instalacja/aktualizacja zaleznosci...
pip install -r requirements.txt -q
if errorlevel 1 (
    echo [BLAD] Nie mozna zainstalowac zaleznosci.
    pause
    exit /b 1
)

:: ── Uruchomienie serwera ──────────────────────────────────────────
echo.
echo  Serwer uruchomiony:
echo    http://localhost:5001
echo    http://127.0.0.1:5001
echo.
echo  Aby zatrzymac serwer nacisnij Ctrl+C
echo.

:: Otwieranie przegladarki po 2 sekundach (w tle)
start "" /b timeout /t 2 >nul && start http://localhost:5001

set FLASK_DEBUG=1
python app.py

pause
