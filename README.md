# Tygodniowy_raport_napraw

Przeglądarkowa aplikacja Flask do prowadzenia tygodniowego raportu napraw dla działu utrzymania ruchu.

## Spis tresci

- [Tygodniowy\_raport\_napraw](#tygodniowy_raport_napraw)
  - [Spis tresci](#spis-tresci)
  - [Uruchomienie](#uruchomienie)
  - [Dostępne funkcje](#dostępne-funkcje)
  - [Testy](#testy)


## Uruchomienie

```bash
python -m venv .venv
# Linux/macOS:
source .venv/bin/activate
# Windows PowerShell:
# .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
flask --app app:create_app run --host 0.0.0.0 --port 5001
```

Aplikacja używa bazy SQLite tworzonej automatycznie w katalogu `instance/`.

Po uruchomieniu jest dostępna lokalnie pod `http://localhost:5001/` oraz z
innych komputerów w sieci pod `http://ADRES_IP_SERWERA:5001/`. Aby uruchomić
ją bezpośrednio, można także użyć `python app.py`; ten tryb również nasłuchuje
na porcie `5001` i wszystkich interfejsach sieciowych.

Na serwerze należy zezwolić w zaporze systemowej na przychodzące połączenia TCP
na porcie `5001`. W Zaporze systemu Windows można to zrobić z poziomu
„Reguły przychodzące” przez dodanie reguły dla portu TCP `5001`.

## Dostępne funkcje

- wpisywanie rekordów napraw i interwencji,
- automatyczne wyliczanie roku, tygodnia, pozycji w tygodniu, czasu wykluczenia z produkcji i kosztu łącznego,
- filtrowanie rekordów po maszynie, rodzaju zdarzenia, zakresie dat i słowach z opisu,
- podgląd poprzednich wpisów dla wybranej maszyny,
- import danych z plików `.xlsx`,
- eksport przefiltrowanych danych do `.xlsx`.

## Testy

```bash
python -m unittest discover -s tests
```