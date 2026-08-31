# Tygodniowy_raport_napraw

Przeglądarkowa aplikacja Flask do prowadzenia tygodniowego raportu napraw dla działu utrzymania ruchu.

## Uruchomienie

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
flask --app app run
```

Aplikacja używa bazy SQLite tworzonej automatycznie w katalogu `instance/`.

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