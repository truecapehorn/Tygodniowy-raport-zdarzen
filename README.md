# Tygodniowy_raport_napraw

Przeglądarkowa aplikacja Flask do prowadzenia tygodniowego raportu napraw dla działu utrzymania ruchu.

## Spis tresci

- [Tygodniowy\_raport\_napraw](#tygodniowy_raport_napraw)
  - [Spis tresci](#spis-tresci)
  - [Uruchomienie](#uruchomienie)
  - [Konfiguracja](#konfiguracja)
  - [Dostępne funkcje](#dostępne-funkcje)
  - [Testy](#testy)
  - [Zmiany do wprowadzenia](#zmiany-do-wprowadzenia)


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
## Konfiguracja

Aplikacja obsługuje plik `config.json` do elastycznego konfigurowania:

- **DATABASE** - ścieżka do bazy danych (obsługuje dyski sieciowe UNC/NFS)
- **UPLOAD_DIR** - katalog dla przesyłanych plików
- **PORT** - port nasłuchiwania
- **HOST** - adres IP nasłuchiwania
- **DEBUG** - tryb debugowania
- **LABOR_RATE_PLN_PER_HOUR** - stawka godzinowa dla obliczeń
- **MAX_UPLOAD_SIZE_MB** - maksymalny rozmiar pliku
- **REQUIRE_LOGIN** - czy logowanie jest wymagane

**Przykład dla dysku sieciowego (Windows UNC):**
```json
{
  "DATABASE": "//nas-server/share/repair_reports.sqlite3",
  "UPLOAD_DIR": "//nas-server/share/uploads"
}
```

**Przykład dla Linux (NFS):**
```json
{
  "DATABASE": "/mnt/nas/repair_reports.sqlite3",
  "UPLOAD_DIR": "/mnt/nas/uploads"
}
```

Pełna dokumentacja: [CONFIG.md](CONFIG.md), [CONFIG_EXAMPLES.md](CONFIG_EXAMPLES.md), [NETWORK_STORAGE.md](NETWORK_STORAGE.md)
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

## Zmiany do wprowadzenia

- [ ] Dodanie nowego pola do wypełnienia. Dodać pole do osób (po przecinku) biorące udział w interwencji.