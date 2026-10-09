# Tygodniowy_raport_napraw

Przeglądarkowa aplikacja Flask do prowadzenia tygodniowego raportu napraw dla działu utrzymania ruchu.

## Spis tresci

- [Tygodniowy\_raport\_napraw](#tygodniowy_raport_napraw)
  - [Spis tresci](#spis-tresci)
  - [Uruchomienie](#uruchomienie)
  - [Konfiguracja](#konfiguracja)
  - [Dostępne funkcje](#dostępne-funkcje)
  - [Przykładowy wygląd aplikacji](#przykładowy-wygląd-aplikacji)
  - [Panel maszyn](#panel-maszyn)
  - [Panel użytkowników](#panel-użytkowników)
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
- eksport przefiltrowanych danych do `.xlsx`,
- logowanie użytkowników i administratorów,
- zarządzanie uprawnieniami do przeglądania, tworzenia, edycji i usuwania rekordów,
- administracyjne zarządzanie danymi maszyn i ich wpisami.

## Przykładowy wygląd aplikacji

Poniższe zrzuty ekranu pokazują przykładowy widok rejestru oraz filtrowanie
rekordów. Zawierają wyłącznie fikcyjne dane demonstracyjne.

### Strona główna

![Strona główna aplikacji z listą ostatnich rekordów](docs/screenshots/strona-glowna.png)

### Filtrowanie i eksport

![Filtrowanie rekordów i tabela wyników](docs/screenshots/filtrowanie.png)

## Panel maszyn

Po zalogowaniu jako administrator wybierz w menu aplikacji pozycję **Maszyny**
lub otwórz adres `/admin/machines`. Panel wyświetla nazwę, typ, numer seryjny
każdej maszyny oraz liczbę powiązanych z nią wpisów.

Administrator może edytować nazwę, typ i numer seryjny. Zmiana dotyczy wszystkich
wpisów przypisanych do tej maszyny. Numer seryjny musi być unikalny; zapis nie
powiedzie się, jeśli jest już używany przez inną maszynę.

Usunięcie maszyny usuwa również wszystkie jej wpisy i załączone pliki. Przed
operacją aplikacja pokazuje liczbę wpisów do usunięcia i prosi o potwierdzenie.
Usunięcia nie można cofnąć.

## Panel użytkowników

Administrator może otworzyć panel użytkowników pod adresem `/admin/users` i zarządzać kontami bezpośrednio z poziomu aplikacji.

Panel umożliwia:

- tworzenie nowych kont użytkowników,
- edycję istniejących kont i ich uprawnień,
- zmianę roli użytkownika na użytkownika lub administratora,
- zmianę hasła użytkownika,
- nadawanie niezależnych uprawnień do przeglądania, tworzenia, edycji i usuwania rekordów,
- usuwanie kont użytkowników.

W panelu znajduje się również dziennik aktywności. Rejestrowane są między innymi logowania i wylogowania, zmiany haseł, operacje na rekordach i załącznikach, import i eksport danych oraz zarządzanie kontami. Dziennik wyświetla 20 zdarzeń na stronie i posiada paginację.

Administrator ma pełne uprawnienia niezależnie od ustawień poszczególnych pól. Zwykły użytkownik może wykonywać tylko operacje wynikające z przyznanych mu uprawnień.

## Testy

```bash
python -m unittest discover -s tests
```

## Zmiany do wprowadzenia

- [x] Dodano pole osób (oddzielonych przecinkami) biorących udział w interwencji, obsługiwane także przez import i eksport Excel.
- [x] Dodano panel użytkowników, zarządzanie uprawnieniami oraz dziennik aktywności użytkowników i administratorów.
- [ ] 