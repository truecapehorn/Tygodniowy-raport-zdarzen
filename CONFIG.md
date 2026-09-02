# Konfiguracja Aplikacji

Aplikacja obsługuje plik konfiguracyjny `config.json` w katalogu głównym, który pozwala na elastyczne ustawienie wielu parametrów działania.

## Plik `config.json`

Jeśli plik `config.json` nie istnieje, aplikacja będzie używać wartości domyślnych. Jeśli plik istnieje, wartości z niego będą użyte do przesłonięcia domyślnych.

### Dostępne parametry

```json
{
  "DATABASE": "instance/repair_reports.sqlite3",
  "UPLOAD_DIR": "instance/uploads",
  "DEBUG": false,
  "HOST": "0.0.0.0",
  "PORT": 5001,
  "LABOR_RATE_PLN_PER_HOUR": 80,
  "MAX_UPLOAD_SIZE_MB": 50,
  "REQUIRE_LOGIN": true,
  "AUTH_FILE": "auth.json"
}
```

## Parametry

### DATABASE
- **Typ:** string (ścieżka)
- **Domyślnie:** `instance/repair_reports.sqlite3`
- **Opis:** Ścieżka do pliku bazy danych SQLite
- **Obsługiwane formaty ścieżek:**
  - Ścieżka względna: `instance/repair_reports.sqlite3` (względem katalogu głównego aplikacji)
  - Ścieżka bezwzględna Windows: `C:/data/repair_reports.sqlite3` lub `C:\data\repair_reports.sqlite3`
  - UNC path Windows: `//server-name/share/repair_reports.sqlite3` lub `\\server-name\share\repair_reports.sqlite3`
  - Ścieżka Linux: `/mnt/nas/repair_reports.sqlite3` lub `/home/user/shared/repair_reports.sqlite3`

### UPLOAD_DIR
- **Typ:** string (ścieżka)
- **Domyślnie:** `instance/uploads`
- **Opis:** Katalog dla przesyłanych plików
- **Obsługiwane formaty ścieżek:** jak wyżej (DATABASE)

### DEBUG
- **Typ:** boolean
- **Domyślnie:** `false`
- **Opis:** Włącza tryb debugowania Flask (hot-reload, szczegółowe błędy)

### HOST
- **Typ:** string
- **Domyślnie:** `0.0.0.0`
- **Opis:** Adres IP, na którym aplikacja nasłuchuje
- **Wartości:**
  - `0.0.0.0` - dostępna ze wszystkich interfejsów (zalecane dla sieci)
  - `127.0.0.1` - tylko localhost
  - `192.168.x.x` - konkretny adres IP

### PORT
- **Typ:** integer
- **Domyślnie:** `5001`
- **Opis:** Port, na którym aplikacja nasłuchuje

### LABOR_RATE_PLN_PER_HOUR
- **Typ:** float/integer
- **Domyślnie:** `80`
- **Opis:** Stawka godzinowa pracy (PLN) używana do obliczania kosztów pracy

### MAX_UPLOAD_SIZE_MB
- **Typ:** integer
- **Domyślnie:** `50`
- **Opis:** Maksymalny rozmiar pojedynczego przesyłanego pliku w megabajtach

### REQUIRE_LOGIN
- **Typ:** boolean
- **Domyślnie:** `true`
- **Opis:** 
  - `true` - wymagane logowanie do aplikacji
  - `false` - dostęp bez logowania (wszyscy jako goście)

### AUTH_FILE
- **Typ:** string (ścieżka)
- **Domyślnie:** `auth.json`
- **Opis:** Ścieżka do pliku z hasłami i ustawieniami uwierzytelniania

## Przykłady konfiguracji

### 1. Baza danych na dysku sieciowym NAS (Linux/NFS)

```json
{
  "DATABASE": "/mnt/nas/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "/mnt/nas/repair_database/uploads",
  "HOST": "192.168.1.100",
  "PORT": 5001,
  "DEBUG": false,
  "REQUIRE_LOGIN": true
}
```

**Przygotowanie na Linux:**
```bash
# Zamontuj udział NAS
sudo mkdir -p /mnt/nas
sudo mount -t nfs nas-server:/export/repair /mnt/nas

# Lub dodaj do /etc/fstab dla automatycznego montowania:
# nas-server:/export/repair  /mnt/nas  nfs  defaults,_netdev  0  0
```

### 2. Baza danych na udziale sieciowym (Windows UNC)

```json
{
  "DATABASE": "//nas-server/repair/repair_reports.sqlite3",
  "UPLOAD_DIR": "//nas-server/repair/uploads",
  "HOST": "0.0.0.0",
  "PORT": 5001,
  "DEBUG": false
}
```

### 3. Tryb debugowania (development)

```json
{
  "DATABASE": "instance/repair_reports.sqlite3",
  "UPLOAD_DIR": "instance/uploads",
  "DEBUG": true,
  "HOST": "127.0.0.1",
  "PORT": 5001,
  "REQUIRE_LOGIN": false
}
```

### 4. Serwer produkcyjny bez wymaganego logowania

```json
{
  "DATABASE": "/var/lib/repair-app/repair_reports.sqlite3",
  "UPLOAD_DIR": "/var/lib/repair-app/uploads",
  "HOST": "0.0.0.0",
  "PORT": 80,
  "DEBUG": false,
  "REQUIRE_LOGIN": false,
  "LABOR_RATE_PLN_PER_HOUR": 100
}
```

## Kompatybilność Windows i Linux

Aplikacja automatycznie obsługuje różne formaty ścieżek:

- **Ścieżki z `/`** (forward slash) - działają na obu systemach
- **Ścieżki z `\`** (backslash) - działają na Windows
- **UNC paths** (`//server/share`) - konwertowane dla każdego systemu

Rekomendacja: Używaj forward slash `/` w konfiguracji - jest kompatybilna z oboma systemami.

## Dyski sieciowe - obsługa

### Windows (SMB/CIFS)

```json
{
  "DATABASE": "//nas-server/share/db/repair_reports.sqlite3"
}
```

Upewnij się, że udział jest dostępny i zamontowany:
```cmd
net use Z: \\nas-server\share
```

### Linux (NFS/SMB)

Dla NFS:
```bash
sudo mount -t nfs nas-server:/export/database /mnt/nas
```

Dla SMB:
```bash
sudo mount -t cifs //nas-server/share /mnt/nas -o username=user,password=pass
```

Następnie w konfiguracji:
```json
{
  "DATABASE": "/mnt/nas/repair_reports.sqlite3"
}
```

## Uruchomienie aplikacji

```bash
# Linux
python app.py

# Windows
python app.py
```

Aplikacja automatycznie załaduje konfigurację z `config.json` jeśli istnieje.

## Uwagi

1. **Uprawnienia:** Upewnij się, że aplikacja ma uprawnienia do zapisu w katalogach `DATABASE` i `UPLOAD_DIR`
2. **Ścieżki względne:** Odnoszą się do katalogu głównego aplikacji (tam gdzie jest `app.py`)
3. **Komentarze w JSON:** Plik `config.json` nie obsługuje komentarzy (są usuwane przez loader)
4. **Zmiana konfiguracji:** Aplikacja odczytuje `config.json` przy starcie; zmiana wymaga restartu
5. **Bezpieczeństwo:** Nie przechowuj haseł w `config.json` - używaj oddzielnego pliku `auth.json`
