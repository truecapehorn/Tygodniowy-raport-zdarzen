# Przykłady konfiguracji - config.json

## Kopii te przykłady i dostosuj do swoich potrzeb

### Przykład 1: Konfiguracja domyślna (localhost, bez logowania wymagane)

Zapisz jako `config.json`:
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

---

### Przykład 2: Development (localhost, debug mode, brak logowania)

```json
{
  "DATABASE": "instance/repair_reports.sqlite3",
  "UPLOAD_DIR": "instance/uploads",
  "DEBUG": true,
  "HOST": "127.0.0.1",
  "PORT": 5001,
  "LABOR_RATE_PLN_PER_HOUR": 80,
  "MAX_UPLOAD_SIZE_MB": 100,
  "REQUIRE_LOGIN": false,
  "AUTH_FILE": "auth.json"
}
```

---

### Przykład 3: Production - Baza na dysku sieciowym Windows (UNC)

Dla serwera NAS dostępnego jako `\\nas-company\databases`:

```json
{
  "DATABASE": "//nas-company/databases/repair/repair_reports.sqlite3",
  "UPLOAD_DIR": "//nas-company/databases/repair/uploads",
  "DEBUG": false,
  "HOST": "192.168.1.50",
  "PORT": 80,
  "LABOR_RATE_PLN_PER_HOUR": 100,
  "MAX_UPLOAD_SIZE_MB": 100,
  "REQUIRE_LOGIN": true,
  "AUTH_FILE": "auth.json"
}
```

**Wymagane przygotowanie Windows:**
```cmd
# Zamapuj dysk sieciowy (jeśli wolisz)
net use Z: \\nas-company\databases

# Wtedy możesz również użyć:
REM "DATABASE": "Z:/repair/repair_reports.sqlite3"
```

---

### Przykład 4: Production Linux - NAS z NFS

Dla serwera NAS dostępnego jako `nas.company.com:/export/repair`:

```json
{
  "DATABASE": "/mnt/company-nas/repair_reports.sqlite3",
  "UPLOAD_DIR": "/mnt/company-nas/uploads",
  "DEBUG": false,
  "HOST": "0.0.0.0",
  "PORT": 80,
  "LABOR_RATE_PLN_PER_HOUR": 100,
  "MAX_UPLOAD_SIZE_MB": 200,
  "REQUIRE_LOGIN": true,
  "AUTH_FILE": "auth.json"
}
```

**Przygotowanie Linux:**
```bash
# Zainstaluj NFS
sudo apt-get install nfs-common

# Utwórz punkt montowania
sudo mkdir -p /mnt/company-nas

# Zamontuj (tymczasowo - do testów)
sudo mount -t nfs nas.company.com:/export/repair /mnt/company-nas

# Permanentnie - dodaj do /etc/fstab:
# nas.company.com:/export/repair  /mnt/company-nas  nfs  defaults,_netdev,auto  0  0

# Zmieni uprawnienia jeśli potrzeba
sudo chown -R $USER /mnt/company-nas
```

---

### Przykład 5: Production Linux - Udział SMB (Windows)

Dla udziału Samba `\\fileserver\engineering-data`:

```json
{
  "DATABASE": "/mnt/fileserver/engineering-data/repair_reports.sqlite3",
  "UPLOAD_DIR": "/mnt/fileserver/engineering-data/uploads",
  "DEBUG": false,
  "HOST": "0.0.0.0",
  "PORT": 8080,
  "LABOR_RATE_PLN_PER_HOUR": 120,
  "MAX_UPLOAD_SIZE_MB": 150,
  "REQUIRE_LOGIN": true,
  "AUTH_FILE": "auth.json"
}
```

**Przygotowanie Linux:**
```bash
# Zainstaluj CIFS
sudo apt-get install cifs-utils

# Utwórz punkt montowania
sudo mkdir -p /mnt/fileserver

# Utwórz plik z poświadczeniami (bezpieczny)
nano ~/.smb-credentials
# Zawartość:
# username=domain_user
# password=domain_password
# domain=COMPANY
chmod 600 ~/.smb-credentials

# Zamontuj (tymczasowo)
sudo mount -t cifs //fileserver/engineering-data /mnt/fileserver \
  -o credentials=~/.smb-credentials,uid=1000,gid=1000

# Permanentnie - dodaj do /etc/fstab:
# //fileserver/engineering-data  /mnt/fileserver  cifs  credentials=~/.smb-credentials,uid=1000,gid=1000,file_mode=0755,dir_mode=0755  0  0
```

---

### Przykład 6: Wspólna instalacja sieciowa (wiele użytkowników)

```json
{
  "DATABASE": "//shared-server/applications/repair-app/repair_reports.sqlite3",
  "UPLOAD_DIR": "//shared-server/applications/repair-app/uploads",
  "DEBUG": false,
  "HOST": "0.0.0.0",
  "PORT": 5001,
  "LABOR_RATE_PLN_PER_HOUR": 95,
  "MAX_UPLOAD_SIZE_MB": 75,
  "REQUIRE_LOGIN": true,
  "AUTH_FILE": "auth.json"
}
```

**Struktura na serwerze:**
```
//shared-server/applications/repair-app/
├── config.json
├── auth.json
├── app.py
├── instance/
├── templates/
├── repair_reports.sqlite3     (baza danych)
└── uploads/                    (przesyłane pliki)
```

---

### Przykład 7: Wysokie wymagania dotyczące wydajności

```json
{
  "DATABASE": "/fast-ssd-raid/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "/fast-ssd-raid/repair_files/uploads",
  "DEBUG": false,
  "HOST": "0.0.0.0",
  "PORT": 5001,
  "LABOR_RATE_PLN_PER_HOUR": 150,
  "MAX_UPLOAD_SIZE_MB": 500,
  "REQUIRE_LOGIN": true,
  "AUTH_FILE": "auth.json"
}
```

---

### Przykład 8: Dostęp lokalny + backup na NAS

```json
{
  "DATABASE": "/var/lib/repair-app/repair_reports.sqlite3",
  "UPLOAD_DIR": "/var/lib/repair-app/uploads",
  "DEBUG": false,
  "HOST": "127.0.0.1",
  "PORT": 5001,
  "LABOR_RATE_PLN_PER_HOUR": 80,
  "MAX_UPLOAD_SIZE_MB": 50,
  "REQUIRE_LOGIN": false,
  "AUTH_FILE": "/etc/repair-app/auth.json"
}
```

**Backup na NAS (skrypt cron):**
```bash
#!/bin/bash
# /usr/local/bin/backup-repair-db.sh

SOURCE="/var/lib/repair-app/repair_reports.sqlite3"
BACKUP="/mnt/nas/backups/repair-app/"
DATE=$(date +%Y%m%d_%H%M%S)

mkdir -p "$BACKUP"
cp "$SOURCE" "$BACKUP/repair_reports_$DATE.sqlite3"

# Utrzymuj tylko ostatnie 30 kopii
find "$BACKUP" -name "repair_reports_*.sqlite3" -mtime +30 -delete
```

Dodaj do crontab:
```bash
# Backup co godzinę
0 * * * * /usr/local/bin/backup-repair-db.sh

# Lub co dzień o 2:00 rano
0 2 * * * /usr/local/bin/backup-repair-db.sh
```

---

### Przykład 9: Multi-site - Każda lokacja ma swoją bazę

**Lokacja Warszawa:**
```json
{
  "DATABASE": "//warsaw-nas/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "//warsaw-nas/repair_files/uploads",
  "DEBUG": false,
  "HOST": "192.168.1.20",
  "PORT": 5001,
  "LABOR_RATE_PLN_PER_HOUR": 85,
  "MAX_UPLOAD_SIZE_MB": 100,
  "REQUIRE_LOGIN": true,
  "AUTH_FILE": "auth.json"
}
```

**Lokacja Kraków:**
```json
{
  "DATABASE": "//krakow-nas/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "//krakow-nas/repair_files/uploads",
  "DEBUG": false,
  "HOST": "192.168.2.20",
  "PORT": 5001,
  "LABOR_RATE_PLN_PER_HOUR": 80,
  "MAX_UPLOAD_SIZE_MB": 100,
  "REQUIRE_LOGIN": true,
  "AUTH_FILE": "auth.json"
}
```

---

### Przykład 10: Docker - Baza w wolumenie sieciowym

```json
{
  "DATABASE": "/data/repair_reports.sqlite3",
  "UPLOAD_DIR": "/data/uploads",
  "DEBUG": false,
  "HOST": "0.0.0.0",
  "PORT": 5001,
  "LABOR_RATE_PLN_PER_HOUR": 100,
  "MAX_UPLOAD_SIZE_MB": 100,
  "REQUIRE_LOGIN": true,
  "AUTH_FILE": "auth.json"
}
```

**docker-compose.yml:**
```yaml
version: '3.8'
services:
  repair-app:
    image: repair-app:latest
    ports:
      - "5001:5001"
    volumes:
      - type: volume
        source: nfs_storage
        target: /data
        volume:
          nocopy: true
    environment:
      - FLASK_APP=app.py

volumes:
  nfs_storage:
    driver: local
    driver_opts:
      type: nfs
      o: addr=nas.company.com,vers=4,soft,timeo=180,bg,tcp
      device: ":/export/repair"
```

---

## Jak zmienić konfigurację?

1. **Edytuj plik `config.json`** w katalogu głównym aplikacji
2. **Zmień wartości** które chcesz dostosować
3. **Zapisz plik**
4. **Zrestartuj aplikację**

```bash
# Linux/Mac
pkill -f "python app.py"
python app.py

# Windows
# Zamknij okno aplikacji i uruchom ponownie
python app.py
```

---

## Błędy przy konfiguracji

### "Database is locked"
- Sprawdź uprawnienia do katalogów
- Upewnij się że ścieżka jest dostępna
- Spróbuj uruchomić aplikację od nowa

### "Connection refused"
- Sprawdź czy serwer NAS/SMB jest dostępny
- Sprawdź ustawienia zapory
- Sprawdź połączenie sieciowe

### "Permission denied"
- Zmień uprawnienia dostępu do katalogów
- Sprawdź czy użytkownik uruchamiający aplikację ma dostęp

### "File not found"
- Upewnij się że ścieżka w config.json jest prawidłowa
- Sprawdź czy katalog istnieje (jeśli nie - aplikacja go utworzy)
- Dla dysków sieciowych - sprawdź czy są zamonotowane

---

## Automatyczne tworzenie katalogów

Aplikacja automatycznie tworzy katalogi `DATABASE` i `UPLOAD_DIR` jeśli nie istnieją.
Upewnij się że aplikacja ma uprawnienia do zapisu w lokalizacji głównej.
