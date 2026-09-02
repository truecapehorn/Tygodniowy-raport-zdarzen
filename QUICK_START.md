# Szybki start: Baza danych na dysku sieciowym

## Windows - 3 proste kroki

### 1. Zamapuj dysk sieciowy
```cmd
net use Z: \\nas-server\share
```

### 2. Utwórz plik `config.json`
```json
{
  "DATABASE": "Z:/repair_reports.sqlite3",
  "UPLOAD_DIR": "Z:/uploads"
}
```

### 3. Uruchom aplikację
```cmd
python app.py
```

---

## Linux - 4 proste kroki

### 1. Zainstaluj NFS
```bash
sudo apt-get install nfs-common
```

### 2. Zamontuj dysk
```bash
sudo mkdir -p /mnt/nas
sudo mount -t nfs nas-server:/export/repair /mnt/nas
```

### 3. Utwórz plik `config.json`
```json
{
  "DATABASE": "/mnt/nas/repair_reports.sqlite3",
  "UPLOAD_DIR": "/mnt/nas/uploads"
}
```

### 4. Uruchom aplikację
```bash
python app.py
```

---

## Testy połączenia

### Windows
```cmd
# Sprawdź mapowanie
net use

# Test dostępu do pliku
dir Z:\
```

### Linux
```bash
# Sprawdź montowanie
mount | grep nas

# Test dostępu do pliku
ls /mnt/nas/
```

---

## Problemy?

- **Nie mogę się połączyć:** Sprawdź czy serwer jest dostępny (`ping nas-server`)
- **Brak uprawnień:** Upewnij się że masz dostęp do udziału
- **Timeout:** Może być problem z siecią - sprawdź połączenie
- **Plik zablokowany:** Upewnij się że nikt inny nie używa bazy

## Więcej informacji

Pełny przewodnik: [NETWORK_STORAGE.md](NETWORK_STORAGE.md)
Konfiguracja: [CONFIG.md](CONFIG.md)
