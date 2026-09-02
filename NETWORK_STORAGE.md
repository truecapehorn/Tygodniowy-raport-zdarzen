# Przewodnik: Baza danych na dysku sieciowym

Instrukcje konfiguracji bazy danych na udziale sieciowym (NAS, SMB, NFS).

## Windows - UNC Path (\\server\share)

### Krok 1: Połączenie z udziałem sieciowym

#### Opcja A: Mapowanie jako dysk sieciowy (GUI)
1. Otwórz "Eksplorator plików"
2. Kliknij "Ten komputer" > "Mapuj dysk sieciowy"
3. Wpisz ścieżkę: `\\nas-server\share` (zmień nas-server na adres Twojego serwera)
4. Przypisz literę dysku (np. Z:)
5. Kliknij "Dalej" i poczekaj na połączenie

#### Opcja B: Mapowanie z wiersza poleceń
```cmd
net use Z: \\nas-server\share /persistent:yes
```

#### Opcja C: Bezpośrednia ścieżka UNC w config.json
```json
{
  "DATABASE": "//nas-server/share/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "//nas-server/share/repair_database/uploads"
}
```

### Krok 2: Konfiguracja config.json

Jeśli dysk jest zamapowany jako Z::
```json
{
  "DATABASE": "Z:/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "Z:/repair_database/uploads"
}
```

Jeśli używasz bezpośrednio UNC path:
```json
{
  "DATABASE": "//nas-server/share/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "//nas-server/share/repair_database/uploads"
}
```

### Krok 3: Test połączenia
```cmd
# Sprawdź dostęp do ścieżki
dir \\nas-server\share

# Lub jeśli zmapowany dysk
dir Z:
```

---

## Linux - NFS (Network File System)

### Krok 1: Instalacja narzędzi
```bash
# Debian/Ubuntu
sudo apt-get install nfs-common

# Red Hat/CentOS
sudo yum install nfs-utils
```

### Krok 2: Połączenie z serwerem NFS

#### Opcja A: Montowanie ręczne
```bash
# Utwórz punkt montowania
sudo mkdir -p /mnt/nas

# Zamontuj udział NFS
sudo mount -t nfs nas-server:/export/repair /mnt/nas

# Lub z opcjami (rw - read/write, hard - trwałe połączenie)
sudo mount -t nfs -o rw,hard,intr nas-server:/export/repair /mnt/nas
```

#### Opcja B: Automatyczne montowanie (fstab)
1. Edytuj `/etc/fstab`:
```bash
sudo nano /etc/fstab
```

2. Dodaj linię (na końcu pliku):
```
nas-server:/export/repair  /mnt/nas  nfs  defaults,_netdev,auto  0  0
```

3. Test montowania:
```bash
sudo mount -a
```

### Krok 3: Konfiguracja config.json
```json
{
  "DATABASE": "/mnt/nas/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "/mnt/nas/repair_database/uploads"
}
```

### Krok 4: Uprawnienia
```bash
# Sprawdź, czy katalog jest dostępny
ls -la /mnt/nas

# Jeśli trzeba, zmień uprawnienia (dla użytkownika uruchamiającego aplikację)
sudo chown -R $USER /mnt/nas
sudo chmod -R 755 /mnt/nas
```

### Krok 5: Test połączenia
```bash
touch /mnt/nas/test.txt
ls /mnt/nas/test.txt
rm /mnt/nas/test.txt
```

---

## Linux - SMB (Samba/CIFS - udziały Windows)

### Krok 1: Instalacja narzędzi
```bash
# Debian/Ubuntu
sudo apt-get install cifs-utils

# Red Hat/CentOS
sudo yum install cifs-utils
```

### Krok 2: Przygotowanie poświadczeń (opcjonalnie)

Utwórz plik z poświadczeniami (bezpieczny):
```bash
nano ~/.smbcredentials
```

Zawartość:
```
username=samba_user
password=samba_password
domain=WORKGROUP
```

Ustaw uprawnienia:
```bash
chmod 600 ~/.smbcredentials
```

### Krok 3: Montowanie udziału SMB

#### Opcja A: Montowanie ręczne (bez autentykacji)
```bash
sudo mkdir -p /mnt/nas
sudo mount -t cifs //nas-server/share /mnt/nas
```

#### Opcja B: Montowanie ręczne (z autentykacją)
```bash
sudo mount -t cifs //nas-server/share /mnt/nas -o username=user,password=pass
```

#### Opcja C: Montowanie z plikiem poświadczeń
```bash
sudo mount -t cifs //nas-server/share /mnt/nas -o credentials=~/.smbcredentials
```

#### Opcja D: Automatyczne montowanie (fstab)
```bash
sudo nano /etc/fstab
```

Dodaj linię:
```
//nas-server/share  /mnt/nas  cifs  credentials=~/.smbcredentials,uid=1000,gid=1000,file_mode=0755,dir_mode=0755  0  0
```

### Krok 4: Konfiguracja config.json
```json
{
  "DATABASE": "/mnt/nas/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "/mnt/nas/repair_database/uploads"
}
```

---

## MacOS - NFS

### Krok 1: Montowanie NFS
```bash
# Utwórz punkt montowania
sudo mkdir -p /Volumes/nas

# Zamontuj
sudo mount -t nfs -o rw,hard nas-server:/export/repair /Volumes/nas
```

### Krok 2: Konfiguracja config.json
```json
{
  "DATABASE": "/Volumes/nas/repair_database/repair_reports.sqlite3",
  "UPLOAD_DIR": "/Volumes/nas/repair_database/uploads"
}
```

---

## Rozwiązywanie problemów

### Problem: "Permission denied"
```bash
# Linux/Mac - sprawdź uprawnienia
ls -la /mnt/nas
# Zmień uprawnienia jeśli potrzeba
sudo chmod -R 755 /mnt/nas
```

### Problem: "Connection refused"
- Sprawdź czy serwer NAS/SMB jest dostępny: `ping nas-server`
- Sprawdź ustawienia zapory (firewall)
- Sprawdź czy port jest otwarty (445 dla SMB, 2049 dla NFS)

### Problem: Timeout przy dostępie do bazy danych
- Może to być problem z wydajnością sieci
- Rozważ cachowanie danych lokalnie
- Zmniejsz rozmiar bazy danych

### Problem: Montowanie zawiedzie przy starcie Linux
- Dodaj `_netdev` do opcji w fstab: `defaults,_netdev`
- Upewnij się że sieć jest dostępna przed montowaniem

### Test połączenia sieci
```bash
# Linux/Mac
ping nas-server

# Sprawdź trasy
traceroute nas-server

# Sprawdź DNS
nslookup nas-server
```

---

## Zalecenia bezpieczeństwa

1. **Firewall:** Ogranicz dostęp do portów (445 dla SMB, 2049 dla NFS)
2. **Poświadczenia:** Nie przechowuj haseł w config.json
3. **Kopie zapasowe:** Regularnie backupuj bazę danych
4. **Uprawnienia:** Używaj odpowiednich uprawnień dostępu
5. **Monitoring:** Monitoruj dostęp do udziału sieciowego

---

## Wydajność

### Optymalizacja dla NFS
```bash
# Zwiększ rozmiar bufora
sudo mount -t nfs -o rw,hard,intr,rsize=32768,wsize=32768 nas-server:/export/repair /mnt/nas
```

### Optymalizacja dla SMB
```bash
# Zmniejsz opóźnienie
sudo mount -t cifs //nas-server/share /mnt/nas -o cache=loose,rw,hard
```

### SQLite na dysku sieciowym
- SQLite może być wolne na sieciowych lokalizacjach
- Rozważ synchronizowanie lokalnej kopii bazy danych
- Lub użyj zdalnego serwera PostgreSQL zamiast SQLite
