# 📚 Dokumentacja - Indeks

Poniżej znajduje się spis wszystkich plików dokumentacji do konfiguracji i dysków sieciowych.

## 🚀 Szybki Start (zacznij tu!)

| Plik | Opis | Dla kogo |
|------|------|----------|
| [QUICK_START.md](QUICK_START.md) | 3-4 kroky do uruchomienia na dysku sieciowym | Wszyscy chcący szybko zacząć |
| [CHANGES_SUMMARY.md](CHANGES_SUMMARY.md) | Podsumowanie wszystkich zmian | Deweloperzy i administratorzy |

## 📖 Pełna Dokumentacja

| Plik | Opis | Kiedy czytać |
|------|------|--------------|
| [CONFIG.md](CONFIG.md) | **Pełna dokumentacja konfiguracji** - wszystkie parametry i formaty ścieżek | Chcesz zrozumieć wszystkie opcje |
| [CONFIG_EXAMPLES.md](CONFIG_EXAMPLES.md) | **10 praktycznych przykładów** dla różnych scenariuszy | Szukasz gotowego przykładu dla Twojej sytuacji |
| [NETWORK_STORAGE.md](NETWORK_STORAGE.md) | **Szczegółowy przewodnik dysków sieciowych** - instalacja, montowanie, troubleshooting | Chcesz poprawnie skonfigurować dysk sieciowy |
| [README.md](README.md) | **Główny plik** - ogólny opis aplikacji i nowa sekcja o konfiguracji | Chcesz poznać całą aplikację |

## 🎯 Wybierz swoją ścieżkę

### Jestem na Windows i chcę dysk sieciowy
1. Przeczytaj [QUICK_START.md](QUICK_START.md) - sekcja "Windows - 3 proste kroki"
2. Przejdź do [NETWORK_STORAGE.md](NETWORK_STORAGE.md) - sekcja "Windows - UNC Path"
3. Sprawdź [CONFIG_EXAMPLES.md](CONFIG_EXAMPLES.md) - Przykład 3 (Production - Windows UNC)

### Jestem na Linux i chcę dysk sieciowy
1. Przeczytaj [QUICK_START.md](QUICK_START.md) - sekcja "Linux - 4 proste kroki"
2. Przejdź do [NETWORK_STORAGE.md](NETWORK_STORAGE.md) - sekcja "Linux - NFS" lub "Linux - SMB"
3. Sprawdź [CONFIG_EXAMPLES.md](CONFIG_EXAMPLES.md) - Przykład 4 lub 5

### Chcę configu dla development
Przeczytaj [CONFIG_EXAMPLES.md](CONFIG_EXAMPLES.md) - Przykład 2 (Development)

### Chcę configu dla production
Przeczytaj [CONFIG_EXAMPLES.md](CONFIG_EXAMPLES.md) - Przykład 3-10 (Production)

### Chcę szczegółów o wszystkich parametrach
Przeczytaj [CONFIG.md](CONFIG.md)

## 📁 Pliki pomocnicze

### config.json
- **Lokalizacja:** Katalog główny aplikacji (obok app.py)
- **Zawartość:** Ustawienia aplikacji
- **Edytuj:** Gdy chcesz zmienić konfigurację
- **Ignoruj git:** Tak (w .gitignore)

### auth.json
- **Lokalizacja:** Katalog główny aplikacji
- **Zawartość:** Hasło admin i secret key
- **Edytuj:** Rzadko (automatycznie tworzone)
- **Ignoruj git:** Tak (w .gitignore)

### .gitignore
- **Zawiera:** `config.json` i `auth.json`
- **Cel:** Chronić poufne dane przed commitowaniem
- **Edytuj:** Nie (chyba że dodajesz nowe pliki do ignorowania)

## 🔍 Szukaj odpowiedzi na pytania

### P: Jak zmienić port aplikacji?
**O:** Edytuj `config.json`, zmień `"PORT": 5001` na inny numer, np. `"PORT": 8080`

### P: Jak wyłączyć wymagane logowanie?
**O:** W `config.json` zmień `"REQUIRE_LOGIN": true` na `"REQUIRE_LOGIN": false`

### P: Jak zmienić stawkę godzinową?
**O:** W `config.json` zmień `"LABOR_RATE_PLN_PER_HOUR": 80` na inną wartość

### P: Jak przenieść bazę na dysk sieciowy?
**O:** Przeczytaj [NETWORK_STORAGE.md](NETWORK_STORAGE.md) dla Twojego systemu operacyjnego

### P: Jak sprawdzić czy dysk sieciowy jest dostępny?
**O:** Przeczytaj sekcję "Test połączenia" w [NETWORK_STORAGE.md](NETWORK_STORAGE.md)

### P: Co zrobić jeśli aplikacja nie może połączyć się z dykiem sieciowym?
**O:** Przeczytaj sekcję "Rozwiązywanie problemów" w [NETWORK_STORAGE.md](NETWORK_STORAGE.md)

### P: Czy aplikacja obsługuje ścieżki względne?
**O:** Tak! Przeczytaj [CONFIG.md](CONFIG.md) - sekcja "DATABASE" dla formatów ścieżek

### P: Czy muszę edytować app.py żeby zmienić konfigurację?
**O:** Nie! Edytuj tylko `config.json`

### P: Czy mogę mieć różne config.json na różnych serwerach?
**O:** Tak! Każdy serwer może mieć swoją `config.json` (nie jest commitowana do git)

## 📊 Schemat - Jak działa aplikacja

```
Start aplikacji
    ↓
Szuka config.json
    ↓
    ├─ ZNALEZIONY → Wczytuje ustawienia z config.json
    │               ↓
    │               Łączy z domyślnymi wartościami
    │
    └─ NIEZNALEZIONY → Używa wartości domyślnych
                      ↓
                      (można edytować i uruchomić ponownie)
    ↓
Sprawdzuje katalogi
    ↓
    ├─ Jeśli nie istnieją → Tworzy je
    │
    └─ Jeśli istnieją → Używa je
    ↓
Aplikacja gotowa
    ↓
    Dostępna na http://HOST:PORT
```

## 🔐 Bezpieczeństwo

- `config.json` i `auth.json` NIE są commitowane (`.gitignore`)
- Każdy deweloper/serwer ma własną konfigurację
- Hasła są hashowane w `auth.json`
- Zalecenie: Backupuj `config.json` i `auth.json`!

## 💡 Pro Tips

1. **Kopia zapasowa:** Przed zmianą `config.json`, zrób kopię:
   ```bash
   cp config.json config.json.backup
   ```

2. **Test przed zmianą:** Sprawdź dostęp do dysku sieciowego zanim zmienisz config

3. **Logi:** Uruchom z `"DEBUG": true` aby zobaczyć szczegóły błędów

4. **Montowanie automatyczne:** Na Linux dodaj do `/etc/fstab` żeby dysk montował się przy starcie

5. **Reload:** Żeby zastosować nową konfigurację, zrestartuj aplikację

## 📞 Potrzebujesz pomocy?

1. Przeczytaj odpowiednią dokumentację (linki powyżej)
2. Sprawdź [CONFIG_EXAMPLES.md](CONFIG_EXAMPLES.md) - może tam jest Twój scenariusz
3. Sprawdź sekcję "Rozwiązywanie problemów" w [NETWORK_STORAGE.md](NETWORK_STORAGE.md)
4. Przejrzyj [CONFIG.md](CONFIG.md) - tam są szczegóły każdego parametru

---

**Wersja:** 1.0  
**Data:** 2026-09-02  
**Status:** ✅ Pełna dokumentacja  
**Kompatybilność:** Windows, Linux, MacOS
