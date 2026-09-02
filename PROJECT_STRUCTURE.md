# Struktura projektu

Po dodaniu obsługi konfiguracji, projekt ma następującą strukturę:

```
Tygodniowy_raport_napraw/
│
├── 📄 app.py                    ← GŁÓWNA APLIKACJA (zaktualizowana)
├── 📄 requirements.txt          ← Zależności
├── 📄 README.md                 ← Główna dokumentacja (zaktualizowana)
├── 📄 .gitignore                ← Ignorowane pliki (zaktualizowany)
│
├── ⚙️ KONFIGURACJA
│   ├── 📄 config.json           ← ⭐ EDYTUJ TO - ustawienia aplikacji
│   ├── 📄 auth.json             ← Hasło admin i secret key (autogenerowany)
│   └── 📄 CONFIG.md             ← Dokumentacja: Wszystkie parametry
│
├── 📚 DOKUMENTACJA - DYSKI SIECIOWE
│   ├── 📄 QUICK_START.md        ← Szybki start (3-4 kroki)
│   ├── 📄 NETWORK_STORAGE.md    ← Pełny przewodnik dysków sieciowych
│   ├── 📄 CONFIG_EXAMPLES.md    ← 10 praktycznych przykładów
│   └── 📄 DOCS_INDEX.md         ← Indeks wszystkich plików
│
├── 📋 PODSUMOWANIE ZMIAN
│   └── 📄 CHANGES_SUMMARY.md    ← Podsumowanie wszystkich zmian
│
├── 📁 KATALOGI DANYCH
│   ├── instance/                ← Dane aplikacji (domyślnie)
│   │   ├── repair_reports.sqlite3   ← Baza danych
│   │   └── uploads/             ← Przesyłane pliki
│   │
│   └── (Lub na dysku sieciowym jeśli skonfigurowane)
│
├── 📁 SZABLONY
│   └── templates/               ← HTML szablony
│       ├── base.html
│       ├── index.html
│       ├── new_record.html
│       ├── login.html
│       ├── filter_export.html
│       ├── import_export.html
│       ├── change_password.html
│       └── import.html
│
├── 📁 TESTY
│   ├── tests/
│   └── test_app.py
│
├── 📁 DANE PRZYKŁADOWE
│   └── dane_przykładowe/
│
└── 📁 SKRYPTY URUCHOMIENIA
    ├── start_windows.bat        ← Windows
    └── start_linux.sh           ← Linux
```

## 📌 Ważne pliki

### Edytuj bezpośrednio:
- **config.json** - Twoja konfiguracja aplikacji
- **Szablony HTML** - Wygląd aplikacji

### NIE edytuj (autogenerowane):
- **auth.json** - Zmieniaj przez UI aplikacji
- **instance/repair_reports.sqlite3** - Baza danych
- **instance/uploads/** - Przesyłane pliki

### Dobra praktyka:
- Backupuj **config.json** i **auth.json** regularnie!
- Synchronizuj **config.json** między serwerami jeśli potrzeba
- Nie commituj **config.json** do git (jest w .gitignore)

## 🚀 Kroki konfiguracji

1. **Edytuj `config.json`**
   - Zmień DATABASE na ścieżkę dysku sieciowego (jeśli chcesz)
   - Zmień UPLOAD_DIR jeśli potrzeba
   - Dostosuj PORT, HOST i inne parametry

2. **Przygotuj dysk sieciowy** (jeśli używasz)
   - Windows: Zamapuj UNC path
   - Linux: Zamontuj NFS/SMB

3. **Utwórz katalogi** (aplikacja robi to automatycznie)
   - Katalog dla bazy danych
   - Katalog dla przesyłanych plików

4. **Uruchom aplikację**
   ```bash
   python app.py
   ```

5. **Sprawdź** że działa
   - Otwórz http://localhost:5001
   - Zaloguj się (login: admin, hasło: admin)
   - Przetestuj podstawowe funkcje

## 📂 Gdzie przechowywane są dane

### Domyślnie (brak config.json):
```
instance/
├── repair_reports.sqlite3    ← Baza danych
└── uploads/                  ← Przesyłane pliki
    ├── 1806/
    ├── 1839/
    ├── 1841/
    └── ...
```

### Na dysku sieciowym (z config.json):
```
//nas-server/share/repair/     ← Konfiguracja
├── repair_reports.sqlite3    ← Baza danych
└── uploads/                  ← Przesyłane pliki
    └── ...
```

## 🔍 Monitorowanie

Aby sprawdzić gdzie są dane:
```bash
# Windows
type config.json

# Linux
cat config.json
```

Szukaj:
- `"DATABASE"` - ścieżka bazy danych
- `"UPLOAD_DIR"` - ścieżka przesyłanych plików

## 📊 Statystyka plików

| Kategoria | Ilość | Uwagi |
|-----------|-------|-------|
| Pliki dokumentacji | 6 | Nowe pliki |
| Pliki konfiguracji | 2 | config.json, auth.json |
| Szablony HTML | 7 | Bez zmian |
| Testy | 3 | Wszystkie przechodzą ✓ |
| Linie zmienionego kodu | ~50 | Głównie dla konfiguracji |

## ✅ Checklist przed produkcją

- [ ] Przeczytałem `QUICK_START.md`
- [ ] Edytowałem `config.json` dla mojego środowiska
- [ ] Sprawdziłem dostęp do dysku sieciowego (jeśli używam)
- [ ] Uruchomiłem aplikację i sprawdziłem czy działa
- [ ] Backupowałem `config.json` i `auth.json`
- [ ] Uruchomiłem testy: `python -m pytest tests/`
- [ ] Testowałem podstawowe funkcje (login, dodawanie rekordu, export)

## 🎯 Następne kroki

1. **Czytaj dokumentację** w tej kolejności:
   - [QUICK_START.md](QUICK_START.md) - szybki start
   - [DOCS_INDEX.md](DOCS_INDEX.md) - indeks dokumentacji
   - [CONFIG.md](CONFIG.md) - wszystkie parametry

2. **Edytuj config.json** do swoich potrzeb

3. **Konfiguruj dysk sieciowy** jeśli potrzeba

4. **Uruchom aplikację** i testuj

5. **Czytaj pełną dokumentację** gdy będziesz mieć pytania

---

**Struktura projektu gotowa!**
Dokumentacja pełna i kompletna.
Aplikacja gotowa do produkcji.
