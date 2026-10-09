import pandas as pd
import numpy as np

# Wczytanie oryginalnego pliku (musi znajdować się w tym samym katalogu)
# Jeśli Twój plik nazywa się inaczej, zmień nazwę poniżej:
file_path = 'tygodniowy_raport_napraw_dane_przykladowe.xlsx'
df = pd.read_excel(file_path, sheet_name='ZBIORCZE')

# Definicje przykładowych, bezpiecznych danych zastępczych
maszyny = ['Frezarka CNC', 'Prasa Krawędziowa', 'Wtryskarka', 'Robot Spawalniczy', 'Laser Fiber', 'Linia Pakująca', 'Wózek Widłowy', 'Szlifierka', 'Przenośnik Taśmowy', 'Agregat Prądotwórczy']
typy = ['MK-200', 'PK-100', 'WT-50', 'RS-300', 'LF-500', 'LP-10', 'WW-2', 'SZ-80', 'PT-15', 'AP-40']
rodzaje_zdarzenia = ['Awaria', 'Przegląd', 'Konserwacja', 'Regulacja', 'Uszkodzenie mechaniczne']

opisy = [
    'Zatrzymanie maszyny w wyniku błędu czujnika krańcowego.',
    'Spadek ciśnienia w układzie pneumatycznym, nieszczelność na przewodzie.',
    'Okresowy przegląd techniczny, smarowanie prowadnic i łożysk.',
    'Błąd komunikacji na sterowniku PLC, wymagany restart systemu.',
    'Przegrzanie silnika napędowego, czyszczenie radiatora i wentylatora.',
    'Uszkodzenie taśmy transportowej, zacięcie detalu.',
    'Wymiana zużytego elementu tnącego.',
    'Brak reakcji panelu operatorskiego, reset modułu zasilania.'
]

przyczyny = [
    'Naturalne zużycie eksploatacyjne elementu.',
    'Zanieczyszczenie pyłem i opiłkami metalu.',
    'Wahania napięcia w sieci elektrycznej.',
    'Błąd operatora podczas uruchamiania cyklu.',
    'Przeciążenie mechaniczne podzespołu.'
]

sposoby_naprawy = [
    'Wymianę uszkodzonego podzespołu na nowy.',
    'Czyszczenie, regulację oraz ponowne kalibrowanie czujnika.',
    'Reset systemu sterowania i testowe uruchomienie.',
    'Zaciskanie nowego złącza i zabezpieczenie okablowania.',
    'Uzupełnienie poziomu płynów eksploatacyjnych i smaru.'
]

wymienione_czesci = ['Czujnik indukcyjny M12', 'Przekaźnik 24V DC', 'Uszczelka gumowa', 'Filtr powietrza', 'Przewód pneumatyczny 8mm', 'Wkład filtrujący', 'Bezpiecznik 10A']
osoby = ['Jan Kowalski', 'Piotr Nowak', 'Adam Wiśniewski', 'Krzysztof Wójcik', 'Serwis Zewnętrzny']
uwagi = ['Maszyna przekazana do normalnej eksploatacji bez uwag.', 'Zalecana kontrola stanu technicznego przy kolejnym przeglądzie.', 'Brak dalszych usterek w kolejnych dniach roboczych.']

np.random.seed(42)
n_rows = len(df)

# Podmienianie kolumn na dane przykładowe
df['Maszyna'] = np.random.choice(maszyny, size=n_rows)
df['Typ'] = np.random.choice(typy, size=n_rows)
df['Numer'] = [f"SN-{np.random.randint(1000, 9999)}" if str(x).strip() != '-' else '-' for x in df['Numer']]
df['Rodzaj zdarzenia'] = np.random.choice(rodzaje_zdarzenia, size=n_rows, p=[0.5, 0.2, 0.15, 0.1, 0.05])

def random_fill(options, mask_series):
    res = []
    for val in mask_series:
        if pd.isna(val) and np.random.rand() > 0.6:
            res.append(np.nan)
        else:
            res.append(np.random.choice(options))
    return res

df['Opis'] = np.random.choice(opisy, size=n_rows)
df['Przyczyna'] = random_fill(przyczyny, df['Przyczyna'])
df['Sposób naprawy'] = random_fill(sposoby_naprawy, df['Sposób naprawy'])
df['Wymienione części'] = random_fill(wymienione_czesci, df['Wymienione części'])
df['Osoby biorące udział w interwencji'] = random_fill(osoby, df['Osoby biorące udział w interwencji'])
df['Uwagi'] = random_fill(uwagi, df['Uwagi'])

df['Wykluczenie z produkcji [rob.godzina]'] = np.round(np.random.exponential(scale=0.5, size=n_rows), 2)
df['Materiał własny PLN'] = np.random.choice([0.0, 50.0, 120.0, 250.0, 450.0], size=n_rows, p=[0.7, 0.15, 0.1, 0.03, 0.02])
df['Części zamienne PLN'] = np.random.choice([0, 150, 300, 600, 1200, 2500], size=n_rows, p=[0.6, 0.2, 0.1, 0.05, 0.03, 0.02])
df['Serwis PLN'] = np.random.choice([0.0, 500.0, 1200.0, 3000.0], size=n_rows, p=[0.8, 0.1, 0.07, 0.03])
df['Koszt brutto naprawy PLN'] = df['Materiał własny PLN'] + df['Części zamienne PLN'] + df['Serwis PLN']

# Zapis do nowego pliku
output_file = 'raport_napraw_dane_przykladowe.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='ZBIORCZE', index=False)

print(f"Wygenerowano pomyślnie plik: {output_file}")