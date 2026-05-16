# Excel → Word z presetami

> **Status:** TEMP PUBLIC / RELEASE CANDIDATE v0.1.0  
> Repozytorium celowo pozostaje po polsku.  
> Repozytorium nie powinno zawierać danych wejściowych, wynikowych ani prywatnych plików roboczych.

Narzędzie GUI do konwersji wierszy z pliku Excel na osobne mikro-tabele w dokumencie Word (`.docx`). Program korzysta z presetów JSON, które opisują mapowanie kolumn, etykiety, transformacje wartości oraz opcje generowania dokumentu.

Projekt powstał przede wszystkim do szybkiego tworzenia uporządkowanych kart opisowych z tabel Excel, np. dla danych terenowych lub zestawień obiektów. Może jednak działać także z innymi plikami Excel, jeśli kolumny zostaną poprawnie zmapowane w presecie lub ręcznie w GUI.

## Główne możliwości

- wczytywanie plików `.xlsx` i `.xls`,
- wybór presetu JSON po wskazaniu pliku Excel,
- edytowalne mapowanie kolumn na pola w dokumencie Word,
- transformacje wartości, np. formatowanie dat, przeliczanie powierzchni, wartości stałe, łączenie kilku kolumn,
- naturalne sortowanie, np. `Obiekt 2` przed `Obiekt 10`,
- generowanie mikro-tabel w pliku `.docx`,
- opcjonalne ramki na zdjęcie i mapę,
- dwa układy dokumentu: A i B,
- konfiguracja czcionki, marginesów i szerokości kolumn,
- log z przebiegu generowania dokumentu.

## Struktura repozytorium

```text
Excel2Word_with_presets/
├─ README.md
├─ LICENSE
├─ CHANGELOG.md
├─ requirements.txt
├─ .gitignore
├─ .gitattributes
├─ script/
│  └─ excel_rows_to_word_gui.py
├─ presets/
│  ├─ *.json
│  └─ gdy_2_pola.txt
├─ docs/
│  └─ preset_format.md
├─ tools/
│  └─ build_release_zip.py
└─ dist/
   └─ README.md
```

Foldery z danymi wejściowymi, zdjęciami i wynikami powinny być lokalne i nie powinny być commitowane do repozytorium.

## Wymagania

- Python 3.8+
- pakiety z pliku `requirements.txt`:
  - `pandas`
  - `python-docx`
  - `openpyxl`
  - `xlrd>=2.0.1`
  - `Pillow`

Instalacja zależności:

```bash
pip install -r requirements.txt
```

Jeśli skrypt ma własny mechanizm doinstalowywania zależności, można go traktować jako wygodne zabezpieczenie. Dla czystego i powtarzalnego uruchamiania zalecany jest jednak plik `requirements.txt`.

## Szybki start

1. Przenieś działający plik programu do folderu:

   ```text
   script/excel_rows_to_word_gui.py
   ```

2. Uruchom program:

   ```bash
   python script/excel_rows_to_word_gui.py
   ```

3. Wskaż plik Excel.
4. Wczytaj preset JSON z folderu `presets/`.
5. Sprawdź mapowanie pól w GUI.
6. Ustaw opcje dokumentu, zdjęć, mapy, sortowania, marginesów i czcionki.
7. Kliknij generowanie dokumentu `.docx`.

## Presety JSON

Presety opisują, jak kolumny z Excela mają zostać przeniesione do dokumentu Word. Typowy preset zawiera:

```json
{
  "name": "Nazwa presetu",
  "version": 1,
  "options": {
    "decimal_comma": true,
    "page_break": true,
    "photo": { "enabled": true, "height_cm": 6.0 },
    "map": { "enabled": true, "height_cm": 6.0 }
  },
  "aliases": {
    "nazwa_kolumny": ["inna_nazwa", "wariant_nazwy"]
  },
  "mapping": [
    {
      "enabled": true,
      "label": "Etykieta w Wordzie",
      "source": "kolumna_excel",
      "transform": "identity",
      "const": ""
    },
    {
      "enabled": true,
      "label": "Koordynaty",
      "sources": ["xcoord", "ycoord"],
      "transform": "format",
      "const": "{xcoord}, {ycoord}"
    }
  ]
}
```

Więcej informacji: [`docs/preset_format.md`](docs/preset_format.md).

## Wbudowane transformacje

Najważniejsze transformacje używane przez presety:

- `identity` – bez zmiany wartości,
- `m2_to_ha_round2` – konwersja z m² na ha z zaokrągleniem,
- `date_only` – wyciągnięcie samej daty,
- `prelim_to_bedomning` – zamiana wartości technicznej na opis oceny,
- `constant` – wartość stała,
- `format` – złożenie wartości z kilku kolumn według szablonu.

## Zasady dla danych i wyników

Do repozytorium nie powinny trafiać:

- pliki Excel z prawdziwymi danymi,
- zdjęcia terenowe,
- wygenerowane dokumenty Word,
- logi,
- foldery robocze,
- archiwa ZIP z wynikami,
- dane prywatne lub lokalizacyjne, jeśli nie są świadomie przeznaczone do publikacji.

Do pracy lokalnej można używać np. takich folderów:

```text
indata/
output/
TEMP/
```

Są one ignorowane przez `.gitignore`.

## Release

Repozytorium zawiera skrypt pomocniczy:

```bash
python tools/build_release_zip.py
```

Skrypt buduje paczkę ZIP w folderze `dist/`. Paczka release powinna zawierać tylko pliki potrzebne użytkownikowi, bez danych wejściowych i wynikowych.

Zalecany tag release:

```text
v0.1.0
```

Zalecany tytuł release:

```text
Excel2Word z presetami v0.1.0
```

## Roadmap / możliwe dalsze prace

- eksport ustawień GUI jako nowy preset JSON,
- dodatkowe style tabel,
- tryb PDF,
- prostszy kreator presetów,
- podział kodu na mniejsze moduły, jeśli projekt będzie dalej rozwijany.

## Ograniczenia

- Projekt jest narzędziem roboczym, nie pełną aplikacją produkcyjną.
- Presety są przykładami konfiguracji i nie zawierają danych wejściowych ani wynikowych.
- Repozytorium jest tymczasowo publiczne i może później zostać przeniesione do prywatnego trybu.

## Licencja

MIT. Szczegóły w pliku [`LICENSE`](LICENSE).
