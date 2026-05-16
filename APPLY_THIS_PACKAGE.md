# Jak zastosować paczkę porządkującą

Ta paczka jest przygotowana jako bezpieczny overlay dla repozytorium `Excel2Word_with_presets`.

Nie zawiera aktualnego działającego skryptu programu, żeby nie nadpisać Twojej sprawdzonej wersji.

## Co zrobić

1. Skopiuj zawartość tej paczki do root repozytorium.
2. Przenieś aktualny działający plik:

   ```text
   excel_rows_to_word_gui.py
   ```

   do:

   ```text
   script/excel_rows_to_word_gui.py
   ```

3. Plik:

   ```text
   excel_rows_to_word_gui_dev.py
   ```

   przenieś lokalnie do `TEMP/` albo usuń z repo, jeśli nie jest już potrzebny.

4. Nie ruszaj folderu `presets/`, jeśli chcesz zachować obecne presety.
5. Na razie zostaw `presets/gdy_2_pola.txt`, zgodnie z decyzją do późniejszej weryfikacji.
6. Sprawdź w GitHub Desktop, czy do commita trafiają tylko pliki porządkujące i przeniesiony skrypt.
7. Uruchom minimalny test:

   ```bash
   python script/excel_rows_to_word_gui.py
   ```

8. Po teście zbuduj ZIP release:

   ```bash
   python tools/build_release_zip.py
   ```

## Sugerowane commity

### Commit 1

```text
Reorganize repository structure for v0.1.0
```

Zakres:

- `script/`
- przeniesienie głównego skryptu
- usunięcie/przeniesienie wersji dev z root

### Commit 2

```text
Add release and hygiene files
```

Zakres:

- `.gitignore`
- `.gitattributes`
- `requirements.txt`
- `LICENSE`
- `CHANGELOG.md`
- `dist/README.md`
- `tools/build_release_zip.py`
- `docs/`

### Commit 3

```text
Update Polish README for release workflow
```

Zakres:

- `README.md`

## Release notes dla GitHub Releases

Tytuł:

```text
Excel2Word z presetami v0.1.0
```

Treść:

```md
Pierwsza uporządkowana wersja release.

Zawiera:

- GUI do konwersji wierszy Excela na mikro-tabele w Wordzie,
- obsługę presetów JSON,
- mapowanie pól i transformacje wartości,
- obsługę zdjęć oraz ramki mapy,
- układy A/B,
- sortowanie naturalne,
- konfigurację czcionek, marginesów i szerokości kolumn,
- logowanie przebiegu generowania dokumentu,
- paczkę ZIP przygotowaną przez `tools/build_release_zip.py`.

Uwaga:

- repozytorium i dokumentacja pozostają celowo po polsku,
- release nie zawiera danych wejściowych ani wynikowych,
- repozytorium jest tymczasowo publiczne.
```
