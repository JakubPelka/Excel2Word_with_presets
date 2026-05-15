# Checklist przed release

## Pliki i struktura

- [ ] Główny działający skrypt znajduje się w `script/excel_rows_to_word_gui.py`.
- [ ] W root repozytorium nie ma kopii roboczych typu `*_dev.py`, `*_old.py`, `backup.py`.
- [ ] Folder `presets/` zawiera aktualne presety.
- [ ] Plik `gdy_2_pola.txt` został świadomie zostawiony albo przeniesiony po późniejszej weryfikacji.
- [ ] `requirements.txt` zawiera wymagane zależności.
- [ ] `.gitignore` blokuje dane wejściowe, wyniki, logi i paczki release.
- [ ] `README.md` opisuje aktualną strukturę repozytorium.

## Bezpieczeństwo i higiena

- [ ] Repozytorium nie zawiera plików Excel z prawdziwymi danymi.
- [ ] Repozytorium nie zawiera zdjęć terenowych.
- [ ] Repozytorium nie zawiera wygenerowanych plików `.docx`.
- [ ] Repozytorium nie zawiera lokalnych ścieżek użytkownika.
- [ ] Repozytorium nie zawiera tokenów, haseł ani plików `.env`.
- [ ] Repozytorium nie zawiera folderów backupowych.

## Test minimalny

- [ ] Program uruchamia się poleceniem `python script/excel_rows_to_word_gui.py`.
- [ ] Program poprawnie wczytuje plik Excel.
- [ ] Program poprawnie wczytuje preset JSON.
- [ ] Program generuje plik `.docx`.
- [ ] Plik wynikowy trafia do lokalnego folderu ignorowanego przez Git.
- [ ] Po teście w GitHub Desktop nie pojawiają się przypadkowe pliki wynikowe.

## Release

- [ ] Skrypt `python tools/build_release_zip.py` tworzy paczkę ZIP w `dist/`.
- [ ] Paczka ZIP nie zawiera danych wejściowych ani wynikowych.
- [ ] GitHub Release ma tag `v0.1.0`.
- [ ] Paczka ZIP została dodana jako asset release.
