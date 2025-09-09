# Excel → Word: mikro‑tabele z Excela (GUI)

Konwerter, który **rozbija wiersze Excela** na **„mikro‑tabele” w Wordzie**. Działa na `.xlsx` i `.xls`, ma **proste GUI**, edytowalne **mapowanie pól**, kolejność, transformacje (m.in. m²→ha, sama data, logika „Säker/Preliminärt”), oraz **ramki** na zdjęcie i mapę.

> Projekt robiony pod użytek „Naturvärdesbiotop” — nazwy kolumn po szwedzku/angielsku zachowano celowo.

---

## Funkcje

* ✅ Odczyt **Excel** (`.xlsx` – openpyxl, `.xls` – xlrd ≥ 2.0.1)
* ✅ Generacja **DOCX**: dla **każdego wiersza** powstaje mikro‑tabela (Nagłówek ↔ Wartość)
* ✅ **Mapowanie pól**: etykieta (Nowa nazwa), źródło (kolumna Excela), transformacja, wartość stała, **włącz/wyłącz**
* ✅ **Kolejność** (przyciski ↑/↓) + usuwanie/dodawanie pozycji
* ✅ **Transformacje** wbudowane:

  * `identity` – bez zmian
  * `m2_to_ha_round2` – m² → ha, zaokrąglenie 0.01 (opcjonalny przecinek)
  * `date_only` – obcięcie czasu → `YYYY-MM-DD`
  * `prelim_to_bedomning` – `0/nej/false/puste` → **Säker**, inne → **Preliminärt**
  * `constant` – stała wartość (do ręcznego uzupełnienia)
* ✅ Zakładka **„Pozostałe kolumny”** – dodawanie extra pól (domyślnie wyłączone)
* ✅ **Ramki** pod: zdjęcie oraz **mapę** (wysokość w cm)
* ✅ **Podział strony** po zdjęciu/mapie (checkbox)
* ✅ **Auto‑bootstrap** pakietów (`pandas`, `python-docx`, `openpyxl`, `xlrd>=2.0.1`)

---

## Wymagania

* Python 3.8+
* Dostęp do internetu (pierwsze uruchomienie doinstaluje brakujące pakiety). W środowiskach firmowych z blokadami `pip` – zainstaluj paczki ręcznie.

```bash
pip install -U pandas python-docx openpyxl xlrd>=2.0.1
```

---

## Uruchomienie

```bash
python excel_rows_to_word_gui.py
```

1. Wskaż plik **Excel** i **folder wyjściowy**.
2. Kliknij **„Wczytaj i załaduj mapę”**.
3. (Opcjonalnie) Dostosuj mapę w zakładce **Mapowanie pól**.
4. (Opcjonalnie) Zaznacz dodatkowe kolumny w zakładce **Pozostałe kolumny**.
5. Ustaw **Opcje** (ramki, wysokości, podział strony, przecinek/kropka).
6. **Generuj DOCX**.

> Braki wartości ze źródła są zapisywane jako **„ - ”** (nie dotyczy pól `constant`).

---

## GUI — skrót

### Mapowanie pól

* Lista pozycji (kolejność „Kolej.”) — każdy wiersz to jeden wiersz tabeli w Wordzie.
* **Edytuj / dodaj wiersz**:

  * **Włączony** — czy pozycja trafia do Worda.
  * **Nowa nazwa** — etykieta w lewej kolumnie tabeli Word.
  * **Źródło** — kolumna z Excela (może być puste przy `constant`).
  * **Transformacja** — patrz listę wyżej.
  * **Stała** — tekst dla `constant`.
  * **Zastosuj** — aktualizuje wybraną pozycję.
  * **Dodaj z kolumny…** — tworzy pozycję mapującą wybraną kolumnę.
  * **Dodaj wiersz** — tworzy pustą pozycję `constant` (do wypełnienia „Stała”).
* **↑/↓** – zmiana kolejności, **Usuń** — usuwa pozycję z mapy.

### Pozostałe kolumny

* Szybkie dodanie dodatkowych pól (checkboxy). Trafiają na **koniec** tabeli Worda jako `identity`.

### Opcje

* **Dodaj ramkę na zdjęcie** (wysokość w cm)
* **Dodaj ramkę na mapę** (wysokość w cm)
* **Podział strony po każdym rekordzie** (po zdjęciu i mapie)
* **Użyj przecinka dziesiętnego** (np. `0,01` zamiast `0.01`)

---

## Domyślna mapa (preset „Naturvärdesbiotop”)

Wstępnie skonfigurowane pozycje (domyślnie **włączone**):

1. **Naturvärdesbiotop** ← `objektnummer`
2. **Naturvärdesklass** ← `naturvardesklass`
3. **Areal (ha)** ← `Shape__Area` (`m2_to_ha_round2`)
4. **Naturtyp** ← `naturtyp`
5. **Biotoptyp** ← `constant` (puste — do uzupełnienia)
6. **Hydrologisk huvudgrupp** ← `hydromorfologiskTyp`
7. **Natura 2000-naturtyp** ← `constant`
8. **Beskrivning** ← `objektbeskrivning`
9. **Biotopvärde** ← `biotopvarden`
10. **Tidigare kända värdearter** ← `vardearterKandaTidigare`
11. **Inventerade värdearter** ← `vardearterObserverade`
12. **Invasiva främmande arter** ← `invasivaFrammandeArter`
13. **Artvärde** ← `artvarden`
14. **Motivering till naturvärdesklass** ← `constant`
15. **Datum för fältbesök** ← `datumForObjektavgr` (`date_only`)
16. **Inventerare** ← `utforare`
17. **Bedömning** ← `preliminarAvgransning` (`prelim_to_bedomning`)

> Własne pozycje możesz dodać przyciskiem **„Dodaj wiersz”** (stała) lub **„Dodaj z kolumny…”**.

---

## Dostosowania

* **Domyślny podział strony ON**: w kodzie, w `_build_opt_tab()` ustaw
  `self.var_break = tk.BooleanVar(value=True)`
* **Format daty inny niż `YYYY-MM-DD`**: łatwo dodać nową transformację (np. `date_format_ddmmyyyy`).
* **Szablon stylów (.dotx)**: można podpiąć dokument bazowy — zgłoś issue/PR.

---

## Rozwiązywanie problemów

* **„xlrd wymagany ≥ 2.0.1”** – skrypt spróbuje doinstalować. Jeśli brak uprawnień, zainstaluj ręcznie:
  `pip install -U xlrd>=2.0.1`
* **Puste komórki w Word** → pojawia się `" - "` — to normalne (poza `constant`).
* **Szerokość kolumn** – ustalona na \~6 cm (lewa) + reszta (prawa). Zmienisz w kodzie: `left_w_cm`.

---

## Roadmap (pomysły)

* Presety mapy **zapis/odczyt JSON**
* Tryb **1 plik DOCX na wiersz** (np. `{objektnummer}.docx`)
* Automatyczne osadzanie zdjęcia/mapy z kolumn (plik/URL)
* Formatowanie dat (profile SE/ISO)

---

## Licencja

MIT — do użytku wewnętrznego i zewnętrznego. Jeśli używasz w innym projekcie, zostaw wzmiankę o autorze.

---

## Zrzuty (propozycje)

W repo dodaj do `docs/` i podmień ścieżki:

![Główne okno](docs/screenshot_gui_main.png)
![Mapowanie pól](docs/screenshot_mapping.png)
![Wynik DOCX](docs/screenshot_docx.png)
