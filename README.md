# Excel → Word: mikro‑tabele z Excela (GUI)

Konwerter, który **rozbija wiersze Excela** na **„mikro‑tabele” w Wordzie**. Działa na `.xlsx` i `.xls`, ma **GUI**, edytowalne **mapowanie pól**, kolejność, transformacje (m.in. m²→ha, sama data, logika „Säker/Preliminärt”), oraz **ramki** na zdjęcie i mapę. Projekt przygotowany pod „Naturvärdesbiotop”, ale dzięki **presetom JSON** jest uniwersalny.

> Wersja „preset‑first”: brak domyślnej mapy w kodzie. Najpierw wybierasz **Excel**, potem **preset**.

---

## Co nowego (UX)

* ✅ **Automatyczne wczytywanie kolumn** po wybraniu pliku Excel (bez dodatkowego przycisku).
* ✅ **Brak domyślnej mapy** – mapa startuje pusta; wczytujesz **preset** z JSON.
* ✅ **„Pozostałe kolumny”** są **puste do czasu wczytania presetu** (placeholder). Po wczytaniu pokazują tylko **kolumny nieużyte** w mapie i aktualizują się po każdej zmianie.
* ✅ **Ramka na mapę** (po zdjęciu) z osobną wysokością.
* ✅ **Podział strony po każdym rekordzie** – ustawienie w Opcjach (domyślnie **włączone** w przykładowej konfiguracji).
* ✅ **Help** pod edytorem mapy ma zawijanie tekstu.
* ✅ **Brak pop‑upów** po wczytaniu kolumn (czytelny feedback w UI).

---

## Funkcje

* Odczyt **Excel** (`.xlsx` – openpyxl, `.xls` – xlrd ≥ 2.0.1)
* Generacja **DOCX**: dla **każdego wiersza** powstaje mikro‑tabela (Nagłówek ↔ Wartość)
* **Mapowanie pól**: Etykieta (Nowa nazwa), Źródło (kolumna Excela), Transformacja, Stała, Włącz/Wyłącz
* **Kolejność** (↑/↓), **Dodaj z kolumny…**, **Dodaj wiersz** (stała), Usuń
* **Transformacje** wbudowane:

  * `identity` – bez zmian
  * `m2_to_ha_round2` – m² → ha, zaokrąglenie 0.01 (obsługa przecinka)
  * `date_only` – tylko część daty `YYYY‑MM‑DD`
  * `prelim_to_bedomning` – `0/nej/false/puste` → **Säker**, inne → **Preliminärt**
  * `constant` – stała wartość (do ręcznego uzupełnienia)
* **Ramki**: zdjęcie + mapa (wysokość w cm), **podział strony** po nich
* **Lokalizacja liczb**: opcja **przecinka dziesiętnego**
* **Braki wartości**: zapisywane jako **„ - ”** (nie dotyczy pól `constant`)
* **Presety JSON**: konfigurują mapę, aliasy kolumn i opcje dokumentu

---

## Wymagania

* Python 3.8+
* Internet przy pierwszym uruchomieniu (auto‑bootstrap pip). W środowiskach z blokadami – zainstaluj ręcznie:

```bash
pip install -U pandas python-docx openpyxl xlrd>=2.0.1
```

---

## Uruchomienie

```bash
python excel_rows_to_word_gui.py
```

**Przepływ („preset‑first”):**

1. Kliknij **Wybierz…** obok pola *Plik Excel* → kolumny wczytają się **automatycznie** (mapa pozostaje pusta).
2. Kliknij **Wczytaj preset…** i wskaż JSON (np. `presets/naturvardesbiotop.json`).
3. W zakładce **Mapowanie pól** dopracuj ewentualne szczegóły (etykiety, transformacje, kolejność, stałe).
4. W **Pozostałych kolumnach** (po presetcie) zaznacz, co jeszcze dorzucić.
5. W **Opcjach** ustaw ramki (zdjęcie/mapa), wysokości, **podział strony**, przecinek/kropka.
6. **Generuj DOCX**.

> Jeśli mapa jest pusta **i** nic nie wybrano w „Pozostałe kolumny”, skrypt ostrzeże przed generowaniem pustego dokumentu.

---

## Presety (JSON)

Presety są w katalogu `presets/`. Zawierają nazwę, opcje, aliasy kolumn (pomagają trafić w różne nazwy) i listę mapowań.

**Przykład:** `presets/naturvardesbiotop.json`

```json
{
  "name": "Naturvärdesbiotop (preset)",
  "version": 1,
  "options": {
    "decimal_comma": true,
    "page_break": true,
    "photo": { "enabled": true, "height_cm": 6.0 },
    "map":   { "enabled": true, "height_cm": 6.0 }
  },
  "aliases": {
    "Shape__Area": ["Shape_Area", "ShapeArea"],
    "objektnummer": ["objectnumber", "objektnr"],
    "hydromorfologiskTyp": ["hydromorfologi", "hydromorfologisk_typ"],
    "vardearterKandaTidigare": ["vardeArterKandaTidigare", "vardearterKändaTidigare"],
    "vardearterObserverade": ["vardeArterObserverade", "vardearterObserverade"],
    "invasivaFrammandeArter": ["invasivaFrämmandeArter", "invasivaArter"],
    "datumForObjektavgr": ["datumFörObjektavgr", "datumForFaltbesok", "datumFältbesök"],
    "preliminarAvgransning": ["preliminärAvgränsning", "preliminar_avgransning"],
    "naturvardesklass": ["naturvärdesklass", "naturvarde_klass"]
  },
  "mapping": [
    { "enabled": true,  "label": "Naturvärdesbiotop",                "source": "objektnummer",            "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Naturvärdesklass",                 "source": "naturvardesklass",        "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Areal (ha)",                       "source": "Shape__Area",             "transform": "m2_to_ha_round2",      "const": "" },
    { "enabled": true,  "label": "Naturtyp",                         "source": "naturtyp",                "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Biotoptyp",                        "source": "",                        "transform": "constant",             "const": "" },
    { "enabled": true,  "label": "Hydrologisk huvudgrupp",           "source": "hydromorfologiskTyp",     "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Natura 2000-naturtyp",             "source": "",                        "transform": "constant",             "const": "" },
    { "enabled": true,  "label": "Beskrivning",                      "source": "objektbeskrivning",       "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Biotopvärde",                      "source": "biotopvarden",            "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Tidigare kända värdearter",        "source": "vardearterKandaTidigare", "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Inventerade värdearter",           "source": "vardearterObserverade",   "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Invasiva främmande arter",         "source": "invasivaFrammandeArter",  "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Artvärde",                         "source": "artvarden",               "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Motivering till naturvärdesklass", "source": "",                        "transform": "constant",             "const": "" },
    { "enabled": true,  "label": "Datum för fältbesök",              "source": "datumForObjektavgr",      "transform": "date_only",            "const": "" },
    { "enabled": true,  "label": "Inventerare",                      "source": "utforare",                "transform": "identity",             "const": "" },
    { "enabled": true,  "label": "Bedömning",                        "source": "preliminarAvgransning",   "transform": "prelim_to_bedomning",  "const": "" }
  ]
}
```

> **Aliasowanie kolumn** pozwala, by preset zadziałał nawet, gdy w Excelu nazwy są trochę inne.

---

## GUI — skrót

### Mapowanie pól

* Lista pozycji (każda → wiersz w Wordzie).
* **Edytuj / dodaj wiersz**:

  * **Włączony** – czy pozycja trafi do Worda
  * **Nowa nazwa** – etykieta w tabeli Word
  * **Źródło** – kolumna z Excela (może być puste przy `constant`)
  * **Transformacja** – patrz lista wyżej
  * **Stała** – tekst dla `constant`
  * **Zastosuj** – aktualizacja wybranego wiersza
  * **Dodaj z kolumny…** – tworzy pozycję z wybranego źródła
  * **Dodaj wiersz** – tworzy pustą pozycję `constant`
* **↑/↓** – zmiana kolejności, **Usuń** – kasuje pozycję

### Pozostałe kolumny

* **Przed presetem:** placeholder (pusto).
* **Po presetcie:** lista tylko **nieużytych** kolumn – zaznacz, by dodać je na koniec tabeli (transformacja `identity`).

### Opcje dokumentu

* **Ramka zdjęcia** (+ wysokość w cm)
* **Ramka mapy** (+ wysokość w cm)
* **Podział strony po każdym rekordzie** (po zdjęciu i mapie)
* **Użyj przecinka dziesiętnego** (np. `0,01`)

---

## Zachowanie wartości pustych

* Dla pól ze źródła (`identity`, `m2_to_ha_round2`, `date_only`, `prelim_to_bedomning`) brak wartości → **„ - ”**.
* Dla `constant` można zostawić puste – do ręcznego wpisania w GUI.

---

## Rozwiązywanie problemów

* **XLS nie wczytuje się** → potrzebny `xlrd>=2.0.1` (skrypt spróbuje doinstalować). W razie blokad: `pip install -U xlrd>=2.0.1`.
* **Brak treści w DOCX** → upewnij się, że mapa nie jest pusta **albo** zaznaczono „Pozostałe kolumny”.
* **Format daty** potrzebny inny niż `YYYY-MM-DD`? Dodajemy łatwo transformację `date_format`.

---
