# Excel → Word: mikro‑tabele (GUI)

Konwerter, który **rozbija wiersze Excela** na **„mikro‑tabele” w Wordzie**. Obsługuje `.xlsx` i `.xls`, ma wygodne **GUI**, **presety JSON**, edytowalne **mapowanie pól**, transformacje, sortowanie „naturalne”, automatyczne **ramki na zdjęcia i mapę**, układy **A/B**, globalną **czcionkę i marginesy**, kompresję zdjęć oraz **log** z przebiegu.

> Wersja *preset‑first*: najpierw wybierasz plik Excel, potem preset JSON.

---

## Co nowego

* **Układ B (POPRAWIONY)** – prawa kolumna:

  * wiersz 0 to **nagłówek** (szare tło `#A6A6A6`, biały bold);
  * **poniżej** jedna scalona komórka na zdjęcie (**bez** tabeli w tabeli i bez dzielenia pola).
* **Naturalne sortowanie** (human sort) – np. `Plats 2` < `Plats 10`.
* **Transform `format`** – budowanie wartości z wielu kolumn wg szablonu (np. `"{xcoord}, {ycoord}"`).
* **Automatyczne zdjęcia**: wskazujesz folder, kolumnę z ID (domyślnie `objektnummer`).

  * Drugi kadr rozpoznawany po sufiksach: `_2`, `-2`, `(2)`, ` (2)`.
  * **Układ A** może dzielić ramkę na **2 pola** (A: tak, B: nie – zgodnie z powyższą poprawką).
  * Zdjęcia są **centrowane**, skalowane bez nadmiernego powiększania i **kompresowane**.
* **Styl nagłówka** tabeli: tło `#A6A6A6`, biały, pogrubiony tekst.
* **Globalna czcionka i marginesy** (domyślnie: Arial 9 pt; L=2.0, P=6.5, G=3.9, D=3.0 cm).
* **Szerokości kolumn**:

  * Układ A: szerokość kolumny etykiet.
  * Układ B: szer. etykiety, wartości i kolumny zdjęcia.
* **Puste wartości** zapisywane jako `" - "` (nie dotyczy `constant`).
* **Log** z pracy w katalogu wyjściowym (łatwe diagnozowanie problemów).

---

## Wymagania

* Python 3.8+
* Biblioteki: `pandas`, `python-docx`, `openpyxl`, `xlrd>=2.0.1` (dla `.xls`), `Pillow`.
* Skrypt sam doinstaluje zależności (jeśli środowisko na to pozwala). W środowiskach zamkniętych zainstaluj ręcznie.

---

## Szybki start

1. Uruchom: `python excel_rows_to_word_gui.py` (lub wersję *\_dev*).
2. **Plik Excel** → kolumny wczytają się automatycznie.
3. **Wczytaj preset…** (JSON) → pojawi się mapa pól.
4. W zakładce **Mapowanie pól** dopracuj etykiety, kolejność, transformacje.
5. W **Pozostałych kolumnach** zaznacz to, co chcesz dodać *ponad preset*.
6. W **Opcjach** ustaw: zdjęcie/mapę, wysokości, **podział strony**, **sortowanie**, **układ A/B**, **czcionkę**, **marginesy**, **szerokości kolumn**, kompresję zdjęć.
7. (Opcjonalnie) wskaż **folder zdjęć** i **kolumnę ID** do dopasowania.
8. Kliknij **Generuj DOCX**.

---

## Presety JSON

Struktura:

```json
{
  "name": "Nazwa preset",
  "version": 1,
  "options": {
    "decimal_comma": true,
    "page_break": true,
    "photo": { "enabled": true, "height_cm": 6.0 },
    "map":   { "enabled": true, "height_cm": 6.0 }
  },
  "aliases": { "aliasKolumny": ["inne_nazwy"] },
  "mapping": [
    { "enabled": true, "label": "Etykieta", "source": "kolumna", "transform": "identity", "const": "" },
    { "enabled": true, "label": "Koordinater", "sources": ["xcoord","ycoord"], "transform": "format", "const": "{xcoord}, {ycoord}" }
  ]
}
```

**Wyłączone** wiersze: `"enabled": false` (zostają w pliku, ale nie trafiają do Worda).

### Transformy wbudowane

* `identity` – bez zmian
* `m2_to_ha_round2` – m² → ha (zaokrąglenie 0,01; lokalizacja kropka/przecinek)
* `date_only` – tylko data `YYYY‑MM‑DD`
* `prelim_to_bedomning` – 0/"nej"/false/puste → **Säker**, w przeciwnym razie **Preliminärt**
* `constant` – stała wartość
* `format` – składanie z wielu kolumn wg szablonu z `const`

---

## GUI – przegląd

**Główne pola**

* Plik Excel, Folder wyjściowy, Nazwa pliku `.docx`
* Folder zdjęć (opcjonalnie), **Kolumna dopasowania zdjęć** (np. `objektnummer`)
* *Wczytaj preset…* – wybranie presetu JSON

**Zakładki**

* **Mapowanie pól** – lista pozycji (każda → wiersz w Wordzie): włącz/wyłącz, etykieta, źródło, transform, stała; przyciski **↑/↓**, **Usuń**, **Dodaj z kolumny…**, **Dodaj wiersz** (stała). Pomoc pod listą ma zawijanie tekstu.
* **Pozostałe kolumny** – pokaże tylko **nieużyte** kolumny po wczytaniu presetu.
* **Opcje**:

  * **Dokument**: ramka **zdjęcia** (wysokość,
    *A: opcjonalny podział na 2 pola*; *B: zawsze jedna komórka*), ramka **mapy**, **podział strony**, przecinek dziesiętny.
  * **Sortowanie**: wybór kolumny + kierunek (A→Z/Z→A). Dla tekstu używany jest **human sort**.
  * **Układ**: **A** (tabela, potem zdjęcie i mapa) / **B** (tabela z prawą kolumną na zdjęcie; mapa pod spodem).
  * **Czcionka**: rodzina i rozmiar (domyślnie Arial 9 pt).
  * **Szerokości kolumn**:

    * A: szerokość kolumny etykiet;
    * B: szerokości etykiety, wartości, kolumny zdjęcia.
  * **Zdjęcia – kompresja**: jakość JPG, DPI eksportu, „nie powiększaj małych zdjęć”.
  * **Marginesy**: lewy/prawy/górny/dolny w cm.

---

## Zasady zdjęć

* Nazwa pliku == wartość z kolumny ID (np. `Lokal 2.jpg`). Dozwolone także warianty z `_` i numerem bez spacji.
* Drugie zdjęcie (tylko **Układ A** przy *podziale ramki*): sufiksy `_2`, `-2`, `(2)`, ` (2)`.
* Zdjęcia są wyśrodkowane i skalowane do ramki; re‑kodowanie do JPG (domyślnie **90** i **450 DPI**).

---

## Styl tabeli i puste wartości

* Pierwszy wiersz każdej tabeli ma tło `#A6A6A6` i **biały, pogrubiony** tekst.
* Puste wartości zapisywane jako **`" - "`** (nie dotyczy pól `constant`).

---

## Rozwiązywanie problemów

* **.xls nie wczytuje się** – zainstaluj `xlrd>=2.0.1`.
* **Kolejność `1,10,11,2…`** – włącz sortowanie w Opcjach (kolumna + kierunek); skrypt stosuje *human sort* dla tekstu.
* **Zbyt duże pliki DOCX** – zmniejsz *Jakość JPG* lub *DPI eksportu*.
* **Czcionka** nie jest stosowana – upewnij się, że font istnieje w systemie; w razie czego wybierz inną z listy.
* **Diagnoza** – zajrzyj do pliku logu (tworzony obok pliku wynikowego).

---

## Struktura repo (propozycja)

```
Excel2Word_with_presets/
├─ excel_rows_to_word_gui.py
├─ presets/
│  ├─ naturvardesbiotop.json
│  ├─ landskapsomrade.json
│  ├─ vardeelement.json
│  └─ sst.json
├─ indata/
│  └─ foto/               # tu można trzymać zdjęcia
└─ output/
```

---

## Pomysły (roadmap)

* Wtyczki zewnętrznych transformacji (plug‑in pattern).
* Eksport ustawień GUI jako preset (*.json*) jednym kliknięciem.
* Predefiniowane style tabel (np. linie cienkie/grube, kolorystyka).
* Eksport do PDF.

---

**Autorzy / wkład:** automatyzacja oszczędzająca godziny żmudnego copy‑paste – dzięki! 😉
