# Format presetów JSON

Preset JSON opisuje, które kolumny z pliku Excel mają zostać użyte w dokumencie Word oraz jak mają zostać podpisane i przekształcone.

## Minimalna struktura

```json
{
  "name": "Nazwa presetu",
  "version": 1,
  "options": {},
  "aliases": {},
  "mapping": []
}
```

## `name`

Nazwa presetu widoczna dla użytkownika.

```json
"name": "Naturvärdesbiotop"
```

## `version`

Numer wersji formatu presetu. Na razie stosowana jest wartość:

```json
"version": 1
```

## `options`

Opcje dokumentu i formatowania.

Przykład:

```json
"options": {
  "decimal_comma": true,
  "page_break": true,
  "photo": {
    "enabled": true,
    "height_cm": 6.0
  },
  "map": {
    "enabled": true,
    "height_cm": 6.0
  }
}
```

## `aliases`

Alias pozwala dopasować pole nawet wtedy, gdy kolumna w Excelu ma inną nazwę niż oczekiwana.

Przykład:

```json
"aliases": {
  "objektnummer": ["objekt_nr", "id", "lokal_id"]
}
```

## `mapping`

Lista pól, które mają zostać dodane do dokumentu Word.

Przykład pojedynczego pola:

```json
{
  "enabled": true,
  "label": "Objektnummer",
  "source": "objektnummer",
  "transform": "identity",
  "const": ""
}
```

Przykład pola złożonego z kilku kolumn:

```json
{
  "enabled": true,
  "label": "Koordynaty",
  "sources": ["xcoord", "ycoord"],
  "transform": "format",
  "const": "{xcoord}, {ycoord}"
}
```

## Obsługiwane transformacje

- `identity` – zwraca wartość bez zmian,
- `m2_to_ha_round2` – przelicza m² na ha i zaokrągla do dwóch miejsc,
- `date_only` – zwraca samą datę,
- `prelim_to_bedomning` – zamienia wartość logiczną/techniczną na opis oceny,
- `constant` – wstawia stałą wartość z pola `const`,
- `format` – składa tekst z kilku kolumn według szablonu z pola `const`.

## Dobre praktyki

- Nie usuwaj pól z presetu tylko dlatego, że są chwilowo nieużywane — można ustawić `"enabled": false`.
- Nie zapisuj w presetach danych osobowych, lokalizacji wrażliwych ani prawdziwych rekordów terenowych.
- Preset powinien opisywać strukturę danych, a nie zawierać dane.
- Przy zmianie presetu warto zwiększyć numer wersji lub dodać krótką notatkę w `CHANGELOG.md`.
