# Konwerter JSON → XLSX dla HKL

## Opis ogólny

Program `converter.py` obsługuje dwa tryby pracy:

1. **Konwersja JSON → XLSX** — konwertuje pliki JSON z zamówieniami na pliki XLSX
   z formułami Excela i tablicami słownikowymi
2. **Translator EFOR → PROD** — automatycznie tłumaczy parametry zamówień EFOR
   na parametry produkcyjne PROD, ucząc się mapowań z danych treningowych

---

## Część I: Konwersja JSON → XLSX

### Struktura pliku JSON

Każdy plik JSON zawiera jedno zamówienie z następującą strukturą:

```
{
  "orderno": 131,           ← numer zamówienia
  "orderid": 1349,          ← ID zamówienia
  "commission": "...",       ← komisja/nazwa projektu
  "client": "Maxsol v.o.f.",← klient
  "organizationIdent": "HKL",← identyfikator organizacji
  "userIdent": "Maxsol",    ← identyfikator użytkownika
  "created_date": "...",     ← data utworzenia
  "sentDate": "...",         ← data wysłania
  "name", "address", "zip", "city", "country", "email", "phone", "tax"
  "items": [                 ← lista pozycji zamówienia
    {
      "posid": 2685,         ← ID pozycji
      "orderpos": 1,         ← numer pozycji
      "product": "14",       ← ID produktu (klucz do wyboru szablonu)
      "department": "VERTIKAL", ← dział
      "product_description": "VERTICALE JALOEZIE",
      "commission": "C006",
      "parameters": { ... }  ← parametry produktu
    },
    ...
  ]
}
```

### Struktura pliku XLSX (wynikowego)

Każdy element z tablicy `items` generuje osobny plik XLSX.

#### Kolumny A-L: Dane (wypełniane z JSON)

| Kolumna | Opis | Źródło danych |
|---------|------|---------------|
| **A** | Nazwy pól nagłówka | stałe: address, city, client, comment, commission, country, email, name, orderno, organizationIdent, phone, sentDate, tax, userIdent, zip, orderid |
| **B** | Wartości nagłówka | z JSON: odpowiednie pola najwyższego poziomu |
| **C** | Nazwy pól pozycji | stałe: product, comment, orderpos, posid, product_description, commission, department |
| **D** | Wartości pól pozycji | z JSON: odpowiednie pola obiektu item |
| **E** | Nazwy parametrów | stałe, zależne od typu produktu (np. ILOSC, KONFIGURACJA, MODEL...) |
| **F** | Wartości parametrów | z JSON: `parameters[NAZWA]` |
| **G** | Opisy parametrów | z JSON: `parameters[NAZWA___DESCRIPTION]` |
| **H** | Aliasy | z JSON: `parameters[NAZWA_ALIAS]` |
| **I** | Opisy aliasów | z JSON: `parameters[NAZWA_ALIAS___DESCRIPTION]` |
| **J** | Tytuły parametrów | z JSON: `parameters[NAZWA___TITLE]` |
| **K** | Widoczność (1.0/0.0) | z JSON: `parameters[NAZWA___VISIBLE]` |
| **L** | Czy słownik (1.0/0.0) | z JSON: `parameters[NAZWA___DICT]` |

#### Kolumny N-W: Przetworzenie (formuły Excela)

Te kolumny zawierają formuły Excela, które automatycznie przeliczają dane
po otwarciu pliku w Excelu. Formuły odnoszą się do:
- Komórek w kolumnach E-L (dane z JSON)
- Tablic słownikowych w kolumnach X+ (mapowanie wartości)

Przykłady formuł:
- **N2** (`organizacja_kod`): Mapuje `organizationIdent` → kod: HKL="04", Cozy="05", etc.
- **Q2** (`kom`): Składa numer komisji z pól zamówienia
- **S2-S26**: Formuły VLOOKUP mapujące kody z JSON na wartości produkcyjne

#### Kolumny X+: Słowniki (tablice mapowania)

Zawierają tablice referencyjne używane przez formuły VLOOKUP. Przykłady:

| Kolumny | Słownik | Przykład |
|---------|---------|---------|
| AD-AG | RODZAJ → GAT-PROD | S="STANDARD" → ST="Standard" |
| AI-AR | MODEL → MODEL-PROD | 1L="LINKS" → 00="Clasic", sterowanie lewe |
| AT-AW | KOLOR → KOLOR-PROD | 1120-127 → kod 1120, szer. lameli 127 |
| AY-BB | PASEK → BLAM-PROD | B1="50MM" → 50 |
| BD-BG | KOLOR_SYSTEMU → SYST_KOLOR | W="BIAŁY" → 0204="biała" |
| BI-BL | KORALIK_OBSL → KORALIK | KW="PLASTIKOWY BIAŁY" |
| BN-BQ | MONTAZ → Uch-PROD | D="SUFIT" → 1="Klips" |
| CC-CF | OCHRONA_CHILD_SAFETY → CS | CS_NO → N="bez zabezpieczenia" |

---

### Mapowanie nagłówka JSON → Wiersze w kolumnach A-B

| Wiersz | Pole A (nazwa) | Pole B (wartość z JSON) |
|--------|----------------|------------------------|
| 2 | address | `json.address` |
| 3 | city | `json.city` |
| 4 | client | `json.client` |
| 5 | comment | `json.comment` |
| 6 | commission | `json.commission` |
| 7 | country | `json.country` |
| 8 | email | `json.email` |
| 9 | name | `json.name` |
| 10 | orderno | `json.orderno` |
| 11 | organizationIdent | `json.organizationIdent` |
| 12 | phone | `json.phone` |
| 13 | sentDate | `json.sentDate` |
| 14 | tax | `json.tax` |
| 15 | userIdent | `json.userIdent` |
| 16 | zip | `json.zip` |
| 17 | orderid | `json.orderid` |

### Mapowanie pozycji JSON → Wiersze w kolumnach C-D

| Wiersz | Pole C (nazwa) | Pole D (wartość z JSON) |
|--------|----------------|------------------------|
| 2 | product | `item.product` (jako liczba) |
| 3 | comment | `item.comment` |
| 4 | orderpos | `item.orderpos` (jako liczba) |
| 5 | posid | `item.posid` (jako liczba) |
| 6 | product_description | `item.product_description` |
| 7 | commission | `item.commission` |
| 8 | department | `item.department` |

---

### Obsługiwane typy produktów (szablony)

Każdy typ produktu wymaga osobnego szablonu XLSX, ponieważ:
- Parametry (kolumna E) mają inną kolejność i zestaw
- Formuły w kolumnach N-W są inne
- Słowniki w kolumnach X+ są inne

| Product ID | Dział (Department) | Opis produktu | Szablon z pliku |
|-----------|-------------------|---------------|-----------------|
| **14** | VERTIKAL | VERTICALE JALOEZIE | `HKL_Maxsol_131-1(1).xlsx` |
| **71** | PLISSÉGORDIJNEN | EOS | `HKL_Maxsol_157-1(1).xlsx` |
| **24** | RZYMSKIE | RZYMSKIE | `HKL_NaOkniePL_28-1(1).xlsx` |
| **43** | PLISSEE | COSIFLOR | `LUXANGMBH_TEST_LuxanDE_40-1(1).xlsx` |
| **39** | JALOUSIEN | ALUJALOUSIE 16/25 | `LUXANGMBH_TEST_LuxanDE_40-5(1).xlsx` |
| **59** | JALOUSIEN | HOLZJALOUSIEN | `LUXANGMBH_TEST_LuxanDE_40-6(1).xlsx` |
| **75** | ROLLOS | ABSOLUTE | `LUXANGMBH_TEST_LuxanDE_41-1(1).xlsx` |
| **80** | MAGNETISCH FRAME | MFPlus | *wygenerowany automatycznie przez GPT API* |

### Automatyczne generowanie szablonów (GPT API fallback)

```
Pozycja z JSON → Czy istnieje szablon?
  TAK → Konwersja deterministyczna (istniejący kod)
  NIE → GPT API analizuje JSON → Tworzy XLSX + zapisuje szablon do ponownego użycia
```

Gdy konwerter napotyka typ produktu bez szablonu:

1. **Ekstrakcja nazw parametrów** — z pominięciem metadanych (`___DESCRIPTION`, `_ALIAS`, etc.)
2. **Zapytanie do GPT** — ustala kolejność parametrów w szablonie
3. **Tworzenie szablonu** — nagłówki, formuły rejestracyjne (N-O), formuły pozycji (P-Q)
4. **Zapis** — `templates/template_{ID}.xlsx` do ponownego użycia

**Uwaga:** Szablony GPT nie zawierają formuł VLOOKUP (kolumny S+) ani słowników (kolumny X+) —
te elementy należy uzupełnić ręcznie.

---

## Część II: Translator EFOR → PROD

### Opis

System EFOR generuje parametry zamówień w formacie JSON. Muszą one zostać
przetłumaczone na parametry produkcyjne (PROD) dla produkcji. Translator uczy się
mapowań automatycznie z danych treningowych dostarczonych przez Tomasza.

### Przepływ danych

```
Faza uczenia:
  10.xlsx → parsowanie par → nauka mapowań kluczy/wartości → zapis do mappings/*.json
                                    ↓ (niezmapowane klucze)
                              GPT sugestie → gpt_suggestions w mappings/*.json

Faza tłumaczenia:
  JSON → ekstrakcja parametrów EFOR → dopasowanie ASORTMENT → zastosowanie mapowań → wynik.xlsx
                                                            ↓ (jeśli nieznany)     ↑ (kolumny GPT = czerwone)
                                                      GPT fallback → sugestia mapowania
```

### Algorytm uczenia

Dane treningowe (`10/10.xlsx`) zawierają 10 000 par wejście→wyjście z 32 typów
asortymentu. Dla każdego klucza wyjściowego (PROD) algorytm:

1. Sprawdza czy wartość jest stała (identyczna we wszystkich wierszach) → **stała**
2. Dla każdego klucza wejściowego (EFOR) oblicza:
   - Dopasowanie dokładne (`input == output`) → **kopiowanie (copy)**
   - Dopasowanie po dzieleniu (`input / 10 == output`) → **dzielenie (divide10)**
   - Spójne mapowanie wartości → **słownik (lookup)**
3. Wybiera najlepszy klucz źródłowy z wynikiem > 60% par
4. Zapisuje mapowanie do pliku JSON

### Wykrywane transformacje

| Typ | Przykład | Opis |
|-----|---------|------|
| `copy` | SZEROKOSC=3000 → B=3000 | Wartość kopiowana bez zmian |
| `divide10` | SZEROKOSC=985.0 → B=98.5 | Wartość dzielona przez 10 |
| `lookup` | STEROWANIE=S → STER_TYP=SS | Słownik mapujący wartości |
| stała | — → TRANS=T | Zawsze ta sama wartość |

### Dopasowanie ASORTMENT

Pole `department` z JSON (np. "VERTIKAL") musi zostać dopasowane do nazwy ASORTMENT
z danych treningowych (np. "Vertikale"). Strategia:

1. Dokładne dopasowanie nazwy
2. Dopasowanie podciągu (case-insensitive) — "VERTIKAL" pasuje do "Vertikale"
3. Próba z polem `product_description`
4. Jeśli brak dopasowania → GPT fallback

### Struktura plików mapowań

Każdy plik `mappings/{ASORTMENT}.json` zawiera:

```json
{
  "asortment": "Żaluzje drewniane 2010",
  "key_map": {
    "B": {"source": "SZEROKOSC", "transform": "divide10"},
    "H": {"source": "WYSOKOSC", "transform": "divide10"},
    "KOLOR_ZAM": {"source": "KOLOR", "transform": "copy"},
    "STER_TYP": {"source": "STEROWANIE", "transform": "lookup"}
  },
  "value_map": {
    "STER_TYP": {"S": "SS", "D": "DD"}
  },
  "constants": {
    "TRANS": "T",
    "KOLOR2": "-"
  },
  "gpt_suggestions": {
    "MODEL": {"source": "SYSTEM_MODEL", "transform": "lookup", "value_map": {"S ctr2": "S ctr2"}, "description_pl": "Model rolety", "confidence": "medium"},
    "STER": {"source": "manual", "transform": "manual", "description_pl": "Rodzaj sterowania", "confidence": "low"}
  }
}
```

### Sugestie GPT dla niezmapowanych kluczy

Podczas fazy uczenia (`learn`) niektóre klucze wyjściowe PROD nie mogą zostać
automatycznie dopasowane do żadnego klucza wejściowego EFOR (wynik dopasowania = 0).
Przykłady: KOMPONENTDE (konkatenacja), MODEL (obcięty), STER (tłumaczenie na niemiecki),
SYST_KOLOR (podciąg), TUBA/CS (zależne od źródła).

Dla takich kluczy program wysyła zapytanie do GPT, które:
1. Szuka podobnego już zmapowanego klucza PROD i proponuje analogiczne mapowanie
2. Jeśli brak podobnego klucza — tworzy opis w języku polskim
3. Zapisuje sugestię w pliku mapowania JSON w polu `gpt_suggestions`

Każda sugestia GPT zawiera:
- `source` — proponowany klucz źródłowy EFOR (lub `"manual"` jeśli brak dopasowania)
- `transform` — typ transformacji (`copy`, `divide10`, `lookup`, `manual`)
- `description_pl` — opis parametru po polsku
- `confidence` — poziom pewności (`high`, `medium`, `low`)
- `reason` — uzasadnienie wyboru (po angielsku)
- `value_map` — opcjonalny słownik mapowania wartości (dla transformacji `lookup`)

### Plik wynikowy (wynik.xlsx)

- Wiersz 1: nazwy kluczy PROD (nagłówki)
- Wiersz 2+: przetłumaczone wartości (jeden wiersz na pozycję z JSON)
- Puste wartości / NULL → "-"
- Wartości liczbowe zachowują typ numeryczny
- **Kolumny z sugestiami GPT** są podświetlone na **czerwono** (nagłówek + dane)
  w celu łatwej identyfikacji i ręcznej weryfikacji
- Klucze z transformacją `manual` mają wartość `"?"` jako placeholder do uzupełnienia

---

## Użycie programu

### Konwersja JSON → XLSX

```bash
# Przygotowanie szablonów (jednorazowo)
python3 converter.py setup

# Konwersja pojedynczego pliku
python3 converter.py convert HKL_Maxsol_131.json

# Konwersja wszystkich plików JSON
python3 converter.py convert-all

# Lista dostępnych szablonów
python3 converter.py list-templates
```

### Translator EFOR → PROD

```bash
# Nauka mapowań z danych treningowych (domyślnie 10/3.xlsx)
python3 converter.py learn
python3 converter.py learn 10/10.xlsx

# Tłumaczenie zamówienia JSON na parametry PROD
python3 converter.py translate HKL_NaOkniePL_29.json
python3 converter.py translate zamowienie.json wynik_custom.xlsx

# Lista nauczonych mapowań ze statystykami
python3 converter.py list-mappings
```

### Katalogi

| Katalog | Opis |
|---------|------|
| `templates/` | Szablony XLSX (jeden na typ produktu) |
| `importyzefordoprod/` | Pliki wejściowe (JSON + przykłady XLSX) |
| `output/` | Pliki wynikowe (wygenerowane XLSX) |
| `mappings/` | Nauczone mapowania EFOR→PROD (pliki JSON) |
| `10/` | Dane treningowe (3.xlsx, 10.xlsx) |

### Konwencja nazewnictwa plików wynikowych

**Konwersja:** `{nazwa_json}({timestamp})-{nr_sekwencyjny}({nr_w_typie}).xlsx`
```
HKL_Maxsol_131(20260218_233702)-1(1).xlsx   ← pozycja 1, produkt 14 (1. w typie)
```

**Translator:** `wynik_{nazwa_json}.xlsx`
```
wynik_HKL_NaOkniePL_29.xlsx
```

---

## Konwersja wartości

| Wartość w JSON | Wartość w XLSX |
|----------------|----------------|
| `null` | `<NULL>` |
| `""` (pusty string) | `<NULL>` |
| `true` (boolean) | `1.0` |
| `false` (boolean) | `0.0` |
| `665` (liczba) | `665` (liczba) |
| `"S"` (tekst) | `"S"` (tekst) |

---

## Obsługa zagnieżdżonych parametrów

Parametr `WYMIAROWANIE_SLOPOW` może mieć wartość prostą (pusty string)
lub zagnieżdżony obiekt (dict) z podpolami:

```json
"WYMIAROWANIE_SLOPOW": {
  "TYP": "TYP19",
  "WYM_B": 2090,
  "WYM_H1": 2460,
  "WYM_H2": 1870
}
```

W konwersji:
- Wiersz główny (WYMIAROWANIE_SLOPOW) → wartość `<NULL>`
- Podpola (WYMIAROWANIE_SLOPOW.TYP, .WYM_B, etc.) → odpowiednie wartości

W translatorze podpola są spłaszczane do formatu `KLUCZ.PODKLUCZ`.

---

## Ograniczenia i uwagi

1. **Formuły Excela** — zachowywane w pliku XLSX, ale nie przeliczane przez program.
   Przeliczenie nastąpi po otwarciu w Excelu.

2. **Szablony GPT** — nie zawierają formuł VLOOKUP (kolumny S+) ani słowników (X+).
   Te elementy należy uzupełnić ręcznie.

3. **Dane treningowe** — zawierają mieszane formaty wejściowe (EFOR/polski,
   SRCID/zewnętrzne systemy, format niemiecki). Algorytm uczy się z dominującego
   formatu. Dla asortymentów bez danych EFOR, translator korzysta z GPT fallback.

4. **Klucz API** — bez klucza OpenAI API w `.env` lub zmiennej środowiskowej,
   nieznane typy produktów/asortymentów będą pomijane, a sugestie GPT dla
   niezmapowanych kluczy nie zostaną wygenerowane.

5. **Sugestie GPT** — kolumny oznaczone czerwonym tłem w pliku wynikowym wymagają
   ręcznej weryfikacji. Wartości `"?"` oznaczają klucze, dla których GPT nie znalazł
   źródła danych i wymagają ręcznego uzupełnienia.

---

## Wymagania

- Python 3.6+
- Biblioteka `openpyxl` (`pip3 install openpyxl`)
- Klucz OpenAI API (opcjonalnie, dla GPT fallback)

### Pliki konfiguracyjne

| Plik | Opis |
|------|------|
| `.env` | Klucz API: `OPENAI_API_KEY=sk-proj-...` |
| `.gitignore` | Wyklucza `.env` z repozytorium |
