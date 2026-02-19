# Konwerter JSON → XLSX dla HKL

## Opis ogólny

Program `converter.py` konwertuje pliki JSON z zamówieniami na pliki XLSX,
zachowując formuly Excela (kolumny N-W) oraz tablice słownikowe (kolumny X+).
Wymieniane są jedynie dane w kolumnach A-L.

Dla znanych typów produktów (14, 71, 24, 43, 39, 59, 75) konwersja jest
deterministyczna — używane są gotowe szablony. Dla nieznanych typów (np. 80)
program automatycznie wywołuje API OpenAI GPT, aby przeanalizować strukturę
JSON i wygenerować nowy szablon, który zostaje zapisany do ponownego użycia.

---

## Jak działa konwersja

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

## Mapowanie nagłówka JSON → Wiersze w kolumnach A-B

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

## Mapowanie pozycji JSON → Wiersze w kolumnach C-D

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

## Obsługiwane typy produktów (szablony)

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

### Kolejność parametrów w szablonie VERTIKAL (#14)

```
Wiersz  Parametr
  2     ILOSC
  3     KONFIGURACJA
  4     RODZAJ
  5     MODEL
  6     KOLOR_SYSTEMU
  7     KOLOR
  8     PASEK
  9     SZEROKOSC
 10     WYS_RODZ
 11     WYSOKOSC
 12     WYMIAROWANIE_SLOPOW
 13-18  WYMIAROWANIE_SLOPOW.TYP / .WYM_B / .WYM_B1 / .WYM_H / .WYM_H1 / .WYM_H2
 19     ILOSC_PASK
 20     KORALIK_OBSL
 21     DLUGOSC_STER
 22     MOTOR
 23     ZASILANIE
 24     PILOT
 25     AUTOMATYKA
 26     KIESZEN
 27     KORALIK_DOL
 28     MONTAZ
 29     OCHRONA_CHILD_SAFETY
 30     WYSOKOSC_MONTAZU
 31     POW
 32-33  INSTRUKCJA / FILM (puste)
 34     CENA
 35     DOPLATA
 36     DOPLATA_EL
 37     CENA_SUMA
 38     SUMA_BRUTTO
 39     CENA_RABAT
 40     DOPLATA_EL_RABAT
 41     CENA_KONCOWA
 42     WARTOSC_KONCOWA
 43     OPIS_POZYCJI
 44     OPIS_CENY
 45     OPIS_RABATU
```

### Kolejność parametrów w szablonie PLISSÉGORDIJNEN EOS (#71)

```
Wiersz  Parametr
  2     ILOSC
  3     MODEL
  4     KOLOR
  5     KOLOR_DODATKOWY
  6     KOLOR_SYSTE
  7     SZEROKOSC
  8     WYSOKOSC
  9-14  WYMIAROWANIE_SLOPOW.TYP / .WYM_B / .WYM_B1 / .WYM_H / .WYM_H1 / .WYM_H2
 15     STEROWANIE
 16     MOTOR
 17     ZASILANIE
 18     PILOT
 19     AUTOMATYKA
 20     STRONA_STEROWANIA
 21     FUNKCJA
 22     KOLOR_STEROWANIA
 23     STEROWANIE_ELEKTRYCZNE
 24     STRONA_LADOWANIA
 25     STRONA_WYJSCIA_PRZEWODU
 26     OCHRONA_CHILD_SAFETY
 27     WYSOKOSC_MONTAZU
 28     DLUGOSC_STER
 29     PROWADNICE
 30     PROW_KOLOR
 31     MONTAZ
 32     DODATKI
 33     POW
 34-35  INSTRUKCJA / FILM (puste)
 36     CENA
 37     DOPLATA
 38     DOPLATA_EL
 39     CENA_SUMA
 40     SUMA_BRUTTO
 41     CENA_RABAT
 42     DOPLATA_RABAT
 43     DOPLATA_EL_RABAT
 44     CENA_KONCOWA
 45     WARTOSC_KONCOWA
 46     OPIS_POZYCJI
 47     OPIS_CENY
 48     OPIS_RABATU
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
  "WYM_H2": 1870,
  ...
}
```

W takim przypadku:
- Wiersz główny (WYMIAROWANIE_SLOPOW) → wartość `<NULL>`
- Podpola (WYMIAROWANIE_SLOPOW.TYP, .WYM_B, etc.) → odpowiednie wartości

---

## Automatyczne generowanie szablonów (GPT API fallback)

### Schemat działania

```
Pozycja z JSON → Czy istnieje szablon?
  TAK → Konwersja deterministyczna (istniejący kod)
  NIE → GPT API analizuje JSON → Tworzy XLSX + zapisuje szablon do ponownego użycia
```

### Jak to działa

Gdy konwerter napotyka typ produktu bez szablonu (np. produkt 80 — MAGNETISCH FRAME):

1. **Wykrycie braku szablonu** — program wypisuje komunikat:
   `No template for product 80 (MAGNETISCH FRAME). Calling GPT API...`

2. **Ekstrakcja nazw parametrów** — z JSON wyodrębniane są nazwy bazowe
   parametrów (z pominięciem przyrostków metadanych: `___DESCRIPTION`,
   `_ALIAS`, `___TITLE`, `___VISIBLE`, `___DICT`)

3. **Zapytanie do GPT** — program buduje prompt zawierający:
   - Opis oczekiwanej struktury XLSX (kolumny E-L)
   - Przykład kolejności parametrów ze znanego produktu (VERTIKAL #14)
   - Listę parametrów nowego produktu
   - Prośbę o zwrócenie tablicy JSON z parametrami w prawidłowej kolejności

4. **Tworzenie szablonu** — na podstawie odpowiedzi GPT tworzony jest plik XLSX:
   - Wiersz 1: nagłówki (identyczne jak w istniejących szablonach)
   - Kolumna E: nazwy parametrów w kolejności określonej przez GPT
   - Kolumny N-O: uniwersalne formuły rejestracyjne (organizacja, użytkownik, ident)
   - Kolumny P-Q: formuły pozycji z **dynamicznymi odniesieniami do wierszy**
     (program sam znajduje, w którym wierszu wylądowały ILOSC, CENA, CENA_KONCOWA, itp.)

5. **Zapis szablonu** — szablon zapisywany jest jako `templates/template_{ID}.xlsx`

6. **Konwersja** — dane z JSON wypełniane są w nowo utworzonym szablonie
   (identycznie jak dla szablonów deterministycznych)

7. **Ponowne uruchomienie** — przy kolejnym uruchomieniu szablon jest już zapisany,
   więc GPT API nie jest wywoływane ponownie

### Dynamiczne formuły w kolumnach P-Q

Program automatycznie dopasowuje formuły cenowe do pozycji parametrów:

| Wiersz Q | Pole | Formuła (przykład dla produktu 80) |
|----------|------|-------------------------------------|
| Q17 | `liczba` | `=F2` (odniesienie do wiersza ILOSC) |
| Q19 | `cena_podst` | `=IF(ISERROR($F$17/1),0,$F$17)` (odniesienie do CENA) |
| Q22 | `cena` | `=IF(ISERROR($F$23/1),Q19,$F$23)` (odniesienie do CENA_KONCOWA) |
| Q23 | `nazwa` | `=$D$8&" "&$D$6&" "&$F$21` (odniesienie do OPIS_POZYCJI) |
| Q24 | `cena_info` | `=$F$22` (odniesienie do OPIS_CENY) |
| Q25 | `rabat_info` | `=$F$23` (odniesienie do OPIS_RABATU) |

### Konfiguracja klucza API

Klucz OpenAI API należy umieścić w pliku `.env` w katalogu głównym konwertera:

```
OPENAI_API_KEY=sk-proj-...
```

Program szuka klucza w następującej kolejności:
1. Plik `.env` (priorytet)
2. Zmienna środowiskowa `OPENAI_API_KEY`

Plik `.env` jest dodany do `.gitignore` i nie jest śledzony w repozytorium.

### Uwagi dotyczące szablonów GPT

- **Formuły w kolumnach N-W** mogą wymagać ręcznej korekty, szczególnie
  formuły VLOOKUP w kolumnach S+ (mapowanie słownikowe)
- **Kolumny X+ (słowniki)** nie są generowane automatycznie — jeśli produkt
  wymaga mapowania słownikowego, tablice należy dodać ręcznie
- **Model GPT**: domyślnie `gpt-4o-mini` (szybki, tani, wystarczający do
  ustalenia kolejności parametrów)
- **Brak dodatkowych zależności**: komunikacja z API odbywa się przez `urllib`
  (biblioteka standardowa Pythona)

---

## Użycie programu

### 1. Przygotowanie szablonów (jednorazowo)

```bash
python3 converter.py setup
```

Skanuje istniejące pliki xlsx w katalogu `importyzefordoprod/` i wyodrębnia
jeden szablon na typ produktu do katalogu `templates/`.

### 2. Konwersja pojedynczego pliku JSON

```bash
python3 converter.py convert HKL_Maxsol_131.json
```

### 3. Konwersja wszystkich plików JSON

```bash
python3 converter.py convert-all
```

### 4. Lista dostępnych szablonów

```bash
python3 converter.py list-templates
```

### Katalogi

| Katalog | Opis |
|---------|------|
| `templates/` | Szablony XLSX (jeden na typ produktu) |
| `importyzefordoprod/` | Pliki wejściowe (JSON + przykłady XLSX) |
| `output/` | Pliki wynikowe (wygenerowane XLSX) |

### Konwencja nazewnictwa plików wynikowych

```
{nazwa_json}({timestamp})-{nr_sekwencyjny}({nr_w_typie_produktu}).xlsx
```

Przykład:
```
HKL_Maxsol_131(20260218_233702)-1(1).xlsx   ← pozycja 1, produkt 14 (1. w typie)
HKL_Maxsol_131(20260218_233702)-2(2).xlsx   ← pozycja 2, produkt 14 (2. w typie)
```

---

## Ograniczenia i uwagi

1. **Formuły Excela** — formuły są zachowywane w pliku XLSX, ale nie są
   przeliczane przez program. Przeliczenie nastąpi automatycznie po otwarciu
   pliku w Excelu.

2. **Nowy typ produktu** — obsługiwany na dwa sposoby:
   - **Automatycznie (GPT)**: jeśli skonfigurowany jest klucz API, szablon
     zostanie wygenerowany przy pierwszej konwersji. Wymaga późniejszej
     weryfikacji formuł i ewentualnego ręcznego dodania słowników (kolumny X+).
   - **Ręcznie**: umieść przykładowy plik XLSX w `importyzefordoprod/`
     i uruchom `python3 converter.py setup` — szablon zostanie wyodrębniony
     ze wszystkimi formułami i słownikami.

3. **Szablony GPT vs. ręczne** — szablony wygenerowane przez GPT zawierają
   poprawne formuły cenowe (kolumny N-Q), ale **nie mają**:
   - Formuł VLOOKUP mapujących kody (kolumny S+)
   - Tablic słownikowych (kolumny X+)
   Te elementy należy uzupełnić ręcznie, jeśli są potrzebne.

4. **Weryfikacja danych** — po konwersji (zwłaszcza z szablonu GPT) zaleca się
   otwarcie pliku w Excelu i sprawdzenie poprawności formuł.

5. **Klucz API** — bez klucza OpenAI API w pliku `.env` lub zmiennej
   środowiskowej, nieznane typy produktów będą pomijane (SKIP) jak poprzednio.

---

## Wymagania

- Python 3.6+
- Biblioteka `openpyxl` (`pip3 install openpyxl`)
- Klucz OpenAI API (opcjonalnie, dla automatycznego generowania szablonów)

### Pliki konfiguracyjne

| Plik | Opis |
|------|------|
| `.env` | Klucz API: `OPENAI_API_KEY=sk-proj-...` |
| `.gitignore` | Wyklucza `.env` z repozytorium |
