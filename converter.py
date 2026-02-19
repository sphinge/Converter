#!/usr/bin/env python3
"""
HKL JSON to XLSX Converter

Converts JSON order files to XLSX files using product-specific templates.
Templates preserve Excel formulas (columns N-W) and dictionary tables (columns X+).
Only data in columns A-L is replaced from JSON.

Usage:
    python3 converter.py setup                # Extract templates from existing xlsx files
    python3 converter.py convert file.json    # Convert one JSON file
    python3 converter.py convert-all          # Convert all JSON files in input dir

Steps:
    1. Run 'setup' once to extract templates from existing xlsx examples
    2. Run 'convert' or 'convert-all' to convert JSON files
"""

import json
import os
import sys
import urllib.request
import urllib.error
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("openpyxl is required. Install with: pip3 install openpyxl")
    sys.exit(1)


# ── Configuration ──────────────────────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
INPUT_DIR = os.path.join(BASE_DIR, "importyzefordoprod")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Order-level JSON fields → row numbers in columns A-B
HEADER_ROWS = {
    'address': 2,
    'city': 3,
    'client': 4,
    'comment': 5,
    'commission': 6,
    'country': 7,
    'email': 8,
    'name': 9,
    'orderno': 10,
    'organizationIdent': 11,
    'phone': 12,
    'sentDate': 13,
    'tax': 14,
    'userIdent': 15,
    'zip': 16,
    'orderid': 17,
}

# Item-level JSON fields → row numbers in columns C-D
ITEM_ROWS = {
    'product': 2,
    'comment': 3,
    'orderpos': 4,
    'posid': 5,
    'product_description': 6,
    'commission': 7,
    'department': 8,
}

# Fields that should be written as numbers in column D
NUMERIC_ITEM_FIELDS = {'product', 'orderpos', 'posid'}


# ── Helper functions ───────────────────────────────────────────────────────────

def to_excel(val):
    """Convert a Python/JSON value to Excel cell value.
    None and empty strings become '<NULL>'.
    Booleans become 1.0 / 0.0.
    Numbers and non-empty strings pass through.
    """
    if val is None:
        return '<NULL>'
    if isinstance(val, bool):
        return 1.0 if val else 0.0
    if isinstance(val, str) and val.strip() == '':
        return '<NULL>'
    return val


def read_param_rows(ws):
    """Read column E of the template to get parameter_name → row mapping.
    Returns two dicts:
      - main_params: parameters without dots (e.g., WYMIAROWANIE_SLOPOW → row)
      - sub_params:  sub-field parameters with dots (e.g., WYMIAROWANIE_SLOPOW.TYP → row)
    """
    main_params = {}
    sub_params = {}
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row=row, column=5).value  # Column E
        if name and isinstance(name, str):
            if '.' in name:
                sub_params[name] = row
            else:
                main_params[name] = row
    return main_params, sub_params


# ── GPT API fallback ──────────────────────────────────────────────────────────

def get_api_key():
    """Load OpenAI API key from .env file or environment variable."""
    env_path = os.path.join(BASE_DIR, '.env')
    if os.path.exists(env_path):
        with open(env_path, 'r') as f:
            for line in f:
                line = line.strip()
                if line.startswith('OPENAI_API_KEY='):
                    key = line.split('=', 1)[1].strip().strip('"').strip("'")
                    if key and not key.startswith('sk-proj-PASTE'):
                        return key
    return os.environ.get('OPENAI_API_KEY')


def extract_base_param_names(parameters):
    """Extract base parameter names from JSON, excluding metadata suffixes."""
    suffixes = ('_ALIAS___DESCRIPTION', '___DESCRIPTION', '___TITLE',
                '___VISIBLE', '___DICT', '_ALIAS')
    base_names = []
    seen = set()
    for key in parameters:
        is_suffix = False
        for s in suffixes:
            if key.endswith(s):
                is_suffix = True
                break
        if not is_suffix and key not in seen:
            base_names.append(key)
            seen.add(key)
    return base_names


def build_gpt_prompt(item):
    """Build a GPT prompt to determine parameter ordering for a new product."""
    params = item.get('parameters', {})
    base_names = extract_base_param_names(params)

    prompt = (
        "You are helping create an Excel template for an HKL order system.\n"
        "I need the parameters ordered logically for a product spreadsheet.\n"
        "Parameters go in column E starting at row 2.\n\n"
        "Example ordering for VERTIKAL #14:\n"
        "ILOSC, KONFIGURACJA, RODZAJ, MODEL, KOLOR_SYSTEMU, KOLOR, PASEK, "
        "SZEROKOSC, WYS_RODZ, WYSOKOSC, WYMIAROWANIE_SLOPOW, ILOSC_PASK, "
        "KORALIK_OBSL, DLUGOSC_STER, MOTOR, ZASILANIE, PILOT, AUTOMATYKA, "
        "KIESZEN, KORALIK_DOL, MONTAZ, OCHRONA_CHILD_SAFETY, WYSOKOSC_MONTAZU, "
        "POW, INSTRUKCJA, FILM, CENA, CENA_SUMA, CENA_RABAT, DOPLATA, "
        "SUMA_BRUTTO, WARTOSC_KONCOWA, CENA_KONCOWA, OPIS_POZYCJI, OPIS_CENY, "
        "OPIS_RABATU\n\n"
        "The general pattern:\n"
        "1. ILOSC (quantity) first\n"
        "2. Product type/model params (MODEL, PROD, etc.)\n"
        "3. Color params (KOLOR, KOLOR_RAMY, KOLOR_SYSTE, etc.)\n"
        "4. Dimension params (SZEROKOSC, WYSOKOSC)\n"
        "5. Technical/feature params (STEROWANIE, etc.)\n"
        "6. Pricing: CENA, CENA_SUMA, CENA_RABAT, DOPLATA, SUMA_BRUTTO, "
        "WARTOSC_KONCOWA, CENA_KONCOWA\n"
        "7. Descriptions last: OPIS_POZYCJI, OPIS_CENY, OPIS_RABATU\n\n"
        f"New product: {item.get('department', '?')} (#{item.get('product', '?')})\n"
        f"Parameters: {', '.join(base_names)}\n\n"
        "Return ONLY a JSON array with these exact parameter names in the "
        "correct order. No explanation, just the JSON array."
    )
    return prompt


def call_gpt_api(prompt, api_key, model="gpt-4o-mini"):
    """Call OpenAI API using urllib (no extra dependencies)."""
    url = "https://api.openai.com/v1/chat/completions"

    payload = json.dumps({
        "model": model,
        "messages": [
            {"role": "system", "content": "You return only valid JSON."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0
    }).encode('utf-8')

    req = urllib.request.Request(url, data=payload, method='POST')
    req.add_header('Content-Type', 'application/json')
    req.add_header('Authorization', f'Bearer {api_key}')

    try:
        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode('utf-8'))
            content = result['choices'][0]['message']['content'].strip()
            # Strip markdown code fences if present
            if content.startswith('```'):
                content = content.split('\n', 1)[1]
                content = content.rsplit('```', 1)[0].strip()
            return json.loads(content)
    except urllib.error.HTTPError as e:
        error_body = e.read().decode('utf-8')
        print(f"  GPT API error ({e.code}): {error_body[:200]}")
        return None
    except Exception as e:
        print(f"  GPT API error: {e}")
        return None


def create_template_from_gpt(param_order, item):
    """Create a template xlsx from GPT-determined parameter ordering."""
    wb = openpyxl.Workbook()
    ws = wb.active

    # Row 1 headers (matching existing templates)
    headers = {
        1: 'Nagłówek', 3: 'Pozycja', 5: 'Identyfikator', 6: 'Wart',
        7: 'Wart_opis', 8: 'Alias', 9: 'Alias_opis', 10: 'Tytuł',
        11: 'Widoczny', 12: 'Słownik', 14: 'Rej', 16: 'Poz',
    }
    for col, val in headers.items():
        ws.cell(row=1, column=col, value=val)

    # Column E: parameter names in GPT-determined order
    params = item.get('parameters', {})
    suffixes = ('_ALIAS___DESCRIPTION', '___DESCRIPTION', '___TITLE',
                '___VISIBLE', '___DICT', '_ALIAS')
    row = 2
    param_rows = {}

    for param_name in param_order:
        ws.cell(row=row, column=5, value=param_name)
        param_rows[param_name] = row
        row += 1

        # If the value is a dict, add sub-field rows
        val = params.get(param_name)
        if isinstance(val, dict):
            for sub_key in val:
                is_suffix = any(sub_key.endswith(s) for s in suffixes)
                if not is_suffix:
                    ws.cell(row=row, column=5, value=f"{param_name}.{sub_key}")
                    row += 1

    # ── Columns N-O: registration header formulas (universal) ──
    n_o = [
        (2, 'organizacja_kod',
         '=IF($B$11="Cozy","05",IF($B$11="HKL","04",'
         'IF($B$11="LuxanGmbH","03",IF($B$11="FENIX","02",'
         'IF($B$11="LuxanEwaKrawczyk","01","00")))))'),
        (3, 'organizacja', '=B11'),
        (4, 'uzytkownik', '=B15'),
        (5, 'ident', '=B15&"-"&B10'),
        (6, 'uwagi', '=$B$6'),
    ]
    for r, label, formula in n_o:
        ws.cell(row=r, column=14, value=label)
        ws.cell(row=r, column=15, value=formula)

    # ── Columns P-Q: position formulas ──
    # Universal formulas (reference only header/item rows in A-D)
    pq_universal = [
        (2,  'kom',                '="EO/"&$B$10&"/"&D4&"/"&$B$6&"-"&$D$7'),
        (3,  'odbiorca',           '=$B$15'),
        (4,  'dealer',             None),
        (5,  'kooperant',          None),
        (6,  'sprzedawca',         '=B11'),
        (7,  'adres_ident',        '=$B$15'),
        (8,  'adres_nazwa',        None),
        (9,  'adres_adres',        '=$B$2'),
        (10, 'adres_kodpoczt',     '=$B$16'),
        (11, 'adres_miejscowosc',  '=$B$3'),
        (12, 'adres_kraj',         '=$B$7'),
        (13, 'adres_nrtel',        '=$B$12'),
        (14, 'adres_email',        '=$B$8'),
        (15, 'adres_region',       None),
        (16, 'uwagi',              '=$D$3'),
    ]
    for r, label, formula in pq_universal:
        ws.cell(row=r, column=16, value=label)
        if formula:
            ws.cell(row=r, column=17, value=formula)

    # Dynamic formulas (adjusted to actual parameter row positions)
    ilosc_r = param_rows.get('ILOSC')
    cena_r = param_rows.get('CENA')
    cena_konc_r = param_rows.get('CENA_KONCOWA') or param_rows.get('WARTOSC_KONCOWA')
    opis_poz_r = param_rows.get('OPIS_POZYCJI')
    opis_ceny_r = param_rows.get('OPIS_CENY')
    opis_rab_r = param_rows.get('OPIS_RABATU')

    ws.cell(row=17, column=16, value='liczba')
    if ilosc_r:
        ws.cell(row=17, column=17, value=f'=F{ilosc_r}')

    ws.cell(row=18, column=16, value='ilosc')
    ws.cell(row=18, column=17, value=1)

    ws.cell(row=19, column=16, value='cena_podst')
    if cena_r:
        ws.cell(row=19, column=17,
                value=f'=IF(ISERROR($F${cena_r}/1),0,$F${cena_r})')

    ws.cell(row=20, column=16, value='rabat_proc')
    ws.cell(row=20, column=17,
            value='=ROUND(IF(ISERROR($Q$21/$Q$19),0,$Q$21/$Q$19),4)')

    ws.cell(row=21, column=16, value='rabat')
    ws.cell(row=21, column=17, value='=$Q$19-$Q$22')

    ws.cell(row=22, column=16, value='cena')
    if cena_konc_r:
        ws.cell(row=22, column=17,
                value=f'=IF(ISERROR($F${cena_konc_r}/1),Q19,$F${cena_konc_r})')
    else:
        ws.cell(row=22, column=17, value='=Q19')

    ws.cell(row=23, column=16, value='nazwa')
    if opis_poz_r:
        ws.cell(row=23, column=17,
                value=f'=$D$8&" "&$D$6&" "&$F${opis_poz_r}')
    else:
        ws.cell(row=23, column=17, value='=$D$8&" "&$D$6')

    ws.cell(row=24, column=16, value='cena_info')
    if opis_ceny_r:
        ws.cell(row=24, column=17, value=f'=$F${opis_ceny_r}')

    ws.cell(row=25, column=16, value='rabat_info')
    if opis_rab_r:
        ws.cell(row=25, column=17, value=f'=$F${opis_rab_r}')

    ws.cell(row=26, column=16, value='grupa_cenowa')
    ws.cell(row=26, column=17,
            value='=IF(ISERROR(FIND("#",$I$7,1)),"-",'
                  'MID($I$7,FIND("#",$I$7,1)+1,20))')

    return wb


# ── Template setup ─────────────────────────────────────────────────────────────

def setup_templates():
    """Scan existing xlsx files and extract one template per product type."""
    os.makedirs(TEMPLATES_DIR, exist_ok=True)

    if not os.path.isdir(INPUT_DIR):
        print(f"Input directory not found: {INPUT_DIR}")
        return {}

    xlsx_files = sorted(f for f in os.listdir(INPUT_DIR) if f.endswith('.xlsx'))
    found = {}

    for fname in xlsx_files:
        fpath = os.path.join(INPUT_DIR, fname)
        try:
            wb = openpyxl.load_workbook(fpath)
            ws = wb.active

            product_raw = ws.cell(row=2, column=4).value  # D2 = product ID
            department = ws.cell(row=8, column=4).value    # D8 = department

            if product_raw is not None:
                product_key = str(int(float(product_raw)))

                if product_key not in found:
                    import shutil
                    dest = os.path.join(TEMPLATES_DIR, f"template_{product_key}.xlsx")
                    wb.close()
                    shutil.copy2(fpath, dest)
                    found[product_key] = {'source': fname, 'department': department}
                    print(f"  Product {product_key:>3} ({department}) <- {fname}")
                    continue

            wb.close()
        except Exception as e:
            print(f"  Skip {fname}: {e}")

    print(f"\n{len(found)} templates extracted to {TEMPLATES_DIR}/")
    return found


# ── Data filling ───────────────────────────────────────────────────────────────

def fill_worksheet(ws, order, item):
    """Fill columns A-L of a template worksheet with JSON data."""

    # ── Columns A-B: order header ──
    for field, row in HEADER_ROWS.items():
        ws.cell(row=row, column=1, value=field)           # A
        ws.cell(row=row, column=2, value=to_excel(order.get(field)))  # B

    # ── Columns C-D: item metadata ──
    for field, row in ITEM_ROWS.items():
        val = item.get(field)
        ws.cell(row=row, column=3, value=field)            # C

        if field in NUMERIC_ITEM_FIELDS and val is not None:
            try:
                ws.cell(row=row, column=4, value=float(val))  # D as number
            except (ValueError, TypeError):
                ws.cell(row=row, column=4, value=to_excel(val))
        else:
            ws.cell(row=row, column=4, value=to_excel(val))  # D

    # ── Columns E-L: parameters ──
    params = item.get('parameters', {})
    main_param_rows, sub_param_rows = read_param_rows(ws)

    def write_param_row(row, param_name, val, desc, alias, alias_desc, title, visible, is_dict):
        """Write a single parameter row to columns E-L."""
        ws.cell(row=row, column=5,  value=param_name)               # E: name
        ws.cell(row=row, column=6,  value=to_excel(val))             # F: value
        ws.cell(row=row, column=7,  value=to_excel(desc))            # G: description
        ws.cell(row=row, column=8,  value=to_excel(alias))           # H: alias
        ws.cell(row=row, column=9,  value=to_excel(alias_desc))      # I: alias desc
        ws.cell(row=row, column=10, value=title or '')               # J: title
        ws.cell(row=row, column=11, value=1.0 if visible else 0.0)   # K: visible
        ws.cell(row=row, column=12, value=1.0 if is_dict else 0.0)   # L: dict

    for param_name, row in main_param_rows.items():
        if param_name not in params:
            continue

        val = params.get(param_name)

        # Handle nested dict values (e.g., WYMIAROWANIE_SLOPOW with sub-fields)
        if isinstance(val, dict):
            # Write '<NULL>' for the main parameter row
            desc   = params.get(f'{param_name}___DESCRIPTION')
            alias  = params.get(f'{param_name}_ALIAS')
            adesc  = params.get(f'{param_name}_ALIAS___DESCRIPTION')
            title  = params.get(f'{param_name}___TITLE')
            vis    = params.get(f'{param_name}___VISIBLE')
            isdict = params.get(f'{param_name}___DICT')
            write_param_row(row, param_name, '<NULL>', desc, alias, adesc, title, vis, isdict)

            # Fill sub-field rows (e.g., WYMIAROWANIE_SLOPOW.TYP)
            for sub_key, sub_row in sub_param_rows.items():
                if not sub_key.startswith(param_name + '.'):
                    continue
                field_name = sub_key.split('.', 1)[1]  # e.g., "TYP"
                sub_val    = val.get(field_name)
                sub_desc   = val.get(f'{field_name}___DESCRIPTION')
                sub_alias  = val.get(f'{field_name}_ALIAS')
                sub_adesc  = val.get(f'{field_name}_ALIAS___DESCRIPTION')
                sub_title  = val.get(f'{field_name}___TITLE')
                sub_vis    = val.get(f'{field_name}___VISIBLE')
                sub_isdict = val.get(f'{field_name}___DICT')
                write_param_row(sub_row, sub_key, sub_val, sub_desc,
                                sub_alias, sub_adesc, sub_title, sub_vis, sub_isdict)
        else:
            desc   = params.get(f'{param_name}___DESCRIPTION')
            alias  = params.get(f'{param_name}_ALIAS')
            adesc  = params.get(f'{param_name}_ALIAS___DESCRIPTION')
            title  = params.get(f'{param_name}___TITLE')
            vis    = params.get(f'{param_name}___VISIBLE')
            isdict = params.get(f'{param_name}___DICT')
            write_param_row(row, param_name, val, desc, alias, adesc, title, vis, isdict)


# ── Conversion ─────────────────────────────────────────────────────────────────

def convert_json(json_path, output_dir=None):
    """Convert a JSON file to XLSX files (one per item/position)."""
    if output_dir is None:
        output_dir = OUTPUT_DIR
    os.makedirs(output_dir, exist_ok=True)

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    base_name = os.path.splitext(os.path.basename(json_path))[0]
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    items = data.get('items', [])
    if not items:
        print(f"  No items in {json_path}")
        return []

    product_counters = {}
    created = []

    for seq, item in enumerate(items, 1):
        product_id = str(item.get('product', ''))
        department = item.get('department', '?')

        template_path = os.path.join(TEMPLATES_DIR, f"template_{product_id}.xlsx")
        if not os.path.exists(template_path):
            # GPT fallback: create template for unknown product type
            print(f"  No template for product {product_id} ({department}). Calling GPT API...")
            api_key = get_api_key()
            if not api_key:
                print(f"  SKIP item {seq}: no API key (set OPENAI_API_KEY in .env)")
                continue

            prompt = build_gpt_prompt(item)
            param_order = call_gpt_api(prompt, api_key)
            if not param_order:
                print(f"  SKIP item {seq}: GPT API call failed")
                continue

            os.makedirs(TEMPLATES_DIR, exist_ok=True)
            template_wb = create_template_from_gpt(param_order, item)
            template_wb.save(template_path)
            template_wb.close()
            print(f"  Template created: template_{product_id}.xlsx")
            print(f"  Note: formulas in columns N-W may need manual setup.")

        # Per-product-type counter for file naming
        product_counters[product_id] = product_counters.get(product_id, 0) + 1
        ppos = product_counters[product_id]

        # Load template (preserves formulas and dictionaries)
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active

        # Fill data
        fill_worksheet(ws, data, item)

        # Save
        out_name = f"{base_name}({timestamp})-{seq}({ppos}).xlsx"
        out_path = os.path.join(output_dir, out_name)
        wb.save(out_path)
        wb.close()

        created.append(out_name)
        print(f"  [{seq}/{len(items)}] {out_name}  (product {product_id}, {department})")

    print(f"  -> {len(created)}/{len(items)} files created")
    return created


# ── CLI ────────────────────────────────────────────────────────────────────────

def main():
    usage = """
HKL JSON to XLSX Converter

Commands:
  python3 converter.py setup               Extract templates from existing xlsx files
  python3 converter.py convert <file.json>  Convert a single JSON file
  python3 converter.py convert-all          Convert all JSON files in input directory
  python3 converter.py list-templates       Show available templates

Directories:
  Templates:  ./templates/
  Input:      ./importyzefordoprod/
  Output:     ./output/
"""

    if len(sys.argv) < 2:
        print(usage)
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == 'setup':
        print("Extracting templates from existing xlsx files...\n")
        setup_templates()

    elif cmd == 'convert':
        if len(sys.argv) < 3:
            print("Usage: python3 converter.py convert <file.json>")
            sys.exit(1)

        json_path = sys.argv[2]
        # Try as-is, then relative to INPUT_DIR
        if not os.path.exists(json_path):
            alt = os.path.join(INPUT_DIR, json_path)
            if os.path.exists(alt):
                json_path = alt
            else:
                print(f"File not found: {json_path}")
                sys.exit(1)

        if not os.path.isdir(TEMPLATES_DIR) or not os.listdir(TEMPLATES_DIR):
            print("No templates found. Run 'python3 converter.py setup' first.")
            sys.exit(1)

        print(f"Converting: {os.path.basename(json_path)}")
        convert_json(json_path)

    elif cmd == 'convert-all':
        if not os.path.isdir(TEMPLATES_DIR) or not os.listdir(TEMPLATES_DIR):
            print("No templates found. Run 'python3 converter.py setup' first.")
            sys.exit(1)

        json_files = sorted(
            os.path.join(INPUT_DIR, f)
            for f in os.listdir(INPUT_DIR)
            if f.endswith('.json')
        )

        if not json_files:
            print(f"No JSON files found in {INPUT_DIR}")
            sys.exit(1)

        print(f"Found {len(json_files)} JSON files\n")
        total = 0
        for jf in json_files:
            print(f"Converting: {os.path.basename(jf)}")
            created = convert_json(jf)
            total += len(created)
            print()

        print(f"Done. {total} xlsx files created in {OUTPUT_DIR}/")

    elif cmd == 'list-templates':
        if not os.path.isdir(TEMPLATES_DIR):
            print("No templates directory. Run 'python3 converter.py setup' first.")
            sys.exit(1)

        templates = sorted(f for f in os.listdir(TEMPLATES_DIR) if f.endswith('.xlsx'))
        if not templates:
            print("No templates found.")
        else:
            print("Available templates:")
            for t in templates:
                product_id = t.replace('template_', '').replace('.xlsx', '')
                try:
                    wb = openpyxl.load_workbook(os.path.join(TEMPLATES_DIR, t), read_only=True)
                    ws = wb.active
                    dept = ws.cell(row=8, column=4).value or '?'
                    desc = ws.cell(row=6, column=4).value or '?'
                    wb.close()
                    print(f"  Product {product_id:>3}: {dept} / {desc}")
                except Exception:
                    print(f"  Product {product_id:>3}: {t}")

    else:
        print(f"Unknown command: {cmd}")
        print(usage)
        sys.exit(1)


if __name__ == '__main__':
    main()
