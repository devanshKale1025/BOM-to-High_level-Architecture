"""
parser.py — Robust BOM Parser  (v2)
Handles every edge case: encoding issues, empty files, no headers,
multi-sheet Excel, section-header rows, merged cells, etc.
"""

import pandas as pd
import re
import os

AREA_ALIASES  = ['area', 'room', 'location', 'zone', 'section', 'place']
DESC_ALIASES  = ['description', 'desc', 'item description', 'material description',
                 'component', 'equipment', 'name', 'item name', 'particulars',
                 'material', 'items', 'item']
QTY_ALIASES   = ['qty', 'quantity', 'count', 'no.', 'nos', 'number',
                 'no of units', 'nos.', 'qty.', 'units']
CAT_ALIASES   = ['category', 'cat', 'type', 'class', 'group', 'tag', 'cat.']
PART_ALIASES  = ['part number', 'part no', 'part no.', 'part#', 'pn',
                 'item code', 'material code', 'model', 'part', 'catalog',
                 'cat no', 'article', 'part_number', 'partno', 'part_no']
SRNO_ALIASES  = ['sr', 'sr.', 'sr. no', 'sr no', 'sr.no', 's.no', 'serial',
                 'no', '#', 'item no', 'sl no', 'sr_no', 'sno', 'sl.no',
                 'item#', 'sr. no.']

SECTION_KEYWORDS = ['pdc room', 'operator room', 'control room', 'field',
                    'junction', 'server room', 'mcc room', 'marshalling',
                    'total', 'subtotal', 'note', 'remarks']

CSV_ENCODINGS = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']


def _norm(text):
    return str(text).strip().lower()


def _find_col(df, aliases):
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for alias in aliases:
        if alias in cols_lower:
            return cols_lower[alias]
    for alias in aliases:
        for col_low, col_orig in cols_lower.items():
            if alias in col_low:
                return col_orig
    for alias in aliases:
        for col_low, col_orig in cols_lower.items():
            if col_low in alias and len(col_low) > 2:
                return col_orig
    return None


def _detect_header_row(raw_df):
    all_aliases = set(DESC_ALIASES + QTY_ALIASES + SRNO_ALIASES +
                      PART_ALIASES + CAT_ALIASES + AREA_ALIASES)
    best_row, best_score = 0, 0
    for i in range(min(15, len(raw_df))):
        row_vals = [_norm(v) for v in raw_df.iloc[i].tolist()
                    if str(v).strip() not in ('', 'nan', 'NaN', 'None')]
        score = sum(1 for v in row_vals if v in all_aliases)
        if score > best_score:
            best_score = score
            best_row   = i
    return best_row


def _is_section_header(desc):
    if not desc or _norm(desc) in ('', 'nan', 'none'):
        return False
    d = _norm(desc)
    for kw in SECTION_KEYWORDS:
        if d == kw or d.startswith(kw + ' ') or d.endswith(' ' + kw):
            return True
    return False


def _canonical_area(text):
    t = str(text).strip().upper()
    if 'OPERATOR' in t:
        return 'OPERATOR ROOM'
    if any(k in t for k in ('PDC', 'PROCESS', 'CONTROL ROOM', 'SERVER ROOM')):
        return 'PDC ROOM'
    if t in ('', 'NAN', 'NONE'):
        return ''
    return t


def _read_raw(filepath, ext):
    """Read raw file with no header. Returns dataframe or raises clear error."""
    if ext in ('.xlsx', '.xls'):
        try:
            xl = pd.ExcelFile(filepath)
        except Exception as e:
            raise ValueError(f"Cannot open Excel file: {e}. Make sure it is not corrupted.")

        best_df, best_rows = None, 0
        for sheet in xl.sheet_names:
            try:
                df = xl.parse(sheet, header=None, dtype=str, keep_default_na=False)
                df = df.dropna(how='all').reset_index(drop=True)
                if len(df) > best_rows:
                    best_df   = df
                    best_rows = len(df)
            except Exception:
                continue

        if best_df is None or best_df.empty:
            raise ValueError("Excel file is empty or has no readable data in any sheet.")
        return best_df, xl.sheet_names[0]

    else:  # CSV
        for enc in CSV_ENCODINGS:
            try:
                df = pd.read_csv(filepath, header=None, dtype=str, encoding=enc,
                                 keep_default_na=False, on_bad_lines="skip", engine="python")
                if not df.empty and len(df.columns) > 0:
                    return df, None
            except (UnicodeDecodeError, pd.errors.EmptyDataError):
                continue
            except Exception:
                continue
        raise ValueError(
            "Could not read CSV file. Try saving as UTF-8 from Excel "
            "(File → Save As → CSV UTF-8) and re-upload."
        )


def _read_with_header(filepath, ext, hdr_row, sheet_name):
    """Re-read file using the detected header row."""
    if ext in ('.xlsx', '.xls'):
        return pd.read_excel(filepath, header=hdr_row, dtype=str,
                             keep_default_na=False, sheet_name=sheet_name)
    else:
        for enc in CSV_ENCODINGS:
            try:
                df = pd.read_csv(filepath, header=hdr_row, dtype=str,
                                 encoding=enc, keep_default_na=False,
                                 on_bad_lines='skip', engine='python')
                if not df.empty:
                    return df
            except Exception:
                continue
        raise ValueError("Could not re-read CSV with detected header row.")


def parse_bom(filepath: str) -> pd.DataFrame:
    """
    Parse any BOM file into a normalized dataframe.
    Returns DataFrame with: sr_no, area, description, qty, category, part_number
    """
    # ── Basic file checks ─────────────────────────────────────────────────────
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    size = os.path.getsize(filepath)
    if size == 0:
        raise ValueError(
            "The uploaded file is empty (0 bytes). "
            "Please check your file and re-upload."
        )

    ext = os.path.splitext(filepath)[1].lower()
    if ext not in ('.xlsx', '.xls', '.csv'):
        raise ValueError(
            f"Unsupported file type '{ext}'. "
            "Please upload .xlsx, .xls, or .csv"
        )

    # ── Read raw ──────────────────────────────────────────────────────────────
    raw, sheet_name = _read_raw(filepath, ext)

    if raw.empty or len(raw.columns) == 0:
        raise ValueError(
            "File has no columns. Make sure your BOM has at least a "
            "Description column and is not completely empty."
        )

    # ── Detect header ─────────────────────────────────────────────────────────
    hdr_row = _detect_header_row(raw)

    # ── Re-read with header ───────────────────────────────────────────────────
    df = _read_with_header(filepath, ext, hdr_row, sheet_name)

    # Clean column names — remove unnamed/empty
    df.columns = [str(c).strip() for c in df.columns]
    df = df[[c for c in df.columns
             if c and c.lower() not in ('nan', 'none')
             and not c.lower().startswith('unnamed')]]

    if df.empty or len(df.columns) == 0:
        raise ValueError(
            "No usable columns found. Make sure your BOM has column headers "
            "like 'Description', 'Qty', 'Part Number', etc."
        )

    # ── Map to canonical columns ──────────────────────────────────────────────
    col_map = {}
    for canonical, aliases in [
        ('sr_no',       SRNO_ALIASES),
        ('description', DESC_ALIASES),
        ('qty',         QTY_ALIASES),
        ('category',    CAT_ALIASES),
        ('part_number', PART_ALIASES),
        ('area',        AREA_ALIASES),
    ]:
        found = _find_col(df, aliases)
        if found and found not in col_map:
            col_map[found] = canonical

    df = df.rename(columns=col_map)

    # Last resort: if description still missing, use longest text column
    if 'description' not in df.columns:
        remaining = [c for c in df.columns if c not in col_map.values()]
        if remaining:
            avg_len = {c: df[c].fillna('').apply(lambda x: len(str(x))).mean()
                       for c in remaining}
            best = max(avg_len, key=avg_len.get)
            df = df.rename(columns={best: 'description'})

    # Ensure all required columns exist
    for col in ('sr_no', 'description', 'qty', 'category', 'part_number', 'area'):
        if col not in df.columns:
            df[col] = ''

    # ── Propagate area + filter rows ─────────────────────────────────────────
    current_area = 'PDC ROOM'
    rows = []

    for _, row in df.iterrows():
        desc = str(row.get('description', '')).strip()
        sr   = str(row.get('sr_no',      '')).strip()

        # Skip completely blank rows
        vals = [str(v).strip() for v in row.values]
        if all(v in ('', 'nan', 'None', 'NaN') for v in vals):
            continue

        # Section header → update area, skip row
        if _is_section_header(desc):
            current_area = _canonical_area(desc) or current_area
            continue

        # Skip rows with no meaningful description
        if desc in ('', 'nan', 'None', 'NaN'):
            continue

        # Propagate / canonicalize area
        row_area = str(row.get('area', '')).strip()
        row = row.copy()
        if row_area in ('', 'nan', 'None', 'NaN'):
            row['area'] = current_area
        else:
            canonical    = _canonical_area(row_area)
            row['area']  = canonical if canonical else row_area.upper()
            current_area = row['area']

        rows.append(row)

    if not rows:
        raise ValueError(
            "No data rows found after parsing. "
            "Make sure your file has a 'Description' column with item names."
        )

    result = pd.DataFrame(rows)
    result = result[['sr_no', 'area', 'description', 'qty', 'category', 'part_number']]
    result = result.reset_index(drop=True)

    # ── Clean values ──────────────────────────────────────────────────────────
    def clean_qty(v):
        try:
            return max(1, int(float(str(v).strip().replace(',', ''))))
        except Exception:
            return 1

    def clean_str(v):
        s = str(v).strip()
        return '' if s.lower() in ('nan', 'none', 'nat') else s

    result['qty']         = result['qty'].apply(clean_qty)
    result['description'] = result['description'].apply(clean_str)
    result['part_number'] = result['part_number'].apply(clean_str)
    result['category']    = result['category'].apply(clean_str)

    result = result[result['description'] != ''].reset_index(drop=True)

    if result.empty:
        raise ValueError(
            "All rows were filtered out as empty or headers. "
            "Make sure your BOM file has actual item descriptions."
        )

    return result


if __name__ == '__main__':
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else 'bom_full.csv'
    try:
        df = parse_bom(path)
        print(df.to_string())
        print(f"\n✅  {len(df)} rows parsed successfully.")
    except Exception as e:
        print(f"❌  Error: {e}")