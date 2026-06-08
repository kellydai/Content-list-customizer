import streamlit as st
import openpyxl
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import math
from rapidfuzz import fuzz, process
import fitz  # PyMuPDF, used by the Label colorizer

# Fuzzy-match thresholds (0–100)
FUZZY_HIGH   = 88   # >= this: HIGH confidence (auto-replace)
FUZZY_MEDIUM = 70   # >= this but < HIGH: MEDIUM (user must confirm)
                    # < MEDIUM: LOW/NONE (no match)

st.set_page_config(page_title="ICRC Content List Customizer", page_icon="🏥", layout="wide")

GREEN_HEX  = "00A176"
green_fill = PatternFill(fill_type="solid", fgColor=GREEN_HEX)

ID_KEYWORDS   = ["item code", "customer id", "customer_id", "code", "id", "reference", "ref", "article no"]
DESC_KEYWORDS = ["description", "item description", "designation", "name", "libelle", "libellé"]

# ── helpers ────────────────────────────────────────────────────────────────

def normalize(s):
    if not s: return ""
    s = str(s).lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("×", "x")
    s = re.sub(r"(\d)\s*x\s*(\d)", r"\1x\2", s)
    return s.rstrip(".,;")

def col_score(col, keywords):
    c = str(col).lower().strip()
    return max((2 if kw == c else 1 if kw in c else 0) for kw in keywords)

def auto_detect_cols(columns):
    id_sc   = {c: col_score(c, ID_KEYWORDS)   for c in columns}
    desc_sc = {c: col_score(c, DESC_KEYWORDS) for c in columns}
    best_id   = max(id_sc,   key=id_sc.get)
    best_desc = max(desc_sc, key=desc_sc.get)
    if best_id == best_desc:
        others  = [c for c in columns if c != best_desc]
        best_id = others[0] if others else best_id
    confident = id_sc[best_id] > 0 and desc_sc[best_desc] > 0
    return best_id, best_desc, confident

# ── reference list loading (Excel OR PDF) ─────────────────────────────────

def load_ref_df_excel(file_bytes):
    wb   = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws   = wb.active
    rows = list(ws.iter_rows(values_only=True))
    all_entries, header = [], None
    for row in rows:
        str_vals = [v for v in row if isinstance(v, str) and v.strip()]
        lower    = " ".join(v.lower() for v in str_vals)
        if len(str_vals) >= 2 and any(kw in lower for kw in ID_KEYWORDS + DESC_KEYWORDS):
            header = [str(v).strip() if v else f"col_{i}" for i, v in enumerate(row)]
            continue
        if header and any(v is not None for v in row):
            entry = {header[i]: row[i] for i in range(min(len(header), len(row)))}
            all_entries.append(entry)
    if not all_entries:
        return pd.DataFrame()
    return pd.DataFrame(all_entries).dropna(axis=1, how="all")

def load_ref_df_pdf(file_bytes):
    """Extract reference table(s) from a PDF reference list."""
    all_rows, header = [], None
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                for row in table:
                    clean = [str(c).strip() if c else "" for c in row]
                    if not any(clean): continue
                    lower = " ".join(c.lower() for c in clean)
                    if any(kw in lower for kw in ID_KEYWORDS + DESC_KEYWORDS):
                        header = clean
                        continue
                    if not header:
                        continue

                    entry = {header[i]: clean[i] for i in range(min(len(header), len(clean)))}

                    # Merge wrapped-text continuation rows (only one cell has content)
                    non_empty_cols = [k for k, v in entry.items() if v.strip()]
                    desc_key = next((k for k in header if any(
                        kw in k.lower() for kw in DESC_KEYWORDS)), None)
                    is_continuation = (
                        all_rows
                        and desc_key
                        and non_empty_cols == [desc_key]
                    )
                    if is_continuation:
                        all_rows[-1][desc_key] = (
                            all_rows[-1].get(desc_key, "") + "\n" + entry[desc_key]
                        ).strip()
                    else:
                        all_rows.append(entry)

    if not all_rows:
        return pd.DataFrame()
    return pd.DataFrame(all_rows).replace("", pd.NA).dropna(axis=1, how="all")

def load_ref_df(file_bytes, filename):
    if filename.lower().endswith(".pdf"):
        return load_ref_df_pdf(file_bytes)
    return load_ref_df_excel(file_bytes)

def build_lookup(ref_df, id_col, desc_col):
    """
    Returns a dict:
      {
        "exact":   {normalized_desc: [customer_id, ...]},
        "entries": [(normalized_desc, original_desc, customer_id), ...],
      }
    `entries` is used for fuzzy matching when exact lookup fails.
    """
    exact = {}
    entries = []
    for _, row in ref_df.iterrows():
        cid  = row.get(id_col)
        desc = row.get(desc_col)
        if pd.isna(cid)  if isinstance(cid,  float) else cid  is None: continue
        if pd.isna(desc) if isinstance(desc, float) else desc is None: continue
        nd = normalize(str(desc))
        if nd:
            exact.setdefault(nd, []).append(str(cid))
            entries.append((nd, str(desc), str(cid)))
    return {"exact": exact, "entries": entries}

# ── ERP content list loading (Excel OR PDF) ───────────────────────────────

def find_erp_table_excel(ws):
    """Return (header_row_1based, desc_col_1based, item_code_col_1based_or_None)."""
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        vals = [str(v).lower().strip() if v else "" for v in row]
        if any("description" in v for v in vals):
            desc_col      = next(j + 1 for j, v in enumerate(vals) if "description" in v)
            item_code_col = next((j + 1 for j, v in enumerate(vals)
                                  if "item code" in v or v == "code"), None)
            return i, desc_col, item_code_col
    return None, None, None

def load_erp_items_excel(erp_bytes):
    """Returns list of dicts with keys: row, description. row is 1-based sheet row."""
    wb = openpyxl.load_workbook(BytesIO(erp_bytes), data_only=False)
    ws = wb.active
    header_row, desc_col, _ = find_erp_table_excel(ws)
    if not header_row:
        return []
    items = []
    for r in range(header_row + 1, ws.max_row + 1):
        val = ws.cell(r, desc_col).value
        if not val: break
        items.append({"row": r, "description": str(val)})
    return items

def _extract_tables_robust(page):
    """
    Try multiple pdfplumber extraction strategies, return the first that yields tables.
    Using 'lines' respects actual PDF borders (best for bordered tables).
    Falling back to 'text' handles tables with no visible borders.
    """
    for strategy in (
        {"vertical_strategy": "lines",  "horizontal_strategy": "lines"},
        {"vertical_strategy": "lines",  "horizontal_strategy": "text"},
        {"vertical_strategy": "text",   "horizontal_strategy": "text"},
    ):
        tables = page.extract_tables(strategy)
        if tables:
            return tables
    return []

# Columns that identify a real data row (a row must have at least one of these)
_DATA_COL_KEYWORDS = ["expiry", "batch", "lot", "uom", "unit", "quantity",
                      "qty", "manufacturer", "country", "mfr"]

def _has_data_columns(entry, desc_key):
    """Return True if ANY column other than description has a non-empty value."""
    for k, v in entry.items():
        if k == desc_key:
            continue
        if str(v).strip():
            return True
    return False

def load_erp_items_pdf(file_bytes):
    """
    Extract ERP content list rows from a PDF — preserving batch structure.
    A new batch starts whenever a page contains a 'CONTENT LIST' heading.
    Continuation pages (no heading) attach to the current batch.

    Returns a list of batches:
      [{ "meta": {...}, "columns": [...], "items": [pdf_row_dict, ...] }, ...]
    """
    batches = []
    current = None  # active batch dict

    def new_batch():
        return {"meta": {}, "columns": [], "items": []}

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            # ── Detect a new batch by 'CONTENT LIST' heading on this page
            if "CONTENT LIST" in text.upper():
                if current and current["items"]:
                    batches.append(current)
                current = new_batch()

            if current is None:
                current = new_batch()

            # ── title-block metadata for the active batch
            for line in text.splitlines():
                l = line.strip()
                if re.match(r"^KM\w+", l):
                    current["meta"].setdefault("kit_code", l.split()[0])
                if "assembly batch" in l.lower():
                    current["meta"].setdefault("batch", l)
                if "po number" in l.lower():
                    current["meta"].setdefault("po", l)
                if "content list" in l.lower():
                    current["meta"].setdefault("title", "CONTENT LIST")
                if re.match(r"^SET,", l, re.I) and "kit_name" not in current["meta"]:
                    current["meta"]["kit_name"] = l

            # ── table rows
            for table in _extract_tables_robust(page):
                for row in table:
                    clean = [str(c).strip() if c is not None else "" for c in row]
                    if not any(clean):
                        continue

                    # Header row detection — establishes columns for this batch
                    if any("description" in c.lower() for c in clean):
                        current["columns"] = clean
                        continue

                    if not current["columns"]:
                        continue

                    padded = clean + [""] * max(0, len(current["columns"]) - len(clean))
                    entry = {current["columns"][i]: padded[i] for i in range(len(current["columns"]))}

                    desc_key = next(
                        (k for k in current["columns"] if "description" in k.lower()), None
                    )
                    desc_val = entry.get(desc_key, "").strip() if desc_key else ""

                    if not desc_val:
                        continue

                    is_continuation = (
                        bool(current["items"])
                        and desc_key is not None
                        and not _has_data_columns(entry, desc_key)
                    )

                    if is_continuation:
                        prev_desc = current["items"][-1].get(desc_key, "")
                        current["items"][-1][desc_key] = (prev_desc + "\n" + desc_val).strip()
                    else:
                        current["items"].append(entry)

    if current and current["items"]:
        batches.append(current)

    return batches

def load_erp_items(file_bytes, filename):
    """
    Returns (items, source_format, pdf_meta, pdf_cols, pdf_batches).
    items: flat list of dicts with 'row' (globally unique 1-based id),
           'description', '_pdf_row', and 'batch_idx' (PDF only).
    For Excel: pdf_batches is None.
    For PDF: pdf_batches is the list returned by load_erp_items_pdf, with each
             batch's items annotated with the same 'row' id as the flat list,
             so generate_output_from_pdf can look up the customer ID per item.
    """
    if filename.lower().endswith(".pdf"):
        batches = load_erp_items_pdf(file_bytes)
        # First batch's meta+cols serve as fallback for legacy callers
        first_meta = batches[0]["meta"] if batches else {}
        first_cols = batches[0]["columns"] if batches else []

        items = []
        global_row = 0
        for bi, batch in enumerate(batches):
            desc_key = next(
                (k for k in batch["columns"] if "description" in k.lower()),
                batch["columns"][0] if batch["columns"] else "description",
            )
            for pdf_row in batch["items"]:
                global_row += 1
                pdf_row["_row_id"] = global_row  # back-reference for output
                items.append({
                    "row": global_row,
                    "description": pdf_row.get(desc_key, ""),
                    "_pdf_row": pdf_row,
                    "batch_idx": bi,
                })
        return items, "pdf", first_meta, first_cols, batches
    else:
        items = load_erp_items_excel(file_bytes)
        return items, "excel", {}, [], None

# ── matching ───────────────────────────────────────────────────────────────

def match_item(desc_bilingual, lookup):
    """
    Match strategy:
      1. EXACT normalized match per description fragment (bilingual = split on \n).
      2. FUZZY match (rapidfuzz token_set_ratio) against all reference entries
         when exact match fails. token_set_ratio handles extra/missing words well
         (e.g. content list has 'with paper label and Red cap', ref doesn't).

    Returns (best_id, confidence_label, candidate_list).
      confidence_label ∈ {"HIGH", "MEDIUM", "LOW/NONE"}
    """
    if not desc_bilingual:
        return None, "LOW/NONE", []

    exact   = lookup["exact"]
    entries = lookup["entries"]

    # ── 1. exact normalized match ────────────────────────────────────────
    candidates = {}
    for frag in str(desc_bilingual).split("\n"):
        for cid in exact.get(normalize(frag.strip()), []):
            candidates[cid] = True
    unique = list(candidates.keys())
    if unique:
        if len(unique) == 1:
            return unique[0], "HIGH", unique
        return unique[0], "MEDIUM", unique

    # ── 2. fuzzy fallback ────────────────────────────────────────────────
    if not entries:
        return None, "LOW/NONE", []

    norm_choices = [e[0] for e in entries]
    best_score = 0
    best_idx   = None
    for frag in str(desc_bilingual).split("\n"):
        n = normalize(frag.strip())
        if not n:
            continue
        result = process.extractOne(n, norm_choices, scorer=fuzz.token_set_ratio)
        if result and result[1] > best_score:
            best_score = result[1]
            best_idx   = result[2]

    if best_idx is None or best_score < FUZZY_MEDIUM:
        return None, "LOW/NONE", []

    best_cid = entries[best_idx][2]
    label    = "HIGH" if best_score >= FUZZY_HIGH else "MEDIUM"
    # surface the matched ID as the single candidate so the review UI can show it
    return best_cid, label, [best_cid]

# ── output generation ──────────────────────────────────────────────────────

def apply_text_format(ws):
    """Apply wrap-text + black font to every used cell. No fill (white background)."""
    wrap_align = Alignment(wrap_text=True, vertical="center")
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            f = cell.font
            cell.font = Font(name=f.name, size=f.size, bold=f.bold,
                             italic=f.italic, color="000000")
            cell.alignment = wrap_align


# Kept for reference; the content list output no longer fills cells green.
def apply_green(ws):
    apply_text_format(ws)
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).fill = green_fill


# Rough character widths used for column auto-sizing & row height estimation.
_COL_WIDTHS = {
    "item code":   18,
    "description": 55,
    "quantity":     9,
    "uom":         12,
    "vendor":      22,
    "batch":       16,
    "expiry":      12,
    "manufacturer":18,
    "country":     14,
}
_DEFAULT_COL_WIDTH = 14
_MAX_COL_WIDTH     = 60


def _column_width_for(header):
    h = str(header or "").lower()
    for kw, w in _COL_WIDTHS.items():
        if kw in h:
            return w
    return _DEFAULT_COL_WIDTH


def _apply_print_layout(ws, header_row):
    """
    Make the sheet print cleanly:
      • column widths sized to header type
      • row heights tall enough that wrapped text isn't clipped
      • print area restricted to the used range
      • fit-to-width landscape, narrow margins
    """
    n_rows = ws.max_row
    n_cols = ws.max_column
    if n_rows == 0 or n_cols == 0:
        return

    # 1. Column widths
    col_widths = {}
    for ci in range(1, n_cols + 1):
        header_val = ws.cell(header_row, ci).value if header_row else None
        width = _column_width_for(header_val)
        col_widths[ci] = width
        ws.column_dimensions[get_column_letter(ci)].width = width

    # 2. Row heights — based on wrapped-line count per cell
    for r in range(1, n_rows + 1):
        max_lines = 1
        for ci in range(1, n_cols + 1):
            val = ws.cell(r, ci).value
            if val is None or val == "":
                continue
            col_w = col_widths.get(ci, _DEFAULT_COL_WIDTH)
            # Each \n forces a new line; long lines wrap based on column width
            for raw_line in str(val).split("\n"):
                lines = max(1, math.ceil(len(raw_line) / max(col_w - 2, 1)))
                max_lines = max(max_lines, lines)
            max_lines = max(max_lines, str(val).count("\n") + 1)
        # ~15pt per text line
        ws.row_dimensions[r].height = max(18, max_lines * 15)

    # 3. Print area = exactly the used range
    last_col_letter = get_column_letter(n_cols)
    ws.print_area = f"A1:{last_col_letter}{n_rows}"

    # 4. Page setup: landscape, fit-to-width, narrow margins, centered
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    # Repeat header row on every printed page
    if header_row:
        ws.print_title_rows = f"{header_row}:{header_row}"

def generate_output_excel(erp_bytes, confirmed_matches):
    wb = openpyxl.load_workbook(BytesIO(erp_bytes), data_only=False)
    ws = wb.active
    header_row, _, item_code_col = find_erp_table_excel(ws)
    ic = item_code_col or 1
    for row_num, cid in confirmed_matches.items():
        ws.cell(row_num, ic).value = cid or None
    apply_text_format(ws)
    _apply_print_layout(ws, header_row or 1)
    out = BytesIO(); wb.save(out)
    return out.getvalue()

def _safe_sheet_name(name, existing):
    """Excel sheet name: max 31 chars, no \\/*?:[] and must be unique."""
    name = re.sub(r"[\\/*?:\[\]]", "-", str(name)).strip() or "Content list"
    name = name[:31]
    base, n = name, 2
    while name in existing:
        suffix = f" ({n})"
        name = base[: 31 - len(suffix)] + suffix
        n += 1
    return name


def _write_batch_sheet(ws, batch, confirmed_matches):
    """Write a single batch's title block + data table into the given sheet."""
    meta = batch.get("meta", {})
    cols = batch.get("columns", []) or []
    items = batch.get("items", [])

    row_ptr = 1

    def write_title_cell(r, c, val, bold=False):
        cell = ws.cell(r, c, val)
        if bold:
            cell.font = Font(bold=True)

    write_title_cell(row_ptr, 1, meta.get("title", "CONTENT LIST"), bold=True); row_ptr += 1
    if "kit_code" in meta:
        write_title_cell(row_ptr, 1, meta["kit_code"], bold=True); row_ptr += 1
    if "kit_name" in meta:
        write_title_cell(row_ptr, 1, meta["kit_name"]); row_ptr += 2
    if "batch" in meta:
        write_title_cell(row_ptr, 1, meta["batch"]); row_ptr += 1
    if "po" in meta:
        write_title_cell(row_ptr, 1, meta["po"]); row_ptr += 2

    # Header row — prepend "Item code" if not present
    has_ic = any("item code" in k.lower() or k.lower() == "code" for k in cols)
    out_cols = cols if has_ic else ["Item code"] + cols

    for ci, col in enumerate(out_cols, 1):
        cell = ws.cell(row_ptr, ci, col)
        cell.font = Font(bold=True)
    header_row = row_ptr
    row_ptr += 1

    ic_col_idx = next(
        (i + 1 for i, k in enumerate(out_cols) if "item code" in k.lower() or k.lower() == "code"),
        1,
    )
    for pdf_row in items:
        row_id = pdf_row.get("_row_id")
        cid = confirmed_matches.get(row_id)
        for ci, col in enumerate(out_cols, 1):
            if ci == ic_col_idx:
                # Item-code column: write the confirmed customer ID
                ws.cell(row_ptr, ci, cid or "")
            else:
                ws.cell(row_ptr, ci, pdf_row.get(col, ""))
        row_ptr += 1

    apply_text_format(ws)
    _apply_print_layout(ws, header_row)


def _sheet_name_from_batch(batch, fallback):
    """Derive a friendly sheet name from batch metadata."""
    meta = batch.get("meta", {})
    # Prefer the part after 'Assembly Batch:' if present
    batch_str = meta.get("batch", "")
    if batch_str:
        after = batch_str.split(":", 1)[-1].strip()
        if after:
            return after
    return meta.get("kit_code") or fallback


# ── Label colorizer (separate tool) ───────────────────────────────────────

def colorize_pdf_pages_green(pdf_bytes, green_hex=GREEN_HEX):
    """
    Paint every page of `pdf_bytes` with a green background, leaving the
    existing text/lines on top untouched. Returns the modified PDF as bytes.

    Implementation: insert a full-page filled rectangle BELOW existing
    content (overlay=False) using PyMuPDF.
    """
    r = int(green_hex[0:2], 16) / 255
    g = int(green_hex[2:4], 16) / 255
    b = int(green_hex[4:6], 16) / 255

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for page in doc:
            page.draw_rect(page.rect, color=None, fill=(r, g, b), overlay=False)
        out = BytesIO()
        doc.save(out)
    finally:
        doc.close()
    return out.getvalue()


def generate_output_from_pdf(erp_bytes, confirmed_matches, pdf_batches):
    """Build a fresh Excel from PDF-extracted batches — one sheet per batch."""
    wb = openpyxl.Workbook()
    # Remove the auto-created default sheet so we control sheet order
    default_ws = wb.active
    wb.remove(default_ws)

    used = set()
    for bi, batch in enumerate(pdf_batches, 1):
        sheet_name = _safe_sheet_name(
            _sheet_name_from_batch(batch, f"Content list {bi}"), used
        )
        used.add(sheet_name)
        ws = wb.create_sheet(sheet_name)
        _write_batch_sheet(ws, batch, confirmed_matches)

    if not pdf_batches:
        wb.create_sheet("Content list")  # empty fallback

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# ══════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════

DEFAULTS = dict(step=1, erp_bytes=None, erp_filename="", ref_bytes=None, ref_filename="",
                ref_df=None, id_col=None, desc_col=None, lookup=None,
                erp_items=None, erp_format="excel", pdf_meta={}, pdf_cols=[],
                pdf_batches=None,
                matches=None, output_bytes=None)
for k, v in DEFAULTS.items():
    if k not in st.session_state: st.session_state[k] = v

def reset():
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# SIDEBAR — choose tool
# ══════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.header("🛠️  Tools")
    tool = st.radio(
        "Pick a tool",
        options=["Content List Customizer", "Label Colorizer (PDF → green)"],
        label_visibility="collapsed",
    )
    st.caption(
        "• **Content List Customizer** — replace Sofia item codes with "
        "customer codes and export a clean Excel.\n\n"
        "• **Label Colorizer** — apply Pantone 340U green background to "
        "every page of a label PDF."
    )

# ══════════════════════════════════════════════════════════════════════════
# TOOL 2 — Label Colorizer  (early return so the rest of the file is skipped)
# ══════════════════════════════════════════════════════════════════════════
if tool.startswith("Label"):
    st.title("🟩 Label Colorizer")
    st.caption(
        "Upload a label PDF; every page background is filled with "
        f"Pantone 340U (#{GREEN_HEX}). Text and lines stay readable on top."
    )

    label_file = st.file_uploader("Upload label PDF", type=["pdf"])

    if label_file:
        with st.spinner("Applying green background…"):
            try:
                green_bytes = colorize_pdf_pages_green(label_file.read())
            except Exception as e:
                st.error(f"Failed to colorize PDF: {e}")
                st.stop()

        st.success("Done!")
        stem = re.sub(r"\.pdf$", "", label_file.name, flags=re.I)
        st.download_button(
            label="⬇️  Download green-background PDF",
            data=green_bytes,
            file_name=f"{stem}_green.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
    else:
        st.info("Pick a PDF above to colorize.")

    st.stop()  # don't render the Content List flow below


# ══════════════════════════════════════════════════════════════════════════
# HEADER + PROGRESS  (Content List Customizer)
# ══════════════════════════════════════════════════════════════════════════

st.title("🏥 ICRC Content List Customizer")
st.caption("Upload an ERP content list (Excel or PDF) and a customer reference list (Excel or PDF) "
           "to produce a clean Excel output.")

STEP_LABELS = ["1 · Upload", "2 · Confirm Columns", "3 · Review Matches", "4 · Download"]
pcols = st.columns(len(STEP_LABELS))
for i, (col, label) in enumerate(zip(pcols, STEP_LABELS), 1):
    with col:
        if   i < st.session_state.step:  st.success(f"✓ {label}")
        elif i == st.session_state.step: st.info(f"▶ {label}")
        else: st.write(f"○ {label}")

st.divider()

# ══════════════════════════════════════════════════════════════════════════
# STEP 1 — Upload
# ══════════════════════════════════════════════════════════════════════════
if st.session_state.step == 1:
    st.subheader("Upload both files")
    st.write("Both **Excel (.xlsx)** and **PDF (.pdf)** formats are accepted.")
    c1, c2 = st.columns(2)
    with c1:
        erp_file = st.file_uploader("📄 ERP Content List", type=["xlsx", "pdf"])
    with c2:
        ref_file = st.file_uploader("📋 Customer Reference List", type=["xlsx", "pdf"])

    if erp_file and ref_file:
        if st.button("Analyse →", type="primary"):
            st.session_state.erp_bytes    = erp_file.read()
            st.session_state.erp_filename = erp_file.name
            st.session_state.ref_bytes    = ref_file.read()
            st.session_state.ref_filename = ref_file.name

            ref_df = load_ref_df(st.session_state.ref_bytes, ref_file.name)
            if ref_df.empty:
                st.error("Could not parse the reference list — no recognisable header row found.")
                st.stop()

            st.session_state.ref_df = ref_df
            bid, bdesc, _           = auto_detect_cols(list(ref_df.columns))
            st.session_state.id_col   = bid
            st.session_state.desc_col = bdesc
            st.session_state.step     = 2
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# STEP 2 — Confirm columns
# ══════════════════════════════════════════════════════════════════════════
elif st.session_state.step == 2:
    st.subheader("Confirm reference list columns")

    ref_df    = st.session_state.ref_df
    cols_list = list(ref_df.columns)

    ref_fmt = "PDF" if st.session_state.ref_filename.lower().endswith(".pdf") else "Excel"
    erp_fmt = "PDF" if st.session_state.erp_filename.lower().endswith(".pdf") else "Excel"
    st.write(f"ERP file: **{st.session_state.erp_filename}** ({erp_fmt}) · "
             f"Reference file: **{st.session_state.ref_filename}** ({ref_fmt})")

    st.write("**Reference list preview** (first 10 rows):")
    st.dataframe(ref_df.head(10), use_container_width=True)
    st.write("---")
    st.write("Confirm which column is the **Customer ID** and which is the **Item Description**.")

    c1, c2 = st.columns(2)
    with c1:
        id_col = st.selectbox(
            "🔑 Customer ID column",
            options=cols_list,
            index=cols_list.index(st.session_state.id_col) if st.session_state.id_col in cols_list else 0,
            help="Column A is the Customer ID — confirm or change"
        )
    with c2:
        desc_col = st.selectbox(
            "📝 Item Description column",
            options=cols_list,
            index=cols_list.index(st.session_state.desc_col) if st.session_state.desc_col in cols_list else min(1, len(cols_list)-1),
            help="Column B is the Item Description — confirm or change"
        )

    if id_col == desc_col:
        st.warning("⚠️ Customer ID and Description must be different columns.")
    else:
        st.write("**Sample of selected columns:**")
        st.dataframe(ref_df[[id_col, desc_col]].dropna(how="all").head(6), use_container_width=True)

        bc, cc = st.columns([1, 5])
        with bc:
            if st.button("← Back"): st.session_state.step = 1; st.rerun()
        with cc:
            if st.button("✅ Confirm & run matching →", type="primary"):
                st.session_state.id_col   = id_col
                st.session_state.desc_col = desc_col
                st.session_state.lookup   = build_lookup(ref_df, id_col, desc_col)

                with st.spinner("Parsing ERP file and matching…"):
                    items, fmt, pdf_meta, pdf_cols, pdf_batches = load_erp_items(
                        st.session_state.erp_bytes, st.session_state.erp_filename)

                if not items:
                    st.error("Could not find a data table in the ERP content list.")
                    st.stop()

                lookup = st.session_state.lookup
                for item in items:
                    cid, conf, cands = match_item(item["description"], lookup)
                    item["matched_id"]  = cid
                    item["confidence"]  = conf
                    item["candidates"]  = cands

                st.session_state.erp_items   = items
                st.session_state.erp_format  = fmt
                st.session_state.pdf_meta    = pdf_meta
                st.session_state.pdf_cols    = pdf_cols
                st.session_state.pdf_batches = pdf_batches
                st.session_state.matches     = {i["row"]: i["matched_id"] for i in items}
                st.session_state.step       = 3
                st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# STEP 3 — Review matches
# ══════════════════════════════════════════════════════════════════════════
elif st.session_state.step == 3:
    st.subheader("Review matches")
    items   = st.session_state.erp_items
    matches = dict(st.session_state.matches)

    n_high   = sum(1 for i in items if i["confidence"] == "HIGH")
    n_medium = sum(1 for i in items if i["confidence"] == "MEDIUM")
    n_low    = sum(1 for i in items if i["confidence"] == "LOW/NONE")

    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("🟢 High confidence", n_high)
    mc2.metric("🟡 Needs review",    n_medium)
    mc3.metric("🔴 No match",        n_low)

    if n_medium == 0 and n_low == 0:
        st.success("All rows matched cleanly — no review needed.")
    else:
        st.info("Expand 🟡/🔴 rows below to confirm or correct the customer ID.")

    for item in items:
        conf  = item["confidence"]
        icon  = {"HIGH": "🟢", "MEDIUM": "🟡", "LOW/NONE": "🔴"}[conf]
        # Show up to 2 lines so bilingual / wrapped descriptions are visible
        desc_lines = str(item["description"]).split("\n")
        preview    = " ↵ ".join(l[:60] for l in desc_lines[:2] if l.strip())
        label = f"{icon} Row {item['row']} — {preview}"

        with st.expander(label, expanded=(conf != "HIGH")):
            current = matches.get(item["row"])
            options = item["candidates"] + ["— leave blank —"]
            if not item["candidates"]: options = ["— leave blank —"]
            default_idx = options.index(current) if current in options else len(options) - 1

            chosen = st.selectbox("Customer ID", options=options,
                                  index=default_idx, key=f"sel_{item['row']}")
            custom = st.text_input("Or type a custom ID (overrides selection above)",
                                   value="", key=f"txt_{item['row']}")
            resolved = custom.strip() if custom.strip() else (
                None if chosen == "— leave blank —" else chosen)
            matches[item["row"]] = resolved

    st.session_state.matches = matches

    bc, gc = st.columns([1, 5])
    with bc:
        if st.button("← Back"): st.session_state.step = 2; st.rerun()
    with gc:
        if st.button("⚙️ Generate output file →", type="primary"):
            with st.spinner("Generating…"):
                if st.session_state.erp_format == "pdf":
                    out = generate_output_from_pdf(
                        st.session_state.erp_bytes, matches,
                        st.session_state.pdf_batches or [])
                else:
                    out = generate_output_excel(st.session_state.erp_bytes, matches)
            st.session_state.output_bytes = out
            st.session_state.step = 4
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# STEP 4 — Download
# ══════════════════════════════════════════════════════════════════════════
elif st.session_state.step == 4:
    st.subheader("Download")
    matches       = st.session_state.matches
    matched_count = sum(1 for v in matches.values() if v)
    blank_count   = sum(1 for v in matches.values() if not v)

    src_fmt = "PDF" if st.session_state.erp_format == "pdf" else "Excel"
    st.success(f"✅ File ready — converted from **{src_fmt}** · "
               f"**{matched_count}** rows with customer ID · **{blank_count}** left blank.")
    st.caption(
        "White background · black text · landscape fit-to-page · "
        "print area restricted to the data range. "
        "Need a green-branded label? Pick the **Label Colorizer** tool in the sidebar."
    )

    # Derive a sensible output filename
    stem = re.sub(r"\.(xlsx|pdf)$", "", st.session_state.erp_filename, flags=re.I)
    out_name = f"{stem}_customer.xlsx"

    st.download_button(
        label="⬇️  Download customer content list",
        data=st.session_state.output_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.write("")
    if st.button("🔄 Process another file"):
        reset()
