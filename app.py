import streamlit as st
import openpyxl
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill, Font

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
    lookup = {}
    for _, row in ref_df.iterrows():
        cid  = row.get(id_col)
        desc = row.get(desc_col)
        if pd.isna(cid)  if isinstance(cid,  float) else cid  is None: continue
        if pd.isna(desc) if isinstance(desc, float) else desc is None: continue
        nd = normalize(str(desc))
        if nd:
            lookup.setdefault(nd, []).append(str(cid))
    return lookup

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
    Extract ERP content list rows from a PDF.
    Returns (meta_dict, items_list, columns_list).
    """
    meta    = {}
    items   = []
    columns = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            # ── title-block metadata ─────────────────────────────────────
            text = page.extract_text() or ""
            for line in text.splitlines():
                l = line.strip()
                if re.match(r"^KM\w+", l):          meta.setdefault("kit_code", l.split()[0])
                if "assembly batch" in l.lower():    meta.setdefault("batch",    l)
                if "po number" in l.lower():         meta.setdefault("po",       l)
                if "content list" in l.lower():      meta.setdefault("title",    "CONTENT LIST")
                if re.match(r"^SET,", l, re.I) and "kit_name" not in meta:
                    meta["kit_name"] = l

            # ── table rows ───────────────────────────────────────────────
            for table in _extract_tables_robust(page):
                for row in table:
                    # Normalise: None → "", collapse internal newlines within
                    # a single pdfplumber cell (pdfplumber can return multi-line
                    # strings for wrapped cells — keep them as \n).
                    clean = [str(c).strip() if c is not None else "" for c in row]
                    if not any(clean):
                        continue

                    # Header row detection
                    if any("description" in c.lower() for c in clean):
                        columns = clean
                        continue

                    if not columns:
                        continue

                    # Build entry dict; pad short rows with ""
                    padded = clean + [""] * max(0, len(columns) - len(clean))
                    entry  = {columns[i]: padded[i] for i in range(len(columns))}

                    desc_key = next(
                        (k for k in columns if "description" in k.lower()), None
                    )
                    desc_val = entry.get(desc_key, "").strip() if desc_key else ""

                    if not desc_val:
                        continue

                    # ── continuation detection ───────────────────────────
                    # A continuation row is one where the description cell has
                    # content but ALL other columns are empty.
                    # We do NOT rely on key-name comparison — we simply check
                    # whether any non-description column has a non-empty value.
                    is_continuation = (
                        bool(items)
                        and desc_key is not None
                        and not _has_data_columns(entry, desc_key)
                    )

                    if is_continuation:
                        # Append wrapped text to the previous item's description
                        prev_desc = items[-1].get(desc_key, "")
                        items[-1][desc_key] = (prev_desc + "\n" + desc_val).strip()
                    else:
                        items.append(entry)

    return meta, items, columns

def load_erp_items(file_bytes, filename):
    """
    Returns (items, source_format).
    items: list of dicts with at least 'row' (int) and 'description' (str).
      For PDF items, 'row' is a 1-based index (no real sheet row).
    source_format: 'excel' or 'pdf'
    """
    if filename.lower().endswith(".pdf"):
        meta, pdf_items, cols = load_erp_items_pdf(file_bytes)
        desc_key = next((k for k in cols if "description" in k.lower()), cols[0] if cols else "description")
        items = [{"row": i + 1,
                  "description": row.get(desc_key, ""),
                  "_pdf_row": row}
                 for i, row in enumerate(pdf_items)]
        return items, "pdf", meta, cols
    else:
        items = load_erp_items_excel(file_bytes)
        return items, "excel", {}, []

# ── matching ───────────────────────────────────────────────────────────────

def match_item(desc_bilingual, lookup):
    if not desc_bilingual: return None, "LOW/NONE", []
    candidates = {}
    for frag in str(desc_bilingual).split("\n"):
        for cid in lookup.get(normalize(frag.strip()), []):
            candidates[cid] = True
    seen, unique = set(), []
    for c in candidates:
        if c not in seen: unique.append(c); seen.add(c)
    if not unique:      return None,      "LOW/NONE", []
    if len(unique) == 1: return unique[0], "HIGH",    unique
    return unique[0], "MEDIUM", unique

# ── output generation ──────────────────────────────────────────────────────

def apply_green(ws):
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.fill = green_fill
            f = cell.font
            cell.font = Font(name=f.name, size=f.size, bold=f.bold,
                             italic=f.italic, color="000000")

def generate_output_excel(erp_bytes, confirmed_matches):
    wb = openpyxl.load_workbook(BytesIO(erp_bytes), data_only=False)
    ws = wb.active
    _, _, item_code_col = find_erp_table_excel(ws)
    ic = item_code_col or 1
    for row_num, cid in confirmed_matches.items():
        ws.cell(row_num, ic).value = cid or None
    apply_green(ws)
    out = BytesIO(); wb.save(out)
    return out.getvalue()

def generate_output_from_pdf(erp_bytes, confirmed_matches, pdf_meta, pdf_cols, erp_items):
    """Build a fresh Excel from PDF-extracted data + confirmed customer IDs."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Content list"

    row_ptr = 1

    # Title block
    def write_title_cell(r, c, val, bold=False):
        cell = ws.cell(r, c, val)
        if bold: cell.font = Font(bold=True)

    write_title_cell(row_ptr, 1, pdf_meta.get("title", "CONTENT LIST"), bold=True); row_ptr += 1
    if "kit_code" in pdf_meta:
        write_title_cell(row_ptr, 1, pdf_meta["kit_code"], bold=True); row_ptr += 1
    if "kit_name" in pdf_meta:
        write_title_cell(row_ptr, 1, pdf_meta["kit_name"]); row_ptr += 2
    if "batch" in pdf_meta:
        write_title_cell(row_ptr, 1, pdf_meta["batch"]); row_ptr += 1
    if "po" in pdf_meta:
        write_title_cell(row_ptr, 1, pdf_meta["po"]); row_ptr += 2

    # Header row — prepend "Item code" if not already present
    desc_key  = next((k for k in pdf_cols if "description" in k.lower()), None)
    has_ic    = any("item code" in k.lower() or k.lower() == "code" for k in pdf_cols)
    out_cols  = (pdf_cols if has_ic else ["Item code"] + pdf_cols)

    for ci, col in enumerate(out_cols, 1):
        ws.cell(row_ptr, ci, col)
    header_row = row_ptr
    row_ptr += 1

    # Data rows
    ic_col_idx = next((i + 1 for i, k in enumerate(out_cols) if "item code" in k.lower() or k.lower() == "code"), 1)
    for item in erp_items:
        cid = confirmed_matches.get(item["row"])
        pdf_row = item.get("_pdf_row", {})
        for ci, col in enumerate(out_cols, 1):
            if col == out_cols[ic_col_idx - 1] and not has_ic:
                ws.cell(row_ptr, ci, cid or "")
            else:
                ws.cell(row_ptr, ci, pdf_row.get(col, ""))
        row_ptr += 1

    apply_green(ws)
    out = BytesIO(); wb.save(out)
    return out.getvalue()

# ══════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════

DEFAULTS = dict(step=1, erp_bytes=None, erp_filename="", ref_bytes=None, ref_filename="",
                ref_df=None, id_col=None, desc_col=None, lookup=None,
                erp_items=None, erp_format="excel", pdf_meta={}, pdf_cols=[],
                matches=None, output_bytes=None)
for k, v in DEFAULTS.items():
    if k not in st.session_state: st.session_state[k] = v

def reset():
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# HEADER + PROGRESS
# ══════════════════════════════════════════════════════════════════════════

st.title("🏥 ICRC Content List Customizer")
st.caption("Upload an ERP content list (Excel or PDF) and a customer reference list (Excel or PDF) "
           "to produce a branded Excel output.")

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
                    items, fmt, pdf_meta, pdf_cols = load_erp_items(
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

                st.session_state.erp_items  = items
                st.session_state.erp_format = fmt
                st.session_state.pdf_meta   = pdf_meta
                st.session_state.pdf_cols   = pdf_cols
                st.session_state.matches    = {i["row"]: i["matched_id"] for i in items}
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
                        st.session_state.pdf_meta,
                        st.session_state.pdf_cols,
                        st.session_state.erp_items)
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
    st.write(f"Green fill: `#{GREEN_HEX}` (Pantone 340 U) · Text: black")

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
