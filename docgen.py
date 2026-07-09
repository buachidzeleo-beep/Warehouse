"""
docgen.py — Warehouse pick list & delivery note generator.

Reads the daily produce export (same layout as the example file) and produces:
  • Pick list  : one Word file per VEHICLE  (shops broken down within each file)
  • Delivery   : one Word file per STORE

Only 5 source columns are used (by position, robust to header text changes):
  A (0) Vehicle | C (2) Store | E (4) Barcode | F (5) Item | I (8) Boxes
"""

import io
import re
import zipfile

import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ----------------------------------------------------------------------------
# Column positions in the source file (0-based). "Exactly like the example."
# ----------------------------------------------------------------------------
COL_VEHICLE = 0   # A
COL_STORE   = 2   # C
COL_BARCODE = 4   # E
COL_ITEM    = 5   # F
COL_BOXES   = 8   # I
MIN_COLUMNS = 9

FONT = "Arial"
C_HEADER_FILL = "1F3864"
C_ALT_FILL    = "F2F2F2"
C_TOTAL_FILL  = "DDEBF7"
C_SHOP_FILL   = "E8EEF7"
C_GREY        = "595959"
C_ACCENT      = "1F3864"

ALIGN = {
    "left":   WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right":  WD_ALIGN_PARAGRAPH.RIGHT,
}

# column widths (inches) — sum ≈ 7.5" content width (Letter, 0.5" margins)
PICK_WIDTHS = [0.36, 3.81, 1.67, 0.83, 0.83]          # №, item, barcode, boxes, ✓
DEL_WIDTHS  = [0.39, 4.47, 1.67, 0.97]                # №, item, barcode, boxes


# ============================================================================
# 1. DATA LOADING & AGGREGATION
# ============================================================================
def load_and_prepare(file_like):
    """Read the export, keep the 5 used columns, clean, and return a tidy frame.

    Returns (df, stats) where df has columns:
        Vehicle, Store, Barcode, Item, Boxes
    and stats is a dict with row counts for a data-quality summary.
    """
    raw = pd.read_excel(file_like, header=None, skiprows=1, dtype=str)
    if raw.shape[1] < MIN_COLUMNS:
        raise ValueError(
            f"ფაილში მოსალოდნელია მინიმუმ {MIN_COLUMNS} სვეტი (A–I), "
            f"ნაპოვნია {raw.shape[1]}. შეამოწმეთ, რომ ფაილი ნიმუშის იდენტურია."
        )

    df = pd.DataFrame({
        "Vehicle": raw.iloc[:, COL_VEHICLE],
        "Store":   raw.iloc[:, COL_STORE],
        "Barcode": raw.iloc[:, COL_BARCODE],
        "Item":    raw.iloc[:, COL_ITEM],
        "Boxes":   raw.iloc[:, COL_BOXES],
    })

    total_read = len(df)
    for c in ["Vehicle", "Store", "Barcode", "Item"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["Boxes"] = pd.to_numeric(df["Boxes"], errors="coerce")

    # drop unusable rows
    bad_mask = (
        (df["Vehicle"] == "") | (df["Store"] == "") |
        (df["Barcode"] == "") | (df["Item"] == "") |
        df["Boxes"].isna() | (df["Boxes"] <= 0)
    )
    dropped = int(bad_mask.sum())
    df = df[~bad_mask].copy()
    df["Boxes"] = df["Boxes"].round().astype(int)

    if df.empty:
        raise ValueError("გამოსაყენებელი მონაცემები ვერ მოიძებნა (ცარიელი ან არასწორი ფაილი).")

    # defensive: collapse any duplicate Vehicle×Store×Barcode
    df = (df.groupby(["Vehicle", "Store", "Barcode", "Item"], as_index=False, sort=False)
            ["Boxes"].sum())

    stats = {
        "rows_read": total_read,
        "rows_used": len(df),
        "rows_dropped": dropped,
        "vehicles": df["Vehicle"].nunique(),
        "stores": df["Store"].nunique(),
        "skus": df["Barcode"].nunique(),
        "total_boxes": int(df["Boxes"].sum()),
    }
    return df, stats


def _store_source_order(df):
    """Return list of (vehicle, store) in the order stores first appear."""
    order = (df.assign(_i=range(len(df)))
               .groupby(["Vehicle", "Store"], sort=False)["_i"].min()
               .reset_index().sort_values("_i"))
    return list(zip(order["Vehicle"], order["Store"]))


def build_pick_nested(df):
    """vehicle -> {shops -> {rows}}  (shops in source order, items alphabetical)."""
    pairs = _store_source_order(df)
    out = []
    for veh in sorted(df["Vehicle"].unique()):
        dv = df[df["Vehicle"] == veh]
        shops = []
        for (v, st) in pairs:
            if v != veh:
                continue
            ds = dv[dv["Store"] == st].sort_values("Item", kind="stable")
            rows = [{"item": r.Item, "barcode": r.Barcode, "boxes": int(r.Boxes)}
                    for r in ds.itertuples()]
            shops.append({"store": st, "boxes": int(ds["Boxes"].sum()), "rows": rows})
        out.append({
            "vehicle": veh,
            "shop_count": len(shops),
            "total_boxes": int(dv["Boxes"].sum()),
            "shops": shops,
        })
    return out


def build_delivery_flat(df):
    """One entry per store (vehicle -> store order), items alphabetical."""
    pairs = _store_source_order(df)
    out = []
    for (veh, st) in pairs:
        ds = df[(df["Vehicle"] == veh) & (df["Store"] == st)].sort_values("Item", kind="stable")
        rows = [{"item": r.Item, "barcode": r.Barcode, "boxes": int(r.Boxes)}
                for r in ds.itertuples()]
        out.append({"vehicle": veh, "store": st,
                    "boxes": int(ds["Boxes"].sum()), "rows": rows})
    return out


# ============================================================================
# 2. LOW-LEVEL DOCX HELPERS
# ============================================================================
def _set_run_font(run, name=FONT, size=None, bold=None, color=None):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(attr), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)


def _shade(el_pr, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    el_pr.append(shd)


def _cell_bg(cell, hex_color):
    _shade(cell._tc.get_or_add_tcPr(), hex_color)


def _para_bg(paragraph, hex_color):
    _shade(paragraph._p.get_or_add_pPr(), hex_color)


def _vcenter(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    v = OxmlElement("w:vAlign")
    v.set(qn("w:val"), "center")
    tcPr.append(v)


def _para_bottom_border(paragraph, color=C_ACCENT, sz=10):
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    pPr.append(pbdr)


def _table_borders(table, color="BFBFBF", sz=4):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def _fixed_layout(table, widths_in):
    table.autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)

    # overall table width
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(round(sum(widths_in) * 1440))))

    # rewrite the grid — LibreOffice/Word use tblGrid under fixed layout
    old_grid = tbl.find(qn("w:tblGrid"))
    if old_grid is not None:
        tbl.remove(old_grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths_in:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(round(w * 1440))))
        grid.append(gc)
    tblPr.addnext(grid)

    # per-cell widths too
    for row in table.rows:
        for i, w in enumerate(widths_in):
            row.cells[i].width = Inches(w)


def _repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)


def _cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    cs = OxmlElement("w:cantSplit")
    cs.set(qn("w:val"), "true")
    trPr.append(cs)


def _cell_text(cell, text, align="left", bold=False, color="000000", size=10, fill=None):
    if fill:
        _cell_bg(cell, fill)
    _vcenter(cell)
    p = cell.paragraphs[0]
    p.alignment = ALIGN[align]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run("" if text is None else str(text))
    _set_run_font(r, FONT, size, bold, color)


def _setup_page(doc):
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(0.6)
    sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.5)
    sec.right_margin = Inches(0.5)
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(10)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(attr), FONT)
    pf = style.paragraph_format
    pf.space_after = Pt(2)
    pf.line_spacing = 1.0


def _add_para(doc, runs, space_after=4, space_before=0, keep_next=True,
              bottom_border=False, fill=None):
    """runs = list of (text, dict-of-run-kwargs)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.keep_with_next = keep_next
    if fill:
        _para_bg(p, fill)
    if bottom_border:
        _para_bottom_border(p)
    for text, kw in runs:
        r = p.add_run(text)
        _set_run_font(r, FONT, kw.get("size", 10), kw.get("bold", False), kw.get("color", "000000"))
    return p


# ============================================================================
# 3. TABLE + DOCUMENT BUILDERS
# ============================================================================
def _item_table(doc, rows, widths, with_check, total_value=None):
    ncols = len(widths)
    table = doc.add_table(rows=1, cols=ncols)
    _table_borders(table)

    # header
    hdr = table.rows[0]
    _repeat_header(hdr)
    heads = ["№", "ნომენკლატურა", "შტრიხ-კოდი", "ყუთი"] + (["✓"] if with_check else [])
    aligns = ["center", "left", "left", "center"] + (["center"] if with_check else [])
    for c, (t, a) in enumerate(zip(heads, aligns)):
        _cell_text(hdr.cells[c], t, align=a, bold=True, color="FFFFFF", size=10, fill=C_HEADER_FILL)

    # data
    for i, r in enumerate(rows):
        row = table.add_row()
        _cant_split(row)
        fill = C_ALT_FILL if i % 2 == 1 else None
        _cell_text(row.cells[0], i + 1, align="center", size=10, fill=fill)
        _cell_text(row.cells[1], r["item"], align="left", size=(11 if not with_check else 10), fill=fill)
        _cell_text(row.cells[2], r["barcode"], align="left", color=C_GREY, size=10, fill=fill)
        _cell_text(row.cells[3], r["boxes"], align="center", bold=True,
                   size=(11 if not with_check else 10), fill=fill)
        if with_check:
            _cell_text(row.cells[4], "", align="center", fill=fill)

    # total row
    if total_value is not None:
        row = table.add_row()
        _cant_split(row)
        _cell_text(row.cells[0], "", fill=C_TOTAL_FILL)
        _cell_text(row.cells[1], "სულ ყუთი", align="right", bold=True, fill=C_TOTAL_FILL)
        _cell_text(row.cells[2], "", fill=C_TOTAL_FILL)
        _cell_text(row.cells[3], total_value, align="center", bold=True, size=12, fill=C_TOTAL_FILL)
        if with_check:
            _cell_text(row.cells[4], "", fill=C_TOTAL_FILL)

    _fixed_layout(table, widths)
    return table


def make_picklist_doc(veh, date_str):
    """One vehicle's pick list: per-shop breakdown with a tick column."""
    doc = Document()
    _setup_page(doc)

    _add_para(doc, [(f"შესაგროვებელი სია — {veh['vehicle']}", {"size": 15, "bold": True})],
              space_after=2)
    _add_para(doc, [(f"თარიღი: {date_str}    ·    მაღაზია: {veh['shop_count']}"
                     f"    ·    სულ ყუთი: {veh['total_boxes']}", {"size": 9, "color": C_GREY})],
              space_after=6)

    for shop in veh["shops"]:
        _add_para(doc,
                  [(shop["store"], {"size": 10, "bold": True}),
                   (f"      ·  ყუთი: {shop['boxes']}", {"size": 10, "bold": True, "color": C_GREY})],
                  space_before=6, space_after=2, fill=C_SHOP_FILL)
        _item_table(doc, shop["rows"], PICK_WIDTHS, with_check=True)
    return doc


def make_delivery_doc(note, date_str, signature=False):
    """One store's delivery note: big shop name, small vehicle line."""
    doc = Document()
    _setup_page(doc)

    _add_para(doc, [(f"მანქანა: {note['vehicle']}     ·     თარიღი: {date_str}",
                     {"size": 10, "color": C_GREY})], space_after=2)
    _add_para(doc, [(note["store"], {"size": 20, "bold": True})],
              space_after=3, bottom_border=True)
    _add_para(doc, [(f"მიღების ფურცელი  ·  სულ ყუთი: {note['boxes']}",
                     {"size": 9, "color": C_GREY})], space_before=2, space_after=6)

    _item_table(doc, note["rows"], DEL_WIDTHS, with_check=False, total_value=note["boxes"])

    if signature:
        _add_para(doc, [("", {})], space_before=10)
        _add_para(doc, [("მიმღები (სახელი/გვარი): ____________________________        ",
                         {"size": 10}),
                        ("ხელმოწერა: ______________        თარიღი: ____________",
                         {"size": 10})], space_before=14, keep_next=False)
    return doc


# ============================================================================
# 4. PACKAGING (bytes + zip)
# ============================================================================
def _doc_bytes(doc):
    b = io.BytesIO()
    doc.save(b)
    return b.getvalue()


def sanitize_filename(s, maxlen=120):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return (s[:maxlen]).strip() or "unnamed"


def _zip_bytes(files):
    """files: list of (filename, bytes). Ensures unique names."""
    buf = io.BytesIO()
    seen = {}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files:
            base = name
            if name in seen:
                seen[name] += 1
                stem = name[:-5] if name.lower().endswith(".docx") else name
                base = f"{stem}_{seen[name]}.docx"
            else:
                seen[name] = 1
            z.writestr(base, data)
    return buf.getvalue()


def build_picklist_zip(df, date_str):
    """One .docx per vehicle -> zip bytes."""
    files = []
    for veh in build_pick_nested(df):
        fname = f"{sanitize_filename(veh['vehicle'])}.docx"
        files.append((fname, _doc_bytes(make_picklist_doc(veh, date_str))))
    return _zip_bytes(files), len(files)


def build_delivery_zip(df, date_str, signature=False, vehicle_prefix=True):
    """One .docx per store -> zip bytes.

    vehicle_prefix keeps each truck's notes grouped together when the zip is
    sorted by name; set False for plain store-named files.
    """
    files = []
    for note in build_delivery_flat(df):
        if vehicle_prefix:
            fname = f"{sanitize_filename(note['vehicle'])} - {sanitize_filename(note['store'], 90)}.docx"
        else:
            fname = f"{sanitize_filename(note['store'])}.docx"
        files.append((fname, _doc_bytes(make_delivery_doc(note, date_str, signature))))
    return _zip_bytes(files), len(files)
