import streamlit as st
import pdfplumber
from docx import Document
import os
import zipfile
from datetime import datetime
import re

st.set_page_config(page_title="SN Tools", page_icon="🧪", layout="wide")

# ==============================
# Utility Functions
# ==============================

def is_micro_test(name):
    keywords = [
        "Total Aerobic",
        "Mold",
        "Yeast",
        "Coliform",
        "Salmonella",
        "E.coli",
        "E. coli",
        "Staphylococcus"
    ]
    return any(k.lower() in name.lower() for k in keywords)


# -------------------------------------------------------
# Label aliases: maps known label variants from PDFs
# to our internal session_state keys.
# -------------------------------------------------------
HEADER_LABEL_MAP = {
    # Product name
    "product name":             "product_name",
    "product":                  "product_name",
    "item name":                "product_name",
    "item":                     "product_name",

    # Brand
    "brand":                    "brand",
    "manufacturer":             "brand",

    # Lot / Batch number
    "lot no":                   "lot_no",
    "lot no.":                  "lot_no",
    "lot number":               "lot_no",
    "lot#":                     "lot_no",
    "batch no":                 "lot_no",
    "batch no.":                "lot_no",
    "batch number":             "lot_no",
    "batch#":                   "lot_no",
    "batch":                    "lot_no",

    # Quantity / Batch size
    "quantity":                 "quantity",
    "qty":                      "quantity",
    "batch size":               "quantity",
    "size":                     "quantity",
    "net weight":               "quantity",
    "net wt":                   "quantity",
    "net wt.":                  "quantity",
    "total quantity":           "quantity",

    # Manufacturing date
    "mfg. date":                "mfg_date",
    "mfg.date":                 "mfg_date",
    "mfg date":                 "mfg_date",
    "mfg":                      "mfg_date",
    "manufacturing date":       "mfg_date",
    "manufacture date":         "mfg_date",
    "manufactured date":        "mfg_date",
    "date of manufacture":      "mfg_date",
    "date of manufacturing":    "mfg_date",
    "production date":          "mfg_date",
    "produced date":            "mfg_date",

    # Expiry date
    "expiry date":              "expiry_date",
    "expiry":                   "expiry_date",
    "expire date":              "expiry_date",
    "expiration date":          "expiry_date",
    "expiration":               "expiry_date",
    "exp. date":                "expiry_date",
    "exp.date":                 "expiry_date",
    "exp date":                 "expiry_date",
    "exp":                      "expiry_date",
    "best before":              "expiry_date",
    "use by":                   "expiry_date",

    # Shelf life
    "shelf life":               "shelf_life",
    "shelflife":                "shelf_life",

    # Plant part
    "plant part":               "plant_part",
    "part":                     "plant_part",
    "part used":                "plant_part",
    "parts used":               "plant_part",
    "used part":                "plant_part",

    # Latin / botanical name
    "plant latin name":         "latin_name",
    "latin name":               "latin_name",
    "botanical name":           "latin_name",
    "botanical":                "latin_name",
    "scientific name":          "latin_name",
    "plant name":               "latin_name",

    # Country of origin
    "country of origin":        "origin",
    "country":                  "origin",
    "origin":                   "origin",
    "place of origin":          "origin",
    "source country":           "origin",

    # Solvent
    "solvent":                  "solvent",
    "extraction solvent":       "solvent",
    "solvent of extraction":    "solvent",
}

FIELD_DEFAULTS = {
    "product_name": "",
    "brand":        "Skyherb®",
    "lot_no":       "",
    "quantity":     "",
    "mfg_date":     "",
    "shelf_life":   "3 years",
    "plant_part":   "",
    "latin_name":   "",
    "origin":       "China",
    "solvent":      "",
    "expiry_date":  "",
}


def normalise(text):
    """Lowercase + collapse whitespace for fuzzy label matching."""
    return re.sub(r"\s+", " ", text.strip().lower())


def extract_header_info(pdf_file):
    found = {}
    weight_re = re.compile(r"^\d[\d,. ]*\s*(kg|g|lbs?|lb)", re.IGNORECASE)

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    cells = [c.strip() if c else "" for c in row]
                    for i in range(0, len(cells) - 1, 2):
                        label = normalise(cells[i])
                        value = cells[i + 1].strip()
                        if not value:
                            continue
                        if label in HEADER_LABEL_MAP:
                            key = HEADER_LABEL_MAP[label]
                            if key not in found:
                                found[key] = value
                        elif "quantity" not in found and weight_re.match(value):
                            found["quantity"] = value

    return found


def extract_spec_tables(pdf_file):
    general_rows = []
    micro_rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue
                    if len(row) >= 3:
                        row_data = {
                            "Characteristic": row[0].strip(),
                            "Specification":  row[1].strip() if row[1] else "",
                            "Method":         row[2].strip() if row[2] else "",
                        }
                        if is_micro_test(row_data["Characteristic"]):
                            micro_rows.append(row_data)
                        else:
                            general_rows.append(row_data)

    return general_rows, micro_rows


def extract_coa_tables(pdf_file):
    general_rows = []
    micro_rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue
                    if len(row) >= 4:
                        row_data = {
                            "Characteristic": row[0].strip(),
                            "Standard":       row[1].strip() if row[1] else "",
                            "Result":         row[2].strip() if row[2] else "",
                            "Method":         row[3].strip() if row[3] else "",
                        }
                        if is_micro_test(row_data["Characteristic"]):
                            micro_rows.append(row_data)
                        else:
                            general_rows.append(row_data)

    return general_rows, micro_rows


def replace_placeholders_in_doc(doc, header_data):
    def replace_in_paragraph(paragraph, data):
        full_text = "".join(run.text for run in paragraph.runs)
        replaced = False
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in full_text:
                full_text = full_text.replace(placeholder, value)
                replaced = True
        if replaced:
            for i, run in enumerate(paragraph.runs):
                run.text = full_text if i == 0 else ""

    for p in doc.paragraphs:
        replace_in_paragraph(p, header_data)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, header_data)


def find_table_by_header(doc, keywords):
    for table in doc.tables:
        if not table.rows:
            continue
        first_row_text = " ".join(
            cell.text.strip().lower() for cell in table.rows[0].cells
        )
        if any(kw.lower() in first_row_text for kw in keywords):
            return table
    return None


def safe_write_row(cells, values):
    for i, val in enumerate(values):
        if i < len(cells):
            cells[i].text = val


def fill_spec_template(header_data, general_rows, micro_rows):
    doc = Document("templates/Internal_SPEC_Template.docx")
    replace_placeholders_in_doc(doc, header_data)

    main_table  = find_table_by_header(doc, ["characteristic", "specification", "method"])
    micro_table = find_table_by_header(doc, ["microbiological", "microbiology", "aerobic", "yeast", "mold", "coliform"])

    if main_table is None:
        if doc.tables:
            main_table = doc.tables[0]
            st.warning("Could not identify the SPEC main table by header — using the first table as fallback.")
        else:
            st.error("No tables found in the SPEC template.")
            return None

    for row_data in general_rows:
        safe_write_row(
            main_table.add_row().cells,
            [row_data.get("Characteristic", ""),
             row_data.get("Specification", ""),
             row_data.get("Method", "")]
        )

    if micro_table is not None:
        for row_data in micro_rows:
            safe_write_row(
                micro_table.add_row().cells,
                [row_data.get("Characteristic", ""),
                 row_data.get("Specification", ""),
                 row_data.get("Method", "")]
            )
    elif micro_rows:
        st.warning("No microbiological table found in the SPEC template — micro rows were skipped.")

    output_path = "outputs/Generated_SPEC.docx"
    doc.save(output_path)
    return output_path


def fill_coa_template(header_data, general_rows, micro_rows):
    doc = Document("templates/Internal_COA_Template.docx")
    replace_placeholders_in_doc(doc, header_data)

    main_table  = find_table_by_header(doc, ["characteristic", "standard", "result", "method"])
    micro_table = find_table_by_header(doc, ["microbiological", "microbiology", "aerobic", "yeast", "mold", "coliform"])

    if main_table is None:
        if doc.tables:
            main_table = doc.tables[0]
            st.warning("Could not identify the COA main table by header — using the first table as fallback.")
        else:
            st.error("No tables found in the COA template.")
            return None

    for row_data in general_rows:
        safe_write_row(
            main_table.add_row().cells,
            [row_data.get("Characteristic", ""),
             row_data.get("Standard", ""),
             row_data.get("Result", ""),
             row_data.get("Method", "")]
        )

    if micro_table is not None:
        for row_data in micro_rows:
            safe_write_row(
                micro_table.add_row().cells,
                [row_data.get("Characteristic", ""),
                 row_data.get("Standard", ""),
                 row_data.get("Result", ""),
                 row_data.get("Method", "")]
            )
    elif micro_rows:
        st.warning("No microbiological table found in the COA template — micro rows were skipped.")

    output_path = "outputs/Generated_COA.docx"
    doc.save(output_path)
    return output_path


# ==============================
# Watermark / Header / Scan Functions
# ==============================

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream, NameObject, DecodedStreamObject
import io
import numpy as np
from PIL import Image as PILImage

import sys
FITZ_ERROR = ""
try:
    import pymupdf as fitz  # PyMuPDF — page rendering for preview + colour removal
except Exception as _e1:
    try:
        import fitz
        if not hasattr(fitz, "open"):   # the unrelated "fitz" package from PyPI
            raise ImportError("a different package called 'fitz' is installed instead of PyMuPDF")
    except Exception as _e2:
        fitz = None
        FITZ_ERROR = f"{_e1} / {_e2}"

try:  # optional: lets users click on the preview to pick a colour
    from streamlit_image_coordinates import streamlit_image_coordinates
except ImportError:
    streamlit_image_coordinates = None

try:  # optional: reads text on scanned pages for the source-info check
    import pytesseract
    pytesseract.get_tesseract_version()
except Exception:
    pytesseract = None

HEADER_IMAGE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "SN_Header.png")
PREVIEW_DPI = 110
PICK_WIDTH = 450   # fallback display width (px) of the clickable preview


def _stretch_call(fn, *args, **kw):
    """Call a Streamlit element full-width on both new (width=) and old (use_container_width=) versions."""
    try:
        return fn(*args, width="stretch", **kw)
    except TypeError:
        return fn(*args, use_container_width=True, **kw)


def _extract_image_thumbnail(xobj):
    try:
        width  = int(xobj.get("/Width",  0))
        height = int(xobj.get("/Height", 0))
        if width == 0 or height == 0:
            return None

        data = xobj.get_data()

        filter_val = xobj.get("/Filter", "")
        filters = [str(f) for f in filter_val] if hasattr(filter_val, "__iter__") and not isinstance(filter_val, str) else [str(filter_val)]

        if "/DCTDecode" in filters:
            img = PILImage.open(io.BytesIO(data))
        else:
            cs = xobj.get("/ColorSpace", "/DeviceRGB")
            cs_str = str(cs[0] if isinstance(cs, list) else cs)
            if "RGB"  in cs_str: mode, channels = "RGB",  3
            elif "CMYK" in cs_str: mode, channels = "CMYK", 4
            else:                  mode, channels = "L",    1
            expected = width * height * channels
            if len(data) < expected:
                return None
            img = PILImage.frombytes(mode, (width, height), data[:expected])

        img.thumbnail((180, 180), PILImage.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None


def scan_xobjects(pdf_bytes):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    seen   = {}

    for page_num, page in enumerate(reader.pages):
        content = page.get_contents()
        if content is None:
            continue

        try:
            content_obj = ContentStream(content, reader)
        except Exception:
            continue

        do_names = set()
        for operands, operator in content_obj.operations:
            if operator == b"Do" and operands:
                do_names.add(str(operands[0]))

        if not do_names:
            continue

        resources = page.get("/Resources", {})
        xobj_dict = resources.get("/XObject", {})

        for name_obj, xobj_ref in xobj_dict.items():
            name = str(name_obj)
            if name not in do_names:
                continue

            try:
                xobj = xobj_ref.get_object() if hasattr(xobj_ref, "get_object") else xobj_ref
            except Exception:
                continue

            subtype = str(xobj.get("/Subtype", "/Unknown")).lstrip("/")

            if name not in seen:
                thumbnail = _extract_image_thumbnail(xobj) if subtype == "Image" else None
                seen[name] = {
                    "name":      name,
                    "subtype":   subtype,
                    "pages":     [page_num + 1],
                    "thumbnail": thumbnail,
                }
            else:
                pg = page_num + 1
                if pg not in seen[name]["pages"]:
                    seen[name]["pages"].append(pg)

    return list(seen.values())


def _set_page_contents(page, writer, content_obj):
    """Replace a writer-owned page's content stream (works on old and new pypdf)."""
    if hasattr(page, "replace_contents"):
        page.replace_contents(content_obj)
    else:  # very old pypdf
        new_stream = DecodedStreamObject()
        new_stream.set_data(content_obj.get_data())
        page[NameObject("/Contents")] = writer._add_object(new_stream)


@st.cache_data(show_spinner=False, max_entries=16)
def remove_selected_bytes(input_pdf_bytes, names_to_remove, remove_gs=False):
    """Strip the chosen XObjects (and optionally gs operators). bytes -> bytes."""
    names_to_remove = set(names_to_remove)
    if not names_to_remove and not remove_gs:
        return input_pdf_bytes

    # Pages are edited after they belong to the writer (required by pypdf 7+)
    writer = PdfWriter(clone_from=io.BytesIO(input_pdf_bytes))

    for page in writer.pages:
        content = page.get_contents()
        if content is None:
            continue
        try:
            content_obj = ContentStream(content, writer)
        except Exception:
            continue

        new_ops = []
        changed = False
        for operands, operator in content_obj.operations:
            if operator == b"Do" and operands and str(operands[0]) in names_to_remove:
                changed = True
                continue
            if operator == b"gs" and remove_gs:
                changed = True
                continue
            new_ops.append((operands, operator))

        if changed:
            content_obj.operations = new_ops
            _set_page_contents(page, writer, content_obj)

    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def remove_selected(input_pdf_bytes, output_path, names_to_remove, remove_gs=False):
    """File-writing wrapper kept for backwards compatibility."""
    data = remove_selected_bytes(input_pdf_bytes, tuple(sorted(names_to_remove)), remove_gs)
    with open(output_path, "wb") as f:
        f.write(data)
    return output_path


def make_header_overlay(page_width, page_height, header_img_path,
                        img_width_pt, top_margin_pt):
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c   = canvas.Canvas(buf, pagesize=(page_width, page_height))
    img = ImageReader(header_img_path)
    orig_w, orig_h = img.getSize()
    img_h_pt = img_width_pt * (orig_h / orig_w)
    c.drawImage(header_img_path, 0,
                page_height - top_margin_pt - img_h_pt,
                width=img_width_pt, height=img_h_pt, mask="auto")
    c.save()
    buf.seek(0)
    return buf.read()


def add_header_bytes(pdf_bytes, img_width_frac=0.98, top_margin_pt=10):
    # Pages are edited after they belong to the writer (required by pypdf 7+)
    writer = PdfWriter(clone_from=io.BytesIO(pdf_bytes))

    for page in writer.pages:
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        overlay_bytes = make_header_overlay(pw, ph, HEADER_IMAGE_PATH,
                                            pw * img_width_frac, top_margin_pt)
        overlay_page  = PdfReader(io.BytesIO(overlay_bytes)).pages[0]
        page.merge_page(overlay_page)

    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def add_header_to_pdf(input_path, output_path, img_width_frac=0.98, top_margin_pt=10):
    """File-writing wrapper kept for backwards compatibility."""
    with open(input_path, "rb") as f:
        data = add_header_bytes(f.read(), img_width_frac, top_margin_pt)
    with open(output_path, "wb") as f:
        f.write(data)
    return output_path


# ------------------------------------------------------------------
# Colour (stamp / signature / seal) removal
# ------------------------------------------------------------------

def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def rgb_to_hex(c):
    return "#%02x%02x%02x" % tuple(int(v) for v in c)


def _grow(mask, r):
    """Expand a boolean mask by r pixels in every direction."""
    out = mask.copy()
    for dy in range(-r, r + 1):
        for dx in range(-r, r + 1):
            if dx or dy:
                out |= np.roll(np.roll(mask, dy, axis=0), dx, axis=1)
    return out


def remove_colors(img, targets, hue_tol=25, sat_min=10, whiten=75, edge_clean=55):
    """
    Remove coloured ink (stamps, seals, signatures) while keeping black text under it.

    Ink separation: each matched pixel is explained as  paper × ink^a × text,
    where `ink` is the stamp's typical colour and `a` its local density. Solving
    for `text` per pixel takes the stamp out whether its ink is pale or dense,
    so no ghost of the letters is left. Removed areas get the page's own paper
    tone (scans are rarely pure white), and blended/grey pixels along the ink
    edges are cleaned too.
      hue_tol    - hue distance (degrees) still counted as the same ink
      sat_min    - minimum colour strength (%) for a pixel to count as ink
      whiten     - what's left lighter than this (%) becomes paper
      edge_clean - along the ink edges, anything lighter than this (%) becomes paper
    """
    rgb = np.array(img.convert("RGB")).astype(np.float32)
    hsv = np.array(img.convert("HSV")).astype(np.int16)
    h, s = hsv[..., 0], hsv[..., 1]
    colored = s >= sat_min * 255 / 100

    paper_px = (rgb.mean(axis=2) > 200) & (s < 15)
    paper = (np.median(rgb[paper_px], axis=0) if paper_px.sum() > 100
             else np.array([255, 255, 255], np.float32))
    edge_px = max(2, int(round(min(img.size) / 900)))     # ~3 px at 300 dpi
    out = rgb.copy()
    tol = hue_tol * 255 / 360

    for color in targets:
        th = int(np.array(PILImage.new("RGB", (1, 1), tuple(color)).convert("HSV"))[0, 0, 0])
        d = np.abs(h - th)
        d = np.minimum(d, 256 - d)              # hue wraps around (red is at both ends)
        core = colored & (d <= tol)
        if not core.any():
            continue

        ink = np.median(rgb[core], axis=0)      # this stamp's typical colour
        x = np.log(np.clip(ink / paper, 0.02, 0.999))
        xc = x - x.mean()
        den = max(float((xc ** 2).sum()), 1e-6)
        y = np.log(np.clip(rgb[core] / paper, 1 / 255, 1.0))
        ym = y.mean(axis=1, keepdims=True)
        a = np.clip(((y - ym) * xc).sum(axis=1) / den, 0, None)    # local ink density
        k = np.clip(np.exp(ym[:, 0] - a * x.mean()), 0, 1)        # text darkness left
        val = k[:, None] * paper[None, :]
        val[k >= whiten / 100] = paper
        out[core] = val

        edge = _grow(core, edge_px) & ~core
        e = out[edge]
        e[e.mean(axis=1) >= edge_clean * 255 / 100] = paper
        out[edge] = e

    return PILImage.fromarray(np.clip(out, 0, 255).astype(np.uint8))


def detect_ink_colors(img, max_colors=8, min_share=0.00005):
    """
    Find the coloured inks on a page (ignores black/grey/white).
    Groups neighbouring hues so a thin stamp still counts as one ink,
    and keeps even small marks (a stamp may cover well under 0.1% of a page).
    """
    hsv = np.array(img.convert("HSV")).reshape(-1, 3).astype(np.int32)
    rgb = np.array(img.convert("RGB")).reshape(-1, 3)
    mask = (hsv[:, 1] >= 50) & (hsv[:, 2] >= 50)
    if not mask.any():
        return []
    nb = 24                                     # hue buckets of 15°
    bins = hsv[mask, 0] * nb // 256
    cols = rgb[mask]
    counts = np.bincount(bins, minlength=nb)
    threshold = max(min_share * hsv.shape[0], 30)

    found, used = [], set()
    for b in np.argsort(counts)[::-1]:
        b = int(b)
        if counts[b] == 0 or b in used:
            continue
        group = {(b + k) % nb for k in (-1, 0, 1)} - used   # peak + neighbours
        total = sum(int(counts[g]) for g in group)
        used |= group
        if total < threshold:
            continue
        sel = cols[np.isin(bins, list(group))]
        found.append((total, tuple(int(v) for v in np.median(sel, axis=0))))
        if len(found) >= max_colors:
            break
    return [c for _, c in found]


def pick_color_at(img, x, y, radius=4):
    """Most saturated pixel around (x, y), or None if the spot is grey/white/black."""
    box = img.crop((max(0, x - radius), max(0, y - radius),
                    min(img.width, x + radius + 1), min(img.height, y + radius + 1)))
    hsv = np.array(box.convert("HSV")).reshape(-1, 3)
    rgb = np.array(box.convert("RGB")).reshape(-1, 3)
    if hsv.size == 0:
        return None
    i = int(hsv[:, 1].argmax())
    if hsv[i, 1] < 40:
        return None
    return tuple(int(v) for v in rgb[i])


def rasterize_remove_colors(pdf_bytes, targets, params, dpi=300):
    """Render every page, remove the colours, rebuild a PDF with the same page sizes."""
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = fitz.open()
    for page in src:
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        img = PILImage.frombytes("RGB", (pix.width, pix.height), pix.samples)
        img = remove_colors(img, targets, *params)
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        new_page = out.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, stream=buf.getvalue())
    data = out.tobytes(deflate=True, garbage=3)
    out.close()
    src.close()
    return data


# ------------------------------------------------------------------
# Preview helpers
# ------------------------------------------------------------------

@st.cache_data(show_spinner=False, max_entries=64)
def pdf_page_count(pdf_bytes):
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        return doc.page_count


@st.cache_data(show_spinner=False, max_entries=64)
def render_page(pdf_bytes, page_index, dpi=PREVIEW_DPI):
    """Return (PIL image, page width pt, page height pt)."""
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        page = doc[page_index]
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        img = PILImage.frombytes("RGB", (pix.width, pix.height), pix.samples)
        return img, page.rect.width, page.rect.height


def paste_header(img, page_width_pt, dpi, width_frac, top_margin_pt):
    """Raster version of the header overlay, used only for the preview."""
    hdr = PILImage.open(HEADER_IMAGE_PATH).convert("RGBA")
    scale = dpi / 72
    w = max(1, int(page_width_pt * width_frac * scale))
    h = max(1, int(w * hdr.height / hdr.width))
    hdr = hdr.resize((w, h), PILImage.LANCZOS)
    out = img.convert("RGB").copy()
    out.paste(hdr, (0, int(top_margin_pt * scale)), hdr)
    return out


def areas_for(areas, fname, page_index):
    """White-out boxes (fractions of the page) that apply to this file + page."""
    return [a["box"] for a in areas
            if a["file"] in ("*", fname) and a["page"] in ("*", page_index)]


def whiteout_image(img, boxes):
    """Preview version: paint white rectangles (boxes are 0-1 fractions)."""
    if not boxes:
        return img
    from PIL import ImageDraw
    out = img.convert("RGB").copy()
    d = ImageDraw.Draw(out)
    W, H = out.size
    for x0, y0, x1, y1 in boxes:
        d.rectangle([x0 * W, y0 * H, x1 * W, y1 * H], fill="white")
    return out


def whiteout_pdf(pdf_bytes, fname, areas):
    """
    Real white-out with PyMuPDF redactions: text, images and drawings under
    each box are removed and the box is filled white. Rest of the PDF stays vector.
    """
    if not any(a["file"] in ("*", fname) for a in areas):
        return pdf_bytes
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for pno, page in enumerate(doc):
        boxes = areas_for(areas, fname, pno)
        if not boxes:
            continue
        W, H = page.rect.width, page.rect.height
        for x0, y0, x1, y1 in boxes:
            r = fitz.Rect(x0 * W, y0 * H, x1 * W, y1 * H) * page.derotation_matrix
            page.add_redact_annot(r.normalize(), fill=(1, 1, 1))
        try:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS)
        except (AttributeError, TypeError):
            page.apply_redactions()
    data = doc.tobytes(deflate=True, garbage=3)
    doc.close()
    return data


def process_pdf(pdf_bytes, names_to_remove, remove_gs, colors, color_params, dpi,
                add_header, header_width_frac, header_top_pt, fname="", areas=(),
                return_stages=False):
    """Full pipeline: overlays -> white-out -> colours -> header. bytes -> bytes."""
    data = remove_selected_bytes(pdf_bytes, tuple(sorted(names_to_remove)), remove_gs)
    if areas:
        data = whiteout_pdf(data, fname, areas)
    text_stage = data                   # still has its text layer (used by the source check)
    if colors:
        data = rasterize_remove_colors(data, colors, color_params, dpi)
    image_stage = data                  # what the page looks like before the header
    if add_header:
        data = add_header_bytes(data, header_width_frac, header_top_pt)
    if return_stages:
        return data, text_stage, image_stage
    return data


# ------------------------------------------------------------------
# Source-info check (flags only — never changes the document)
# ------------------------------------------------------------------

SOURCE_RULES = [
    (re.compile(r"\b(co\.?\s*,?\s*ltd|company\s+limited|limited|ltd\.?|inc\.|incorporated|llc|"
                r"corp\.|corporation|gmbh|pvt\.?|plc)(?![a-z])", re.I), "Company name"),
    (re.compile(r"(有限公司|股份|公司|集团|制药|株式会社|주식회사)"), "Company name"),
    (re.compile(r"\b(manufacturer|manufactured\s+by|producer|produced\s+by|supplier|supplied\s+by|"
                r"factory|made\s+by|distributed\s+by)\b", re.I), "Manufacturer / supplier wording"),
    (re.compile(r"\b(tel|fax|phone|telephone|mobile)\b\s*[:.：]?\s*\+?\d", re.I), "Phone / fax"),
    (re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"), "Email"),
    (re.compile(r"\b(www\.\S+|https?://\S+|[\w-]+\.(com|cn|net|org|kr|jp|in|de|co\.kr|com\.cn)\b)", re.I),
     "Website"),
    (re.compile(r"\b(address|addr\.?|add\s*:|road|street|avenue|district|province|industrial\s+park|"
                r"development\s+zone|chemical\s+zone|zip\s*code|postcode)\b", re.I), "Address"),
]


def parse_terms(text):
    return tuple(t.strip().lower() for t in re.split(r"[,;\n]", text or "") if t.strip())


def classify_line(text, ignore_terms, watch_terms):
    t = text.strip()
    if len(t) < 3:
        return None
    low = t.lower()
    for w in watch_terms:
        if w in low:
            return "Supplier name"
    if any(i in low for i in ignore_terms):
        return None
    for rx, reason in SOURCE_RULES:
        if rx.search(t):
            return reason
    return None


def _pdf_text_lines(page):
    """(text, Rect in displayed page coordinates) for every text line on the page."""
    out = []
    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0:
            continue
        for line in b["lines"]:
            txt = "".join(span["text"] for span in line["spans"])
            r = fitz.Rect(line["bbox"])
            if page.rotation:
                r = r * page.rotation_matrix
            out.append((txt, r))
    return out


def _ocr_lines(img):
    """(text, (x0, y0, x1, y1) in pixels) for every text line found by OCR."""
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    lines = {}
    for i, word in enumerate(data["text"]):
        if not word.strip():
            continue
        k = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
        if k not in lines:
            lines[k] = [[word], [x, y, x + w, y + h]]
        else:
            lines[k][0].append(word)
            bb = lines[k][1]
            lines[k][1] = [min(bb[0], x), min(bb[1], y), max(bb[2], x + w), max(bb[3], y + h)]
    return [(" ".join(words), tuple(bb)) for words, bb in lines.values()]


def check_page(page, ignore_terms, watch_terms, raster=None):
    """
    Look for company / source info on one page.
    raster: optional PIL image of how the page finally looks (used for OCR on scanned pages).
    Returns {"spots": [...], "scanned": bool, "ocr": bool, "images": int}
    """
    W, H = page.rect.width, page.rect.height
    lines = [(t, (r.x0 / W, r.y0 / H, r.x1 / W, r.y1 / H)) for t, r in _pdf_text_lines(page)]
    scanned = sum(len(t.strip()) for t, _ in lines) < 20
    used_ocr = False
    if scanned and raster is not None and pytesseract is not None:
        try:
            rw, rh = raster.size
            lines = [(t, (b[0] / rw, b[1] / rh, b[2] / rw, b[3] / rh)) for t, b in _ocr_lines(raster)]
            used_ocr = True
        except Exception:
            pass

    spots = []
    for txt, box in lines:
        reason = classify_line(txt, ignore_terms, watch_terms)
        if reason:
            spots.append({"reason": reason, "text": " ".join(txt.split())[:90], "box": box})

    # images that aren't the whole page (logos, stamps, signatures pasted in)
    n_images = 0
    try:
        for info in page.get_image_info():
            r = fitz.Rect(info["bbox"]) & page.rect
            if not r.is_empty and r.width * r.height < 0.6 * W * H:
                n_images += 1
    except Exception:
        pass
    return {"spots": spots, "scanned": scanned, "ocr": used_ocr, "images": n_images}


def check_pdf(text_pdf_bytes, ignore_terms, watch_terms, image_pdf_bytes=None):
    """Run check_page on every page; OCR scanned pages from the final-looking version."""
    report = []
    with fitz.open(stream=text_pdf_bytes, filetype="pdf") as doc:
        img_doc = fitz.open(stream=image_pdf_bytes, filetype="pdf") if image_pdf_bytes else None
        for pno, page in enumerate(doc):
            raster = None
            if pytesseract is not None and len(page.get_text().strip()) < 20:
                src = img_doc[pno] if img_doc else page
                pix = src.get_pixmap(dpi=200, alpha=False)
                raster = PILImage.frombytes("RGB", (pix.width, pix.height), pix.samples)
            report.append(check_page(page, ignore_terms, watch_terms, raster))
        if img_doc:
            img_doc.close()
    return report


@st.cache_data(show_spinner=False, max_entries=32)
def preview_check(base_bytes, fname, areas_key, page_index, ignore_terms, watch_terms, raster):
    areas = [{"file": f, "page": p, "box": b} for f, p, b in areas_key]
    data = whiteout_pdf(base_bytes, fname, areas) if areas else base_bytes
    with fitz.open(stream=data, filetype="pdf") as doc:
        return check_page(doc[page_index], ignore_terms, watch_terms, raster)


def highlight_spots(img, spots):
    """Draw orange boxes around flagged spots (preview only)."""
    if not spots:
        return img
    from PIL import ImageDraw
    out = img.convert("RGB").copy()
    d = ImageDraw.Draw(out)
    W, H = out.size
    for i, s in enumerate(spots, 1):
        x0, y0, x1, y1 = s["box"]
        box = [x0 * W - 3, y0 * H - 2, x1 * W + 3, y1 * H + 2]
        d.rectangle(box, outline=(255, 140, 0), width=2)
        d.rectangle([box[0], box[1] - 12, box[0] + 14, box[1]], fill=(255, 140, 0))
        d.text((box[0] + 3, box[1] - 12), str(i), fill="white")
    return out


def summarize_report(report):
    """One-line summary like '3 spot(s) on p.1, 1 on p.2; 2 image(s) on p.1'."""
    parts = [f"{len(r['spots'])} on p.{i + 1}" for i, r in enumerate(report) if r["spots"]]
    total = sum(len(r["spots"]) for r in report)
    imgs = [f"{r['images']} on p.{i + 1}" for i, r in enumerate(report) if r["images"]]
    scanned = [str(i + 1) for i, r in enumerate(report) if r["scanned"] and not r["ocr"]]
    return total, parts, imgs, scanned


def swatch_html(hex_color, label=""):
    return (f"<div style='display:flex;align-items:center;gap:8px;margin:4px 0'>"
            f"<div style='width:34px;height:22px;border-radius:4px;border:1px solid #999;"
            f"background:{hex_color}'></div><code>{label or hex_color}</code></div>")


# ==============================
# PDF Unlock Functions
# ==============================

def _try_unlock_reader(pdf_bytes: bytes, password: str) -> tuple[PdfReader | None, str]:
    """
    Attempt to open a PDF and decrypt it.
    Returns (reader, status_string).
    status: "not_encrypted" | "unlocked" | "wrong_password" | "error"
    """
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
    except Exception as e:
        return None, f"error:{e}"

    if not reader.is_encrypted:
        return reader, "not_encrypted"

    # Try supplied password first, then empty string (owner lock / no user pw)
    passwords_to_try = [password] if password else []
    passwords_to_try.append("")  # always try blank as fallback

    for pw in passwords_to_try:
        try:
            result = reader.decrypt(pw)
            if result.value > 0:          # 1 = user pw, 2 = owner pw
                return reader, "unlocked"
        except Exception:
            pass

    return None, "wrong_password"


def unlock_pdf(pdf_bytes: bytes, password: str = "") -> tuple[bytes | None, str]:
    """
    Decrypt a PDF and return the raw bytes of the unlocked file.
    Returns (unlocked_bytes, status).
    status: "not_encrypted" | "unlocked" | "wrong_password" | "error:..."
    """
    reader, status = _try_unlock_reader(pdf_bytes, password)

    if status == "not_encrypted":
        # File is not encrypted — clone it as-is so the user still gets a clean copy
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        buf = io.BytesIO()
        writer.write(buf)
        return buf.getvalue(), "not_encrypted"

    if status == "unlocked":
        writer = PdfWriter()
        writer.append_pages_from_reader(reader)
        # Explicitly clear encryption metadata
        writer._encrypt = None
        buf = io.BytesIO()
        writer.write(buf)
        return buf.getvalue(), "unlocked"

    return None, status


# ==============================
# Session State Initialisation
# ==============================

for field, default in FIELD_DEFAULTS.items():
    if field not in st.session_state:
        st.session_state[field] = default

os.makedirs("outputs", exist_ok=True)


# ==============================
# Sidebar Navigation
# ==============================

with st.sidebar:
    st.title("🧪 SN Tools")
    st.caption("Skyherb® internal document tools")
    st.markdown("---")
    _tools = ["🚿 Watermark Remover", "🔬 COA Review", "📄 SPEC & COA Generator", "🔓 PDF Unlocker"]
    _tool_help = ["Clean supplier PDFs: logos, stamps, white-out, header",
                  "Check supplier COAs against our specs; draft supplier questions",
                  "Turn supplier SPEC/COA PDFs into Skyherb Word documents",
                  "Remove passwords and print/copy restrictions"]
    try:
        page = st.radio("Tool", options=_tools, captions=_tool_help,
                        label_visibility="collapsed", key="nav_tool")
    except TypeError:   # older Streamlit without captions
        page = st.radio("Tool", options=_tools, label_visibility="collapsed", key="nav_tool")
    st.markdown("---")
    st.caption("Files are processed in this session only.")

# Shared look & feel
st.markdown("""
<style>
.block-container {padding-top: 2.2rem; padding-bottom: 3rem;}
h1 {font-size: 1.85rem !important; margin-bottom: 0.2rem !important;}
h3 {font-size: 1.15rem !important;}
.stTabs [data-baseweb="tab-list"] {gap: 4px; flex-wrap: wrap;}
.stTabs [data-baseweb="tab"] {font-weight: 600; padding: 6px 12px; border-radius: 8px 8px 0 0;}
div[data-testid="stExpander"] details {border-radius: 10px;}
.sn-chip {display: inline-block; padding: 2px 10px; margin: 2px 6px 2px 0; border-radius: 999px;
          font-size: 0.82rem; background: rgba(14,140,114,0.10); color: #0b6e5a;
          border: 1px solid rgba(14,140,114,0.35);}
.sn-chip.muted {background: rgba(128,128,128,0.08); color: #777; border-color: rgba(128,128,128,0.3);}
.sn-hint {font-size: 0.86rem; color: #6b6b6b; margin: -0.3rem 0 0.6rem 0;}
</style>""", unsafe_allow_html=True)


def chips(items, empty="No changes yet"):
    if not items:
        return f"<span class='sn-chip muted'>{empty}</span>"
    return "".join(f"<span class='sn-chip'>{i}</span>" for i in items)



# ==============================
# PAGE 1 — SPEC & COA Generator
# ==============================

if page == "📄 SPEC & COA Generator":
    import html as _html
    import pandas as pd
    import coa_review as cr
    import coa_generator as cg

    def _secret(name):
        try:
            return st.secrets.get(name)
        except Exception:
            return None

    API_KEY = _secret("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")
    MODEL = _secret("COA_MODEL") or os.environ.get("COA_MODEL") or cr.DEFAULT_MODEL
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    TEMPLATES = {"COA": os.path.join(APP_DIR, "templates", "Internal_COA_Template.docx"),
                 "SPEC": os.path.join(APP_DIR, "templates", "Internal_SPEC_Template.docx")}
    LIB_PATH = os.path.join(APP_DIR, "spec_library.xlsx")
    FIELDS = [("ProductName", "Product name"), ("LotNo", "Lot No."), ("Quantity", "Quantity"),
              ("ManuDate", "Mfg. date"), ("ExpiryDate", "Expiry date"), ("ShelfLife", "Shelf life"),
              ("PlantPart", "Plant part"), ("LatinName", "Plant Latin name"),
              ("Origin", "Country of origin"), ("Solvent", "Extraction solvent")]
    NONE = "— not in spec library —"
    HIDEABLE = [("Brand", "Brand"), ("LotNo", "Lot No."), ("Quantity", "Quantity"), ("ManuDate", "Mfg. date"),
                ("ShelfLife", "Shelf life"), ("PlantPart", "Plant part"), ("LatinName", "Plant Latin name"),
                ("Origin", "Country of origin"), ("Solvent", "Extraction solvent")]
    HIDE_LABEL = dict(HIDEABLE)
    HIDE_KEY = {v: k for k, v in HIDEABLE}
    NA = re.compile(r"^\s*(n/?a|none|not applicable|-+|—)\s*$", re.I)

    st.session_state.setdefault("gen_rows", [])
    st.session_state.setdefault("gen_rows_v", 0)
    st.session_state.setdefault("gen_msgs", [])
    st.session_state.setdefault("gen_doctype", "COA")
    st.session_state.setdefault("gen_hide", [])
    for k, _ in FIELDS:
        st.session_state.setdefault(f"gen_f_{k}", "China" if k == "Origin" else "")

    @st.cache_data(show_spinner=False, max_entries=8)
    def _gen_lib(path, mtime):
        return cr.load_spec_library(path)

    specs_df = products_df = None
    if os.path.isfile(LIB_PATH):
        try:
            specs_df, products_df = _gen_lib(LIB_PATH, os.path.getmtime(LIB_PATH))
        except Exception:
            pass
    lib_products = list(products_df["Product"]) if products_df is not None else []

    @st.cache_data(show_spinner=False, max_entries=32)
    def _gen_read_ai(data, filename, api_key, model):
        return cr.ai_extract_coa(cr.make_client(api_key), model, data, filename)

    def _gen_reset():
        for k, _ in FIELDS:
            st.session_state[f"gen_f_{k}"] = "China" if k == "Origin" else ""
        st.session_state["gen_rows"] = []
        st.session_state["gen_rows_v"] += 1
        st.session_state["gen_msgs"] = []
        st.session_state.pop("gen_out", None)
        st.session_state["gen_product"] = NONE
        st.session_state["gen_doctype"] = "COA"
        st.session_state["gen_hide"] = []

    def _gen_read():
        """Read the uploaded supplier document into the form (runs before the widgets are drawn)."""
        f = st.session_state.get("gen_upload")
        msgs = []
        if f is None:
            return
        data, name = f.getvalue(), f.name
        coa = None
        if st.session_state.get("gen_use_ai") and API_KEY:
            try:
                coa = _gen_read_ai(data, name, API_KEY, MODEL)
            except Exception as e:
                msgs.append(("warning", f"AI reading failed ({e}); used the text reader instead."))
        if coa is None and name.lower().endswith(".pdf"):
            try:
                coa = cr.extract_text_coa(data)
            except Exception as e:
                msgs.append(("error", f"Couldn't read the PDF: {e}"))
        if coa is None or (not coa["tests"] and coa.get("scanned", True)):
            msgs.append(("error", "No text could be read — this looks like a scan. Turn on “Read with AI”."
                          if not API_KEY or not st.session_state.get("gen_use_ai") else
                          "No test rows could be read from this document."))
            st.session_state["gen_msgs"] = msgs
            return
        info = coa["info"]

        # our product in the spec library
        prod, score = (cr.match_product(info.get("product_name", ""), products_df)
                       if products_df is not None else (None, 0))
        st.session_state["gen_product"] = prod or NONE
        pinfo = cr.product_info(products_df, prod) if prod else {}

        shelf = cg.shelf_life_text(info.get("shelf_life", ""),
                                   pinfo.get("Shelf life (months)") if re.fullmatch(
                                       r"\d+(\.0)?", pinfo.get("Shelf life (months)", "") or "") else None)
        values = {
            "ProductName": prod or info.get("product_name", ""),
            "LotNo": info.get("lot_no", ""), "Quantity": info.get("quantity", ""),
            "ManuDate": info.get("mfg_date", ""),
            "ExpiryDate": info.get("expiry_date", "") or info.get("retest_date", ""),
            "ShelfLife": shelf,
            "PlantPart": pinfo.get("Plant part") or info.get("plant_part", ""),
            "LatinName": pinfo.get("Latin name") or info.get("latin_name", ""),
            "Origin": st.session_state.get("gen_f_Origin") or "China",
            "Solvent": pinfo.get("Extraction solvent") or info.get("solvent", ""),
        }
        values["PlantPart"] = re.sub(r"^(used|part)\s*:?\s*", "", values["PlantPart"], flags=re.I)
        if not values["ExpiryDate"]:
            exp = cg.expiry_from(values["ManuDate"], shelf)
            if exp:
                values["ExpiryDate"] = exp.isoformat()
                msgs.append(("info", "Expiry date calculated from the mfg. date and shelf life."))
        hide = [k for k in ("PlantPart", "LatinName", "Solvent") if NA.match(pinfo.get(
            {"PlantPart": "Plant part", "LatinName": "Latin name", "Solvent": "Extraction solvent"}[k], "") or "x")]
        if not values["PlantPart"].strip() and not values["LatinName"].strip():
            hide = sorted(set(hide) | {"PlantPart", "LatinName"}, key=list(HIDE_LABEL).index)
        for k in hide:
            values[k] = ""
        st.session_state["gen_hide"] = [HIDE_LABEL[k] for k in hide]
        if hide:
            msgs.append(("info", "Left off the document: " + ", ".join(HIDE_LABEL[k] for k in hide)
                         + " (no value for this product). Change it under ② if needed."))
        for k, v in values.items():
            st.session_state[f"gen_f_{k}"] = cr.norm_text(v)

        terms = cg.supplier_terms_from(info)
        rows = cg.rows_from_supplier(coa["tests"], st.session_state.get("gen_our_names", True), terms)
        if prod and st.session_state.get("gen_our_spec", False):
            rows = cg.apply_our_spec(rows, cr.spec_rows_for(specs_df, prod))
        st.session_state["gen_rows"] = rows
        st.session_state["gen_rows_v"] += 1
        st.session_state.pop("gen_out", None)
        has_results = any(str(r.get("Results", "")).strip() for r in rows)
        st.session_state["gen_doctype"] = "COA" if has_results else "SPEC"
        if not has_results:
            msgs.append(("info", "No results on this document, so it's treated as a specification → creating a SPEC."))
        src = "AI" if coa["source"] == "ai" else "text reader"
        msgs.insert(0, ("success", f"Read {len(rows)} test rows with the {src}"
                        + (f"; matched to “{prod}” in the spec library." if prod else ".")))
        if info.get("product_name") and prod and cr.clean_name(info["product_name"]) != cr.clean_name(prod):
            msgs.append(("info", f"Supplier calls it “{info['product_name']}”; using our name “{prod}”."))
        if info.get("supplier"):
            msgs.append(("info", f"Supplier on the document: {info['supplier']} — left out of our document."))
        for w in coa.get("warnings", []):
            msgs.append(("warning", f"AI note: {w}"))
        st.session_state["gen_msgs"] = msgs

    def _gen_apply_spec():
        prod = st.session_state.get("gen_product")
        if not prod or prod == NONE or specs_df is None:
            return
        rows = st.session_state.get("gen_rows_current") or st.session_state["gen_rows"]
        st.session_state["gen_rows"] = cg.apply_our_spec(rows, cr.spec_rows_for(specs_df, prod))
        st.session_state["gen_rows_v"] += 1
        st.session_state["gen_msgs"] = [("success", f"Standards replaced with our spec for “{prod}”; "
                                                    "tests our spec requires but the supplier didn't report were added.")]

    # sticky right column
    st.markdown("""
        <style>
        div[data-testid="stColumn"]:has(.st-key-gen_panel), div[data-testid="column"]:has(.st-key-gen_panel) {
            position: sticky; top: 3.75rem; align-self: flex-start;
            max-height: calc(100vh - 4.5rem); overflow-y: auto; padding-right: 6px;
        }
        .sn-doc {font-family: 'Times New Roman', serif; font-size: 12px; color: #111; background: #fff;
                 border: 1px solid rgba(128,128,128,.35); border-radius: 4px; padding: 14px 16px;}
        .sn-doc h4 {text-align: center; margin: 0 0 8px 0; font-size: 16px; color: #111;}
        .sn-doc table {width: 100%; border-collapse: collapse; margin-bottom: 6px;}
        .sn-doc td, .sn-doc th {border: 1px solid #2e8b57; padding: 2px 5px; font-size: 11px; color: #111;}
        .sn-doc th {background: #c0c0c0; font-size: 12px;}
        .sn-doc td.lbl {font-weight: bold; width: 18%;}
        .sn-doc tr.grp td {background: #eef4f1; font-style: italic; font-weight: bold;}
        .sn-doc tr.ban td {font-weight: bold; font-size: 13px; padding-left: 14px;}
        .sn-doc tr.bad td {background: #fde2e1;}
        .sn-doc td.c {text-align: center;}
        </style>""", unsafe_allow_html=True)

    # ── header ────────────────────────────────────────────────────────────
    h1, h2 = st.columns([5, 1])
    with h1:
        st.title("📄 SPEC & COA Generator")
        st.markdown("<div class='sn-hint'>Upload a supplier COA or specification → check the fields and rows → "
                    "download a Skyherb COA / SPEC built from your Word templates.</div>", unsafe_allow_html=True)
    with h2:
        st.write("")
        st.button("↺ Start over", key="gen_reset", on_click=_gen_reset)
    if API_KEY:
        st.markdown(chips([f"🤖 AI reading · {MODEL}"] + ([f"📚 {len(lib_products)} product spec(s)"]
                                                           if lib_products else [])), unsafe_allow_html=True)

    left, right = st.columns([1.15, 1], gap="large")

    # ════════════════════════════ left: inputs ═══════════════════════════
    with left:
        st.subheader("① Supplier document")
        st.file_uploader("Supplier COA or specification (PDF or image)", type=["pdf", "png", "jpg", "jpeg"],
                         key="gen_upload")
        r1, r2, r3 = st.columns([1.1, 1.3, 1])
        with r1:
            st.toggle("Read with AI", value=bool(API_KEY), disabled=not API_KEY, key="gen_use_ai",
                      help="Reads any layout, including scans. Without AI, only text PDFs with tables can be read.")
        with r2:
            st.toggle("Use Skyherb test names", value=True, key="gen_our_names",
                      help="E.g. 'Pb' → 'Lead (Pb)', 'Aerobic bacterial count' → 'Total Plate Count'.")
        with r3:
            _stretch_call(st.button, "📖 Read document", type="primary", key="gen_read", on_click=_gen_read,
                          disabled=st.session_state.get("gen_upload") is None)
        for level, msg in st.session_state["gen_msgs"]:
            getattr(st, level)(msg)

        st.subheader("② Product details")
        if hasattr(st, "segmented_control"):
            doc_type = st.segmented_control("Create", ["COA", "SPEC", "COA + SPEC"],
                                            key="gen_doctype") or "COA"
        else:
            doc_type = st.radio("Create", ["COA", "SPEC", "COA + SPEC"], horizontal=True, key="gen_doctype")
        kinds = ["COA", "SPEC"] if doc_type == "COA + SPEC" else [doc_type]
        BATCH = {"LotNo", "Quantity", "ManuDate", "ExpiryDate"}
        pc1, pc2 = st.columns([2, 1])
        with pc1:
            st.selectbox("Our product (spec library)", [NONE] + lib_products, key="gen_product",
                         help="Used for the product details and, if you choose, our spec limits.")
        with pc2:
            date_fmt = st.selectbox("Date format", list(cg.DATE_FORMATS), key="gen_datefmt")
        def _not_botanical():
            st.session_state["gen_hide"] = sorted(set(st.session_state.get("gen_hide", []))
                                                  | {"Plant part", "Plant Latin name"}, key=list(HIDE_KEY).index)

        hc1, hc2 = st.columns([3, 1.2])
        with hc1:
            hide_opts = [l for k, l in HIDEABLE if "COA" in kinds or k not in BATCH]
            st.session_state["gen_hide"] = [h for h in st.session_state.get("gen_hide", []) if h in hide_opts]
            st.multiselect("Leave off the document", hide_opts, key="gen_hide",
                           help="These rows are removed from the document's header table; the rest close up.")
        with hc2:
            st.write("")
            st.button("🧪 Not plant-derived", key="gen_not_bot", on_click=_not_botanical,
                      help="Leaves off Plant part and Plant Latin name (e.g. Berberine HCl, synthetic or "
                           "fermented ingredients).")
        hidden = {HIDE_KEY[h] for h in st.session_state.get("gen_hide", [])}

        fc1, fc2 = st.columns(2)
        shown_fields = [(k, l) for k, l in FIELDS if ("COA" in kinds or k not in BATCH) and k not in hidden]
        for i, (k, label) in enumerate(shown_fields):
            with (fc1 if i % 2 == 0 else fc2):
                st.text_input(label, key=f"gen_f_{k}")
        if "COA" not in kinds:
            st.caption("A SPEC has no lot, quantity or dates — those fields are hidden.")

        st.subheader("③ Tests")
        o1, o2 = st.columns([1.3, 1])
        with o1:
            st.button("📐 Use our spec limits for Standards", key="gen_apply_spec", on_click=_gen_apply_spec,
                      disabled=st.session_state.get("gen_product", NONE) == NONE,
                      help="Replaces the supplier's limits with the spec library's, and adds required tests "
                           "the supplier didn't report (fill in their results).")
        with o2:
            group_headings = st.toggle("Extra sub-headings", value=False, key="gen_groups",
                                       help="Skyherb style puts headings only on Solvents Residue and "
                                            "Microbiological. Turn on to also head Physical, Assay and Heavy Metals.")
        cols = ["Include", "Group", "Characteristic", "Standards", "Results", "Methods", "Supplier info?"]
        base = pd.DataFrame(st.session_state["gen_rows"] or [], columns=cols)
        base["Include"] = base["Include"].fillna(True).astype(bool)
        base["Supplier info?"] = base["Supplier info?"].fillna(False).astype(bool)
        edited = st.data_editor(
            base, num_rows="dynamic", hide_index=True, key=f"gen_rows_ed_{st.session_state['gen_rows_v']}",
            column_config={
                "Include": st.column_config.CheckboxColumn(default=True, width="small"),
                "Group": st.column_config.SelectboxColumn(options=cg.GROUPS, default="Physical & Chemical",
                                                          required=True),
                "Characteristic": st.column_config.TextColumn(width="medium"),
                "Standards": st.column_config.TextColumn(width="medium"),
                "Results": st.column_config.TextColumn(width="small"),
                "Supplier info?": st.column_config.CheckboxColumn("⚠️ Supplier info", disabled=True,
                                                                  width="small"),
            })
        edited = edited.astype(object).where(edited.notna(), "")
        rows = [r for r in edited.to_dict("records") if str(r.get("Characteristic", "")).strip()]
        for r in rows:
            r["Include"] = bool(r.get("Include", True)) if r.get("Include", "") != "" else True
        order = {g: i for i, g in enumerate(cg.GROUPS)}
        rows.sort(key=lambda r: order.get(r.get("Group"), 0))
        st.session_state["gen_rows_current"] = rows
        st.caption("Untick rows you don't want, edit any cell, or add rows at the bottom. "
                   "Microbiological rows go under the template's micro section.")

    # ════════════════════════════ right: checks, preview, export ═════════
    fields = {k: st.session_state.get(f"gen_f_{k}", "").strip() for k, _ in FIELDS}
    for k in ("ManuDate", "ExpiryDate"):
        fields[k] = cg.format_date(fields[k], date_fmt) if fields[k] else ""
    fields["Brand"] = "Skyherb®"
    for k in hidden:
        fields[k] = ""
    fields["IssueDate"] = cg.format_date(datetime.today().strftime("%Y-%m-%d"), date_fmt)

    with right:
        try:
            panel = st.container(key="gen_panel")
        except TypeError:
            panel = st.container()
    with panel:
        st.markdown(f"**Creating: {' + '.join(kinds)}**")
        missing_tpl = [k for k in kinds if not os.path.isfile(TEMPLATES[k])]
        with st.expander("Templates", expanded=bool(missing_tpl)):
            for k in ("COA", "SPEC"):
                up = st.file_uploader(f"Use a different {k} template for this session (.docx)", type="docx",
                                      key=f"gen_tpl_{k}")
                if up is not None:
                    st.session_state[f"gen_tpl_bytes_{k}"] = up.getvalue()
            st.caption("Templates use placeholders such as {{ProductName}}, {{LotNo}}, {{Quantity}}, {{ManuDate}}, "
                       "{{ExpiryDate}}, {{ShelfLife}}, {{PlantPart}}, {{LatinName}}, {{Origin}}, {{Solvent}}, "
                       "{{IssueDate}}, and a table row with 'Characteristic | Standards | (Results) | Methods'.")

        def _tpl(kind):
            return st.session_state.get(f"gen_tpl_bytes_{kind}") or (
                open(TEMPLATES[kind], "rb").read() if os.path.isfile(TEMPLATES[kind]) else None)

        # ── checks ──
        st.markdown("**✅ Checks before issuing**")
        checks = []
        for k in kinds:
            checks += [(lvl, (f"[{k}] " if len(kinds) > 1 else "") + m)
                       for lvl, m in cg.check_rows(rows, fields, k, hide=hidden)]
        for k in kinds:
            tb = _tpl(k)
            if tb:
                try:
                    from docx import Document as _D
                    have = cg.placeholders_in(_D(io.BytesIO(tb)))
                    if "ExpiryDate" not in have and fields.get("ExpiryDate") and k == "COA" \
                            and "ExpiryDate" not in hidden:
                        checks.append(("info", "The COA template has no {{ExpiryDate}} placeholder, so the expiry "
                                               "date won't appear. Add it to the template in Word to show it."))
                except Exception:
                    pass
            else:
                checks.append(("error", f"No {k} template found (templates/Internal_{k}_Template.docx)."))
        errors = [m for lvl, m in checks if lvl == "error"]
        if not rows:
            st.caption("Read a document (or add rows) to see the checks.")
        elif not checks:
            st.success("No problems found.")
        else:
            icon = {"error": "❌", "warning": "⚠️", "info": "ℹ️"}
            st.markdown("\n".join(f"- {icon[l]} {m.replace('{{', '{{ ').replace('}}', ' }}')}"
                                   for l, m in checks))

        # ── preview ──
        st.markdown("**👁️ Preview**")
        if rows:
            bad = {m.split(":")[0].replace("[COA] ", "") for lvl, m in checks if lvl == "error"}
            show_res = "COA" in kinds and doc_type != "SPEC"
            e = lambda x: _html.escape(str(x or ""))
            if show_res:
                hdr = [("Product Name", fields["ProductName"], "Brand", "Skyherb®"),
                       ("Lot No.", fields["LotNo"], "Quantity", fields["Quantity"]),
                       ("Mfg. Date", fields["ManuDate"], "Shelf Life", fields["ShelfLife"]),
                       ("Plant Part", fields["PlantPart"], "Plant Latin Name", fields["LatinName"]),
                       ("Country of Origin", fields["Origin"], "Solvent", fields["Solvent"])]
            else:
                hdr = [("Product Name", fields["ProductName"], "Brand", "Skyherb®"),
                       ("Extraction Solvent", fields["Solvent"], "Shelf Life", fields["ShelfLife"]),
                       ("Plant Part", fields["PlantPart"], "Plant Latin Name", fields["LatinName"]),
                       ("Country of origin", fields["Origin"], "", "")]
            pairs = [(a, b) for a, b, c, d in hdr for a, b in ((a, b), (c, d)) if a]
            key_of = {"Brand": "Brand", "Lot No.": "LotNo", "Quantity": "Quantity", "Mfg. Date": "ManuDate",
                      "Shelf Life": "ShelfLife", "Plant Part": "PlantPart", "Plant Latin Name": "LatinName",
                      "Country of Origin": "Origin", "Country of origin": "Origin", "Solvent": "Solvent",
                      "Extraction Solvent": "Solvent"}
            pairs = [(a, b) for a, b in pairs if key_of.get(a) not in hidden]
            pairs += [("", "")] * (len(pairs) % 2)
            hdr = [(pairs[i][0], pairs[i][1], pairs[i + 1][0], pairs[i + 1][1]) for i in range(0, len(pairs), 2)]
            h = ["<div class='sn-doc'><h4>" + ("CERTIFICATE OF ANALYSIS" if show_res else "PRODUCT SPECIFICATIONS")
                 + "</h4><table>"]
            h += [f"<tr><td class='lbl'>{e(a)}</td><td>{e(b)}</td><td class='lbl'>{e(c)}</td><td>{e(d)}</td></tr>"
                  for a, b, c, d in hdr]
            h.append("</table><table><tr><th>Characteristic</th><th>Standards</th>"
                     + ("<th>Results</th>" if show_res else "") + "<th>Methods</th></tr>")
            last = None
            ncol = 4 if show_res else 3
            for r in [r for r in rows if r.get("Include", True)]:
                if r.get("Group") != last:
                    label = "Microbiological Specifications" if r.get("Group") == "Microbiological" else r.get("Group")
                    if r.get("Group") == "Microbiological" or r.get("Group") in cg.BANNER_GROUPS:
                        h.append(f"<tr class='ban'><td colspan='{ncol}'>{e(label)}</td></tr>")
                    elif group_headings:
                        h.append(f"<tr class='grp'><td colspan='{ncol}'>{e(label)}</td></tr>")
                    last = r.get("Group")
                cls = " class='bad'" if r.get("Characteristic") in bad else ""
                h.append(f"<tr{cls}><td><b>{e(r.get('Characteristic'))}</b></td><td class='c'>{e(r.get('Standards'))}</td>"
                         + (f"<td class='c'>{e(r.get('Results'))}</td>" if show_res else "")
                         + f"<td class='c'>{e(r.get('Methods'))}</td></tr>")
            h.append("</table></div>")
            st.markdown("".join(h), unsafe_allow_html=True)
            st.caption("Quick preview of the content. The downloaded file uses your Word template's exact layout.")
        else:
            st.caption("Nothing to preview yet.")

        # ── generate ──
        st.markdown("**📦 Download**")
        ok_anyway = True
        if errors:
            ok_anyway = st.checkbox(f"Create anyway ({len(errors)} problem(s) above)", key="gen_override")
        if _stretch_call(st.button, "⚙️ Create document" + ("s" if len(kinds) > 1 else ""), type="primary",
                         key="gen_go", disabled=not rows or not ok_anyway):
            out = {}
            safe = lambda s: re.sub(r"[^\w.\-]+", "_", s).strip("_")[:60]
            for k in kinds:
                tb = _tpl(k)
                if not tb:
                    continue
                try:
                    docx_bytes = cg.fill_template(tb, fields, rows, k, group_headings, hide=hidden)
                    stem = f"SN-{k}-{safe(fields['ProductName'] or 'product')}" + (
                        f"-{safe(fields['LotNo'])}" if k == "COA" and fields["LotNo"] else "")
                    out[k] = {"docx": docx_bytes, "stem": stem, "pdf": cg.docx_to_pdf(docx_bytes)}
                    with open(os.path.join("outputs", stem + ".docx"), "wb") as f:
                        f.write(docx_bytes)
                except Exception as ex:
                    st.error(f"Couldn't create the {k}: {ex}")
            st.session_state["gen_out"] = out

        out = st.session_state.get("gen_out") or {}
        for k, o in out.items():
            c1, c2 = st.columns(2)
            with c1:
                _stretch_call(st.download_button, f"⬇️ {k} (Word)", o["docx"], file_name=o["stem"] + ".docx",
                              mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                              key=f"gen_dl_docx_{k}")
            with c2:
                if o["pdf"]:
                    _stretch_call(st.download_button, f"⬇️ {k} (PDF)", o["pdf"], file_name=o["stem"] + ".pdf",
                                  mime="application/pdf", key=f"gen_dl_pdf_{k}")
            if o["pdf"]:
                with st.expander(f"Exact preview — {k}"):
                    try:
                        import fitz as _fz
                    except ImportError:
                        import pymupdf as _fz
                    d = _fz.open(stream=o["pdf"], filetype="pdf")
                    for pg in d:
                        pix = pg.get_pixmap(dpi=80)
                        _stretch_call(st.image, pix.tobytes("png"))
        if out and not any(o["pdf"] for o in out.values()):
            st.caption("PDF download needs LibreOffice on the server; open the Word file and save as PDF instead.")


# ==============================
# PAGE 2 — Watermark Remover
# ==============================

elif page == "🚿 Watermark Remover":

    # ── state ─────────────────────────────────────────────────────────────
    for _k, _v in {"wm_scan_results": [], "wm_remove_set": set(), "cr_colors": [],
                   "pv_last_event": None, "wo_areas": [], "pv_page": 1}.items():
        if _k not in st.session_state:
            st.session_state[_k] = _v

    def _wm_reset():
        st.session_state["wm_remove_set"] = set()
        st.session_state["wo_areas"] = []
        st.session_state["cr_colors"] = []
        st.session_state["wm_results"] = None
        for k in [k for k in st.session_state.keys() if str(k).startswith("chk_")]:
            del st.session_state[k]

    def _pv_step(delta, n):
        st.session_state["pv_page"] = min(max(1, int(st.session_state.get("pv_page", 1)) + delta), n)

    def _pv_file_changed():
        st.session_state["pv_page"] = 1

    def _set_all_overlays(names, value):
        for n in names:
            st.session_state[f"chk_{n}"] = value

    # Sticky preview column
    st.markdown("""
        <style>
        div[data-testid="stColumn"]:has(.st-key-wm_preview_panel),
        div[data-testid="column"]:has(.st-key-wm_preview_panel) {
            position: sticky; top: 3.75rem; align-self: flex-start;
            max-height: calc(100vh - 4.5rem); overflow-y: auto; padding-right: 6px;
        }
        .st-key-wm_preview_panel [data-testid="stImage"] img,
        .st-key-wm_preview_panel iframe {
            border: 1px solid rgba(128,128,128,0.35); border-radius: 4px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08); background: #fff;
        }
        </style>""", unsafe_allow_html=True)

    # ── header ────────────────────────────────────────────────────────────
    t1, t2 = st.columns([5, 1])
    with t1:
        st.title("🚿 Watermark Remover")
        st.markdown("<div class='sn-hint'>Upload → choose edits in the tabs → check the live preview "
                    "→ process & download.</div>", unsafe_allow_html=True)
    with t2:
        st.write("")
        st.button("↺ Reset all edits", on_click=_wm_reset, key="wm_reset",
                  help="Clear removed logos, white-out boxes and stamp colours.")

    uploaded_files = st.file_uploader(
        "Upload one or more PDF files", type="pdf", accept_multiple_files=True,
        key="wm_uploader", label_visibility="collapsed")

    if not uploaded_files:
        st.info("⬆️ Upload one or more PDFs to get started. You can drop several files at once.")
        st.stop()

    # getvalue() always returns the full file, even after earlier reads
    file_bytes = {f.name: f.getvalue() for f in uploaded_files}
    upload_sig = tuple((n, len(b)) for n, b in file_bytes.items())
    preview_ok = fitz is not None

    # ── auto-scan for embedded logos / watermarks when the uploads change ─
    if st.session_state.get("wm_scan_sig") != upload_sig:
        combined, scan_errors = {}, []
        with st.spinner("Scanning for embedded logos & watermarks…"):
            for fname, data in file_bytes.items():
                try:
                    for item in scan_xobjects(data):
                        name = item["name"]
                        if name not in combined:
                            combined[name] = item.copy()
                        else:
                            combined[name]["pages"] = sorted(set(combined[name]["pages"] + item["pages"]))
                            if combined[name]["thumbnail"] is None and item["thumbnail"]:
                                combined[name]["thumbnail"] = item["thumbnail"]
                except Exception as e:
                    scan_errors.append(f"{fname}: {e}")
        st.session_state["wm_scan_results"] = list(combined.values())
        wm_hints = {"wm", "watermark", "mark", "stamp", "logo", "bg", "background"}
        st.session_state["wm_remove_set"] = {
            it["name"] for it in combined.values()
            if it["subtype"] == "Form" or any(h in it["name"].lstrip("/").lower() for h in wm_hints)}
        for k in [k for k in st.session_state.keys() if str(k).startswith("chk_")]:
            del st.session_state[k]
        st.session_state["wm_scan_errors"] = scan_errors
        st.session_state["wm_scan_sig"] = upload_sig
        st.session_state["wm_results"] = None
        st.session_state["pv_page"] = 1

    for err in st.session_state.get("wm_scan_errors", []):
        st.warning(f"Could not scan {err}")

    scan_results = st.session_state["wm_scan_results"]
    if preview_ok:
        pages_info = ", ".join(f"{n} ({pdf_page_count(b)} p.)" for n, b in file_bytes.items())
    else:
        pages_info = ", ".join(file_bytes.keys())
    st.caption(f"📎 {len(file_bytes)} file(s): {pages_info}")

    left, right = st.columns([1.05, 1], gap="large")

    # ── right panel, part 1: preview selectors ───────────────────────────
    with right:
        try:
            panel = st.container(key="wm_preview_panel")   # key -> CSS hook for the sticky panel
        except TypeError:
            panel = st.container()
    with panel:
        st.subheader("👁️ Live Preview")
        if preview_ok:
            if len(file_bytes) > 1:
                preview_name = st.selectbox("File", list(file_bytes.keys()), key="pv_file",
                                            on_change=_pv_file_changed, label_visibility="collapsed")
            else:
                preview_name = next(iter(file_bytes))
            n_pages = pdf_page_count(file_bytes[preview_name])
            if st.session_state["pv_page"] > n_pages:
                st.session_state["pv_page"] = 1
            if n_pages > 1:
                n1, n2, n3, n4 = st.columns([1, 2, 1, 3])
                with n1:
                    st.button("◀", key="pv_prev", on_click=_pv_step, args=(-1, n_pages),
                              disabled=st.session_state["pv_page"] <= 1)
                with n2:
                    st.number_input("Page", min_value=1, max_value=n_pages, step=1, key="pv_page",
                                    label_visibility="collapsed")
                with n3:
                    st.button("▶", key="pv_next", on_click=_pv_step, args=(1, n_pages),
                              disabled=st.session_state["pv_page"] >= n_pages)
                with n4:
                    st.caption(f"of {n_pages} pages")
            preview_page = int(st.session_state["pv_page"]) - 1
        else:
            st.warning("Preview, stamp removal and white-out need PyMuPDF, which isn't available to "
                       "the Python running this app. Add `pymupdf` to requirements.txt (or run the "
                       "command below locally) and restart the app.")
            st.code(f'"{sys.executable}" -m pip install pymupdf numpy', language="bash")
            if FITZ_ERROR:
                st.caption(f"Import error: {FITZ_ERROR}")

    # ── left panel: edits in tabs ────────────────────────────────────────
    with left:
        # Tab labels must never change, otherwise Streamlit jumps back to the first tab
        tab_logo, tab_wo, tab_col, tab_hdr, tab_src = st.tabs([
            "🧩 Logos", "⬜ White-out", "🎨 Stamps", "🏷️ Header", "🔎 Source check"])

        # Logos & watermarks ------------------------------------------------
        remove_gs = False
        with tab_logo:
            st.markdown("<div class='sn-hint'>Logos and watermarks stored as separate objects in the "
                        "PDF. Ticked items are removed cleanly — the text stays selectable.</div>",
                        unsafe_allow_html=True)
            if not scan_results:
                st.info("No embedded logos or watermarks found. A stamp or logo that's part of a "
                        "scanned page can be removed in the **Stamps** or **White-out** tab.")
            else:
                all_names = [it["name"] for it in scan_results]
                b1, b2, _ = st.columns([1, 1, 2])
                with b1:
                    _stretch_call(st.button, "☑ Tick all", key="ovl_all",
                                  on_click=_set_all_overlays, args=(all_names, True))
                with b2:
                    _stretch_call(st.button, "☐ Untick all", key="ovl_none",
                                  on_click=_set_all_overlays, args=(all_names, False))

                COLS = 3
                remove_set = set()
                for row_start in range(0, len(scan_results), COLS):
                    cols = st.columns(COLS)
                    for col, item in zip(cols, scan_results[row_start: row_start + COLS]):
                        with col:
                            name, subtype, pages, thumb = (item["name"], item["subtype"],
                                                           item["pages"], item["thumbnail"])
                            with st.container(border=True):
                                if thumb:
                                    _stretch_call(st.image, thumb)
                                else:
                                    icon, label = ("📐", "Vector / Form") if subtype == "Form" else ("❓", subtype)
                                    st.markdown(
                                        f"<div style='background:rgba(128,128,128,0.08);border-radius:6px;"
                                        f"padding:22px 0;text-align:center;font-size:26px'>{icon}<br>"
                                        f"<small style='font-size:11px;color:#777'>{label}</small></div>",
                                        unsafe_allow_html=True)
                                page_str = (f"p.{pages[0]}" if len(pages) == 1
                                            else f"p.{pages[0]}–{pages[-1]}" if pages == list(range(pages[0], pages[-1] + 1))
                                            else f"{len(pages)} pages")
                                key = f"chk_{name}"
                                if key not in st.session_state:
                                    st.session_state[key] = name in st.session_state["wm_remove_set"]
                                if st.checkbox(f"Remove · {page_str}", key=key, help=f"{name} · {subtype}"):
                                    remove_set.add(name)
                st.session_state["wm_remove_set"] = remove_set
                with st.expander("Advanced"):
                    remove_gs = st.checkbox(
                        "Also strip `gs` (graphic state) operators — can remove stubborn transparent "
                        "watermarks, but may affect page rendering", value=False, key="wm_remove_gs")

        names_to_remove = tuple(sorted(st.session_state["wm_remove_set"])) if scan_results else ()

        # Preview base image (after overlays + white-out)
        if preview_ok:
            base_bytes = remove_selected_bytes(file_bytes[preview_name], names_to_remove, remove_gs)
            base_img, page_w_pt, page_h_pt = render_page(base_bytes, preview_page)
            page_boxes = areas_for(st.session_state["wo_areas"], preview_name, preview_page)
            wo_img = whiteout_image(base_img, page_boxes)

        draw_ok = preview_ok and streamlit_image_coordinates is not None

        # White-out ---------------------------------------------------------
        with tab_wo:
            st.markdown("<div class='sn-hint'>Cover any text, logo or image with white. What's under "
                        "the box is deleted from the PDF, not just hidden.</div>", unsafe_allow_html=True)
            if not preview_ok:
                st.caption("Needs PyMuPDF — see the message in the preview panel.")
            elif not draw_ok:
                st.warning("Drawing on the preview needs `streamlit-image-coordinates` in "
                           "requirements.txt. Restart the app after adding it.")
            else:
                st.markdown("**How:** in the preview, switch to **⬜ White out**, then drag a box.")
            wo_scope = st.radio(
                "New boxes apply to",
                ["This page only", "All pages of this file", "All pages of all files"],
                horizontal=True, key="wo_scope",
                help="'All pages of all files' is handy for a logo in the same spot on every document.")
            if st.session_state["wo_areas"]:
                st.markdown("**Boxes**")
                for i, a in enumerate(list(st.session_state["wo_areas"])):
                    where = (f"{a['file']} · p.{a['page'] + 1}" if a["file"] != "*" and a["page"] != "*"
                             else f"{a['file']} · all pages" if a["file"] != "*"
                             else "All files · all pages")
                    x0, y0, x1, y1 = a["box"]
                    c1, c2 = st.columns([6, 1])
                    with c1:
                        st.markdown(f"<span class='sn-chip'>#{i + 1}</span> {where} "
                                    f"<span style='color:#888;font-size:0.82rem'>· "
                                    f"{int((x1 - x0) * 100)}% × {int((y1 - y0) * 100)}% of page</span>",
                                    unsafe_allow_html=True)
                    with c2:
                        if st.button("✕", key=f"wo_del_{i}", help="Delete this box"):
                            st.session_state["wo_areas"].pop(i)
                            st.rerun()
                if st.button("Clear all boxes", key="wo_clear"):
                    st.session_state["wo_areas"] = []
                    st.rerun()

        # Stamp colours -----------------------------------------------------
        color_params = (25, 10, 75, 55)
        raster_dpi = 300
        with tab_col:
            st.markdown("<div class='sn-hint'>Removes coloured ink — red QC seals, blue/violet stamps, "
                        "signatures — while keeping black text underneath.</div>", unsafe_allow_html=True)
            if not preview_ok:
                st.caption("Needs PyMuPDF — see the message in the preview panel.")
            else:
                detected = [rgb_to_hex(c) for c in detect_ink_colors(wo_img)]
                detected = [h for h in detected if h not in st.session_state["cr_colors"]]
                st.markdown("**1. Add the stamp's colour**")
                if detected:
                    st.caption("Colours found on this page:")
                    dcols = st.columns(min(len(detected), 4))
                    for i, hx in enumerate(detected):
                        with dcols[i % len(dcols)]:
                            st.markdown(swatch_html(hx), unsafe_allow_html=True)
                            if st.button("＋ Add", key=f"cr_add_det_{hx}_{i}"):
                                st.session_state["cr_colors"].append(hx)
                                st.rerun()
                else:
                    st.caption("No other coloured ink found on this page.")
                m1, m2 = st.columns([1, 2])
                with m1:
                    manual = st.color_picker("Or pick any colour", "#7a4fc0", key="cr_manual")
                with m2:
                    st.write("")
                    st.write("")
                    if st.button("＋ Add this colour", key="cr_add_manual"):
                        if manual not in st.session_state["cr_colors"]:
                            st.session_state["cr_colors"].append(manual)
                        st.rerun()
                if draw_ok:
                    st.caption("…or switch the preview to **🎨 Pick colour** and click on the stamp.")

                st.markdown("**2. Colours being removed**")
                if st.session_state["cr_colors"]:
                    scols = st.columns(min(len(st.session_state["cr_colors"]), 4))
                    for i, hx in enumerate(list(st.session_state["cr_colors"])):
                        with scols[i % len(scols)]:
                            st.markdown(swatch_html(hx), unsafe_allow_html=True)
                            if st.button("✕ Remove", key=f"cr_del_{hx}_{i}"):
                                st.session_state["cr_colors"].remove(hx)
                                st.rerun()
                    st.caption("ℹ️ With stamp removal on, pages are saved as images, so their text "
                               "won't be selectable.")
                else:
                    st.caption("None yet — stamp removal is off.")

                with st.expander("Fine-tune — if traces of the stamp remain"):
                    s1, s2 = st.columns(2)
                    with s1:
                        hue_tol = st.slider("Colour range", 1, 90, 25, key="cr_hue",
                                            help="Raise if parts of the stamp stay; lower if other colours disappear.")
                        whiten = st.slider("Whiten %", 30, 100, 75, key="cr_whiten",
                                           help="Lower if a faint ghost of the letters remains; "
                                                "raise if text or lines under the stamp fade.")
                    with s2:
                        sat_min = st.slider("Min strength %", 1, 80, 10, key="cr_sat",
                                            help="Lower to also catch pale, washed-out stamp ink.")
                        edge_clean = st.slider("Edge cleanup %", 20, 95, 55, key="cr_edge",
                                               help="Lower to remove grey outlines around the stamp letters; "
                                                    "raise if nearby text gets thinner.")
                    raster_dpi = st.select_slider("Output resolution (dpi)", options=[150, 200, 300, 400],
                                                  value=300, key="cr_dpi",
                                                  help="Higher = sharper but larger files.")
                    color_params = (hue_tol, sat_min, whiten, edge_clean)

        active_colors = [hex_to_rgb(h) for h in st.session_state["cr_colors"]] if preview_ok else []

        # Header ------------------------------------------------------------
        with tab_hdr:
            header_exists = os.path.isfile(HEADER_IMAGE_PATH)
            st.markdown("<div class='sn-hint'>Places the SN company header at the top of every page.</div>",
                        unsafe_allow_html=True)
            if not header_exists:
                st.warning("⚠️ `SN_Header.png` not found next to `app.py`. Add it to the repo to enable this.")
            add_header = st.toggle("Add SN company header", value=False,
                                   disabled=not header_exists, key="wm_add_header")
            img_width_pct, top_margin = 98, 10
            if add_header and header_exists:
                hcol1, hcol2 = st.columns(2)
                with hcol1:
                    img_width_pct = st.slider("Width (% of page)", 20, 100, 98, 1, key="wm_header_width")
                with hcol2:
                    top_margin = st.number_input("Top margin (pt)", 0, 100, 10, 1, key="wm_top_margin")

        # Source check ------------------------------------------------------
        with tab_src:
            st.markdown("<div class='sn-hint'>Flags spots that may reveal the manufacturer or supplier "
                        "(company names, “manufactured by”, phone/fax, email, website, address, logos) "
                        "so you can review before sending. Nothing is changed automatically.</div>",
                        unsafe_allow_html=True)
            sc1, sc2 = st.columns(2)
            with sc1:
                ignore_text = st.text_area("Our own names / address (ignored)",
                                           "Skyherb, Loker Ave, Carlsbad, CA92010", height=90,
                                           key="si_ignore", help="Comma-separated.")
            with sc2:
                watch_text = st.text_area("Supplier names to look for", "", height=90, key="si_watch",
                                          placeholder="e.g. Xinhua, Dae Hyung",
                                          help="Comma-separated. Always flagged wherever they appear.")
            ignore_terms, watch_terms = parse_terms(ignore_text), parse_terms(watch_text)
            si_highlight = st.toggle("Highlight flagged spots on the preview", value=True, key="si_hl")
            if pytesseract is None:
                st.caption("Scanned pages can only be checked by eye. To check them automatically, add "
                           "`pytesseract` to requirements.txt and `tesseract-ocr` to packages.txt.")

    # ── right panel, part 2: live result + export ────────────────────────
    with panel:
        changes = []
        if names_to_remove: changes.append(f"🧩 {len(names_to_remove)} logo(s) removed")
        if preview_ok and st.session_state["wo_areas"]:
            changes.append(f"⬜ {len(st.session_state['wo_areas'])} white-out box(es)")
        if active_colors: changes.append(f"🎨 {len(active_colors)} stamp colour(s)")
        if add_header and header_exists: changes.append("🏷️ header")

        if preview_ok:
            after_img = wo_img
            if active_colors:
                after_img = remove_colors(after_img, active_colors, *color_params)
            if add_header and header_exists:
                after_img = paste_header(after_img, page_w_pt, PREVIEW_DPI, img_width_pct / 100, top_margin)

            raster_for_ocr = None
            if pytesseract is not None:
                raster_for_ocr = remove_colors(wo_img, active_colors, *color_params) if active_colors else wo_img
            areas_key = tuple((a["file"], a["page"], tuple(a["box"])) for a in st.session_state["wo_areas"])
            try:
                pv_report = preview_check(base_bytes, preview_name, areas_key, preview_page,
                                          ignore_terms, watch_terms, raster_for_ocr)
            except Exception:
                pv_report = None
            if pv_report and si_highlight:
                after_img = highlight_spots(after_img, pv_report["spots"])

            # interaction mode
            modes = ["👁️ View", "⬜ White out (drag)", "🎨 Pick colour (click)"] if draw_ok else ["👁️ View"]
            mode = modes[0]
            if draw_ok:
                if hasattr(st, "segmented_control"):
                    mode = st.segmented_control("Mode", modes, default=modes[0], key="pv_mode",
                                                label_visibility="collapsed") or modes[0]
                else:
                    mode = st.radio("Mode", modes, horizontal=True, key="pv_mode",
                                    label_visibility="collapsed")

            st.markdown(chips(changes), unsafe_allow_html=True)

            if mode == modes[0]:
                _stretch_call(st.image, after_img)
            else:
                import inspect
                _sig = str(inspect.signature(streamlit_image_coordinates))
                if "stretch" in _sig:
                    _kw = {"width": "stretch"}
                elif "use_column_width" in _sig:
                    _kw = {"use_column_width": "always"}
                else:
                    _kw = {"width": PICK_WIDTH}
                drag = mode.startswith("⬜")
                st.caption("Drag a box over the area to white out." if drag
                           else "Click directly on the stamp's coloured ink.")
                event = streamlit_image_coordinates(
                    after_img, key=f"pv_{'drag' if drag else 'pick'}_{preview_name}_{preview_page}",
                    click_and_drag=drag, cursor="crosshair", **_kw)

                if event and event.get("unix_time") != st.session_state["pv_last_event"]:
                    st.session_state["pv_last_event"] = event.get("unix_time")
                    shown_w = event.get("width") or after_img.width
                    shown_h = event.get("height") or after_img.height
                    if drag and "x1" in event:
                        fx = sorted([event["x1"] / shown_w, event["x2"] / shown_w])
                        fy = sorted([event["y1"] / shown_h, event["y2"] / shown_h])
                        box = (max(0.0, fx[0]), max(0.0, fy[0]), min(1.0, fx[1]), min(1.0, fy[1]))
                        if (box[2] - box[0]) * shown_w >= 4 and (box[3] - box[1]) * shown_h >= 4:
                            st.session_state["wo_areas"].append({
                                "file": "*" if wo_scope == "All pages of all files" else preview_name,
                                "page": preview_page if wo_scope == "This page only" else "*",
                                "box": box,
                            })
                            st.rerun()
                        else:
                            st.toast("Drag to draw a box — a single click doesn't white out anything.")
                    elif not drag and "x" in event:
                        sx = wo_img.width / shown_w
                        picked = pick_color_at(wo_img, int(event["x"] * sx), int(event["y"] * sx))
                        if picked is None:
                            st.toast("That spot isn't coloured ink — click directly on the stamp.")
                        else:
                            hx = rgb_to_hex(picked)
                            if hx not in st.session_state["cr_colors"]:
                                st.session_state["cr_colors"].append(hx)
                                st.toast(f"Added stamp colour {hx}")
                                st.rerun()

            if pv_report:
                n_spots = len(pv_report["spots"])
                notes = []
                if n_spots:
                    notes.append(f"**{n_spots} spot(s)** may show company/source info"
                                 + (" (orange boxes)" if si_highlight else ""))
                if pv_report["images"]:
                    notes.append(f"{pv_report['images']} image(s) remain — check none is a supplier logo")
                if pv_report["scanned"] and not pv_report["ocr"]:
                    notes.append("scanned page — review text by eye")
                if notes:
                    st.warning("🔎 " + " · ".join(notes))
                    if n_spots:
                        with st.expander("Show flagged spots"):
                            for n, spot in enumerate(pv_report["spots"], 1):
                                st.markdown(f"#{n} · *{spot['reason']}*: `{spot['text']}`")
                else:
                    st.success("🔎 No company/source info detected on this page.")
        else:
            st.markdown(chips(changes), unsafe_allow_html=True)

        # ── export ───────────────────────────────────────────────────────
        st.divider()
        st.subheader("📦 Process & Download")

        if st.session_state.get("wm_results_sig") != upload_sig:
            st.session_state["wm_results"] = None

        if not changes:
            st.caption("No edits selected yet — processing would return the files unchanged.")
        if _stretch_call(st.button, f"⚙️ Process {len(file_bytes)} file(s)", type="primary", key="wm_process"):
            results, errors, reports = [], [], {}
            progress = st.progress(0, text="Processing…")
            for i, (fname, data) in enumerate(file_bytes.items()):
                try:
                    output_name = f"SN-{fname.replace(' ', '_')}"
                    out_bytes, text_stage, image_stage = process_pdf(
                        data, names_to_remove, remove_gs,
                        active_colors, color_params, raster_dpi,
                        add_header and header_exists, img_width_pct / 100, top_margin,
                        fname=fname, areas=st.session_state["wo_areas"] if preview_ok else [],
                        return_stages=True,
                    )
                    with open(f"outputs/{output_name}", "wb") as f:
                        f.write(out_bytes)
                    results.append((output_name, out_bytes))
                    if preview_ok:
                        try:
                            reports[output_name] = check_pdf(text_stage, ignore_terms, watch_terms,
                                                             image_stage if active_colors else None)
                        except Exception:
                            pass
                except Exception as e:
                    errors.append((fname, str(e)))
                progress.progress((i + 1) / len(file_bytes), text=f"Processed {i + 1}/{len(file_bytes)}")
            progress.empty()

            zip_bytes = None
            if len(results) > 1:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for output_name, out_bytes in results:
                        zf.writestr(output_name, out_bytes)
                zip_bytes = zip_buf.getvalue()
            st.session_state["wm_results"] = {"files": results, "errors": errors, "ops": list(changes),
                                              "zip": zip_bytes, "reports": reports}
            st.session_state["wm_results_sig"] = upload_sig

        res = st.session_state.get("wm_results")
        if res:
            for name, err in res["errors"]:
                st.error(f"❌ {name}: {err}")
            files = res["files"]
            if files:
                st.success(f"✅ {len(files)} file(s) ready" +
                           (f" — {', '.join(res['ops'])}" if res["ops"] else "") + ".")

                def _download(label, data, file_name, mime, key, stretch=False, **kw):
                    # on_click="ignore" (Streamlit 1.43+) stops the page reloading on download
                    base = dict(file_name=file_name, mime=mime, key=key, **kw)
                    for extra in ({"on_click": "ignore"}, {}):
                        try:
                            if stretch:
                                return _stretch_call(st.download_button, label, data, **base, **extra)
                            return st.download_button(label, data, **base, **extra)
                        except TypeError:
                            continue

                if res["zip"]:
                    _download(f"⬇️ Download all ({len(files)} files) as ZIP", res["zip"],
                              "SN_processed_pdfs.zip", "application/zip", "dl_wm_all_zip",
                              stretch=True, type="primary")
                    with st.expander("Download files individually"):
                        for output_name, out_bytes in files:
                            _download(f"⬇️ {output_name}", out_bytes, output_name,
                                      "application/pdf", f"dl_{output_name}")
                else:
                    output_name, out_bytes = files[0]
                    _download(f"⬇️ Download {output_name}", out_bytes, output_name,
                              "application/pdf", f"dl_{output_name}", stretch=True, type="primary")

                reports = res.get("reports") or {}
                flagged = []
                for out_name, rep_ in reports.items():
                    total, parts, imgs, scanned = summarize_report(rep_)
                    bits = []
                    if total: bits.append(f"{total} spot(s) ({', '.join(parts)})")
                    if imgs: bits.append(f"image(s) to check ({', '.join(imgs)})")
                    if scanned: bits.append(f"scanned p.{', '.join(scanned)} — check by eye")
                    if bits:
                        flagged.append((out_name, bits, rep_))
                if flagged:
                    st.warning(f"🔎 Review before sending — {len(flagged)} file(s) may show source info:")
                    for out_name, bits, rep_ in flagged:
                        with st.expander(f"{out_name} — " + "; ".join(bits)):
                            any_spot = False
                            for pno, r in enumerate(rep_):
                                for n, spot in enumerate(r["spots"], 1):
                                    any_spot = True
                                    st.markdown(f"p.{pno + 1} · #{n} · *{spot['reason']}*: `{spot['text']}`")
                            if not any_spot:
                                st.caption("No flagged text — only images or scanned pages to check.")
                elif reports:
                    st.info("🔎 No company/source info detected in the processed files.")


# ==============================
# PAGE — COA Review
# ==============================

elif page == "🔬 COA Review":
    import hashlib
    import json
    import pandas as pd
    import coa_review as cr

    def _secret(name):
        try:
            return st.secrets.get(name)
        except Exception:
            return None

    API_KEY = _secret("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")
    MODEL = _secret("COA_MODEL") or os.environ.get("COA_MODEL") or cr.DEFAULT_MODEL
    SPEC_LIB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "spec_library.xlsx")
    AUTO = "Auto-detect"
    SEV_ICON = {"High": "🔴 High", "Medium": "🟠 Medium", "Low": "🔵 Low"}

    for _k, _v in {"cr_results": {}, "cr_custom_rows": [], "cr_custom_name": "Customer spec"}.items():
        if _k not in st.session_state:
            st.session_state[_k] = _v
    if "cr_lib_bytes" not in st.session_state:
        st.session_state["cr_lib_bytes"] = open(SPEC_LIB_PATH, "rb").read() if os.path.isfile(SPEC_LIB_PATH) else None
        st.session_state["cr_lib_source"] = "spec_library.xlsx in the app" if st.session_state["cr_lib_bytes"] else ""

    @st.cache_data(show_spinner=False, max_entries=8)
    def _load_lib(data):
        return cr.load_spec_library(data)

    @st.cache_data(show_spinner=False, max_entries=64)
    def _read_text(data):
        return cr.extract_text_coa(data)

    @st.cache_data(show_spinner=False, max_entries=64)
    def _read_ai(data, filename, api_key, model):
        return cr.ai_extract_coa(cr.make_client(api_key), model, data, filename)

    @st.cache_data(show_spinner=False, max_entries=64)
    def _judge_ai(coa_json, spec_key, prod_json, rules_json, api_key, model, _coa, _rows, _prod, _rules):
        return cr.ai_review(cr.make_client(api_key), model, _coa, _rows, _prod, _rules)

    def _lib_to_bytes(specs_df, products_df):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            specs_df.to_excel(xw, sheet_name="Specs", index=False)
            products_df.to_excel(xw, sheet_name="Products", index=False)
        return buf.getvalue()

    specs_df = products_df = None
    lib_error = ""
    if st.session_state["cr_lib_bytes"]:
        try:
            specs_df, products_df = _load_lib(st.session_state["cr_lib_bytes"])
        except Exception as e:
            lib_error = str(e)
    product_names = list(products_df["Product"]) if products_df is not None else []

    # ── header ────────────────────────────────────────────────────────────
    h1, h2 = st.columns([5, 1])
    with h1:
        st.title("🔬 COA Review")
        st.markdown("<div class='sn-hint'>Upload supplier COAs → pick the spec → get a list of problems and a "
                    "follow-up email for the supplier.</div>", unsafe_allow_html=True)
    with h2:
        st.write("")
        if st.button("↺ Clear results", key="cr_clear"):
            st.session_state["cr_results"] = {}
            for k in [k for k in st.session_state.keys() if str(k).startswith(("cr_iss_", "cr_mail_", "cr_subj_"))]:
                del st.session_state[k]
            st.rerun()

    _chips = ([f"🤖 AI on · {MODEL}"] if API_KEY else []) + \
             ([f"📚 {len(product_names)} product spec(s)"] if product_names else [])
    if _chips:
        st.markdown(chips(_chips), unsafe_allow_html=True)
    if not API_KEY:
        st.caption("🤖 AI is off: text-based PDFs are checked by the rules; scanned COAs and judgement checks "
                   "(methods, vague wording, product match) need an `ANTHROPIC_API_KEY` in the app's secrets.")
    if lib_error:
        st.error(f"Couldn't read the spec library: {lib_error}")

    tab_review, tab_lib = st.tabs(["🔍 Review", "📚 Spec library"])

    # ════════════════════════════ Spec library tab ═══════════════════════
    with tab_lib:
        st.markdown("<div class='sn-hint'>One row per product per test. Changes made here last for this "
                    "session — download the file and replace <code>spec_library.xlsx</code> in GitHub to keep "
                    "them.</div>", unsafe_allow_html=True)
        st.caption(f"Using: {st.session_state.get('cr_lib_source') or 'no library loaded'}")
        up = st.file_uploader("Use a different library (.xlsx)", type="xlsx", key="cr_lib_upload")
        if up is not None and st.session_state.get("cr_lib_upload_sig") != (up.name, up.size):
            try:
                _load_lib(up.getvalue())
                st.session_state["cr_lib_bytes"] = up.getvalue()
                st.session_state["cr_lib_source"] = f"{up.name} (uploaded this session)"
                st.session_state["cr_lib_upload_sig"] = (up.name, up.size)
                st.rerun()
            except Exception as e:
                st.error(f"That file doesn't look like a spec library: {e}")

        if specs_df is not None:
            prod_filter = st.selectbox("Show product", ["All products"] + product_names, key="cr_lib_filter")
            shown = specs_df if prod_filter == "All products" else specs_df[specs_df["Product"] == prod_filter]
            st.markdown("**Specs**")
            edited_specs = st.data_editor(
                shown, num_rows="dynamic", key=f"cr_lib_specs_{prod_filter}", hide_index=True,
                column_config={
                    "Product": st.column_config.SelectboxColumn(options=product_names, required=True),
                    "Category": st.column_config.SelectboxColumn(options=[
                        "Assay", "Identification", "Physical", "Heavy metals", "Microbiology",
                        "Residual solvents", "Pesticides", "Contaminants"]),
                    "Min": st.column_config.NumberColumn(format="%g"),
                    "Max": st.column_config.NumberColumn(format="%g"),
                    "Required": st.column_config.SelectboxColumn(options=["Y", "N"]),
                })
            st.markdown("**Products**")
            edited_products = st.data_editor(products_df, num_rows="dynamic", key="cr_lib_products",
                                             hide_index=True)
            c1, c2 = st.columns(2)
            with c1:
                if st.button("💾 Apply changes for this session", key="cr_lib_apply"):
                    others = specs_df if prod_filter == "All products" else specs_df[specs_df["Product"] != prod_filter]
                    merged = edited_specs if prod_filter == "All products" else pd.concat([others, edited_specs])
                    st.session_state["cr_lib_bytes"] = _lib_to_bytes(merged.dropna(subset=["Product", "Test"]),
                                                                     edited_products.dropna(subset=["Product"]))
                    st.session_state["cr_lib_source"] = "edited in this session — download to keep"
                    st.rerun()
            with c2:
                _stretch_call(st.download_button, "⬇️ Download spec_library.xlsx", st.session_state["cr_lib_bytes"],
                              file_name="spec_library.xlsx",
                              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              key="cr_lib_dl")
        else:
            st.info("No spec library yet. Add `spec_library.xlsx` next to `app.py` in GitHub, or upload one above.")

    # ════════════════════════════ Review tab ═════════════════════════════
    with tab_review:
        left, right = st.columns([1.1, 1], gap="large")

        with left:
            st.subheader("① Supplier COAs")
            coa_files = st.file_uploader("PDF or scanned images — several at once is fine",
                                         type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True,
                                         key="cr_upload")

        with right:
            st.subheader("② Specification")
            spec_mode = st.radio("Compare against", ["Spec library", "Customer spec file"], horizontal=True,
                                 key="cr_spec_mode", label_visibility="collapsed")
            if spec_mode == "Customer spec file":
                spec_file = st.file_uploader("Customer specification (PDF, image or .xlsx)",
                                             type=["pdf", "png", "jpg", "jpeg", "xlsx"], key="cr_spec_upload")
                if spec_file is not None and st.session_state.get("cr_spec_sig") != (spec_file.name, spec_file.size):
                    with st.spinner("Reading the specification…"):
                        rows, pname, err = [], "", ""
                        try:
                            if spec_file.name.lower().endswith(".xlsx"):
                                sdf, _ = cr.load_spec_library(spec_file.getvalue())
                                rows = sdf.drop(columns=["Product"]).to_dict("records")
                                pname = str(sdf["Product"].iloc[0]) if len(sdf) else ""
                            elif API_KEY:
                                pname, rows = cr.ai_extract_spec(cr.make_client(API_KEY), MODEL,
                                                                 spec_file.getvalue(), spec_file.name)
                            elif spec_file.name.lower().endswith(".pdf"):
                                rows = cr.text_spec_rows(spec_file.getvalue())
                            else:
                                err = "Reading an image needs the AI (add an API key), or upload a PDF/.xlsx."
                        except Exception as e:
                            err = str(e)
                    st.session_state["cr_custom_rows"] = rows
                    st.session_state["cr_custom_name"] = pname or os.path.splitext(spec_file.name)[0]
                    st.session_state["cr_spec_sig"] = (spec_file.name, spec_file.size)
                    st.session_state.pop("cr_custom_editor", None)
                    if err:
                        st.error(err)
                    elif not rows:
                        st.warning("No test rows found — add them in the table below.")
                st.text_input("Product name", key="cr_custom_name")
                base_rows = st.session_state["cr_custom_rows"] or [
                    {c: None for c in cr.SPEC_COLUMNS if c != "Product"}]
                custom_df = st.data_editor(
                    pd.DataFrame(base_rows, columns=[c for c in cr.SPEC_COLUMNS if c != "Product"]),
                    num_rows="dynamic", hide_index=True, key="cr_custom_editor",
                    column_config={"Min": st.column_config.NumberColumn(format="%g"),
                                   "Max": st.column_config.NumberColumn(format="%g"),
                                   "Required": st.column_config.SelectboxColumn(options=["Y", "N"])})
                st.caption("Check the limits read from the file before reviewing.")
                if specs_df is not None and st.button("➕ Add this spec to the library (this session)",
                                                      key="cr_custom_save"):
                    name = st.session_state["cr_custom_name"].strip() or "Customer spec"
                    add = custom_df.dropna(subset=["Test"]).assign(Product=name)[cr.SPEC_COLUMNS]
                    new_specs = pd.concat([specs_df[specs_df["Product"] != name], add], ignore_index=True)
                    new_prods = products_df if name in product_names else pd.concat(
                        [products_df, pd.DataFrame([{"Product": name}])], ignore_index=True)
                    st.session_state["cr_lib_bytes"] = _lib_to_bytes(new_specs, new_prods)
                    st.session_state["cr_lib_source"] = "edited in this session — download to keep"
                    st.rerun()
            elif not product_names:
                st.warning("No spec library loaded — see the 📚 Spec library tab, or use a customer spec file.")
            else:
                st.caption("Each COA is matched to a product automatically by its name. Override it in the "
                           "table below if needed.")

        if not coa_files:
            st.info("⬆️ Upload one or more COAs to start.")
            st.stop()

        file_bytes = {f.name: f.getvalue() for f in coa_files}

        # per-file spec choice
        if spec_mode == "Spec library" and product_names:
            plan = pd.DataFrame({"COA file": list(file_bytes),
                                 "Spec": [st.session_state.get("cr_choice", {}).get(n, AUTO) for n in file_bytes]})
            plan = st.data_editor(plan, hide_index=True, key="cr_plan", disabled=["COA file"],
                                  column_config={"Spec": st.column_config.SelectboxColumn(
                                      options=[AUTO] + product_names, required=True)})
            st.session_state["cr_choice"] = dict(zip(plan["COA file"], plan["Spec"]))

        o1, o2, o3 = st.columns([1, 1.4, 1.2])
        with o1:
            use_ai = st.toggle("Use AI", value=bool(API_KEY), disabled=not API_KEY, key="cr_use_ai",
                               help="Reads scans and adds judgement checks. COAs are sent to the Claude API.")
        with o2:
            supplier_name = st.text_input("Supplier name for the email (optional)", key="cr_supplier")
        with o3:
            st.write("")
            go = _stretch_call(st.button, f"🔍 Review {len(file_bytes)} COA(s)", type="primary", key="cr_go")

        # ── run ───────────────────────────────────────────────────────────
        if go:
            results = {}
            prog = st.progress(0, text="Reviewing…")
            for i, (fname, data) in enumerate(file_bytes.items()):
                prog.progress(i / len(file_bytes), text=f"Reading {fname}…")
                rec = {"file": fname, "error": "", "warnings": [], "source": ""}
                is_pdf = fname.lower().endswith(".pdf")
                coa = None
                if is_pdf:
                    try:
                        coa = _read_text(data)
                    except Exception as e:
                        rec["warnings"].append(f"Text reading failed: {e}")
                if use_ai and API_KEY:
                    try:
                        coa = _read_ai(data, fname, API_KEY, MODEL)
                    except Exception as e:
                        rec["warnings"].append(f"AI reading failed ({e}); used the text reader instead.")
                if coa is None or (not coa["tests"] and coa.get("scanned", not is_pdf)):
                    rec["error"] = ("This looks like a scanned COA (no text to read). Turn on AI to read it."
                                    if not (use_ai and API_KEY) else "No test results could be read from this file.")
                    rec["short"] = "Scanned — needs AI" if not (use_ai and API_KEY) else "Couldn't read results"
                    rec.update(info=(coa or {}).get("info", {}), coa=coa or {"info": {}, "tests": []},
                               product="", result={"tests": [], "findings": [], "counts": {}}, findings=[])
                    results[fname] = rec
                    continue
                rec["source"] = "AI" if coa["source"] == "ai" else "Text"

                # which spec
                if spec_mode == "Customer spec file":
                    pname = st.session_state["cr_custom_name"].strip() or "Customer spec"
                    cdf = custom_df.dropna(subset=["Test"]).assign(Product=pname)
                    spec_rows = cr.spec_rows_for(cr._norm_cols(cdf, cr.SPEC_COLUMNS), pname)
                    pinfo = {}
                else:
                    choice = st.session_state.get("cr_choice", {}).get(fname, AUTO)
                    pname = choice
                    if choice == AUTO:
                        pname, score = cr.match_product(coa["info"].get("product_name", ""), products_df) \
                            if products_df is not None else (None, 0)
                    spec_rows = cr.spec_rows_for(specs_df, pname) if pname else []
                    pinfo = cr.product_info(products_df, pname) if pname else {}
                if not spec_rows:
                    rec["error"] = (f"No matching spec for '{coa['info'].get('product_name', '?')}'. Pick the "
                                    "product in the table above and review again.")
                    rec["short"] = "No matching spec — pick one"
                prog.progress((i + 0.5) / len(file_bytes), text=f"Checking {fname}…")
                result = cr.review_coa(coa, spec_rows, pinfo, pname or "")
                findings = result["findings"]
                if use_ai and API_KEY and spec_rows:
                    try:
                        ai = _judge_ai(json.dumps(coa, sort_keys=True, default=str), pname,
                                       json.dumps(pinfo, sort_keys=True),
                                       json.dumps(findings, sort_keys=True), API_KEY, MODEL,
                                       coa, spec_rows, pinfo, findings)
                        findings = cr.merge_ai(findings, ai, coa.get("warnings", []))
                    except Exception as e:
                        rec["warnings"].append(f"AI checks failed: {e}")
                rec.update(info=coa["info"], coa=coa, product=pname or "", result=result, findings=findings)
                results[fname] = rec
            prog.empty()
            st.session_state["cr_results"] = results
            for k in [k for k in st.session_state.keys() if str(k).startswith(("cr_iss_", "cr_mail_", "cr_subj_"))]:
                del st.session_state[k]

        results = {k: v for k, v in st.session_state["cr_results"].items() if k in file_bytes}
        if not results:
            st.stop()

        # ── ③ results ─────────────────────────────────────────────────────
        st.divider()
        st.subheader("③ Results")

        def _included(rec):
            df = st.session_state.get(f"cr_iss_{rec['file']}")
            if isinstance(df, pd.DataFrame):
                df = df[df["Include"].fillna(False).astype(bool)].drop(columns=["Include"])
                rows = df.astype(object).where(df.notna(), "").to_dict("records")
                for r in rows:
                    r["Severity"] = str(r.get("Severity") or "Low").split()[-1]
                    r["Status"] = r.get("Status") or "Review"
                return [r for r in rows if str(r.get("Item", "")).strip() or str(r.get("Detail", "")).strip()]
            return rec["findings"]

        summary_slot = st.container()
        sel = st.selectbox("Details for", list(results), key="cr_sel") if len(results) > 1 else next(iter(results))

        def _fill_summary():
            summary = []
            for rec in results.values():
                inc = _included(rec)
                summary.append({
                    "COA file": rec["file"], "Spec": rec["product"] or "—", "Lot": rec["info"].get("lot_no", ""),
                    "Verdict": ("⚠️ " + rec.get("short", rec["error"])) if rec["error"] else
                    ("❌ " if any(f["Severity"] == "High" for f in inc) else
                     "⚠️ " if any(f["Severity"] == "Medium" for f in inc) else "✅ ") + cr.verdict(inc),
                    "High": sum(f["Severity"] == "High" for f in inc),
                    "Medium": sum(f["Severity"] == "Medium" for f in inc),
                    "Low": sum(f["Severity"] == "Low" for f in inc), "Read by": rec["source"],
                })
            with summary_slot:
                _stretch_call(st.dataframe, pd.DataFrame(summary), hide_index=True)
                reviewed = [dict(r, findings=_included(r)) for r in results.values()
                            if r["result"]["tests"] and not r["error"]]
                if len(reviewed) > 1:
                    _stretch_call(st.download_button, f"⬇️ Download all {len(reviewed)} reviews (Excel)",
                                  cr.to_excel(reviewed), file_name=f"COA_review_{datetime.today():%Y-%m-%d}.xlsx",
                                  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  key="cr_dl_all")

        rec = results[sel]
        for w in rec["warnings"]:
            st.warning(w)
        if rec["error"]:
            st.error(rec["error"])
        if not rec["result"]["tests"] and not rec["findings"]:
            _fill_summary()
            st.stop()

        info = rec["info"]
        d1, d2 = st.columns([1.15, 1], gap="large")
        with d1:
            st.markdown(f"**{info.get('product_name') or sel}**  \n"
                        f"Spec: {rec['product'] or '—'} · Lot {info.get('lot_no') or '—'} · "
                        f"Mfg {info.get('mfg_date') or '—'} · Exp {info.get('expiry_date') or info.get('retest_date') or '—'}")
        with d2:
            cnt = rec["result"].get("counts", {})
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("❌ Fail", cnt.get("Fail", 0))
            m2.metric("⚠️ Missing", cnt.get("Missing", 0))
            m3.metric("🔍 Review", cnt.get("Review", 0))
            m4.metric("✅ Pass", cnt.get("Pass", 0))

        t_iss, t_tests, t_mail, t_info = st.tabs(["🚩 Issues", "📋 Test by test", "✉️ Supplier email", "ℹ️ COA details"])

        with t_iss:
            st.markdown("<div class='sn-hint'>Untick anything that isn't a real problem, edit the wording, "
                        "or add a row for something the tool missed (it's kept in the export as “Manual”).</div>",
                        unsafe_allow_html=True)
            key = f"cr_iss_{sel}"
            base = pd.DataFrame(rec["findings"], columns=["Severity", "Status", "Item", "Detail",
                                                          "Question for supplier", "Category", "Source"])
            base.insert(0, "Include", True)
            base["Severity"] = base["Severity"].map(lambda s: SEV_ICON.get(s, s))
            edited = st.data_editor(
                base, num_rows="dynamic", hide_index=True, key=key + "_ed",
                column_order=["Include", "Severity", "Status", "Item", "Detail", "Question for supplier", "Source"],
                column_config={
                    "Include": st.column_config.CheckboxColumn(default=True, width="small"),
                    "Severity": st.column_config.SelectboxColumn(options=list(SEV_ICON.values()),
                                                                 default="🟠 Medium", width="small"),
                    "Status": st.column_config.SelectboxColumn(options=["Fail", "Missing", "Review"],
                                                               default="Review", width="small"),
                    "Detail": st.column_config.TextColumn(width="large"),
                    "Question for supplier": st.column_config.TextColumn(width="large"),
                    "Source": st.column_config.TextColumn(default="Manual", width="small"),
                })
            edited["Source"] = edited["Source"].fillna("Manual")
            edited["Category"] = edited["Category"].fillna("Manual")
            st.session_state[key] = edited
            if edited.empty:
                st.success("No issues found against this specification.")

        inc = _included(rec)
        with t_tests:
            only = st.toggle("Only show problems", value=False, key=f"cr_only_{sel}")
            tdf = pd.DataFrame(rec["result"]["tests"])
            if not tdf.empty:
                if only:
                    tdf = tdf[tdf["Status"].isin(["Fail", "Missing", "Review"])]
                tdf["Status"] = tdf["Status"].map(lambda s: cr.STATUS_ICON.get(s, s))
                _stretch_call(st.dataframe, tdf[["Status", "Test (spec)", "Our limit", "COA result", "Test (COA)",
                                                 "COA spec", "Method", "Note"]], hide_index=True)

        with t_mail:
            subj, body = cr.build_email(info, inc, rec["product"], st.session_state.get("cr_supplier", "")
                                        or info.get("supplier", ""))
            sig = hashlib.md5(json.dumps([f.get("Question for supplier") for f in inc]).encode()).hexdigest()
            if st.session_state.get(f"cr_mailsig_{sel}") != sig:      # issues changed -> rebuild the draft
                st.session_state[f"cr_subj_{sel}"] = subj
                st.session_state[f"cr_mail_{sel}"] = body
                st.session_state[f"cr_mailsig_{sel}"] = sig
            st.text_input("Subject", key=f"cr_subj_{sel}")
            st.text_area("Email", key=f"cr_mail_{sel}", height=320)
            def _polish(f=sel):
                try:
                    st.session_state[f"cr_mail_{f}"] = cr.ai_polish_email(
                        cr.make_client(API_KEY), MODEL, st.session_state[f"cr_subj_{f}"],
                        st.session_state[f"cr_mail_{f}"])
                except Exception as e:
                    st.session_state["cr_polish_error"] = str(e)

            e1, e2 = st.columns(2)
            with e1:
                if API_KEY:
                    st.button("✨ Polish wording with AI", key=f"cr_polish_{sel}", on_click=_polish)
                if st.session_state.get("cr_polish_error"):
                    st.error(f"AI polishing failed: {st.session_state.pop('cr_polish_error')}")
            with e2:
                st.download_button("⬇️ Download email (.txt)",
                                   f"Subject: {st.session_state[f'cr_subj_{sel}']}\n\n{st.session_state[f'cr_mail_{sel}']}",
                                   file_name=f"COA_email_{os.path.splitext(sel)[0]}.txt", key=f"cr_dl_mail_{sel}")
            st.caption("The draft follows the ticked issues in 🚩 Issues. Edits here are kept until the issues change.")

        with t_info:
            _stretch_call(st.dataframe, pd.DataFrame([{"Field": k.replace("_", " ").capitalize(), "Value": v}
                                                      for k, v in info.items()]), hide_index=True)
            signed = rec["coa"].get("signed")
            st.caption(f"Read by: {rec['source']} · Signature: "
                       f"{'yes' if signed else 'not found' if signed is False else 'not checked (check by eye)'}")

        one = dict(rec, findings=inc)
        stem = os.path.splitext(sel)[0]
        x1, x2 = st.columns(2)
        with x1:
            _stretch_call(st.download_button, "⬇️ Report (PDF)", cr.to_pdf(one, f"COA Review — {rec['product'] or stem}"),
                          file_name=f"COA_review_{stem}.pdf", mime="application/pdf", key=f"cr_dl_pdf_{sel}")
        with x2:
            _stretch_call(st.download_button, "⬇️ Report (Excel)", cr.to_excel([one]),
                          file_name=f"COA_review_{stem}.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          key=f"cr_dl_xlsx_{sel}")

        _fill_summary()


# ==============================
# PAGE 3 — PDF Unlocker
# ==============================

elif page == "🔓 PDF Unlocker":

    st.title("🔓 PDF Unlocker")
    st.caption(
        "Remove encryption and permission restrictions from PDFs. "
        "Supports owner-locked files (no password needed) and user-password-protected files."
    )

    # ── STEP 1 — Upload ───────────────────────────────────────────────────────
    st.subheader("① Upload PDF(s)")

    unlock_files = st.file_uploader(
        "Upload one or more locked PDF files",
        type="pdf",
        accept_multiple_files=True,
        key="unlock_uploader"
    )

    if not unlock_files:
        st.info("Upload one or more PDFs to get started.")
        st.stop()

    st.write(f"{len(unlock_files)} file(s) uploaded.")

    # ── STEP 2 — Password (optional) ─────────────────────────────────────────
    st.divider()
    st.subheader("② Password (optional)")
    st.caption(
        "Leave blank if the PDFs are owner-locked only (print/copy restrictions) "
        "with no user-open password. Enter a password only if the file asks for one to open."
    )

    unlock_password = st.text_input(
        "Password",
        type="password",
        placeholder="Leave blank if no open password is required",
        key="unlock_password"
    )

    same_pw_for_all = True
    if len(unlock_files) > 1:
        same_pw_for_all = st.checkbox(
            "Use the same password for all files",
            value=True,
            key="unlock_same_pw"
        )
        if not same_pw_for_all:
            st.info(
                "Per-file passwords are not supported in batch mode. "
                "Enter a shared password above, or process files one at a time."
            )

    # ── STEP 3 — Output naming ────────────────────────────────────────────────
    st.divider()
    st.subheader("③ Output Options")

    prefix = st.text_input(
        "Output filename prefix",
        value="unlocked-",
        help="This prefix is prepended to each original filename.",
        key="unlock_prefix"
    )

    # ── STEP 4 — Unlock ───────────────────────────────────────────────────────
    st.divider()

    if st.button("🔓 Unlock PDF(s)", type="primary"):

        results  = []   # list of (output_name, output_path, status)
        errors   = []   # list of (fname, reason)
        skipped  = []   # list of (fname, "not_encrypted")

        progress = st.progress(0, text="Unlocking...")

        for i, uploaded in enumerate(unlock_files):
            fname = uploaded.name
            try:
                pdf_bytes = uploaded.read()
                unlocked_bytes, status = unlock_pdf(pdf_bytes, password=unlock_password)

                safe_name   = fname.replace(" ", "_")
                output_name = f"{prefix}{safe_name}"
                output_path = f"outputs/{output_name}"

                if status in ("unlocked", "not_encrypted"):
                    with open(output_path, "wb") as f:
                        f.write(unlocked_bytes)
                    results.append((output_name, output_path, status))
                    if status == "not_encrypted":
                        skipped.append(fname)

                elif status == "wrong_password":
                    errors.append((fname, "Wrong or missing password — could not decrypt."))

                else:
                    # status starts with "error:"
                    errors.append((fname, status.replace("error:", "", 1)))

            except Exception as e:
                errors.append((fname, str(e)))

            progress.progress(
                (i + 1) / len(unlock_files),
                text=f"Processing {i + 1}/{len(unlock_files)}…"
            )

        progress.empty()

        # ── Results summary ────────────────────────────────────────────────
        unlocked_count  = len([r for r in results if r[2] == "unlocked"])
        passthru_count  = len([r for r in results if r[2] == "not_encrypted"])
        error_count     = len(errors)

        if unlocked_count:
            st.success(f"✅ {unlocked_count} file(s) successfully unlocked.")
        if passthru_count:
            st.info(
                f"ℹ️ {passthru_count} file(s) were not encrypted — "
                "copied through unchanged: "
                + ", ".join(f"`{n}`" for n in skipped)
            )
        for fname, reason in errors:
            st.error(f"❌ **{fname}**: {reason}")

        # ── Download buttons ───────────────────────────────────────────────
        if results:
            st.divider()
            st.subheader("⬇️ Download Unlocked Files")

            # Batch ZIP download (shown whenever there are 2+ files)
            if len(results) > 1:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for output_name, output_path, _ in results:
                        zf.write(output_path, arcname=output_name)
                zip_buf.seek(0)
                _stretch_call(
                    st.download_button,
                    label=f"⬇️ Download All ({len(results)} files) as ZIP",
                    data=zip_buf,
                    file_name="unlocked_pdfs.zip",
                    mime="application/zip",
                    key="dl_unlock_all_zip",
                    type="primary",
                )
                st.caption("Or download files individually:")

            for output_name, output_path, status in results:
                label_suffix = " *(was not encrypted)*" if status == "not_encrypted" else ""
                with open(output_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️ {output_name}{label_suffix}",
                        data=f,
                        file_name=output_name,
                        mime="application/pdf",
                        key=f"dl_unlock_{output_name}"
                    )

    # ── Helper note ───────────────────────────────────────────────────────────
    with st.expander("ℹ️ How does this work?"):
        st.markdown("""
**Two kinds of PDF locks:**

| Type | What it does | Password needed to unlock? |
|------|-------------|---------------------------|
| **User password** (open password) | Prevents opening the file | ✅ Yes — enter it above |
| **Owner password** (permissions lock) | Restricts printing, copying, editing | ❌ No — unlocked automatically |

**Batch processing:** Upload multiple files at once. All will be processed with the same password (if any).

**Output files** are prefixed with the text you set in ③ and are immediately available for download after processing.

> *This tool uses [pypdf](https://pypdf.readthedocs.io/) for decryption. It cannot brute-force unknown user passwords.*
        """)
