import streamlit as st
import pdfplumber
from docx import Document
import os
import zipfile
from datetime import datetime
import re

st.set_page_config(page_title="SN Tools", layout="wide")

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


@st.cache_data(show_spinner=False, max_entries=16)
def remove_selected_bytes(input_pdf_bytes, names_to_remove, remove_gs=False):
    """Strip the chosen XObjects (and optionally gs operators). bytes -> bytes."""
    names_to_remove = set(names_to_remove)
    if not names_to_remove and not remove_gs:
        return input_pdf_bytes

    reader = PdfReader(io.BytesIO(input_pdf_bytes))
    writer = PdfWriter()

    for page in reader.pages:
        content = page.get_contents()
        if content is None:
            writer.add_page(page)
            continue

        try:
            content_obj = ContentStream(content, reader)
        except Exception:
            writer.add_page(page)
            continue

        new_ops = []
        for operands, operator in content_obj.operations:
            if operator == b"Do" and operands:
                if str(operands[0]) in names_to_remove:
                    continue
            if operator == b"gs" and remove_gs:
                continue
            new_ops.append((operands, operator))

        content_obj.operations = new_ops
        new_stream = DecodedStreamObject()
        new_stream.set_data(content_obj.get_data())
        page[NameObject("/Contents")] = new_stream
        writer.add_page(page)

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
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()

    for page in reader.pages:
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        overlay_bytes = make_header_overlay(pw, ph, HEADER_IMAGE_PATH,
                                            pw * img_width_frac, top_margin_pt)
        overlay_page  = PdfReader(io.BytesIO(overlay_bytes)).pages[0]
        page.merge_page(overlay_page)
        writer.add_page(page)

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


def remove_colors(img, targets, hue_tol=25, sat_min=10, whiten=70):
    """
    Remove pixels whose hue is close to any target colour.

    In the channel where an ink is brightest (red for a red seal, blue for a
    purple one) the ink nearly vanishes while black text underneath still
    shows, so matched pixels are replaced by that channel as grey. The stamp
    disappears and overlapping text survives.
      hue_tol - hue distance (degrees) still counted as the same ink
      sat_min - minimum colour strength (%) for a pixel to count as ink
      whiten  - leftover grey lighter than this (%) becomes pure white
    """
    rgb = np.array(img.convert("RGB")).astype(np.float32)
    hsv = np.array(img.convert("HSV")).astype(np.int16)
    h, s = hsv[..., 0], hsv[..., 1]
    colored = s >= sat_min * 255 / 100
    out = rgb.copy()
    tol = hue_tol * 255 / 360

    for color in targets:
        th = int(np.array(PILImage.new("RGB", (1, 1), tuple(color)).convert("HSV"))[0, 0, 0])
        d = np.abs(h - th)
        d = np.minimum(d, 256 - d)              # hue wraps around (red is at both ends)
        core = colored & (d <= tol)
        if not core.any():
            continue
        match = _grow(core, 1)                  # catch the soft edges of the ink
        ch = int(np.argmax(color))              # channel where this ink is brightest
        ref = max(np.percentile(rgb[..., ch][core], 75), 1)
        gray = np.clip(rgb[..., ch] * (255.0 / ref), 0, 255)
        gray[gray >= whiten * 255 / 100] = 255
        out[match] = gray[match][:, None]

    return PILImage.fromarray(out.astype(np.uint8))


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
    st.markdown("---")
    page = st.radio(
        "Navigate to",
        options=["🚿 Watermark Remover", "📄 SPEC & COA Generator", "🔓 PDF Unlocker"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.caption("Skyherb® Internal Tools")


# ==============================
# PAGE 1 — SPEC & COA Generator
# ==============================

if page == "📄 SPEC & COA Generator":

    st.title("📄 SPEC & COA Generator")

    st.subheader("Upload Supplier Documents")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        spec_pdf = st.file_uploader("Upload Supplier SPEC (PDF)", type="pdf", key="spec_uploader")
    with col_up2:
        coa_pdf  = st.file_uploader("Upload Supplier COA (PDF)",  type="pdf", key="coa_uploader")

    extract_col1, extract_col2 = st.columns(2)

    with extract_col1:
        if spec_pdf and st.button("🔍 Extract Info from SPEC PDF"):
            with st.spinner("Extracting..."):
                extracted = extract_header_info(spec_pdf)
            if extracted:
                for key, value in extracted.items():
                    st.session_state[key] = value
                st.success(f"Extracted {len(extracted)} field(s) from SPEC PDF. Review and edit below.")
            else:
                st.warning("No recognisable header fields found in the SPEC PDF.")

    with extract_col2:
        if coa_pdf and st.button("🔍 Extract Info from COA PDF"):
            with st.spinner("Extracting..."):
                extracted = extract_header_info(coa_pdf)
            if extracted:
                for key, value in extracted.items():
                    st.session_state[key] = value
                st.success(f"Extracted {len(extracted)} field(s) from COA PDF. Review and edit below.")
            else:
                st.warning("No recognisable header fields found in the COA PDF.")

    st.divider()

    st.subheader("Product & Batch Information")
    st.caption("Fields are pre-filled from the PDF after clicking Extract. You can edit any value before generating.")

    col1, col2 = st.columns(2)

    with col1:
        product_name = st.text_input("Product Name",       value=st.session_state["product_name"], key="in_product_name")
        brand        = st.text_input("Brand",               value=st.session_state["brand"],        key="in_brand")
        lot_no       = st.text_input("Lot No.",             value=st.session_state["lot_no"],        key="in_lot_no")
        quantity     = st.text_input("Quantity",            value=st.session_state["quantity"],      key="in_quantity")
        mfg_date     = st.text_input("Mfg. Date",          value=st.session_state["mfg_date"],      key="in_mfg_date")
        expiry_date  = st.text_input("Expiry Date",        value=st.session_state["expiry_date"],   key="in_expiry_date")

    with col2:
        shelf_life   = st.text_input("Shelf Life",         value=st.session_state["shelf_life"],    key="in_shelf_life")
        plant_part   = st.text_input("Plant Part",         value=st.session_state["plant_part"],    key="in_plant_part")
        latin_name   = st.text_input("Plant Latin Name",   value=st.session_state["latin_name"],    key="in_latin_name")
        origin       = st.text_input("Country of Origin",  value=st.session_state["origin"],        key="in_origin")
        solvent      = st.text_input("Extraction Solvent", value=st.session_state["solvent"],       key="in_solvent")

    header_data = {
        "ProductName": product_name,
        "Brand":       brand,
        "Origin":      origin,
        "Solvent":     solvent,
        "PlantPart":   plant_part,
        "LatinName":   latin_name,
        "ShelfLife":   shelf_life,
        "LotNo":       lot_no,
        "Quantity":    quantity,
        "ManuDate":    mfg_date,
        "ExpiryDate":  expiry_date,
        "IssueDate":   datetime.today().strftime("%Y-%m-%d"),
    }

    st.divider()

    gen_col1, gen_col2 = st.columns(2)

    with gen_col1:
        if spec_pdf and st.button("⚙️ Generate Internal SPEC"):
            with st.spinner("Generating SPEC..."):
                general_rows, micro_rows = extract_spec_tables(spec_pdf)
                output_file = fill_spec_template(header_data, general_rows, micro_rows)
            if output_file:
                with open(output_file, "rb") as f:
                    st.download_button(
                        "⬇️ Download SPEC",
                        f,
                        file_name="Internal_SPEC.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

    with gen_col2:
        if coa_pdf and st.button("⚙️ Generate Internal COA"):
            with st.spinner("Generating COA..."):
                general_rows, micro_rows = extract_coa_tables(coa_pdf)
                output_file = fill_coa_template(header_data, general_rows, micro_rows)
            if output_file:
                with open(output_file, "rb") as f:
                    st.download_button(
                        "⬇️ Download COA",
                        f,
                        file_name="Internal_COA.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )


# ==============================
# PAGE 2 — Watermark Remover
# ==============================

elif page == "🚿 Watermark Remover":

    st.title("🚿 Watermark Remover")
    st.caption("Remove logos, watermarks and coloured stamps, white out any area, and add your header. "
               "The preview on the right updates with every change.")

    for _k, _v in {"wm_scan_results": [], "wm_remove_set": set(), "cr_colors": [],
                   "pv_last_event": None, "wo_areas": []}.items():
        if _k not in st.session_state:
            st.session_state[_k] = _v

    # Keep the preview column pinned on screen while the controls scroll
    st.markdown("""
        <style>
        div[data-testid="stColumn"]:has(.st-key-wm_preview_panel),
        div[data-testid="column"]:has(.st-key-wm_preview_panel) {
            position: sticky; top: 3.75rem; align-self: flex-start;
            max-height: calc(100vh - 4.5rem); overflow-y: auto;
        }
        </style>""", unsafe_allow_html=True)

    # ── ① Upload (full width) ─────────────────────────────────────────────
    st.subheader("① Upload PDF(s)")
    uploaded_files = st.file_uploader(
        "Upload one or more PDF files",
        type="pdf",
        accept_multiple_files=True,
        key="wm_uploader"
    )

    if not uploaded_files:
        st.info("Upload one or more PDFs to get started.")
        st.stop()

    # getvalue() always returns the full file, even after earlier reads
    file_bytes = {f.name: f.getvalue() for f in uploaded_files}
    preview_ok = fitz is not None

    left, right = st.columns([1.15, 1], gap="large")

    # ── Right panel, part 1: which page to preview ───────────────────────
    with right:
        try:
            panel = st.container(key="wm_preview_panel")   # key -> CSS hook for the sticky panel
        except TypeError:                                  # older Streamlit: no sticky panel
            panel = st.container()
    with panel:
        st.subheader("👁️ Live Preview")
        if preview_ok:
            pv1, pv2 = st.columns([3, 1])
            with pv1:
                preview_name = st.selectbox("File", list(file_bytes.keys()), key="pv_file")
            n_pages = pdf_page_count(file_bytes[preview_name])
            with pv2:
                preview_page = st.number_input("Page", min_value=1, max_value=n_pages, value=1,
                                               step=1, key="pv_page") - 1
        else:
            st.warning(
                "Preview, colour removal and white-out need PyMuPDF, which isn't available to the "
                "Python running this app. Add `pymupdf` to requirements.txt (or run the command "
                "below locally), then restart the app."
            )
            st.code(f'"{sys.executable}" -m pip install pymupdf numpy', language="bash")
            if FITZ_ERROR:
                st.caption(f"Import error: {FITZ_ERROR}")

    # ── Left panel: all controls ─────────────────────────────────────────
    with left:
        st.write(f"{len(uploaded_files)} file(s) uploaded.")

        # ② Scan
        st.subheader("② Embedded Logos & Watermarks")
        st.caption("Finds logos and watermarks stored as separate objects. Stamps that are part "
                   "of a scanned page won't show here — use ③ or ④ for those.")

        if st.button("🔍 Scan PDFs for embedded elements"):
            combined = {}
            with st.spinner("Scanning..."):
                for fname, data in file_bytes.items():
                    try:
                        for item in scan_xobjects(data):
                            name = item["name"]
                            if name not in combined:
                                combined[name] = item.copy()
                            else:
                                combined[name]["pages"] = sorted(
                                    set(combined[name]["pages"] + item["pages"]))
                                if combined[name]["thumbnail"] is None and item["thumbnail"]:
                                    combined[name]["thumbnail"] = item["thumbnail"]
                    except Exception as e:
                        st.warning(f"Could not scan {fname}: {e}")

            st.session_state["wm_scan_results"] = list(combined.values())
            wm_hints = {"wm", "watermark", "mark", "stamp", "logo", "bg", "background"}
            st.session_state["wm_remove_set"] = {
                it["name"] for it in st.session_state["wm_scan_results"]
                if it["subtype"] == "Form" or any(h in it["name"].lstrip("/").lower() for h in wm_hints)
            }
            if st.session_state["wm_scan_results"]:
                st.success(f"Found {len(st.session_state['wm_scan_results'])} embedded element(s). "
                           "Tick what to remove.")
            else:
                st.warning("No embedded overlay objects found in these PDFs.")

        scan_results = st.session_state["wm_scan_results"]
        remove_gs = False

        if scan_results:
            COLS = 3
            remove_set = st.session_state["wm_remove_set"].copy()
            for row_start in range(0, len(scan_results), COLS):
                cols = st.columns(COLS)
                for col, item in zip(cols, scan_results[row_start: row_start + COLS]):
                    with col:
                        name, subtype, pages, thumb = (item["name"], item["subtype"],
                                                       item["pages"], item["thumbnail"])
                        if thumb:
                            _stretch_call(st.image, thumb)
                        else:
                            icon, label = ("📐", "Vector / Form") if subtype == "Form" else ("❓", subtype)
                            st.markdown(
                                f"<div style='background:#f0f0f0;border-radius:6px;"
                                f"padding:28px 0;text-align:center;font-size:28px'>"
                                f"{icon}<br><small style='font-size:11px;color:#555'>{label}</small></div>",
                                unsafe_allow_html=True)
                        page_str = (f"p.{pages[0]}" if len(pages) == 1
                                    else f"p.{pages[0]}–{pages[-1]}" if pages == list(range(pages[0], pages[-1] + 1))
                                    else f"{len(pages)} pages")
                        st.caption(f"`{name}` · {subtype} · {page_str}")
                        if st.checkbox("Remove", value=(name in remove_set), key=f"chk_{name}"):
                            remove_set.add(name)
                        else:
                            remove_set.discard(name)
            st.session_state["wm_remove_set"] = remove_set

            remove_gs = st.checkbox(
                "Also strip `gs` (graphic state) operators  ⚠️ may affect some page rendering",
                value=False, key="wm_remove_gs")

        names_to_remove = tuple(sorted(st.session_state["wm_remove_set"])) if scan_results else ()

        # Preview base image (after overlays + white-out), used for colour detection too
        if preview_ok:
            base_bytes = remove_selected_bytes(file_bytes[preview_name], names_to_remove, remove_gs)
            base_img, page_w_pt, page_h_pt = render_page(base_bytes, preview_page)
            page_boxes = areas_for(st.session_state["wo_areas"], preview_name, preview_page)
            wo_img = whiteout_image(base_img, page_boxes)

        # ③ White out areas
        st.divider()
        st.subheader("③ White Out Areas")
        if not preview_ok:
            st.caption("Needs PyMuPDF — see the message on the right.")
        elif streamlit_image_coordinates is None:
            st.warning("Drawing on the preview needs one more package: add "
                       "`streamlit-image-coordinates` to requirements.txt and restart the app.")
        else:
            st.caption("Choose **White out an area** above the preview, then drag a box over any "
                       "text, logo or image. It's covered with white and removed from the PDF.")
        wo_scope = st.radio(
            "New boxes apply to",
            ["This page only", "All pages of this file", "All pages of all files"],
            horizontal=True, key="wo_scope",
            help="Use 'All pages of all files' for a logo that sits in the same spot on every document.")

        if st.session_state["wo_areas"]:
            for i, a in enumerate(list(st.session_state["wo_areas"])):
                fkey = "*" if a["file"] == "*" else "f"
                pkey = "*" if a["page"] == "*" else "p"
                where = (f"{a['file']} p.{a['page'] + 1}" if fkey == "f" and pkey == "p"
                         else f"{a['file']} — all pages" if fkey == "f"
                         else "all files — all pages")
                c1, c2 = st.columns([5, 1])
                with c1:
                    x0, y0, x1, y1 = a["box"]
                    st.caption(f"Box {i + 1}: {where} · {int((x1 - x0) * 100)}% × {int((y1 - y0) * 100)}% "
                               f"of the page")
                with c2:
                    if st.button("✕", key=f"wo_del_{i}"):
                        st.session_state["wo_areas"].pop(i)
                        st.rerun()
            if st.button("Clear all white-out boxes", key="wo_clear"):
                st.session_state["wo_areas"] = []
                st.rerun()
        else:
            st.caption("No white-out boxes yet.")

        # ④ Colour removal
        st.divider()
        st.subheader("④ Remove Coloured Stamps / Marks")
        colors_on = False
        color_params = (25, 10, 70)
        raster_dpi = 300

        if preview_ok:
            colors_on = st.checkbox("Remove ink colours (red QC seal, blue/purple stamp, signature…)",
                                    value=False, key="cr_enable")
        if colors_on:
            st.caption("Pages are converted to images for this step, so text in the output won't "
                       "be selectable. Black text under a stamp is kept.")

            detected = detect_ink_colors(wo_img)
            st.markdown("**Colours found on this page** — Add each stamp colour:")
            if detected:
                dcols = st.columns(min(len(detected), 4))
                for i, c in enumerate(detected):
                    hx = rgb_to_hex(c)
                    with dcols[i % len(dcols)]:
                        st.markdown(swatch_html(hx), unsafe_allow_html=True)
                        if st.button("Add", key=f"cr_add_det_{hx}_{i}",
                                     disabled=hx in st.session_state["cr_colors"]):
                            st.session_state["cr_colors"].append(hx)
                            st.rerun()
            else:
                st.caption("No coloured ink detected on this page.")

            mcol1, mcol2 = st.columns([1, 2])
            with mcol1:
                manual = st.color_picker("Or choose a colour", "#7a4fc0", key="cr_manual")
            with mcol2:
                st.write("")
                st.write("")
                if st.button("Add this colour", key="cr_add_manual"):
                    if manual not in st.session_state["cr_colors"]:
                        st.session_state["cr_colors"].append(manual)
                    st.rerun()
            if streamlit_image_coordinates is not None:
                st.caption("Tip: choose **Pick a stamp colour** above the preview and click on the stamp.")

            st.markdown("**Colours to remove:**")
            if st.session_state["cr_colors"]:
                scols = st.columns(min(len(st.session_state["cr_colors"]), 4))
                for i, hx in enumerate(list(st.session_state["cr_colors"])):
                    with scols[i % len(scols)]:
                        st.markdown(swatch_html(hx), unsafe_allow_html=True)
                        if st.button("✕ Remove", key=f"cr_del_{hx}_{i}"):
                            st.session_state["cr_colors"].remove(hx)
                            st.rerun()
            else:
                st.info("No colours selected yet.")

            with st.expander("Fine-tune", expanded=False):
                s1, s2, s3 = st.columns(3)
                with s1:
                    hue_tol = st.slider("Colour range", 1, 90, 25, key="cr_hue",
                                        help="Raise if parts of the stamp stay; lower if other colours disappear.")
                with s2:
                    sat_min = st.slider("Min strength %", 1, 80, 10, key="cr_sat",
                                        help="Lower to also remove the faint, pale edges of the stamp.")
                with s3:
                    whiten = st.slider("Whiten %", 30, 100, 70, key="cr_whiten",
                                       help="Lower if a faint ghost of the stamp remains; "
                                            "raise if text under the stamp looks thin.")
                raster_dpi = st.select_slider("Output resolution (dpi)", options=[150, 200, 300, 400],
                                              value=300, key="cr_dpi",
                                              help="Higher = sharper but larger files.")
                color_params = (hue_tol, sat_min, whiten)

        active_colors = [hex_to_rgb(h) for h in st.session_state["cr_colors"]] if colors_on else []

        # ⑤ Header
        st.divider()
        st.subheader("⑤ Company Header")
        header_exists = os.path.isfile(HEADER_IMAGE_PATH)
        if not header_exists:
            st.warning("⚠️ `SN_Header.png` not found next to `app.py`. "
                       "Place the file there to enable this option.")
        add_header = st.checkbox("Stamp SN company header on every page", value=False,
                                 disabled=not header_exists, key="wm_add_header")
        img_width_pct, top_margin = 98, 10
        if add_header and header_exists:
            hcol1, hcol2 = st.columns(2)
            with hcol1:
                img_width_pct = st.slider("Header width (% of page width)", 20, 100, 98, 1,
                                          key="wm_header_width")
            with hcol2:
                top_margin = st.number_input("Top margin (pt)", 0, 100, 10, 1, key="wm_top_margin")

        # ⑥ Source-info check (flags only, changes nothing)
        st.divider()
        st.subheader("⑥ Source-Info Check")
        st.caption("Flags spots that may reveal the manufacturer or supplier — company names, "
                   "\"manufactured by\", phone/fax, email, website, address, logos — so you can "
                   "review before sending to a customer. Nothing is changed automatically.")
        sc1, sc2 = st.columns(2)
        with sc1:
            ignore_text = st.text_area("Our own names / address (ignored)",
                                       "Skyherb, Loker Ave, Carlsbad, CA92010", height=80,
                                       key="si_ignore", help="Comma-separated. Lines containing these are not flagged.")
        with sc2:
            watch_text = st.text_area("Supplier names to look for (optional)", "", height=80,
                                      key="si_watch", placeholder="e.g. Xinhua, Dae Hyung",
                                      help="Comma-separated. Always flagged wherever they appear.")
        ignore_terms, watch_terms = parse_terms(ignore_text), parse_terms(watch_text)
        si_highlight = st.checkbox("Highlight flagged spots on the preview", value=True, key="si_hl")
        if pytesseract is None:
            st.caption("Scanned pages (no text layer) can only be checked by eye. To check them "
                       "automatically, add `pytesseract` to requirements.txt and `tesseract-ocr` "
                       "to packages.txt.")

        # ⑦ Process
        st.divider()
        st.subheader("⑦ Process Files")
        if colors_on and not active_colors:
            st.warning("Colour removal is on but no colours are selected — it will be skipped.")

        # Results live in session_state so the download buttons survive reruns
        upload_sig = tuple((n, len(b)) for n, b in file_bytes.items())
        if st.session_state.get("wm_results_sig") != upload_sig:
            st.session_state["wm_results"] = None

        if st.button("⚙️ Process Files", type="primary"):
            results, errors, reports = [], [], {}
            progress = st.progress(0, text="Processing...")
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
                            reports[output_name] = check_pdf(
                                text_stage, ignore_terms, watch_terms,
                                image_stage if active_colors else None)
                        except Exception:
                            pass
                except Exception as e:
                    errors.append((fname, str(e)))
                progress.progress((i + 1) / len(file_bytes),
                                  text=f"Processed {i + 1}/{len(file_bytes)}")
            progress.empty()

            ops = []
            if names_to_remove: ops.append(f"{len(names_to_remove)} element(s) removed")
            if st.session_state["wo_areas"] and preview_ok:
                ops.append(f"{len(st.session_state['wo_areas'])} area(s) whited out")
            if active_colors: ops.append(f"{len(active_colors)} colour(s) removed")
            if add_header and header_exists: ops.append("header added")

            zip_bytes = None
            if len(results) > 1:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for output_name, out_bytes in results:
                        zf.writestr(output_name, out_bytes)
                zip_bytes = zip_buf.getvalue()

            st.session_state["wm_results"] = {"files": results, "errors": errors,
                                              "ops": ops, "zip": zip_bytes, "reports": reports}
            st.session_state["wm_results_sig"] = upload_sig

        res = st.session_state.get("wm_results")
        if res:
            for name, err in res["errors"]:
                st.error(f"❌ {name}: {err}")
            files = res["files"]
            if files:
                st.success(f"✅ {len(files)} file(s) done" +
                           (f" ({', '.join(res['ops'])})" if res["ops"] else "") + ".")

                reports = res.get("reports") or {}
                if reports:
                    flagged = []
                    for out_name, rep_ in reports.items():
                        total, parts, imgs, scanned = summarize_report(rep_)
                        bits = []
                        if total:
                            bits.append(f"**{total} spot(s)** with possible company/source info ({', '.join(parts)})")
                        if imgs:
                            bits.append(f"image(s)/logo(s) to check ({', '.join(imgs)})")
                        if scanned:
                            bits.append(f"scanned page(s) {', '.join(scanned)} — check by eye")
                        if bits:
                            flagged.append((out_name, bits, rep_))
                    if flagged:
                        st.warning(f"🔎 Before sending to a customer, review {len(flagged)} file(s):")
                        for out_name, bits, rep_ in flagged:
                            with st.expander(f"{out_name} — " + "; ".join(b.replace('**', '') for b in bits)):
                                for pno, r in enumerate(rep_):
                                    for n, spot in enumerate(r["spots"], 1):
                                        st.markdown(f"p.{pno + 1} · #{n} · *{spot['reason']}*: `{spot['text']}`")
                    else:
                        st.info("🔎 No company/source info detected in the processed files.")

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
                    _download(f"⬇️ Download All ({len(files)} files) as ZIP", res["zip"],
                              "SN_processed_pdfs.zip", "application/zip", "dl_wm_all_zip",
                              stretch=True, type="primary")
                    with st.expander("Download files individually"):
                        for output_name, out_bytes in files:
                            _download(f"⬇️ {output_name}", out_bytes, output_name,
                                      "application/pdf", f"dl_{output_name}")
                else:
                    output_name, out_bytes = files[0]
                    _download(f"⬇️ {output_name}", out_bytes, output_name,
                              "application/pdf", f"dl_{output_name}", type="primary")

    # ── Right panel, part 2: the live result ─────────────────────────────
    if preview_ok:
        with panel:
            after_img = wo_img
            if active_colors:
                after_img = remove_colors(after_img, active_colors, *color_params)
            if add_header and header_exists:
                after_img = paste_header(after_img, page_w_pt, PREVIEW_DPI,
                                         img_width_pct / 100, top_margin)

            # Source-info check for this page (flags only)
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

            modes = ["View only", "White out an area (drag)"]
            if colors_on:
                modes.append("Pick a stamp colour (click)")
            if streamlit_image_coordinates is not None:
                mode = st.radio("On the preview:", modes, horizontal=True, key="pv_mode")
            else:
                mode = "View only"

            if mode == "View only":
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
                drag = mode.startswith("White out")
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
                                st.rerun()

            changes = []
            if names_to_remove: changes.append(f"{len(names_to_remove)} overlay(s) removed")
            if page_boxes: changes.append(f"{len(page_boxes)} area(s) whited out")
            if active_colors: changes.append(f"{len(active_colors)} colour(s) removed")
            if add_header and header_exists: changes.append("header added")
            st.caption(f"{preview_name} — page {preview_page + 1} of {n_pages}. "
                       "Showing: " + (", ".join(changes) if changes else "no changes yet"))

            if pv_report:
                n_spots = len(pv_report["spots"])
                notes = []
                if n_spots:
                    notes.append(f"**{n_spots} spot(s)** on this page may contain company/source info"
                                 + (" (orange boxes)" if si_highlight else ""))
                if pv_report["images"]:
                    notes.append(f"{pv_report['images']} image(s) remain — check none is a supplier logo or stamp")
                if pv_report["scanned"] and not pv_report["ocr"]:
                    notes.append("this page is a scan, so its text can't be checked automatically — review by eye")
                if notes:
                    st.warning("🔎 " + " · ".join(notes))
                    if n_spots:
                        with st.expander("Show flagged spots"):
                            for n, spot in enumerate(pv_report["spots"], 1):
                                st.markdown(f"#{n} · *{spot['reason']}*: `{spot['text']}`")
                else:
                    st.success("🔎 No company/source info detected on this page.")


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
