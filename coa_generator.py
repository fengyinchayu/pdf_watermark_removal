"""
SPEC & COA Generator engine for SN Tools.

Turns a supplier COA / specification (already read into rows by coa_review) into a
Skyherb COA or SPEC by filling the Word templates in templates/.

    rows_from_supplier()  clean rows: Skyherb test names, groups, junk removed
    apply_our_spec()      optional: replace the Standards column with our spec library limits
    check_rows()          consistency checks before a document is issued
    fill_template()       write header fields + rows into the .docx template
"""
import copy
import datetime as dt
import io
import re

import coa_review as cr

# ======================================================================
# Rows
# ======================================================================

# Skyherb house names (as on our SPEC sheets); None = keep the supplier's wording
DISPLAY_NAMES = {
    "lead": "Lead", "arsenic": "Arsenic", "cadmium": "Cadmium", "mercury": "Mercury",
    "heavy metals": "Heavy Metals", "total plate count": "Total Aerobic", "yeast & mold": "Mold and Yeast",
    "yeast": "Yeast", "mold": "Mold", "e. coli": "E. coli", "salmonella": "Salmonella",
    "staphylococcus aureus": "Staphylococcus aureus", "coliforms": "Coliform",
    "enterobacteriaceae": "Enterobacteriaceae", "listeria": "Listeria", "loss on drying": "Loss on drying",
    "ash": None, "sulfated ash": None, "bulk density": "Bulk Density", "tapped density": "Tapped Density",
    "particle size": "Particle Size", "appearance": "Appearance*", "odor & taste": None,
    "identification": "Identification", "pesticides": "Pesticide Residues", "residual solvents": "Residual Solvents",
    "aflatoxins": "Aflatoxins", "pah": "PAHs", "ph": "pH", "solubility": "Solubility",
}

# order of sections on our documents; the banner groups always get a heading row (house style)
GROUPS = ["Physical & Chemical", "Assay & Identification", "Heavy Metals", "Solvents Residue",
          "Pesticides & Contaminants", "Microbiological"]
BANNER_GROUPS = {"Solvents Residue", "Pesticides & Contaminants"}
_GROUP_OF_CAT = {"assay": "Assay & Identification", "identification": "Assay & Identification",
                 "physical": "Physical & Chemical", "heavy metals": "Heavy Metals",
                 "residual solvents": "Solvents Residue", "pesticides": "Pesticides & Contaminants",
                 "contaminants": "Pesticides & Contaminants", "microbiology": "Microbiological"}
_MICRO_UNITS = {"cfu/g": "CFU/g", "mpn/g": "MPN/g", "cfu/ml": "CFU/mL", "mpn/100g": "MPN/100g"}

# rows that are really header information / column titles, not tests
_JUNK = re.compile(
    r"^(product\s*name|trade\s*name|product|item\s*name|batch\s*(?:no|number|size)?|lot\s*(?:no|number)?|"
    r"manu\w*\.?\s*date|mfg\.?\s*date|production\s*date|exp\w*\.?\s*date|expiration\s*date|report\s*date|"
    r"test\s*date|issue\s*date|quantity|batch\s*size|plant\s*source|botanical\s*(?:name|source)|latin\s*name|"
    r"used\s*part|parts?\s*used|plant\s*part|shelf\s*life|storage|packing|package|country\s*of\s*origin|origin|"
    r"conclusion|remarks?|note|items?|test\s*items?|analy\w*|specifications?|results?|methods?|standards?|"
    r"manufacturer|manufactured\s*by|supplier|producer|address|tel|fax|phone|e-?mail|website|"
    r"analy[sz]ed\s*by|approved\s*by|checked\s*by|tested\s*by|qc|qa)\s*:?$",
    re.I)
_SUPPLIER_INFO = re.compile(
    r"\b(co\.?\s*,?\s*ltd|limited|inc\.|llc|corp\.|gmbh|有限公司|公司)\b|"
    r"\b(tel|fax|phone)\b\s*[:.]?\s*\+?\d|[\w.+-]+@[\w-]+\.[\w.]+|www\.\S+|https?://", re.I)


def display_name(name, use_ours=True):
    raw = cr.norm_text(name).strip(" :：")
    if not use_ours:
        return raw
    c = cr.canon(raw)
    if re.search(r"\((?:cfu|mpn)", raw, re.I):        # already in house form, e.g. "Total Aerobic (CFU/g)"
        return raw
    if DISPLAY_NAMES.get(c):
        return DISPLAY_NAMES[c]
    # keep the supplier's wording for assay markers etc., tidied
    return raw[:1].upper() + raw[1:] if raw else raw


def group_of(name, category=""):
    known = cr.CATEGORY_OF.get(cr.canon(name))
    if known:                                  # our standard group for known tests, whatever the supplier used
        return _GROUP_OF_CAT.get(known, "Physical & Chemical")
    if re.search(cr._SOLVENTS, cr.clean_name(name)):
        return "Solvents Residue"
    cat = cr.category_of(name, category or "")
    if cat not in _GROUP_OF_CAT:
        cat = cr.category_of(name, "")
    return _GROUP_OF_CAT.get(cat, "Assay & Identification")


def supplier_flag(text, supplier_terms=()):
    t = str(text or "")
    if not t.strip():
        return False
    low = t.lower()
    if any(s and s in low for s in supplier_terms):
        return True
    return bool(_SUPPLIER_INFO.search(t))


def supplier_terms_from(info):
    """Distinctive words of the supplier's name (e.g. 'xinhua') to look for elsewhere."""
    terms = []
    for key in ("supplier",):
        name = cr.clean_name(info.get(key, ""))
        name = re.sub(r"\b(co|ltd|limited|inc|llc|corp|company|group|bio\w*|pharma\w*|tech\w*|"
                      r"natural|nutrition\w*|ingredients?|health|industr\w*|trading|the|and)\b", " ", name)
        terms += [w for w in name.split() if len(w) >= 4]
    return tuple(dict.fromkeys(terms))


def rows_from_supplier(tests, use_our_names=True, supplier_terms=()):
    """Supplier test rows -> clean rows for our document (list of dicts)."""
    rows = []
    for t in tests:
        name = cr.norm_text(t.get("name"))
        std, res = cr.norm_text(t.get("spec")), cr.norm_text(t.get("result"))
        unit = cr.norm_text(t.get("unit"))
        if unit and unit.lower() not in res.lower() and res and re.search(r"\d", res):
            res = f"{res} {unit}"
        if not name or _JUNK.match(name):
            continue
        if not std and not res:
            continue
        if re.fullmatch(r"(specifications?|standards?|limits?)", std, re.I) or \
                re.fullmatch(r"(results?|values?)", res, re.I):
            continue                                           # repeated column-title row
        flagged = any(supplier_flag(v, supplier_terms) for v in (name, std, res, t.get("method")))
        rows.append({
            "Include": not flagged,                             # supplier details start unticked
            "Group": group_of(name, t.get("category", "")),
            "Characteristic": display_name(name, use_our_names),
            "Standards": pretty_limit_text(std),
            "Results": pretty_result_text(res),
            "Methods": cr.norm_text(t.get("method")),
            "Supplier info?": flagged,
        })
    rows = [house_micro(r) for r in rows]
    order = {g: i for i, g in enumerate(GROUPS)}
    rows.sort(key=lambda r: order.get(r["Group"], 0))          # stable: keeps supplier order inside a group
    return rows


def house_micro(r):
    """Microbiology in house style: 'Total Aerobic (CFU/g)' | '≤1000' | '<10'."""
    if r.get("Group") != "Microbiological":
        return r
    unit = cr.find_unit(r.get("Standards")) or cr.find_unit(r.get("Results"))
    if unit not in _MICRO_UNITS:
        return r
    r = dict(r)
    if not re.search(r"\((?:cfu|mpn)", r["Characteristic"], re.I):
        r["Characteristic"] = f"{r['Characteristic']} ({_MICRO_UNITS[unit]})"
    rx = re.compile(r"\s*" + r"\s*".join(re.escape(ch) for ch in unit), re.I)
    for k in ("Standards", "Results"):
        r[k] = rx.sub("", str(r.get(k) or "")).strip()
    return r


def pretty_limit_text(s):
    s = cr.norm_text(s)
    s = re.sub(r"10\^(\d)", lambda m: "10" + "⁰¹²³⁴⁵⁶⁷⁸⁹"[int(m.group(1))], s)
    return s


def pretty_result_text(s):
    return pretty_limit_text(s)


def format_limit(lim):
    """Our spec limit -> COA wording: '≥ 50%', '≤ 1 ppm', '4.5-7.0', 'Negative/25g'."""
    def num(v):
        return cr.fmt_num(v).replace("10^", "10")
    us = {"g/ml": "g/mL", "cfu/g": "CFU/g", "mpn/g": "MPN/g"}.get(lim.unit or "", lim.unit or "")
    if lim.min is not None and lim.max is not None:
        return f"{num(lim.min)}-{num(lim.max)}{us}"
    if lim.max is not None:
        return f"{'≤' if lim.max_incl else '<'}{num(lim.max)}{us}"
    if lim.min is not None:
        return f"{'≥' if lim.min_incl else '>'}{num(lim.min)}{us}"
    return lim.text or lim.expected or ""


def apply_our_spec(rows, spec_rows, add_missing=True):
    """Replace Standards (and empty Methods) with our spec library; add spec tests the supplier lacks."""
    out = [dict(r) for r in rows]
    used = set()
    for sr in spec_rows:
        best, best_s = None, 0.0
        for i, r in enumerate(out):
            if i in used:
                continue
            s = cr.name_score(r["Characteristic"], sr["test"])
            if s > best_s:
                best, best_s = i, s
        if (best is None or best_s < 0.82) and cr.canon(sr["test"]) == "yeast & mold":
            parts = [i for i, r in enumerate(out) if i not in used and cr.canon(r["Characteristic"]) in ("yeast", "mold")]
            if parts:                                # supplier reports yeast and mold separately
                for i in parts:
                    used.add(i)
                    out[i]["Standards"] = format_limit(sr["limit"])
                    out[i]["From our spec"] = True
                continue
        if best is not None and best_s >= 0.82:
            used.add(best)
            out[best]["Standards"] = format_limit(sr["limit"])
            if sr["method"] and not out[best]["Methods"]:
                out[best]["Methods"] = sr["method"]
            out[best]["From our spec"] = True
        elif add_missing and sr["required"]:
            out.append({"Include": True, "Group": _GROUP_OF_CAT.get(sr["category"], "Assay & Identification"),
                        "Characteristic": display_name(sr["test"]), "Standards": format_limit(sr["limit"]),
                        "Results": "", "Methods": sr["method"], "Supplier info?": False, "From our spec": True})
    out = [house_micro(r) for r in out]
    order = {g: i for i, g in enumerate(GROUPS)}
    out.sort(key=lambda r: order.get(r["Group"], 0))
    return out


# ======================================================================
# Checks before issuing
# ======================================================================

def check_rows(rows, fields, doc_type="COA", today=None, hide=()):
    """Return a list of (level, message); level is 'error', 'warning' or 'info'."""
    today = today or dt.date.today()
    fields = {k: ("" if k in set(hide or ()) else v) for k, v in fields.items()}
    hide = set(hide or ())
    out = []
    inc = [r for r in rows if r.get("Include", True)]
    for key, label in (("ProductName", "Product name"), ("LotNo", "Lot No."), ("ManuDate", "Mfg. date")):
        if (doc_type == "SPEC" and key in ("LotNo", "ManuDate")) or key in hide:
            continue
        if not str(fields.get(key, "")).strip():
            out.append(("warning", f"{label} is empty."))
    if not inc:
        out.append(("error", "No test rows are included."))
    for r in inc:
        name = r.get("Characteristic") or "(unnamed row)"
        if r.get("Supplier info?"):
            out.append(("warning", f"“{name}” may contain supplier/manufacturer details — check before sending."))
        if doc_type != "COA":
            continue
        if not str(r.get("Results") or "").strip():
            out.append(("error", f"{name}: no result."))
            continue
        lim = cr.parse_limit(r.get("Standards"))
        res = cr.parse_result(r.get("Results"), lim.unit)
        status, note, _ = cr.compare(name, cr.category_of(name), lim, res)
        if status == "Fail":
            out.append(("error", f"{name}: result {r.get('Results')} does not meet the standard {r.get('Standards')}."))
        elif status == "Review" and "no value" not in note and "rather than negative" not in note:
            out.append(("info", f"{name}: {note}"))
    if doc_type == "COA":
        vague = [r.get("Characteristic") for r in inc if cr.parse_result(r.get("Results")).qual == "pass"
                 and cr.parse_limit(r.get("Standards")).numeric]
        if vague:
            out.append(("info", f"{len(vague)} result(s) say only 'Pass/Conforms' where the standard is a number "
                                f"({', '.join(vague[:6])}{'…' if len(vague) > 6 else ''}). Customers may ask for values."))
    for key, label in (("ProductName", "Product name"), ("LatinName", "Latin name"), ("Solvent", "Solvent"),
                       ("Origin", "Country of origin")):
        if supplier_flag(fields.get(key, "")):
            out.append(("warning", f"{label} looks like it contains company details."))
    mfg = cr.parse_date(fields.get("ManuDate", ""))
    exp = cr.parse_date(fields.get("ExpiryDate", ""), end_of_month=True)
    if mfg.date and exp.date and exp.date <= mfg.date:
        out.append(("error", "Expiry date is not after the manufacture date."))
    if doc_type == "COA" and exp.date and exp.date < today:
        out.append(("error", f"This lot expired on {exp.date:%Y-%m-%d}."))
    if mfg.date and mfg.date > today:
        out.append(("warning", "Manufacture date is in the future."))
    return out


# ======================================================================
# Dates / shelf life
# ======================================================================

DATE_FORMATS = {"Aug. 19, 2025": "%b. %d, %Y", "2025-08-19": "%Y-%m-%d", "08/19/2025": "%m/%d/%Y",
                "19 Aug 2025": "%d %b %Y", "August 19, 2025": "%B %d, %Y"}


def format_date(text, fmt_label):
    p = cr.parse_date(text)
    if not p.date:
        return cr.norm_text(text)
    out = p.date.strftime(DATE_FORMATS.get(fmt_label, "%Y-%m-%d"))
    return out.replace("May. ", "May ")


def shelf_life_text(supplier_text="", months=None):
    if months:
        m = int(months)
        return f"{m // 12} years" if m % 12 == 0 and m >= 12 else f"{m} months"
    m = re.search(r"\d+(?:\.\d+)?\s*(?:years?|yrs?|months?)", cr.norm_text(supplier_text), re.I)
    return m.group(0) if m else cr.norm_text(supplier_text)


def expiry_from(mfg_text, shelf_text):
    mfg = cr.parse_date(mfg_text).date
    months = cr.parse_months(shelf_text or "")
    if not mfg or not months:
        return None
    y, m = divmod(mfg.month - 1 + months, 12)
    try:
        return (dt.date(mfg.year + y, m + 1, mfg.day) - dt.timedelta(days=1))
    except ValueError:
        return dt.date(mfg.year + y, m + 1, 28)


# ======================================================================
# Word template filling
# ======================================================================

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(tag):
    return f"{{{W}}}{tag}"


def _replace_in_paragraph(p, data):
    full = "".join(r.text for r in p.runs)
    if "{{" not in full:
        return
    new = re.sub(r"\{\{\s*(\w+)\s*\}\}", lambda m: str(data.get(m.group(1), "")), full)
    if new != full:
        for i, r in enumerate(p.runs):
            r.text = new if i == 0 else ""
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        if p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:      # justified short values get stretched
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    for s in doc.sections:
        for part in (s.header, s.footer):
            for p in part.paragraphs:
                yield p
            for t in part.tables:
                for row in t.rows:
                    for cell in row.cells:
                        yield from cell.paragraphs


def placeholders_in(doc):
    found = set()
    for p in _all_paragraphs(doc):
        found |= set(re.findall(r"\{\{\s*(\w+)\s*\}\}", "".join(r.text for r in p.runs)))
    return found


def _row_text(tr):
    return " ".join("".join(t.text or "" for t in tc.iter(_q("t"))) for tc in tr.findall(_q("tc"))).lower()


def _find_row(doc, test):
    for t in doc.tables:
        for i, tr in enumerate(t._tbl.findall(_q("tr"))):
            if test(_row_text(tr)):
                return t, i, tr
    return None, None, None


def _col_roles(tr):
    roles = []
    for tc in tr.findall(_q("tc")):
        txt = "".join(t.text or "" for t in tc.iter(_q("t"))).lower()
        if re.search(r"result", txt):
            roles.append("Results")
        elif re.search(r"method", txt):
            roles.append("Methods")
        elif re.search(r"standard|spec|limit", txt):
            roles.append("Standards")
        else:
            roles.append("Characteristic")
    return roles


def _set_tc_text(tc, text, bold=False, italic=False, size=21, align=None, font=None, indent=80):
    """Replace a cell's content with one paragraph/run, keeping the cell's own properties."""
    for p in tc.findall(_q("p"))[1:]:
        tc.remove(p)
    p = tc.find(_q("p"))
    if p is None:
        p = tc.makeelement(_q("p"), {})
        tc.append(p)
    for child in list(p):
        if child.tag != _q("pPr"):
            p.remove(child)
    ppr = p.find(_q("pPr"))
    if ppr is None:
        ppr = p.makeelement(_q("pPr"), {})
        p.insert(0, ppr)
    for tag in ("spacing", "ind", "jc", "rPr"):
        for el in ppr.findall(_q(tag)):
            ppr.remove(el)
    sp = ppr.makeelement(_q("spacing"), {_q("before"): "40", _q("after"): "40"})
    ind = ppr.makeelement(_q("ind"), {_q("left"): str(indent), _q("right"): "60"})
    ppr.append(sp)
    ppr.append(ind)
    if align:
        ppr.append(ppr.makeelement(_q("jc"), {_q("val"): align}))
    lines = str(text or "").split("\n")
    for li, line in enumerate(lines):
        r = p.makeelement(_q("r"), {})
        rpr = r.makeelement(_q("rPr"), {})
        if font:
            rpr.append(rpr.makeelement(_q("rFonts"), {_q("ascii"): font, _q("hAnsi"): font, _q("cs"): font}))
        if bold:
            rpr.append(rpr.makeelement(_q("b"), {}))
        if italic:
            rpr.append(rpr.makeelement(_q("i"), {}))
        rpr.append(rpr.makeelement(_q("sz"), {_q("val"): str(size)}))
        rpr.append(rpr.makeelement(_q("szCs"), {_q("val"): str(size)}))
        r.append(rpr)
        if li:
            r.append(r.makeelement(_q("br"), {}))
        t = r.makeelement(_q("t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
        t.text = line
        r.append(t)
        p.append(r)


def _clean_row(proto):
    """Copy a template row and strip shading/fixed heights so it works as a plain data row."""
    tr = copy.deepcopy(proto)
    trpr = tr.find(_q("trPr"))
    if trpr is not None:
        for h in trpr.findall(_q("trHeight")):
            h.set(_q("val"), "340")
            h.set(_q("hRule"), "atLeast")
        for el in trpr.findall(_q("tblHeader")):
            trpr.remove(el)
    for tc in tr.findall(_q("tc")):
        tcpr = tc.find(_q("tcPr"))
        if tcpr is not None:
            for el in tcpr.findall(_q("shd")) + tcpr.findall(_q("vMerge")):
                tcpr.remove(el)
            v = tcpr.find(_q("vAlign"))
            if v is None:
                tcpr.append(tcpr.makeelement(_q("vAlign"), {_q("val"): "center"}))
    return tr


def _data_row(proto, roles, row):
    tr = _clean_row(proto)
    for tc, role in zip(tr.findall(_q("tc")), roles):
        if role == "Characteristic":
            _set_tc_text(tc, row.get(role, ""), bold=True, font="Times New Roman", indent=240)
        else:
            _set_tc_text(tc, row.get(role, ""), align="left" if role == "Methods" else "center",
                         indent=120 if role == "Methods" else 60)
    return tr


def _banner_row(banner_tr, text):
    """Copy of the template's 'Microbiological Specifications' row with new text (same font and size)."""
    tr = copy.deepcopy(banner_tr)
    for p in tr.iter(_q("p")):
        runs = p.findall(_q("r"))
        if not runs:
            continue
        keep = next((r for r in runs if "".join(t.text or "" for t in r.iter(_q("t"))).strip()), runs[0])
        for r in runs:
            if r is not keep:
                p.remove(r)
        ts = keep.findall(_q("t"))
        for t in ts[1:]:
            keep.remove(t)
        if ts:
            ts[0].text = text
        break
    return tr


def _heading_row(proto, text):
    """A full-width sub-heading row (e.g. 'Heavy Metals')."""
    tr = _clean_row(proto)
    tcs = tr.findall(_q("tc"))
    total_span, total_w = 0, 0
    for tc in tcs:
        tcpr = tc.find(_q("tcPr"))
        gs = tcpr.find(_q("gridSpan")) if tcpr is not None else None
        total_span += int(gs.get(_q("val"))) if gs is not None else 1
        w = tcpr.find(_q("tcW")) if tcpr is not None else None
        total_w += int(w.get(_q("w"))) if w is not None and (w.get(_q("w")) or "").isdigit() else 0
    first = tcs[0]
    for tc in tcs[1:]:
        tr.remove(tc)
    tcpr = first.find(_q("tcPr"))
    gs = tcpr.find(_q("gridSpan"))
    if gs is None:
        gs = tcpr.makeelement(_q("gridSpan"), {})
        w = tcpr.find(_q("tcW"))
        (w.addnext(gs) if w is not None else tcpr.insert(0, gs))
    gs.set(_q("val"), str(total_span))
    w = tcpr.find(_q("tcW"))
    if w is not None and total_w:
        w.set(_q("w"), str(total_w))
    shd = tcpr.makeelement(_q("shd"), {_q("val"): "clear", _q("color"): "auto", _q("fill"): "EEF4F1"})
    tcpr.append(shd)
    _set_tc_text(first, text, bold=True, italic=True, size=20)
    return tr


def _widen_single_col_table(tbl, grid_source_tbl):
    """Give a 1-column banner table (e.g. 'Microbiological Specifications') the main table's columns."""
    grid = tbl.find(_q("tblGrid"))
    src = grid_source_tbl.find(_q("tblGrid"))
    if grid is None or src is None or len(grid.findall(_q("gridCol"))) != 1:
        return False
    n = len(src.findall(_q("gridCol")))
    grid.getparent().replace(grid, copy.deepcopy(src))
    for tr in tbl.findall(_q("tr")):
        tcs = tr.findall(_q("tc"))
        if len(tcs) == 1:
            tcpr = tcs[0].find(_q("tcPr"))
            if tcpr is None:
                tcpr = tcs[0].makeelement(_q("tcPr"), {})
                tcs[0].insert(0, tcpr)
            gs = tcpr.find(_q("gridSpan"))
            if gs is None:
                gs = tcpr.makeelement(_q("gridSpan"), {})
                w = tcpr.find(_q("tcW"))
                (w.addnext(gs) if w is not None else tcpr.insert(0, gs))
            gs.set(_q("val"), str(n))
    return True


# header label text -> field key (used when the value cell has no {{placeholder}})
_LABEL_KEYS = [("ProductName", r"product\s*name"), ("Brand", r"brand"), ("LotNo", r"lot|batch"),
               ("Quantity", r"quantity"), ("ManuDate", r"mfg|manufactur|production"),
               ("ExpiryDate", r"exp"), ("ShelfLife", r"shelf"), ("PlantPart", r"plant\s*part|part\s*used"),
               ("LatinName", r"latin|botanical"), ("Origin", r"origin"), ("Solvent", r"solvent")]


def _tc_text(tc):
    return "".join(t.text or "" for t in tc.iter(_q("t"))).strip()


def _pair_key(label_tc, value_tc):
    m = re.search(r"\{\{\s*(\w+)\s*\}\}", _tc_text(value_tc))
    if m:
        return m.group(1)
    label = _tc_text(label_tc).lower()
    for key, rx in _LABEL_KEYS:
        if re.search(rx, label):
            return key
    return None


def _blank_tc(tc):
    ps = tc.findall(_q("p"))
    for p in ps[1:]:
        tc.remove(p)
    if ps:
        for child in list(ps[0]):
            if child.tag != _q("pPr"):
                ps[0].remove(child)


def _put_tc(tc, paragraphs):
    for p in tc.findall(_q("p")):
        tc.remove(p)
    for p in paragraphs:
        tc.append(copy.deepcopy(p))


def hide_header_fields(doc, hide):
    """Remove label/value pairs (e.g. PlantPart, LatinName) from the header block and close up the gaps."""
    hide = set(hide or ())
    if not hide:
        return
    t, i_hdr, hdr = _find_row(doc, lambda s: "characteristic" in s and ("standard" in s or "spec" in s))
    if hdr is None:
        return
    trs = t._tbl.findall(_q("tr"))[:i_hdr]
    slots, pairs = [], []
    for tr in trs:
        tcs = tr.findall(_q("tc"))
        if len(tcs) < 2 or len(tcs) % 2:
            continue
        for j in range(0, len(tcs), 2):
            lt, vt = tcs[j], tcs[j + 1]
            slots.append((tr, lt, vt))
            if _tc_text(lt):
                pairs.append((_pair_key(lt, vt), [copy.deepcopy(p) for p in lt.findall(_q("p"))],
                              [copy.deepcopy(p) for p in vt.findall(_q("p"))]))
    kept = [p for p in pairs if p[0] not in hide]
    if len(kept) == len(pairs):
        return
    used_rows = set()
    for n, (tr, lt, vt) in enumerate(slots):
        if n < len(kept):
            _put_tc(lt, kept[n][1])
            _put_tc(vt, kept[n][2])
            used_rows.add(id(tr))
        else:
            _blank_tc(lt)
            _blank_tc(vt)
    for tr in {id(s[0]): s[0] for s in slots}.values():
        if id(tr) not in used_rows:
            tr.getparent().remove(tr)


def fill_template(template_bytes, fields, rows, doc_type="COA", group_headings=False, hide=()):
    """Fill a Skyherb template. rows: dicts with Group/Characteristic/Standards/Results/Methods/Include.
    hide: field keys to leave off the document, e.g. {"PlantPart", "LatinName"} for a non-botanical product."""
    from docx import Document
    doc = Document(io.BytesIO(template_bytes))
    hide_header_fields(doc, hide)
    for p in _all_paragraphs(doc):
        _replace_in_paragraph(p, fields)

    inc = [r for r in rows if r.get("Include", True) and str(r.get("Characteristic", "")).strip()]
    if doc_type != "COA":
        inc = [dict(r, Results="") for r in inc]
    general = [r for r in inc if r.get("Group") != "Microbiological"]
    micro = [r for r in inc if r.get("Group") == "Microbiological"]

    t_main, i_hdr, hdr = _find_row(doc, lambda s: "characteristic" in s and ("standard" in s or "spec" in s))
    if hdr is None:
        raise ValueError("The template has no 'Characteristic | Standards | …' header row.")
    roles = _col_roles(hdr)
    t_mic, i_mic, mic = _find_row(doc, lambda s: "microbiolog" in s and "characteristic" not in s)

    # general rows go right below the header row (works whether or not more rows follow)
    if mic is not None and t_mic is not t_main:
        _widen_single_col_table(t_mic._tbl, t_main._tbl)
        mic = t_mic._tbl.findall(_q("tr"))[i_mic]
    anchor, last_group = hdr, None
    for r in general:
        grp = r.get("Group", "")
        if grp != last_group:
            if grp in BANNER_GROUPS:
                h = _banner_row(mic, grp) if mic is not None else _heading_row(hdr, grp)
            elif group_headings:
                h = _heading_row(hdr, grp)
            else:
                h = None
            if h is not None:
                anchor.addnext(h)
                anchor = h
            last_group = grp
        new = _data_row(hdr, roles, r)
        anchor.addnext(new)
        anchor = new

    if micro:
        if mic is not None:
            anchor = mic
        else:                                        # no micro banner: add a sub-heading in the main table
            h = _heading_row(hdr, "Microbiological")
            anchor.addnext(h)
            anchor = h
        for r in micro:
            new = _data_row(hdr, roles, r)
            anchor.addnext(new)
            anchor = new
    elif mic is not None and t_mic is not t_main and len(t_mic.rows) == 1:
        t_mic._tbl.getparent().remove(t_mic._tbl)   # no micro tests: drop the empty banner

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def docx_to_pdf(docx_bytes):
    """Convert with LibreOffice when it's installed (returns None otherwise)."""
    import os
    import shutil
    import subprocess
    import tempfile
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return None
    with tempfile.TemporaryDirectory() as d:
        src = os.path.join(d, "doc.docx")
        with open(src, "wb") as f:
            f.write(docx_bytes)
        try:
            subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", d, src],
                           check=True, timeout=120, capture_output=True)
            with open(os.path.join(d, "doc.pdf"), "rb") as f:
                return f.read()
        except Exception:
            return None
