"""
COA Review engine for SN Tools.

Compares a supplier COA against a Skyherb specification and lists the problems.

    Numbers, units, dates and missing items  -> decided by code in this file
    Judgement calls (method suitability,     -> decided by the Claude API
    vague wording, product/spec mismatch)       (optional, see ai_* functions)

Every rule lives in this file, so when the tool misses a problem, add a rule here.
"""
import base64
import datetime as dt
import difflib
import io
import json
import re
from dataclasses import dataclass

# ======================================================================
# Text normalisation
# ======================================================================

_SUP = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻", "0123456789+-")
_CJK = re.compile(r"[　-鿿가-힯]")


def norm_text(s):
    """Unify symbols so the parsers only need to know one spelling."""
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"(?<![\d.^\w])[³‡]\s*(?=\d)", "≥", s)           # Symbol-font ≥ (see below)
    s = re.sub(r"([⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻]+)", lambda m: "^" + m.group(1).translate(_SUP), s)
    for a, b in (("×", "x"), ("＊", "*"), ("μ", "µ"), (" ", " "), ("≦", "≤"), ("≧", "≥"),
                 ("＜", "<"), ("＞", ">"), ("：", ":"), ("％", "%"), ("，", ","), ("（", "("), ("）", ")"),
                 ("⩽", "≤"), ("⩾", "≥"), ("‰", " per mille")):
        s = s.replace(a, b)
    s = re.sub(r"[–—~～−]", "-", s)
    # PDFs made with the Word "Symbol" font: ≤ arrives as £ and ≥ as ³ or ‡
    s = re.sub(r"£\s*(?=\d)", "≤", s)
    s = re.sub(r"(?<![\d.^])[³‡]\s*(?=\d)", "≥", s)
    # superscript exponents flattened by PDF text extraction: "10³cfu/g" arrives as "103cfu/g"
    # ("Less than 10³cfu/g" arrives as "Less than 103cfu/g"; a plain "105 cfu/g" count is left alone)
    s = re.sub(r"(<|≤|than|nmt|max\.?|[x*])\s*10\s?([2-9])(?=\s*(?:cfu|mpn|/g|/ml))", r"\1 10^\2", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    return s


_REF = re.compile(r"\b(usp|nf|ep|bp|jp|chp|aoac|iso|ich|fcc|jecfa)\s*(?:<\s*\d+[a-z]?\s*>|\d+(?:\.\d+)*)", re.I)


def strip_refs(s):
    """'Meets USP <467>' -> 'Meets USP' so chapter numbers aren't read as values."""
    return _REF.sub(lambda m: m.group(1), s)


def clean_name(s):
    s = norm_text(s).lower()
    s = re.sub(r"[^\w%&/+ .-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" .-")
    return s


# ======================================================================
# Numbers, units, limits, results
# ======================================================================

_NUM = (r"10\s*\^\s*(?P<e3>[+-]?\d+)"
        r"|(?P<a>\d+(?:[.,]\d+)*)(?:\s*[x*]\s*10\s*\^\s*(?P<e1>[+-]?\d+)|\s*[eE]\s*(?P<e2>[+-]?\d+))?")
NUM_RE = re.compile(_NUM)
_NUMX = re.sub(r"\?P<\w+>", "", _NUM)      # same pattern without group names, for embedding twice


def _to_float(a):
    if "," in a:
        if re.fullmatch(r"\d{1,3}(,\d{3})+(\.\d+)?", a):
            a = a.replace(",", "")
        elif re.fullmatch(r"\d+,\d{1,2}", a):          # decimal comma, e.g. 1,5
            a = a.replace(",", ".")
        else:
            a = a.replace(",", "")
    try:
        return float(a)
    except ValueError:
        return None


def _match_value(m):
    if m.group("e3") is not None:
        return 10.0 ** int(m.group("e3"))
    v = _to_float(m.group("a"))
    if v is None:
        return None
    e = m.group("e1") or m.group("e2")
    return v * 10.0 ** int(e) if e else v


# unit -> (family, factor to the family's base unit)
UNITS = {
    "%": ("frac", 1e4), "%w/w": ("frac", 1e4), "g/100g": ("frac", 1e4), "ppm": ("frac", 1.0),
    "mg/kg": ("frac", 1.0), "µg/g": ("frac", 1.0), "ug/g": ("frac", 1.0), "mcg/g": ("frac", 1.0),
    "mg/g": ("frac", 1e3), "mg/100g": ("frac", 10.0), "ppb": ("frac", 1e-3), "µg/kg": ("frac", 1e-3),
    "ug/kg": ("frac", 1e-3), "mcg/kg": ("frac", 1e-3), "ppt": ("frac", 1e-6), "ng/g": ("frac", 1e-3),
    "cfu/g": ("cfu", 1.0), "mpn/g": ("mpn", 1.0), "mpn/100g": ("mpn", 0.01),
    "cfu/ml": ("cfu_ml", 1.0), "mpn/ml": ("mpn_ml", 1.0),
    "cfu/10g": ("presence", 1.0), "cfu/25g": ("presence", 1.0), "/25g": ("presence", 1.0),
    "/10g": ("presence", 1.0), "/g": ("presence", 1.0),
    "g/ml": ("density", 1.0), "g/cm3": ("density", 1.0), "g/cc": ("density", 1.0), "kg/l": ("density", 1.0),
    "g/l": ("density", 1e-3),
}
_UNIT_PATTERNS = sorted(UNITS, key=len, reverse=True)
_UNIT_RE = re.compile("|".join(
    r"\s*".join(re.escape(ch) for ch in u.replace(" ", "")) for u in _UNIT_PATTERNS), re.I)


def norm_unit(u):
    if not u:
        return ""
    u = norm_text(u).lower().replace(" ", "")
    u = u.replace("μ", "µ").replace("ug/", "µg/").replace("mcg/", "µg/")
    u = {"cm³": "g/cm3", "g/cm^3": "g/cm3", "w/w%": "%", "%(w/w)": "%"}.get(u, u)
    return u


def find_unit(text):
    m = _UNIT_RE.search(norm_text(text))
    return norm_unit(m.group(0)) if m else ""


def convert(value, from_unit, to_unit):
    """Convert between compatible units. Returns None when the units can't be compared."""
    f, t = norm_unit(from_unit), norm_unit(to_unit)
    if value is None:
        return None
    if not f or not t or f == t:
        return value
    if f in UNITS and t in UNITS and UNITS[f][0] == UNITS[t][0]:
        return value * UNITS[f][1] / UNITS[t][1]
    return None


def same_family(u1, u2):
    a, b = norm_unit(u1), norm_unit(u2)
    if not a or not b or a == b:
        return True
    return a in UNITS and b in UNITS and UNITS[a][0] == UNITS[b][0]


_ND = re.compile(r"\b(not\s*detect\w*|n\.?\s?d\.?|none\s+detected|negative|neg\.?|absent|absence|"
                 r"undetect(?:ed|able)|not\s+found|nil|free|bdl|below\s+(?:detection|lod|loq))\b|<\s*lo[dq]\b", re.I)
_POS = re.compile(r"\b(positive|present|detected|found)\b", re.I)
_PASS = re.compile(r"\b(pass(?:es|ed)?|conform(?:s|ed|ing|ity)?|compl(?:y|ies|ied|iant)|meets?|qualified|ok|"
                   r"in\s+compliance|within\s+(?:the\s+)?(?:limits?|specs?|specification)|satisfactory|acceptable|"
                   r"accord(?:s|ing)?|qualify)\b|[√✓✔]", re.I)

_MAX_OPS = r"(?<!not\s)(?<!no\s)(≤|<=|=<|<|less\s+than|not\s+more\s+than|no\s+more\s+than|nmt|not\s+exceed(?:ing)?|max(?:imum)?\.?|up\s+to|below)"
_MIN_OPS = r"(?<!not\s)(?<!no\s)(≥|>=|=>|>|not\s+less\s+than|no\s+less\s+than|nlt|more\s+than|greater\s+than|min(?:imum)?\.?|at\s+least|above)"


@dataclass
class Limit:
    min: float = None
    min_incl: bool = True
    max: float = None
    max_incl: bool = True
    unit: str = ""
    expected: str = ""        # "negative", or a descriptive text such as "White powder"
    text: str = ""

    @property
    def numeric(self):
        return self.min is not None or self.max is not None

    def describe(self):
        u = f" {self.unit}" if self.unit else ""
        if self.min is not None and self.max is not None:
            return f"{fmt_num(self.min)}–{fmt_num(self.max)}{u}"
        if self.max is not None:
            return f"{'≤' if self.max_incl else '<'} {fmt_num(self.max)}{u}"
        if self.min is not None:
            return f"{'≥' if self.min_incl else '>'} {fmt_num(self.min)}{u}"
        if self.expected == "negative":
            t = (self.text or "").strip()
            return t if t and re.search(r"negative|absent|not\s*detect|n\.?d\.?", t, re.I) else "Negative"
        return self.expected or self.text or "—"


def fmt_num(v):
    if v is None:
        return ""
    if v != 0 and (abs(v) >= 1e5 or abs(v) < 1e-3):
        exp = int(f"{v:e}".split("e")[1])
        mant = v / 10 ** exp
        return f"10^{exp}" if abs(mant - 1) < 1e-9 else f"{mant:g}×10^{exp}"
    return f"{v:,.6g}" if abs(v) >= 1000 else f"{v:.6g}"


def parse_limit(text, default_unit=""):
    """Read a specification text such as 'NMT 1.0 ppm', '≥97.0%', '4.5-7.0', 'Negative/25g'."""
    raw = "" if text is None else str(text)
    t = strip_refs(norm_text(raw))
    lim = Limit(text=raw.strip(), unit=find_unit(t) or norm_unit(default_unit))
    if not t:
        return lim
    low = t.lower()

    # range: "4.5-7.0", "98.0%-102.0%", "between 3 and 5"
    m = re.search(rf"(?<![\w.^])(?:{_NUMX})\s*(?:%|ppm)?\s*(?:-|to|and)\s*(?:{_NUMX})(?![\w.])", low)
    if m and not re.search(r"\bmesh\b", low[: m.start()]):
        nums = [n for n in NUM_RE.finditer(m.group(0))]
        if len(nums) >= 2:
            a, b = _match_value(nums[0]), _match_value(nums[-1])
            if a is not None and b is not None and a < b:
                lim.min, lim.max = a, b
                return lim

    found = False
    for m in re.finditer(rf"{_MAX_OPS}\s*:?\s*({_NUM})", low):
        n = NUM_RE.search(m.group(0)[len(m.group(1)):])
        if n:
            lim.max, lim.max_incl, found = _match_value(n), m.group(1).strip() not in ("<", "less than", "below"), True
            break
    for m in re.finditer(rf"{_MIN_OPS}\s*:?\s*({_NUM})", low):
        n = NUM_RE.search(m.group(0)[len(m.group(1)):])
        if n:
            lim.min, lim.min_incl, found = _match_value(n), m.group(1).strip() not in (">", "more than",
                                                                                     "greater than", "above"), True
            break
    if not found:   # trailing "1.0 ppm max" / "97% min"
        m = re.search(rf"({_NUM})\s*(?:[%\w/µ]+\s*)?\b(?P<mm>max(?:imum)?|min(?:imum)?)\b", low)
        if m:
            n = NUM_RE.search(m.group(0))
            if n:
                if m.group("mm").startswith("max"):
                    lim.max = _match_value(n)
                else:
                    lim.min = _match_value(n)
                found = True
    if found:
        return lim

    if _ND.search(t):
        lim.expected = "negative"
    elif not NUM_RE.search(t) or re.search(r"[a-z]{4,}", low):
        lim.expected = norm_text(raw)     # descriptive spec ("White powder", "Meets USP <467>")
    return lim


@dataclass
class Result:
    op: str = ""               # "=", "<", "≤", ">", "≥"
    value: float = None
    unit: str = ""
    qual: str = ""             # "nd", "pos", "pass", "" (plain number) or "text"
    text: str = ""


def parse_result(text, default_unit=""):
    raw = "" if text is None else str(text)
    t = strip_refs(norm_text(raw))
    low = t.lower()
    r = Result(text=raw.strip(), unit=find_unit(t) or norm_unit(default_unit))
    if not t or low in ("-", "/", "n/a", "na", "--", "none"):
        return r

    if _ND.search(t):
        r.qual = "nd"
    elif _POS.search(t) and not re.search(r"\bnot\b", low):
        r.qual = "pos"

    m = re.search(rf"({_MAX_OPS}|{_MIN_OPS})?\s*:?\s*({_NUM})", low)
    if m and not (r.qual == "nd" and not re.search(r"[<≤]|less|below", low)):
        n = NUM_RE.search(low[m.start():])
        if n:
            r.value = _match_value(n)
            op = (m.group(1) or "").strip()
            if op in ("<", "less than", "below", "≤", "<=", "=<", "nmt", "not more than", "max", "max.", "maximum"):
                r.op = "<" if op in ("<", "less than", "below") else "≤"
            elif op in (">", "more than", "greater than", "above", "≥", ">=", "=>", "nlt", "not less than",
                        "min", "min.", "minimum", "at least"):
                r.op = ">" if op in (">", "more than", "greater than", "above") else "≥"
            else:
                r.op = "="
            if r.qual == "nd" and not r.op.startswith("<") and r.op != "≤":
                r.op = "<"
            return r
    if not r.qual and _PASS.search(t):
        r.qual = "pass"
    elif not r.qual:
        r.qual = "text"
    return r


# ======================================================================
# Test names
# ======================================================================

# canonical name -> patterns (searched on the cleaned, lower-case name)
ALIASES = [
    ("yeast & mold", r"yeasts?\s*(?:&|and|/|\+)\s*mou?lds?|mou?lds?\s*(?:&|and|/|\+)\s*yeasts?|tymc|\by\s*&\s*m\b|combined\s+yeast"),
    ("total plate count", r"total\s+(?:plate|aerobic|bacteri\w*|viable)\s*(?:microbial\s+)?count|aerobic\s+plate|aerobic\s+(?:bacteri\w*|microbial|mesophilic)\s+count|\btpc\b|\bapc\b|\btamc\b|plate\s+count|total\s+aerobic"),
    ("e. coli", r"e\.?\s?coli|escherichia"),
    ("salmonella", r"salmonella"),
    ("staphylococcus aureus", r"staph\w*|s\.\s?aureus"),
    ("coliforms", r"coliform"),
    ("enterobacteriaceae", r"enterobact"),
    ("listeria", r"listeria"),
    ("yeast", r"^yeasts?\b"),
    ("mold", r"^mou?lds?\b"),
    ("heavy metals", r"heavy\s+metals?"),
    ("lead", r"\blead\b|\bpb\b"),
    ("arsenic", r"arsenic|\bas\b"),
    ("cadmium", r"cadmium|\bcd\b"),
    ("mercury", r"mercury|\bhg\b"),
    ("loss on drying", r"loss\s+on\s+drying|\blod\b|moisture|water\s+content|drying\s+loss"),
    ("sulfated ash", r"sul(?:ph|f)ated\s+ash|residue\s+on\s+ignition"),
    ("ash", r"\bash\b"),
    ("bulk density", r"bulk\s+density|loose\s+density|\bdensity\b(?!.*tap)"),
    ("tapped density", r"tap+(?:ed)?\s+density"),
    ("particle size", r"mesh|particle\s+size|sieve|fineness"),
    ("appearance", r"appearance|description|colou?r\b|characters?\b"),
    ("odor & taste", r"odou?r|taste|aroma|smell|flavou?r"),
    ("identification", r"identif\w*|identity|\bid\b"),
    ("pesticides", r"pesticid"),
    ("residual solvents", r"residual\s+solvents?|solvents?\s+residue"),
    ("aflatoxins", r"aflatoxin"),
    ("pah", r"\bpah|benzo"),
    ("ph", r"^ph\b|\bph\s+value"),
    ("solubility", r"solub"),
]
_ALIAS_RE = [(c, re.compile(p, re.I)) for c, p in ALIASES]

SAFETY = {"heavy metals", "microbiology", "pesticides", "residual solvents", "contaminants"}
CATEGORY_OF = {
    "yeast & mold": "microbiology", "total plate count": "microbiology", "e. coli": "microbiology",
    "salmonella": "microbiology", "staphylococcus aureus": "microbiology", "coliforms": "microbiology",
    "enterobacteriaceae": "microbiology", "listeria": "microbiology", "yeast": "microbiology",
    "mold": "microbiology", "heavy metals": "heavy metals", "lead": "heavy metals", "arsenic": "heavy metals",
    "cadmium": "heavy metals", "mercury": "heavy metals", "loss on drying": "physical", "ash": "physical",
    "sulfated ash": "physical", "bulk density": "physical", "tapped density": "physical",
    "particle size": "physical", "appearance": "physical", "odor & taste": "physical",
    "identification": "identification", "pesticides": "pesticides", "residual solvents": "residual solvents",
    "aflatoxins": "contaminants", "pah": "contaminants", "ph": "physical", "solubility": "physical",
}
_SOLVENTS = r"methanol|ethanol|hexane|acetone|ethyl\s+acetate|isopropanol|dichloromethane|chloroform|benzene|toluene"
_HEADING_WORDS = re.compile(r"^(identification|purity|assay|physical|chemical|heavy\s+metals?|microbiolog\w*|"
                            r"micro\w*\s+tests?|contaminants?|chemical\s+contaminants|residual\s+solvents?|"
                            r"pesticides?|physical\s*(?:&|and)\s*chemical\s+(?:properties|tests?)|"
                            r"organoleptic\w*|general|others?|test\s+items?)$", re.I)


_ALIAS_SQ = [(c, re.compile(p.replace(r"\s+", r"\s*"), re.I)) for c, p in ALIASES]


def canon(name):
    n = clean_name(name)
    for c, rx in _ALIAS_RE:
        if rx.search(n):
            return c
    sq = n.replace(" ", "")
    for c, rx in _ALIAS_SQ:                 # names stored without spaces, e.g. "LossonDrying"
        if len(sq) > 3 and rx.search(sq):
            return c
    return n


def category_of(name, given=""):
    g = clean_name(given)
    if g:
        for key in ("microbiolog", "heavy", "pesticid", "solvent", "contamin", "physical", "assay", "identif"):
            if key in g:
                return {"microbiolog": "microbiology", "heavy": "heavy metals", "pesticid": "pesticides",
                        "solvent": "residual solvents", "contamin": "contaminants", "physical": "physical",
                        "assay": "assay", "identif": "identification"}[key]
        return g
    c = canon(name)
    if c in CATEGORY_OF:
        return CATEGORY_OF[c]
    if re.search(_SOLVENTS, clean_name(name)):
        return "residual solvents"
    return "assay"


def _tokens(n):
    return {re.sub(r"s$", "", w) for w in re.findall(r"[a-z0-9]+", n) if w not in ("total", "content", "the", "of")}


def name_score(a, b):
    ca, cb = canon(a), canon(b)
    if ca == cb:
        return 1.0
    if ca in CATEGORY_OF or cb in CATEGORY_OF:   # one is a known test and they differ
        return 0.0
    ta, tb = _tokens(ca), _tokens(cb)
    if ta and tb and (ta <= tb or tb <= ta):
        return 0.9
    return difflib.SequenceMatcher(None, ca, cb).ratio()


# ======================================================================
# Dates
# ======================================================================

_MONTHS = {m: i for i, ms in enumerate(
    [("jan", "january"), ("feb", "february"), ("mar", "march"), ("apr", "april"), ("may",), ("jun", "june"),
     ("jul", "july"), ("aug", "august"), ("sep", "sept", "september"), ("oct", "october"),
     ("nov", "november"), ("dec", "december")], 1) for m in ms}


@dataclass
class ParsedDate:
    date: dt.date = None
    ambiguous: bool = False
    text: str = ""


def _mk(y, m, d):
    try:
        y = int(y)
        y = y + 2000 if y < 100 else y
        return dt.date(y, int(m), int(d))
    except (ValueError, TypeError):
        return None


def parse_date(text, end_of_month=False):
    t = norm_text(text)
    out = ParsedDate(text=t)
    if not t:
        return out
    low = t.lower()
    m = re.search(r"\b((?:19|20)\d{2})\s*[./\-年]\s*(\d{1,2})\s*[./\-月]\s*(\d{1,2})", low)
    if m:
        out.date = _mk(*m.groups())
        return out
    m = re.search(r"\b((?:19|20)\d{2})(\d{2})(\d{2})\b", low)
    if m and 1 <= int(m.group(2)) <= 12 and 1 <= int(m.group(3)) <= 31:
        out.date = _mk(*m.groups())
        return out
    m = re.search(r"\b(\d{1,2})(?:st|nd|rd|th)?[\s\-./,]*([a-z]{3,9})\.?[\s\-./,]*((?:19|20)?\d{2})\b", low)
    if m and m.group(2) in _MONTHS:
        out.date = _mk(m.group(3), _MONTHS[m.group(2)], m.group(1))
        if out.date:
            return out
    m = re.search(r"\b([a-z]{3,9})\.?[\s\-./]*(\d{1,2})(?:st|nd|rd|th)?,?[\s\-./]*((?:19|20)\d{2})\b", low)
    if m and m.group(1) in _MONTHS:
        out.date = _mk(m.group(3), _MONTHS[m.group(1)], m.group(2))
        if out.date:
            return out
    m = re.search(r"\b(\d{1,2})[./\-](\d{1,2})[./\-]((?:19|20)?\d{2})\b", low)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), m.group(3)
        if a > 12:
            out.date = _mk(y, b, a)
        elif b > 12:
            out.date = _mk(y, a, b)
        else:
            out.date, out.ambiguous = _mk(y, b, a), a != b     # read as day/month
        return out
    m = re.search(r"\b([a-z]{3,9})\.?[\s\-./,]*((?:19|20)\d{2})\b", low)
    if m and m.group(1) in _MONTHS:
        y, mo = int(m.group(2)), _MONTHS[m.group(1)]
        d = 1
        if end_of_month:
            d = (dt.date(y + (mo == 12), mo % 12 + 1, 1) - dt.timedelta(days=1)).day
        out.date = _mk(y, mo, d)
        return out
    m = re.search(r"\b(\d{1,2})[./\-]((?:19|20)\d{2})\b", low)
    if m and 1 <= int(m.group(1)) <= 12:
        out.date = _mk(m.group(2), m.group(1), 1)
    return out


def parse_months(text):
    t = clean_name(text)
    m = re.search(r"(\d+(?:\.\d+)?)\s*(years?|yrs?|months?|mos?)\b", t)
    if not m:
        return None
    v = float(m.group(1))
    return int(round(v * 12)) if m.group(2).startswith("y") else int(round(v))


def months_between(a, b):
    return (b.year - a.year) * 12 + (b.month - a.month) + (b.day - a.day) / 30.0


# ======================================================================
# Reading a COA without AI (text PDFs)
# ======================================================================

INFO_FIELDS = [
    # key, label regex, needs a colon after the label
    ("mfg_date", r"manufactur(?:ed|ing|e)?\s+date|date\s+of\s+manufactur\w*|mfg\.?\s*date|mfd\.?\s*date|"
                 r"production\s+date|produced\s+date|date\s+of\s+production|prod\.?\s*date|manu\.?\s*date", False),
    ("expiry_date", r"expir\w*\s+date|exp\.?\s*date|expiry|expiration|best\s+before|use\s+by|valid\s+until", False),
    ("retest_date", r"re-?test\s+date", False),
    ("test_date", r"test(?:ing)?\s+date|date\s+of\s+(?:analysis|test(?:ing)?|examination)|analy[sz](?:is|ed)\s+date|"
                  r"inspection\s+date|examination\s+date|check(?:ing)?\s+date|analy[sz]ed\s+on", False),
    ("report_date", r"report\s+date|issue\s+date|date\s+of\s+issue|issued\s+(?:by\s+[\w &.]{1,40}?\s+)?on|"
                    r"release\s+date|approv(?:al|ed)\s+date|date\s+of\s+report|date\s+of\s+release", False),
    ("lot_no", r"lot\s*(?:no|number|#|num)?\.?|batch\s*(?:no|number|#|num)?\.?(?!\s*(?:quantity|size))", False),
    ("product_name", r"trade\s+name|product\s+name|name\s+of\s+(?:the\s+)?product|item\s+name|commodity|"
                     r"material\s+name|sample\s+name|product", True),
    ("shelf_life", r"shelf\s*-?\s*life", False),
    ("quantity", r"batch\s+quantity|lot\s+quantity|quantity|batch\s+size|lot\s+size|qty", True),
    ("latin_name", r"latin\s+name|botanical\s+(?:name|source)|scientific\s+name|plant\s+source|source\s+plant", False),
    ("plant_part", r"parts?\s+used|plant\s+parts?|used\s+parts?|parts?\s+of\s+(?:the\s+)?plant", False),
    ("solvent", r"extraction\s+solvents?|solvents?\s+(?:used|of\s+extraction|for\s+extraction)|extract(?:ion)?\s+solvents?|(?<!residual\s)solvents?", True),
    ("carrier", r"carriers?|excipients?", True),
    ("ratio", r"extract(?:ion)?\s+ratio|ratio", True),
    ("supplier", r"manufacturer|supplier|producer|manufactured\s+by|produced\s+by", True),
    ("_stop", r"package\s+size|packing|packaging|net\s+weight|storage|specification|conclusion|remarks?", True),
]
_INFO_RE = [(k, re.compile(rf"(?<![\w])(?:{p})(?![\w])\s*(?P<sep>[:#.]|\s)", re.I), strict) for k, p, strict in INFO_FIELDS]
_APPROVAL = re.compile(r"approved\s+by|signature|signed|authori[sz]ed|q[ac]\s+manager|quality\s+(?:assurance|control)"
                       r"\s+(?:manager|supervisor|director|dept\w*)|reviewed\s+by|checked\s+by|issued\s+by|"
                       r"q[ac]\s+dep(?:artmen)?t|analy[sz]ed\s+by|tested\s+by|inspector", re.I)


def _scan_info_line(line, found):
    hits = []
    for key, rx, strict in _INFO_RE:
        for m in rx.finditer(line):
            if strict and m.group("sep") not in (":", "#"):
                hits.append((m.start(), m.end(), "_stop"))     # not a value, but ends the one before
                continue
            hits.append((m.start(), m.end(), key))
    if not hits:
        return
    hits.sort(key=lambda h: (h[0], -(h[1] - h[0])))
    kept, last_end = [], -1
    for h in hits:                       # drop labels that overlap a longer/earlier one
        if h[0] >= last_end:
            kept.append(h)
            last_end = h[1]
    for i, (s, e, key) in enumerate(kept):
        nxt = kept[i + 1][0] if i + 1 < len(kept) else len(line)
        val = line[e:nxt].strip(" :#.,;|\t")
        if val and key != "_stop" and key not in found and len(val) < 120:
            found[key] = val


_HDR = {
    "name": re.compile(r"\b(test(?:ing)?\s*items?|items?|tests?|analy\w*|parameters?|characteristics?|"
                       r"determinations?|check\s*items?|attributes?|properties)\b", re.I),
    "spec": re.compile(r"\b(spec\w*|standards?|limits?|requirements?|criteria|acceptance|norms?)\b", re.I),
    "result": re.compile(r"\b(results?|values?|found|observations?|actual|test\s+data|data)\b", re.I),
    "method": re.compile(r"\b(methods?|references?|procedures?)\b", re.I),
    "unit": re.compile(r"^\s*units?\s*$", re.I),
}


def _header_map(row):
    cols = {}
    for i, cell in enumerate(row):
        c = norm_text(cell)
        if not c or len(c) > 40:
            continue
        for key in ("unit", "method", "result", "spec", "name"):
            if key not in cols and _HDR[key].search(c):
                cols[key] = i
                break
    if "name" not in cols and len(cols) >= 2 and 0 not in cols.values():
        cols["name"] = 0
    return cols if len(cols) >= 3 and "name" in cols and ("result" in cols or "spec" in cols) else None


def _tests_from_tables(tables):
    tests, header, category = [], None, ""
    for table in tables:
        for row in table:
            cells = [norm_text(c) for c in row]
            if not any(cells):
                continue
            hm = _header_map(cells)
            if hm:
                header = hm
                continue
            if header is None:
                continue
            filled = [c for c in cells if c]
            name = cells[header["name"]] if header["name"] < len(cells) else ""
            if len(filled) == 1 and (name or filled[0]):
                category = filled[0]                      # section heading row
                continue
            if not name or _HEADING_WORDS.match(name) and len(filled) <= 1:
                continue
            get = lambda k: cells[header[k]] if k in header and header[k] < len(cells) else ""
            t = {"name": name, "category": category, "spec": get("spec"), "result": get("result"),
                 "unit": get("unit"), "method": get("method")}
            if t["spec"] or t["result"]:
                tests.append(t)
    return tests


def _line_cells(words, gap=7.0):
    """Group the words of one text line into cells separated by wide gaps."""
    cells, cur = [], None
    for w in sorted(words, key=lambda w: w["x0"]):
        if cur and w["x0"] - cur["x1"] <= gap:
            cur["text"] += " " + w["text"]
            cur["x1"] = w["x1"]
        else:
            cur = {"text": w["text"], "x0": w["x0"], "x1": w["x1"]}
            cells.append(cur)
    return cells


def _rows_from_words(page):
    """Rebuild table rows from word positions (for COAs whose tables have no ruling lines)."""
    words = page.extract_words(keep_blank_chars=False, use_text_flow=False, x_tolerance=1.5) or []
    lines = {}
    for w in words:
        key = round(w["top"] / 3)
        lines.setdefault(key, []).append(w)
    rows, bounds = [], None
    for key in sorted(lines):
        cells = _line_cells(lines[key])
        texts = [norm_text(c["text"]) for c in cells]
        if _header_map(texts):
            centers = [(c["x0"] + c["x1"]) / 2 for c in cells]
            bounds = [(centers[i] + centers[i + 1]) / 2 for i in range(len(centers) - 1)]
            rows.append(texts)
            continue
        if bounds is None:
            continue
        out = [""] * (len(bounds) + 1)
        for w in sorted(lines[key], key=lambda w: w["x0"]):
            cx = (w["x0"] + w["x1"]) / 2
            col = sum(cx > b for b in bounds)
            out[col] = (out[col] + " " + w["text"]).strip()
        rows.append(out)
    return rows


def extract_text_coa(pdf_bytes):
    """Read a text-based COA PDF: header info + test rows. Returns a COA dict."""
    import pdfplumber
    info, lines, tables, n_chars = {}, [], [], 0
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n_pages = len(pdf.pages)
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1.5) or ""      # 1.5 keeps spaces some PDFs lose at 3
            n_chars += len(text.strip())
            lines += [norm_text(l) for l in text.splitlines() if l.strip()]
            page_tables = page.extract_tables({"text_x_tolerance": 1.5}) or []
            if not _tests_from_tables(page_tables):
                try:
                    rebuilt = _rows_from_words(page)       # tables without ruling lines
                    if rebuilt:
                        page_tables = page_tables + [rebuilt]
                except Exception:
                    pass
            tables += page_tables
    kv_lines = []
    for table in tables:                                  # key/value cells, e.g. | Lot No. | 2401 |
        for row in table:
            cells = [norm_text(c) for c in row if c]
            for i in range(0, len(cells) - 1, 2):
                kv_lines.append(f"{cells[i]}: {cells[i + 1]}")
    for line in kv_lines + lines:                         # table cells first, then free text
        _scan_info_line(line, info)
    lines += kv_lines
    tests = _tests_from_tables(tables)
    raw = "\n".join(lines)
    return {
        "source": "text", "info": info, "tests": tests, "raw_text": raw, "pages": n_pages,
        "scanned": n_chars < 40, "approval_found": bool(_APPROVAL.search(raw)), "signed": None,
        "warnings": [],
    }


# ======================================================================
# Spec library
# ======================================================================

SPEC_COLUMNS = ["Product", "Test", "Category", "Min", "Max", "Unit", "Expected", "Method", "Required", "Notes"]
PRODUCT_COLUMNS = ["Product", "Aliases", "Latin name", "Plant part", "Extraction solvent", "Carrier",
                   "Shelf life (months)", "Notes"]


def _norm_cols(df, wanted):
    lookup = {re.sub(r"[^a-z]", "", c.lower()): c for c in wanted}
    df = df.rename(columns={c: lookup.get(re.sub(r"[^a-z]", "", str(c).lower()), c) for c in df.columns})
    for c in wanted:
        if c not in df.columns:
            df[c] = None
    return df[wanted]


def load_spec_library(src):
    """src: path or bytes of an .xlsx with a 'Specs' sheet (and optional 'Products' sheet)."""
    import pandas as pd
    data = io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src
    sheets = pd.read_excel(data, sheet_name=None, dtype=object)
    by_name = {k.strip().lower(): v for k, v in sheets.items()}
    specs = by_name.get("specs") if "specs" in by_name else next(iter(sheets.values()))
    specs = _norm_cols(specs, SPEC_COLUMNS).dropna(subset=["Product", "Test"], how="any")
    specs["Product"] = specs["Product"].astype(str).str.strip()
    products = by_name.get("products")
    if products is None:
        products = pd.DataFrame({"Product": sorted(specs["Product"].unique())})
    products = _norm_cols(products, PRODUCT_COLUMNS).dropna(subset=["Product"])
    products["Product"] = products["Product"].astype(str).str.strip()
    missing = set(specs["Product"]) - set(products["Product"])
    if missing:
        products = pd.concat([products, pd.DataFrame({"Product": sorted(missing)})], ignore_index=True)
        products = _norm_cols(products, PRODUCT_COLUMNS)
    return specs.reset_index(drop=True), products.reset_index(drop=True)


def _num(v):
    if v is None:
        return None
    try:
        import math
        f = float(str(v).replace(",", "").strip())
        return None if math.isnan(f) else f
    except ValueError:
        r = parse_result(v)
        return r.value


def _txt(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def spec_rows_for(specs, product):
    rows = []
    for _, r in specs[specs["Product"] == product].iterrows():
        lim = Limit(min=_num(r["Min"]), max=_num(r["Max"]), unit=norm_unit(_txt(r["Unit"])))
        exp = _txt(r["Expected"])
        if exp and not lim.numeric:
            pl = parse_limit(exp, lim.unit)
            if pl.numeric:
                lim = pl
            else:
                lim.expected, lim.text = pl.expected or exp, exp
        elif exp:
            lim.text = exp
        if lim.max is not None and exp.strip().startswith("<") and not exp.strip().startswith("<="):
            lim.max_incl = False
        rows.append({
            "test": _txt(r["Test"]), "category": category_of(_txt(r["Test"]), _txt(r["Category"])),
            "limit": lim, "method": _txt(r["Method"]),
            "required": _txt(r["Required"]).lower() not in ("n", "no", "false", "0", "optional"),
            "notes": _txt(r["Notes"]),
        })
    return rows


def product_info(products, product):
    sel = products[products["Product"] == product]
    if sel.empty:
        return {}
    r = sel.iloc[0]
    return {c: _txt(r[c]) for c in PRODUCT_COLUMNS}


def match_product(coa_product, products):
    """Best spec-library product for a COA product name -> (product, score)."""
    name = clean_name(coa_product)
    if not name:
        return None, 0.0
    best, best_s = None, 0.0
    for _, r in products.iterrows():
        cands = [r["Product"]] + [a for a in re.split(r"[;,|\n]", _txt(r["Aliases"])) if a.strip()]
        for c in cands:
            cn = clean_name(c)
            if not cn:
                continue
            s = difflib.SequenceMatcher(None, name, cn).ratio()
            if cn in name or name in cn:
                s = max(s, 0.85)
            tn, tc = _key_words(name), _key_words(cn)
            if tn and tc:
                s = max(s, len(tn & tc) / len(tc) * 0.85)
                if not (tn & tc):
                    s = min(s, 0.4)       # nothing specific in common
            if s > best_s:
                best, best_s = r["Product"], s
    return (best, best_s) if best_s >= 0.6 else (None, best_s)


_GENERIC = {"extract", "powder", "extracts", "root", "leaf", "leaves", "seed", "fruit", "bean", "herb", "plant",
            "standardized", "standardised", "natural", "organic", "hplc", "uv", "total", "dry", "dried", "fine",
            "water", "soluble", "the", "and", "with"}


def _key_words(n):
    return {w for w in re.findall(r"[a-z]{3,}", n) if w not in _GENERIC}


# ======================================================================
# The review
# ======================================================================

SEV_ORDER = {"High": 0, "Medium": 1, "Low": 2}
STATUS_ICON = {"Pass": "✅ Pass", "Fail": "❌ Fail", "Missing": "⚠️ Missing", "Review": "🔍 Review",
               "Info": "ℹ️ Info", "Not in spec": "· Not in spec"}


def _finding(sev, status, item, detail, question, category="", source="Rule"):
    return {"Severity": sev, "Status": status, "Item": item, "Detail": detail,
            "Question for supplier": question, "Category": category, "Source": source}


def _sev(category, kind):
    safety = category in SAFETY
    if kind == "fail":
        return "High" if safety or category in ("assay", "identification") else "Medium"
    if kind == "missing":
        return "High" if safety else "Medium"
    if kind == "review":
        return "Medium" if safety or category in ("assay", "identification") else "Low"
    return "Low"


def _within(v, lim):
    if lim.min is not None and (v < lim.min or (v == lim.min and not lim.min_incl)):
        return False
    if lim.max is not None and (v > lim.max or (v == lim.max and not lim.max_incl)):
        return False
    return True


def compare(test_name, category, lim, res, coa_spec_text=""):
    """Compare one COA result with our limit -> (status, note, question)."""
    name = test_name
    if not res.text:
        return "Missing", "No result given.", f"{name} has no result on the COA. Please provide the result for this lot."

    if not lim.numeric:
        if lim.expected == "negative":
            if res.qual == "nd":
                return "Pass", "", ""
            if res.qual == "pos":
                return "Fail", f"Result '{res.text}' — expected negative.", \
                    f"{name} is reported as '{res.text}'. Our specification requires negative. Please confirm."
            if res.value is not None and res.op in ("<", "≤"):
                return "Pass", f"Reported as {res.text}; read as not detected.", ""
            if res.qual == "pass":
                return "Review", f"Result says '{res.text}' rather than negative / not detected.", \
                    f"{name} is reported as '{res.text}'. Please state the result as Negative / Not detected, with the sample size tested."
            return "Review", f"Can't read '{res.text}' as negative or positive.", \
                f"Please clarify the {name} result ('{res.text}')."
        if lim.expected:
            if res.qual in ("pass", "nd") or not lim.expected:
                return "Pass", "", ""
            a, b = clean_name(lim.expected), clean_name(res.text)
            if a and b and (a in b or b in a or difflib.SequenceMatcher(None, a, b).ratio() > 0.6):
                return "Pass", "", ""
            return "Review", f"Expected '{lim.expected}', COA says '{res.text}'.", \
                f"{name}: our specification is '{lim.expected}'; the COA states '{res.text}'. Please confirm it complies."
        return "Info", "No limit in our spec.", ""

    want = lim.describe()
    if res.qual == "pass" and res.value is None:
        return "Review", f"Reported as '{res.text}' with no value (limit {want}).", \
            f"{name} is reported as '{res.text}' without a value. Please provide the measured result" + \
            (f" in {lim.unit}." if lim.unit else ".")
    if res.qual == "nd" and res.value is None:
        if lim.min is not None:
            return "Fail", f"Not detected, but our spec needs {want}.", \
                f"{name} is reported as not detected; our specification requires {want}. Please confirm."
        return "Pass", "Not detected.", ""
    if res.value is None:
        return "Review", f"Can't read a value from '{res.text}'.", f"Please clarify the {name} result ('{res.text}')."

    r_unit = res.unit or lim.unit
    note = ""
    if not res.unit and lim.unit:
        note = f"No unit on the COA; read as {lim.unit}. "
    if lim.unit and r_unit and not same_family(r_unit, lim.unit):
        return "Review", f"COA unit {r_unit} can't be compared with our unit {lim.unit}.", \
            f"{name} is reported in {r_unit}; our specification is in {lim.unit}. Please provide the result in {lim.unit}."
    v = convert(res.value, r_unit, lim.unit) if lim.unit else res.value
    if lim.unit and r_unit and norm_unit(r_unit) != norm_unit(lim.unit):
        note += f"Converted {fmt_num(res.value)} {r_unit} → {fmt_num(v)} {lim.unit}. "

    shown = f"{res.op if res.op != '=' else ''}{fmt_num(v)} {lim.unit}".strip()
    if res.op in ("<", "≤"):          # "<10": real value is below 10
        if lim.max is not None and (v < lim.max or (v == lim.max and (res.op == "<" or lim.max_incl))):
            return "Pass", note.strip(), ""
        if lim.min is not None and v <= lim.min:
            return "Fail", (note + f"{shown} is below our minimum {want}.").strip(), \
                f"{name} is reported as {res.text}, below our minimum of {want}. Please confirm."
        return "Review", (note + f"Reported as {res.text}: the detection limit is not low enough to show ≤ "
                          f"{fmt_num(lim.max)} {lim.unit}.").strip(), \
            f"{name} is reported as {res.text}, which can't show compliance with our limit of {want}. " \
            f"Please provide a result with a lower detection limit."
    if res.op in (">", "≥"):
        if lim.max is not None and v >= lim.max:
            return "Fail", (note + f"{shown} exceeds our limit {want}.").strip(), \
                f"{name} is reported as {res.text}, above our limit of {want}. Please confirm."
        if lim.min is not None and v >= lim.min:
            return "Pass", note.strip(), ""
        return "Review", (note + f"Reported as {res.text}.").strip(), f"Please clarify the {name} result ({res.text})."
    if _within(v, lim):
        return "Pass", note.strip(), ""
    side = "above our limit" if lim.max is not None and v >= lim.max else "below our minimum"
    return "Fail", (note + f"{shown} is {side} ({want}).").strip(), \
        f"{name} is reported as {res.text}, {side} ({want}). Please confirm the result and advise."


def _looser(coa_lim, lim):
    """True when the supplier's own spec on the COA is looser than ours."""
    if not coa_lim.numeric or not lim.numeric:
        return False
    if lim.unit and coa_lim.unit and not same_family(coa_lim.unit, lim.unit):
        return False
    cmax = convert(coa_lim.max, coa_lim.unit or lim.unit, lim.unit) if coa_lim.max is not None else None
    cmin = convert(coa_lim.min, coa_lim.unit or lim.unit, lim.unit) if coa_lim.min is not None else None
    if lim.max is not None and cmax is not None and cmax > lim.max * 1.0001:
        return True
    if lim.min is not None and cmin is not None and cmin < lim.min * 0.9999:
        return True
    if lim.max is not None and cmax is None and cmin is not None and lim.min is None:
        return False
    return False


def _method_ok(required, given):
    if not required:
        return True
    squash = lambda x: re.sub(r"[\s\-_.<>()]", "", x)
    g = squash(clean_name(given))
    for alt in re.split(r"\s*(?:/|\bor\b|,|;)\s*", clean_name(required)):
        toks = [squash(t) for t in re.findall(r"[a-z0-9.\-]+", alt) if len(squash(t)) > 1]
        if toks and all(t in g for t in toks):
            return True
    return False


def review_coa(coa, spec_rows, prod_info=None, spec_product="", today=None):
    """Run every code rule. Returns {"findings": [...], "tests": [...], "verdict": str, "counts": {...}}."""
    today = today or dt.date.today()
    prod_info = prod_info or {}
    info = {k: (v or "") for k, v in (coa.get("info") or {}).items()}
    findings, table = [], []

    # ---- basic information ------------------------------------------------
    if not info.get("lot_no"):
        findings.append(_finding("High", "Missing", "Lot / batch number", "No lot or batch number found on the COA.",
                                 "The COA does not show a lot/batch number. Please send a COA that states it.",
                                 "Basic info"))
    if not info.get("mfg_date"):
        findings.append(_finding("Medium", "Missing", "Manufacture date", "No manufacture date found.",
                                 "Please add the manufacture date to the COA.", "Basic info"))
    if not info.get("expiry_date") and not info.get("retest_date"):
        findings.append(_finding("Medium", "Missing", "Expiry date", "No expiry or retest date found.",
                                 "Please add the expiry (or retest) date to the COA.", "Basic info"))
    signed = coa.get("signed")
    if signed is False:
        findings.append(_finding("Medium", "Missing", "Signature / approval", "The COA is not signed or approved.",
                                 "The COA is not signed. Please send a signed copy approved by your QA/QC.",
                                 "Basic info"))
    elif signed is None and not coa.get("approval_found"):
        findings.append(_finding("Low", "Review", "Signature / approval",
                                 "No approval line found in the text — check the signature by eye.",
                                 "Please confirm the COA is signed by your QA/QC department.", "Basic info"))

    # product name vs spec product
    if spec_product and info.get("product_name"):
        names = [spec_product] + [a for a in re.split(r"[;,|\n]", prod_info.get("Aliases", "")) if a.strip()]
        best = max(difflib.SequenceMatcher(None, clean_name(info["product_name"]), clean_name(n)).ratio()
                   for n in names)
        tokens_hit = any(set(re.findall(r"[a-z]{4,}", clean_name(n))) & set(re.findall(r"[a-z]{4,}", clean_name(info["product_name"])))
                         for n in names)
        if best < 0.5 and not tokens_hit:
            findings.append(_finding("Medium", "Review", "Product name",
                                     f"COA product '{info['product_name']}' doesn't look like '{spec_product}'.",
                                     f"The COA is for '{info['product_name']}'. Please confirm this is the product we ordered ({spec_product}).",
                                     "Basic info"))

    # ---- dates --------------------------------------------------------------
    mfg = parse_date(info.get("mfg_date"))
    exp = parse_date(info.get("expiry_date") or info.get("retest_date"), end_of_month=True)
    tst = parse_date(info.get("test_date"))
    rep = parse_date(info.get("report_date"))
    for label, pd_ in (("Manufacture date", mfg), ("Expiry date", exp), ("Test date", tst), ("Report date", rep)):
        if pd_.text and not pd_.date:
            findings.append(_finding("Low", "Review", label, f"Couldn't read the date '{pd_.text}'.",
                                     f"Please state the {label.lower()} in a clear format (e.g. 2026-09-17).", "Dates"))
        elif pd_.ambiguous:
            findings.append(_finding("Low", "Review", label,
                                     f"'{pd_.text}' could be day/month or month/day — read as {pd_.date:%d %b %Y}.",
                                     f"Please confirm the {label.lower()} ({pd_.text}) — is it day/month/year?", "Dates"))
    if mfg.date and exp.date and exp.date <= mfg.date:
        findings.append(_finding("High", "Fail", "Expiry date",
                                 f"Expiry {exp.date:%Y-%m-%d} is not after manufacture {mfg.date:%Y-%m-%d}.",
                                 "The expiry date is earlier than the manufacture date. Please correct the COA.", "Dates"))
    if exp.date and exp.date < today:
        findings.append(_finding("High", "Fail", "Expiry date", f"Expired on {exp.date:%Y-%m-%d}.",
                                 f"This lot expired on {exp.date:%Y-%m-%d}. Please advise or supply a current lot.", "Dates"))
    elif exp.date and months_between(today, exp.date) < 6:
        findings.append(_finding("Low", "Review", "Expiry date",
                                 f"Expires {exp.date:%Y-%m-%d} — less than 6 months left.",
                                 "", "Dates"))
    if mfg.date and mfg.date > today:
        findings.append(_finding("Medium", "Review", "Manufacture date", f"Manufacture date {mfg.date:%Y-%m-%d} is in the future.",
                                 "The manufacture date on the COA is in the future. Please correct it.", "Dates"))
    for label, pd_ in (("Test date", tst), ("Report date", rep)):
        if pd_.date and mfg.date and pd_.date < mfg.date:
            findings.append(_finding("Medium", "Review", label,
                                     f"{label} {pd_.date:%Y-%m-%d} is before manufacture {mfg.date:%Y-%m-%d}.",
                                     f"The {label.lower()} is earlier than the manufacture date. Please confirm the COA belongs to this lot.",
                                     "Dates"))
        if pd_.date and pd_.date > today:
            findings.append(_finding("Low", "Review", label, f"{label} {pd_.date:%Y-%m-%d} is in the future.",
                                     f"Please confirm the {label.lower()} on the COA.", "Dates"))
    sl = (prod_info.get("Shelf life (months)") or "").strip()
    spec_months = int(float(sl)) if re.fullmatch(r"\d+(\.\d+)?", sl) else parse_months(sl)
    if spec_months and mfg.date and exp.date and exp.date > mfg.date:
        coa_months = months_between(mfg.date, exp.date)
        if coa_months > spec_months + 1:
            findings.append(_finding("Low", "Review", "Shelf life",
                                     f"COA gives {coa_months:.0f} months; our spec says {spec_months}.",
                                     f"The COA shelf life ({coa_months:.0f} months) is longer than our specification ({spec_months} months). Please confirm the supporting stability data.",
                                     "Dates"))

    # ---- product details (solvent, carrier, latin name) ---------------------
    if prod_info.get("Extraction solvent") and not info.get("solvent"):
        findings.append(_finding("Low", "Missing", "Extraction solvent",
                                 f"Our spec says {prod_info['Extraction solvent']}; the COA doesn't state a solvent.",
                                 "Please state the extraction solvent on the COA.", "Product details"))
    elif prod_info.get("Extraction solvent") and info.get("solvent"):
        want = [clean_name(w) for w in re.split(r"[,/&+;]|\band\b", prod_info["Extraction solvent"].lower())
                if clean_name(w)]
        have = clean_name(info["solvent"])
        if want and not all(w in have for w in want):
            findings.append(_finding("Medium", "Review", "Extraction solvent",
                                     f"COA: '{info['solvent']}'; our spec: '{prod_info['Extraction solvent']}'.",
                                     f"The COA lists the extraction solvent as '{info['solvent']}', but our specification requires {prod_info['Extraction solvent']}. Please confirm.",
                                     "Product details"))
    if prod_info.get("Carrier") and prod_info["Carrier"].lower() not in ("none", "no", "n/a") and not info.get("carrier"):
        findings.append(_finding("Low", "Missing", "Carrier",
                                 f"Our spec lists carrier '{prod_info['Carrier']}'; the COA doesn't mention one.",
                                 "Please state the carrier/excipient used, if any.", "Product details"))
    if prod_info.get("Latin name") and info.get("latin_name"):
        a = re.findall(r"[a-z]+", clean_name(prod_info["Latin name"]))
        b = re.findall(r"[a-z]+", clean_name(info["latin_name"]))
        q = (f"The COA states the botanical source as '{info['latin_name']}', but our specification is "
             f"{prod_info['Latin name']}. Please confirm the plant species.")
        if a and b and a[0] != b[0]:
            findings.append(_finding("High", "Fail", "Botanical name",
                                     f"Different genus — COA: '{info['latin_name']}'; our spec: '{prod_info['Latin name']}'.",
                                     q, "Product details"))
        elif len(a) > 1 and len(b) > 1 and not set(b[1:]) & set(a[1:]):
            findings.append(_finding("Medium", "Review", "Botanical name",
                                     f"Different species — COA: '{info['latin_name']}'; our spec: '{prod_info['Latin name']}'.",
                                     q, "Product details"))

    # ---- tests ----------------------------------------------------------------
    coa_tests = list(coa.get("tests") or [])
    used, looser = set(), []
    for sr in spec_rows:
        lim, cat = sr["limit"], sr["category"]
        scored = sorted(((name_score(t["name"], sr["test"]), i) for i, t in enumerate(coa_tests) if i not in used),
                        reverse=True)
        matches = [i for s, i in scored if s >= 0.82][:1]
        sc = canon(sr["test"])
        if not matches and sc == "yeast & mold":       # yeast and mold reported separately
            matches = [i for i, t in enumerate(coa_tests) if canon(t["name"]) in ("yeast", "mold") and i not in used]
        if not matches and sc in ("yeast", "mold"):    # spec separate, COA combined
            matches = [i for i, t in enumerate(coa_tests) if canon(t["name"]) == "yeast & mold"][:1]
        if not matches and sc in ("residual solvents", "pesticides", "heavy metals"):
            parts = [t["name"] for t in coa_tests if category_of(t["name"], t.get("category", "")) == sc
                     and canon(t["name"]) != sc]
            if parts:      # our spec has one umbrella row, the COA lists the individual items
                table.append(_row(sr, None, "Info", "COA lists individual results: " + ", ".join(parts[:8])))
                continue
        if not matches:
            umbrella = [t for t in coa_tests if canon(t["name"]) == cat or clean_name(t["name"]) == cat]
            if umbrella and cat in ("residual solvents", "pesticides", "heavy metals"):
                u = umbrella[0]
                note = f"COA only gives a combined '{u['name']}' result ({u['result']}); no separate value."
                q = f"The COA reports {cat} only as '{u['result']}'. Please provide the result for {sr['test']}."
                table.append(_row(sr, None, "Review", note))
                findings.append(_finding(_sev(cat, "review"), "Review", sr["test"], note, q, cat))
            elif sr["required"]:
                table.append(_row(sr, None, "Missing", "Not on the COA."))
                findings.append(_finding(_sev(cat, "missing"), "Missing", sr["test"],
                                         f"Required by our spec ({lim.describe()}) but not on the COA.",
                                         f"The COA does not include {sr['test']}. Please provide the {sr['test']} result for this lot"
                                         + (f" (method: {sr['method']})." if sr["method"] else "."), cat))
            else:
                table.append(_row(sr, None, "Info", "Optional test, not on the COA."))
            continue

        for i in matches:
            if canon(coa_tests[i]["name"]) != "yeast & mold" or sc == "yeast & mold":
                used.add(i)
            t = coa_tests[i]
            res = parse_result(t.get("result"), t.get("unit") or find_unit(t.get("spec")) or "")
            status, note, q = compare(sr["test"], cat, lim, res, t.get("spec"))
            if len(matches) > 1:
                note = (note + f" Reported separately as '{t['name']}'.").strip()
            # supplier's own spec looser than ours
            coa_lim = parse_limit(t.get("spec"), t.get("unit") or "")
            if _looser(coa_lim, lim):
                looser.append(f"{sr['test']} ({coa_lim.describe()} vs our {lim.describe()})")
                note = (note + f" Supplier's own limit {coa_lim.describe()} is looser than ours.").strip()
            if status in ("Fail", "Review", "Missing"):
                kind = {"Fail": "fail", "Review": "review", "Missing": "missing"}[status]
                findings.append(_finding(_sev(cat, kind), status, sr["test"], note, q, cat))
            # method
            if sr["method"] and status != "Missing":
                if not t.get("method"):
                    findings.append(_finding("Low", "Review", f"{sr['test']} — method",
                                             f"Our spec requires {sr['method']}; the COA doesn't state a method.",
                                             f"Please state the test method used for {sr['test']} (our specification requires {sr['method']}).",
                                             cat))
                elif not _method_ok(sr["method"], t["method"]):
                    findings.append(_finding(_sev(cat, "review"), "Review", f"{sr['test']} — method",
                                             f"COA method '{t['method']}'; our spec requires {sr['method']}.",
                                             f"{sr['test']} was tested by {t['method']}; our specification requires {sr['method']}. Please provide results by {sr['method']} or confirm equivalence.",
                                             cat))
            table.append(_row(sr, t, status, note))

    vague = [f for f in findings if f["Status"] == "Review" and f["Source"] == "Rule"
             and re.search(r"with no value|rather than negative", f["Detail"])]
    if len(vague) >= 3:
        names = [f["Item"] for f in vague]
        words = sorted({re.search(r"'([^']+)'", f["Detail"]).group(1) for f in vague
                        if re.search(r"'([^']+)'", f["Detail"])})
        sev = "Medium" if any(f["Severity"] == "Medium" for f in vague) else "Low"
        findings = [f for f in findings if f not in vague]
        findings.append(_finding(sev, "Review", f"Results without values ({len(names)} tests)",
                                 f"Reported only as {' / '.join(repr(w) for w in words)}: " + ", ".join(names) + ".",
                                 "The following results are reported only as " + " / ".join(repr(w) for w in words) +
                                 ": " + ", ".join(names) + ". Please provide the measured values "
                                 "(and Negative/Not detected with sample size for microbiology).", "Vague results"))

    missing = [f for f in findings if f["Status"] == "Missing" and f["Category"] not in
               ("Basic info", "Product details") and f["Source"] == "Rule"]
    if len(missing) >= 3:
        findings = [f for f in findings if f not in missing]
        sev = "High" if any(f["Severity"] == "High" for f in missing) else "Medium"
        names = [f["Item"] for f in missing]
        findings.append(_finding(sev, "Missing", f"Tests missing from the COA ({len(names)})",
                                 "Required by our spec but not on the COA: " + ", ".join(names) + ".",
                                 "The COA does not include results for: " + ", ".join(names) +
                                 ". Please provide these results for this lot.", "Missing tests"))

    if looser:
        findings.append(_finding("Low", "Review", "Supplier's own specification",
                                 "The supplier's limits are looser than ours for: " + "; ".join(looser) + ".",
                                 "Your specification is looser than ours for " +
                                 ", ".join(x.split(" (")[0] for x in looser) +
                                 ". Please confirm every lot shipped to us meets our specification.", "Specification"))

    for i, t in enumerate(coa_tests):
        if i not in used:
            table.append({"Test (spec)": "", "Our limit": "", "Test (COA)": t["name"],
                          "COA spec": t.get("spec", ""), "COA result": _result_text(t),
                          "Method": t.get("method", ""), "Status": "Not in spec", "Note": ""})

    findings.sort(key=lambda f: (SEV_ORDER.get(f["Severity"], 3), f["Category"], f["Item"]))
    counts = {s: sum(1 for r in table if r["Status"] == s) for s in ("Pass", "Fail", "Missing", "Review")}
    sev = {s: sum(1 for f in findings if f["Severity"] == s) for s in ("High", "Medium", "Low")}
    return {"findings": findings, "tests": table, "counts": counts, "severity": sev, "verdict": verdict(findings)}


def _result_text(t):
    r = (t.get("result") or "").strip()
    u = (t.get("unit") or "").strip()
    return f"{r} {u}".strip() if u and u.lower() not in r.lower() else r


def _row(sr, t, status, note):
    return {"Test (spec)": sr["test"], "Our limit": sr["limit"].describe(),
            "Test (COA)": t["name"] if t else "", "COA spec": t.get("spec", "") if t else "",
            "COA result": _result_text(t) if t else "", "Method": t.get("method", "") if t else "",
            "Status": status, "Note": note}


def verdict(findings):
    sevs = {f["Severity"] for f in findings if f.get("Status") != "Info"}
    if "High" in sevs:
        return "Hold — resolve issues with the supplier"
    if "Medium" in sevs:
        return "Check with the supplier"
    if "Low" in sevs:
        return "OK — minor notes"
    return "OK — no issues found"


# ======================================================================
# Supplier email
# ======================================================================

def build_email(info, findings, product="", supplier="", sender="Skyherb Nutritionals — Purchasing"):
    qs = [f["Question for supplier"] for f in findings if (f.get("Question for supplier") or "").strip()]
    lot = info.get("lot_no") or "(lot number not stated)"
    prod = info.get("product_name") or product or "the product"
    subject = f"COA questions — {prod}, lot {lot}"
    greet = f"Dear {supplier} team," if supplier else "Hello,"
    if not qs:
        body = (f"{greet}\n\nThank you for the COA for {prod}, lot {lot}. We have reviewed it and have no "
                f"questions at this time.\n\nBest regards,\n{sender}")
        return subject, body
    items = "\n".join(f"{i}. {q}" for i, q in enumerate(qs, 1))
    body = (f"{greet}\n\nThank you for the COA for {prod}, lot {lot}. Before we can release this lot, "
            f"could you please help with the following:\n\n{items}\n\n"
            f"Please send a corrected or complete COA, or your comments, at your earliest convenience.\n\n"
            f"Best regards,\n{sender}")
    return subject, body


# ======================================================================
# Exports
# ======================================================================

def to_excel(reviews):
    """reviews: list of dicts with keys file, product, info, result (review_coa output), findings (edited)."""
    import pandas as pd
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame([{
            "File": r["file"], "Product (spec)": r["product"], "COA product": r["info"].get("product_name", ""),
            "Lot": r["info"].get("lot_no", ""), "Mfg date": r["info"].get("mfg_date", ""),
            "Expiry": r["info"].get("expiry_date", ""), "Verdict": verdict(r["findings"]),
            "High": sum(f["Severity"] == "High" for f in r["findings"]),
            "Medium": sum(f["Severity"] == "Medium" for f in r["findings"]),
            "Low": sum(f["Severity"] == "Low" for f in r["findings"]),
            "Reviewed": dt.date.today().isoformat(), "Read by": r.get("source", ""),
        } for r in reviews]).to_excel(xw, sheet_name="Summary", index=False)
        pd.DataFrame([{"File": r["file"], **f} for r in reviews for f in r["findings"]]
                     ).to_excel(xw, sheet_name="Issues", index=False)
        pd.DataFrame([{"File": r["file"], **t} for r in reviews for t in r["result"]["tests"]]
                     ).to_excel(xw, sheet_name="Test results", index=False)
        pd.DataFrame([{"File": r["file"], "Field": k, "Value": v} for r in reviews for k, v in r["info"].items()]
                     ).to_excel(xw, sheet_name="COA info", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 70)
            ws.freeze_panes = "A2"
    return buf.getvalue()


def to_pdf(review, title="COA Review"):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        cjk = "STSong-Light"
    except Exception:
        cjk = None

    ss = getSampleStyleSheet()
    cell = ParagraphStyle("cell", parent=ss["BodyText"], fontSize=8, leading=10)
    cell_cjk = ParagraphStyle("cellc", parent=cell, fontName=cjk) if cjk else cell

    def P(text, bold=False):
        t = escape(str(text or ""))
        st = cell_cjk if _CJK.search(t) else cell
        return Paragraph(f"<b>{t}</b>" if bold else t, st)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm, title=title)
    info = review["info"]
    story = [Paragraph(escape(title), ss["Title"]),
             Paragraph(escape(f"{review['file']} — reviewed {dt.date.today():%Y-%m-%d}"), ss["Normal"]),
             Spacer(1, 4 * mm)]
    rows = [[P("Spec product", True), P(review["product"]), P("COA product", True), P(info.get("product_name"))],
            [P("Lot", True), P(info.get("lot_no")), P("Verdict", True), P(verdict(review["findings"]))],
            [P("Mfg date", True), P(info.get("mfg_date")), P("Expiry", True),
             P(info.get("expiry_date") or info.get("retest_date"))]]
    t = Table(rows, colWidths=[30 * mm, 100 * mm, 30 * mm, 100 * mm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                           ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#eef4f1")),
                           ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#eef4f1"))]))
    story += [t, Spacer(1, 5 * mm), Paragraph("Issues", ss["Heading2"])]

    sev_col = {"High": colors.HexColor("#fde2e1"), "Medium": colors.HexColor("#fff3cd"),
               "Low": colors.HexColor("#eef2f7")}
    if review["findings"]:
        data = [[P(h, True) for h in ("Severity", "Status", "Item", "Detail", "Source")]]
        style = [("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dfe9e4")),
                 ("VALIGN", (0, 0), (-1, -1), "TOP")]
        for i, f in enumerate(review["findings"], 1):
            data.append([P(f["Severity"]), P(f["Status"]), P(f["Item"]), P(f["Detail"]), P(f.get("Source", ""))])
            style.append(("BACKGROUND", (0, i), (0, i), sev_col.get(f["Severity"], colors.white)))
        t = Table(data, colWidths=[20 * mm, 20 * mm, 50 * mm, 160 * mm, 18 * mm], repeatRows=1)
        t.setStyle(TableStyle(style))
        story.append(t)
    else:
        story.append(Paragraph("No issues found.", ss["Normal"]))

    story += [Spacer(1, 5 * mm), Paragraph("Test-by-test comparison", ss["Heading2"])]
    data = [[P(h, True) for h in ("Test", "Our limit", "COA result", "Method", "Status", "Note")]]
    for r in review["result"]["tests"]:
        data.append([P(r["Test (spec)"] or r["Test (COA)"]), P(r["Our limit"]), P(r["COA result"]),
                     P(r["Method"]), P(r["Status"]), P(r["Note"])])
    t = Table(data, colWidths=[45 * mm, 35 * mm, 35 * mm, 30 * mm, 22 * mm, 101 * mm], repeatRows=1)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                           ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dfe9e4")),
                           ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(t)
    doc.build(story)
    return buf.getvalue()


# ======================================================================
# Claude API (optional)
# ======================================================================

DEFAULT_MODEL = "claude-sonnet-5-5"

_COA_TOOL = {
    "name": "record_coa",
    "description": "Record everything printed on the certificate of analysis, transcribed exactly.",
    "input_schema": {
        "type": "object",
        "properties": {
            "info": {
                "type": "object",
                "properties": {k: {"type": ["string", "null"]} for k in (
                    "product_name", "supplier", "lot_no", "quantity", "mfg_date", "expiry_date", "retest_date",
                    "test_date", "report_date", "shelf_life", "latin_name", "plant_part", "solvent", "carrier",
                    "ratio")},
            },
            "signed": {"type": ["boolean", "null"],
                       "description": "true if a handwritten/electronic signature or a QA/QC approval stamp is visible, "
                                      "false if the signature field is empty, null if there is no signature field"},
            "signer": {"type": ["string", "null"]},
            "tests": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "name": {"type": "string"}, "category": {"type": ["string", "null"]},
                        "specification": {"type": ["string", "null"]}, "result": {"type": ["string", "null"]},
                        "unit": {"type": ["string", "null"]}, "method": {"type": ["string", "null"]},
                    },
                    "required": ["name"],
                },
            },
            "notes": {"type": ["string", "null"], "description": "Anything unusual: handwritten corrections, "
                                                                 "crossed-out values, unreadable parts, stamps over text"},
        },
        "required": ["info", "tests"],
    },
}

_COA_PROMPT = """Read this certificate of analysis (COA) and record it with the record_coa tool.

Rules:
- Transcribe exactly what is printed. Never guess, calculate or fill in values that are not on the document; use null.
- One entry in `tests` per test row, including section headings as `category` (e.g. "Heavy metals", "Microbiology").
- `specification` and `result` as printed, keeping symbols such as <, ≤, NMT, ND, Negative, Conforms.
- Write powers of ten as 10^3 (e.g. "<10^3 cfu/g"), never as 103.
- If the unit is in its own column, put it in `unit`; otherwise leave it inside the result text.
- Dates exactly as printed (e.g. "September 17, 2026" or "17/09/2026").
- Ignore company letterheads, addresses and logos except for the `supplier` name."""

_FINDINGS_TOOL = {
    "name": "record_findings",
    "description": "Record judgement-based problems found in the COA.",
    "input_schema": {
        "type": "object",
        "properties": {"findings": {"type": "array", "items": {
            "type": "object",
            "properties": {
                "severity": {"type": "string", "enum": ["High", "Medium", "Low"]},
                "category": {"type": "string", "enum": ["Product/spec match", "Method", "Vague result",
                                                        "Solvent/carrier", "Other"]},
                "item": {"type": "string"}, "detail": {"type": "string"},
                "question_for_supplier": {"type": "string"},
            },
            "required": ["severity", "category", "item", "detail", "question_for_supplier"],
        }}},
        "required": ["findings"],
    },
}

_REVIEW_PROMPT = """You are the quality reviewer for Skyherb Nutritionals, a US supplier of botanical extracts and powders.
A COA from a supplier has been transcribed and compared with our specification by code.
The code already checked numbers, units, dates and missing tests — do NOT repeat those findings.

Look only for judgement problems:
1. Product/spec match — does the COA describe the product our spec is for? (e.g. COA says "10:1 extract" but our spec needs a standardised % of a marker compound; different plant part or species; powder vs extract).
2. Method — is each test method suitable for what it measures and for our required method? (e.g. identification by "Organoleptic" when we need HPTLC; assay by UV when we need HPLC; heavy metals by a colour test instead of ICP-MS/AAS).
3. Vague result — wording that doesn't really report a result, e.g. "Conforms", "Complies", "Pass" for tests that need a value, or copy-pasted specification text as the result.
4. Solvent/carrier — extraction solvent or carrier not stated, unusual, or inconsistent with our spec.
5. Other — anything a careful QA reviewer would question (handwritten changes, results identical to the limits, suspicious round numbers on every test, etc.).

Be conservative: only report a problem you can point to in the data. Fewer, solid findings are better than many weak ones.
Write each question_for_supplier as one polite, specific sentence in English.

COA data:
{coa}

Our specification:
{spec}

Findings already made by code (don't repeat):
{rules}"""

_SPEC_TOOL = {
    "name": "record_spec",
    "description": "Record each test requirement of the specification sheet.",
    "input_schema": {
        "type": "object",
        "properties": {
            "product": {"type": ["string", "null"]},
            "rows": {"type": "array", "items": {
                "type": "object",
                "properties": {
                    "test": {"type": "string"}, "category": {"type": ["string", "null"]},
                    "min": {"type": ["number", "null"]}, "max": {"type": ["number", "null"]},
                    "unit": {"type": ["string", "null"]},
                    "expected": {"type": ["string", "null"], "description": "Text requirement such as 'Negative/25g' or 'Brown fine powder'"},
                    "method": {"type": ["string", "null"]},
                    "limit_as_printed": {"type": "string"},
                },
                "required": ["test", "limit_as_printed"],
            }},
        },
        "required": ["rows"],
    },
}


def _content_block(file_bytes, filename):
    name = filename.lower()
    if name.endswith(".pdf"):
        return {"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                                                "data": base64.b64encode(file_bytes).decode()}}
    from PIL import Image
    img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
    if max(img.size) > 2400:
        img.thumbnail((2400, 2400))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                        "data": base64.b64encode(buf.getvalue()).decode()}}


def _call_tool(client, model, content, tool, max_tokens=8000):
    msg = client.messages.create(
        model=model, max_tokens=max_tokens, tools=[tool],
        tool_choice={"type": "tool", "name": tool["name"]},
        messages=[{"role": "user", "content": content}],
    )
    for block in msg.content:
        if getattr(block, "type", "") == "tool_use":
            return block.input
    raise RuntimeError("The model did not return structured data.")


def make_client(api_key):
    import anthropic
    return anthropic.Anthropic(api_key=api_key)


def ai_extract_coa(client, model, file_bytes, filename):
    data = _call_tool(client, model, [_content_block(file_bytes, filename),
                                      {"type": "text", "text": _COA_PROMPT}], _COA_TOOL)
    info = {k: v for k, v in (data.get("info") or {}).items() if v}
    tests = [{"name": t.get("name") or "", "category": t.get("category") or "",
              "spec": t.get("specification") or "", "result": t.get("result") or "",
              "unit": t.get("unit") or "", "method": t.get("method") or ""}
             for t in data.get("tests") or [] if (t.get("name") or "").strip()]
    warnings = [data["notes"]] if data.get("notes") else []
    return {"source": "ai", "info": info, "tests": tests, "raw_text": "", "signed": data.get("signed"),
            "signer": data.get("signer"), "approval_found": bool(data.get("signed")), "warnings": warnings}


def ai_review(client, model, coa, spec_rows, prod_info, rule_findings):
    coa_json = json.dumps({"info": coa.get("info"), "tests": coa.get("tests"), "signed": coa.get("signed")},
                          ensure_ascii=False, indent=1)
    spec_json = json.dumps({"product": prod_info, "tests": [
        {"test": r["test"], "limit": r["limit"].describe(), "method": r["method"]} for r in spec_rows]},
        ensure_ascii=False, indent=1)
    rules = "\n".join(f"- {f['Item']}: {f['Detail']}" for f in rule_findings) or "(none)"
    data = _call_tool(client, model, [{"type": "text", "text": _REVIEW_PROMPT.format(
        coa=coa_json, spec=spec_json, rules=rules)}], _FINDINGS_TOOL, max_tokens=4000)
    return [_finding(f["severity"], "Review", f["item"], f["detail"], f["question_for_supplier"],
                     f["category"], source="AI") for f in data.get("findings") or []]


def ai_extract_spec(client, model, file_bytes, filename):
    data = _call_tool(client, model, [_content_block(file_bytes, filename), {"type": "text", "text":
                      "Record every test requirement in this specification sheet with the record_spec tool. "
                      "Use min/max for numeric limits (e.g. 'NMT 1 ppm' -> max 1, unit ppm; '≥ 50%' -> min 50, "
                      "unit %; write 10^3 as 1000). Use `expected` for text requirements. Never invent limits."}],
                      _SPEC_TOOL)
    return data.get("product"), [{
        "Test": r.get("test"), "Category": r.get("category"), "Min": r.get("min"), "Max": r.get("max"),
        "Unit": r.get("unit"), "Expected": r.get("expected") or ("" if (r.get("min") is not None or r.get("max") is not None)
                                                                 else r.get("limit_as_printed")),
        "Method": r.get("method"), "Required": "Y", "Notes": ""} for r in data.get("rows") or []]


def ai_polish_email(client, model, subject, body):
    msg = client.messages.create(model=model, max_tokens=2000, messages=[{"role": "user", "content":
        "Polish this email to a raw-material supplier. Keep every question and fact, keep it concise, polite and "
        "professional, plain text, no new claims. Return only the email body.\n\nSubject: "
        f"{subject}\n\n{body}"}])
    return "".join(getattr(b, "text", "") for b in msg.content).strip()


def text_spec_rows(file_bytes):
    """Read a customer spec PDF without AI (text PDFs with a table)."""
    import pdfplumber
    tables = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables += page.extract_tables() or []
    rows = []
    for t in _tests_from_tables(tables):
        lim = parse_limit(t["spec"], t.get("unit", ""))
        rows.append({"Test": t["name"], "Category": t.get("category", ""), "Min": lim.min, "Max": lim.max,
                     "Unit": lim.unit, "Expected": "" if lim.numeric else (lim.expected if lim.expected != "negative"
                                                                         else (t["spec"] or "Negative")),
                     "Method": t.get("method", ""), "Required": "Y", "Notes": ""})
    return rows


def merge_ai(rule_findings, ai_findings, coa_warnings=()):
    """Add AI findings that the code rules didn't already make; AI document notes become one finding."""
    def key(item):
        return re.sub(r"\s*[—-]\s*method$", "", clean_name(item))
    have = {(key(f["Item"]), "method" in f["Item"].lower()) for f in rule_findings}
    have_items = {k for k, _ in have}
    out = list(rule_findings)
    for f in ai_findings:
        k, is_method = key(f["Item"]), f["Category"] == "Method" or "method" in f["Item"].lower()
        if (k, is_method) in have or (not is_method and k in have_items and f["Category"] != "Other"):
            continue
        out.append(f)
    notes = [w for w in coa_warnings if w and w.strip()]
    if notes:
        out.append(_finding("Medium", "Review", "Document notes", " ".join(notes),
                            "We noticed the following on the COA: " + " ".join(notes) + " Please confirm or send a clean copy.",
                            "Other", source="AI"))
    out.sort(key=lambda f: (SEV_ORDER.get(f["Severity"], 3), f["Category"], f["Item"]))
    return out
