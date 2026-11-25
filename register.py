"""Read the Master Training Register (Excel) into plain Python objects.

Columns are found by their HEADER TEXT, not by position, so the parser keeps working
when the register gains or loses columns (for example the NFC column moving around).
"""
import datetime as dt
import re
from dataclasses import dataclass, field

import openpyxl

from .companies import merge_company
from .util import clean, to_date


@dataclass
class Training:
    key: int          # column number in the register
    name: str
    ttype: str        # 'expiry' | 'issue' | 'induction'
    inhouse: bool = False


@dataclass
class Register:
    source: str
    employees: list = field(default_factory=list)
    trainings: list = field(default_factory=list)
    has_nfc: bool = False
    has_category: bool = False
    has_medical: bool = False
    unmapped_companies: list = field(default_factory=list)
    columns_info: list = field(default_factory=list)   # what was detected, shown on the Register check page
    excluded: list = field(default_factory=list)      # inactive (red) rows left out of every report
    companies: list = field(default_factory=list)   # parent names, biggest first

    def training_by_regex(self, pattern):
        rx = re.compile(pattern, re.I)
        for t in self.trainings:
            if rx.search(t.name):
                return t
        return None


_EXP_RX = re.compile(r"expir|\bexp\b|valid|\bdue\b|renew|\btill\b|up ?to", re.I)
_ISSUE_RX = re.compile(r"issu|induction date|training date", re.I)

_FIELD_PATTERNS = {
    "name": r"^NAME$",
    "designation": r"^DESIGNATION$",
    "category": r"^EMPLOYEE CATEGORY$",
    "emp_no": r"^EMP\.? ?NO",
    "company": r"^COMPANY$",
    "package": r"^PACKAGE$",
    "nfc": r"NFC|CARD PRINTED",
    "fieldglass": r"FIELD ?GLASS",
}


def _is_red(color):
    """True for red font / fill colours (FF0000, C00000 ... and Excel's indexed red)."""
    if color is None:
        return False
    try:
        if color.type == "rgb" and isinstance(color.rgb, str) and len(color.rgb) == 8:
            r, g, b = (int(color.rgb[i:i + 2], 16) for i in (2, 4, 6))
            return r >= 0xB0 and g <= 0x50 and b <= 0x50
        if color.type == "indexed":
            return color.indexed in (10, 16)
    except Exception:
        pass
    return False


def _cell_red(cell):
    font = getattr(cell, "font", None)
    if font is not None and _is_red(font.color):
        return True
    fill = getattr(cell, "fill", None)
    return bool(fill is not None and fill.fill_type == "solid" and _is_red(fill.fgColor))


def _find_sheet(wb, wanted):
    if wanted in wb.sheetnames:
        return wb[wanted]
    for ws in wb.worksheets:                       # fall back to any sheet that has a NAME header
        for row in ws.iter_rows(min_row=1, max_row=8, values_only=True):
            if any(clean(v).upper() == "NAME" for v in row):
                return ws
    raise ValueError("Register sheet not found. Expected a sheet named '%s'." % wanted)


def parse_register(path, settings, source_name=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = _find_sheet(wb, settings["register_sheet_name"])
    rows, red_cells = [], []
    for r in ws.iter_rows():
        rows.append(tuple(getattr(c, "value", None) for c in r))
        red_cells.append([i for i, c in enumerate(r) if _cell_red(c)])
    wb.close()

    header_idx = next((i for i, r in enumerate(rows[:10])
                       if any(clean(v).upper() == "NAME" for v in r)), None)
    if header_idx is None:
        raise ValueError("Header row with a 'NAME' column was not found in the first 10 rows.")
    head = [clean(v) for v in rows[header_idx]]
    width = len(head)

    cols = {}
    for key, pattern in _FIELD_PATTERNS.items():
        for i, h in enumerate(head):
            if h and re.search(pattern, h.upper()):
                cols[key] = i
                break
    if "name" not in cols or "company" not in cols:
        raise ValueError("NAME and COMPANY columns are required in the register.")

    # ---- sub-header row (Issue Date / Expiry Date ...) is optional ----
    nxt = list(rows[header_idx + 1]) + [None] * width if header_idx + 1 < len(rows) else [None] * width
    name_cell = clean(nxt[cols["name"]])
    has_sub = (not name_cell) or bool(_EXP_RX.search(name_cell) or _ISSUE_RX.search(name_cell) or re.fullmatch(r"(?i)date|name", name_cell))
    sub = [clean(v) for v in nxt[:width]] if has_sub else [""] * width
    data_start = header_idx + (2 if has_sub else 1)

    def label(i):
        """Text that describes column i: its sub-header plus its own header."""
        return (sub[i] + " " + head[i]).strip() if i < width else ""

    # ---- medical block: header is only on its first column, Issue/Expiry are found by their labels ----
    med_issue = med_expiry = med_freq = med_result = None
    med_start = next((i for i, h in enumerate(head)
                      if "MEDICAL" in h.upper() and ("ADNOC" in h.upper() or "ONSHORE" in h.upper() or "PERIODIC" in h.upper())
                      and not re.search(r"FREQ|RESULT|STATUS", h.upper())), None)
    if med_start is not None:
        block = [c for c in range(med_start, min(med_start + 4, width)) if c == med_start or not head[c]]
        for c in block:
            lab = label(c).lower()
            if _EXP_RX.search(lab) and med_expiry is None:
                med_expiry = c
            elif "issu" in lab and med_issue is None:
                med_issue = c
        if med_expiry is None and len(block) >= 2:          # unlabeled pair: first = issue, second = expiry
            med_issue, med_expiry = block[0], block[1]
        elif med_expiry is None and len(block) == 1:        # single column
            med_expiry = block[0]
    for i, h in enumerate(head):
        if "FREQUENCY" in h.upper():
            med_freq = i
        if re.search(r"MEDICAL\s*RESULT", h.upper()):
            med_result = i

    # ---- training columns: start at the HSE induction column, or the first expiry-looking column ----
    start = next((i for i, h in enumerate(head) if "INDUCTION" in h.upper() and "HSE" in h.upper()), None)
    if start is None:
        after = max([c for c in (med_start, med_expiry, med_result) if c is not None] or [0])
        start = next((i for i in range(after + 1, width) if _EXP_RX.search(label(i))), after + 1)

    # a training = its header cell + the following columns with an empty header (merged cell over Issue / Expiry)
    groups = []
    for i in range(start, width):
        if head[i]:
            groups.append({"name": head[i], "cols": [i]})
        elif groups and sub[i] and not re.search(r"remark", sub[i], re.I):
            groups[-1]["cols"].append(i)
    overrides = settings.get("training_type_overrides", {})
    trainings = []
    for g in groups:
        name = re.sub(r"[\s\-\u2013(]*(expiry|expire|expiration|validity)( date)?\)?\s*$", "", g["name"], flags=re.I) or g["name"]
        if re.search(r"remark", name, re.I):
            continue
        exp_cols = [c for c in g["cols"] if _EXP_RX.search(label(c))]
        if "induction" in name.lower():
            ttype, key = "induction", g["cols"][0]
        elif exp_cols:
            ttype, key = "expiry", exp_cols[-1]
        else:
            ttype, key = "issue", g["cols"][0]
        for needle, forced in overrides.items():
            if needle.lower() in name.lower():
                if forced == "issue":
                    ttype, key = "issue", g["cols"][0]
                elif forced == "expiry" and ttype != "expiry":
                    ttype, key = "expiry", g["cols"][-1]
        trainings.append(Training(key=key, name=name, ttype=ttype))

    inhouse_from = next((k for k, t in enumerate(trainings)
                         if t.name.lower().startswith(settings["inhouse_starts_at"].lower())), None)
    if inhouse_from is not None:
        for t in trainings[inhouse_from:]:
            t.inhouse = True

    reg = Register(source=source_name or str(path), trainings=trainings,
                   has_nfc="nfc" in cols, has_category="category" in cols,
                   has_medical=med_expiry is not None)

    from openpyxl.utils import get_column_letter as _L
    reg.columns_info = [{"col": _L(t.key + 1), "name": t.name, "type": t.ttype, "inhouse": t.inhouse,
                         "header": head[t.key] if t.key < width else "", "sub": sub[t.key]} for t in trainings]
    if med_expiry is not None:
        reg.columns_info.insert(0, {"col": _L(med_expiry + 1), "name": "ADNOC MEDICAL (expiry)", "type": "expiry",
                                    "inhouse": False, "header": head[med_start] if med_start is not None else "", "sub": sub[med_expiry]})
    unmapped = {}
    skip_red = settings.get("exclude_red_rows", True)
    min_red = settings.get("red_row_min_cells", 10)
    for ridx in range(data_start, len(rows)):
        r = list(rows[ridx]) + [None] * (width - len(rows[ridx]))
        name = clean(r[cols["name"]])
        if not name:
            continue
        reds = red_cells[ridx]
        if skip_red and (cols["name"] in reds or len(reds) >= min_red):
            reg.excluded.append({"row": ridx + 1, "name": name, "emp_no": clean(r[cols["emp_no"]]) if "emp_no" in cols else "",
                                 "company": clean(r[cols["company"]])})
            continue
        company_raw = clean(r[cols["company"]])
        parent, mapped = merge_company(company_raw, settings)
        if not mapped:
            unmapped[company_raw] = unmapped.get(company_raw, 0) + 1
        cat_raw = clean(r[cols["category"]]) if "category" in cols else ""
        m = re.match(r"^\s*(\d+)", cat_raw)
        nfc = None
        if "nfc" in cols:
            nfc = clean(r[cols["nfc"]]).upper().startswith("P")
        emp = {
            "name": name,
            "designation": clean(r[cols["designation"]]) if "designation" in cols else "",
            "emp_no": clean(r[cols["emp_no"]]) if "emp_no" in cols else "",
            "package": clean(r[cols["package"]]) if "package" in cols else "",
            "category_raw": cat_raw,
            "category_no": int(m.group(1)) if m else None,
            "company_raw": company_raw,
            "company": parent,
            "nfc": nfc,
            "med_issue": to_date(r[med_issue]) if med_issue is not None else None,
            "med_expiry": to_date(r[med_expiry]) if med_expiry is not None else None,
            "med_freq": clean(r[med_freq]) if med_freq is not None else "",
            "med_result": clean(r[med_result]) if med_result is not None else "",
            "t": {t.key: to_date(r[t.key]) for t in trainings},
        }
        reg.employees.append(emp)

    counts = {}
    for e in reg.employees:
        counts[e["company"]] = counts.get(e["company"], 0) + 1
    reg.companies = [c for c, _ in sorted(counts.items(), key=lambda kv: -kv[1])]
    reg.unmapped_companies = sorted(unmapped.items(), key=lambda kv: -kv[1])
    return reg


def resolve_as_of(settings):
    value = settings.get("as_of_date")
    if value:
        return dt.datetime.strptime(value, "%Y-%m-%d").date()
    return dt.date.today()
