"""Fill the HSE Training Compliance Audit Checklist (ATTACHMENT-A) from the register.

The template is edited at XML level, so charts, logos, comments and formatting stay exactly
as they are. Only these cells are touched:
  * header cells (project / contractor / audit date)
  * master category table (NO. OF EMPLOYEE)            -> numbers from the register
  * every category row: TO ATTEND (formula), ATTENDED (number), COMPLIANCE % (formula)
Everything else (sub-totals, COMPLIANCE % CALCULATIONS, COMPLIANCE CHART) recalculates in Excel.
The 18 Yes/No questions in COMPLIANCE CHECKLIST are never touched.
"""
import datetime as dt
import difflib
import io
import re
import zipfile

import openpyxl
from lxml import etree
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, column_index_from_string

from .util import norm, status_of

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
Q = lambda t: "{%s}%s" % (NS, t)
MAIN_SHEET = "ATTACHMENT -A"
CALC_SHEET = "COMPLIANCE % CALCULATIONS"


# ---------------------------------------------------------------- XML helpers
def _split(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


class Sheet:
    def __init__(self, xml_bytes):
        self.root = etree.fromstring(xml_bytes)
        self.data = self.root.find(Q("sheetData"))
        self.rows = {int(r.get("r")): r for r in self.data}
        self.cells = {}
        for r in self.data:
            for c in r:
                self.cells[c.get("r")] = c

    def cell(self, ref):
        c = self.cells.get(ref)
        if c is not None:
            return c
        col, rn = _split(ref)
        row = self.rows.get(rn)
        if row is None:
            raise KeyError("Row %d missing in template" % rn)
        c = etree.SubElement(row, Q("c"), r=ref)
        # keep cells ordered by column
        row[:] = sorted(row, key=lambda x: column_index_from_string(_split(x.get("r"))[0]))
        self.cells[ref] = c
        return c

    @staticmethod
    def _clear(c):
        for ch in list(c):
            c.remove(ch)
        if "t" in c.attrib:
            del c.attrib["t"]

    def set_number(self, ref, value):
        c = self.cell(ref)
        self._clear(c)
        etree.SubElement(c, Q("v")).text = repr(value) if isinstance(value, float) else str(value)

    def set_formula(self, ref, formula):
        c = self.cell(ref)
        self._clear(c)
        etree.SubElement(c, Q("f")).text = formula.lstrip("=")

    def set_text(self, ref, text):
        c = self.cell(ref)
        self._clear(c)
        c.set("t", "inlineStr")
        is_ = etree.SubElement(c, Q("is"))
        etree.SubElement(is_, Q("t")).text = text

    def unshare_formulas(self):
        masters = {}
        for c in self.cells.values():
            f = c.find(Q("f"))
            if f is not None and f.get("t") == "shared" and f.text:
                masters[f.get("si")] = (c.get("r"), f.text)
        for c in self.cells.values():
            f = c.find(Q("f"))
            if f is not None and f.get("t") == "shared":
                origin, text = masters[f.get("si")]
                new = Translator("=" + text, origin=origin).translate_formula(c.get("r"))[1:] \
                    if c.get("r") != origin else text
                for a in ("t", "si", "ref"):
                    f.attrib.pop(a, None)
                f.text = new

    def strip_cached_values(self):
        for c in self.cells.values():
            if c.find(Q("f")) is not None:
                v = c.find(Q("v"))
                if v is not None:
                    c.remove(v)
                if c.get("t") in ("str", "e", "b", "n"):
                    del c.attrib["t"]

    def tobytes(self):
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------- layout discovery
def _read_layout(template_path):
    wb = openpyxl.load_workbook(template_path)
    ws = wb[MAIN_SHEET]
    # master category table
    master = {}                                    # category no -> (row, count column letter, name)
    first_block = next(r for r in range(1, ws.max_row + 1)
                       if str(ws.cell(r, 2).value or "").strip().upper().startswith("TRAINING CATEGORY"))
    for r in range(15, first_block - 1):
        for sn_col, name_col, cnt_col in (("A", "B", "G"), ("J", "K", "Q")):
            sn = ws["%s%d" % (sn_col, r)].value
            if isinstance(sn, (int, float)) and float(sn).is_integer() and 1 <= sn <= 56:
                master[int(sn)] = (r, cnt_col, str(ws["%s%d" % (name_col, r)].value or "").strip())
    # training blocks
    blocks = []
    starts = [r for r in range(1, ws.max_row + 1)
              if str(ws.cell(r, 2).value or "").strip().upper().startswith("TRAINING CATEGORY")]
    for s in starts:
        title = re.sub(r"\s+", " ", str(ws.cell(s, 5).value or "")).strip()
        hdr = next(r for r in range(s, s + 6) if str(ws.cell(r, 1).value or "").strip().upper() == "SN")
        rows, r = [], hdr + 1
        while True:
            i_val = ws.cell(r, 9).value
            if isinstance(i_val, str) and i_val.upper().startswith("=SUM"):
                subtotal = r
                break
            if ws.cell(r, 2).value:
                rows.append((r, str(ws.cell(r, 2).value).strip()))
            r += 1
            if r > s + 80:
                raise ValueError("Sub-total row not found for block '%s'" % title)
        blocks.append({"title": title, "start": s, "rows": rows, "subtotal": subtotal})
    # which side (3rd party / in-house) each sub-total belongs to, from the calculations sheet
    side = {}
    wc = wb[CALC_SHEET]
    for row in wc.iter_rows(min_row=5):
        for c in row:
            v = c.value
            if isinstance(v, str):
                m = re.match(r"='ATTACHMENT -A'!I(\d+)$", v)
                if m:
                    side[int(m.group(1))] = "3rd party" if c.column <= 7 else "In-house"
    for b in blocks:
        b["side"] = side.get(b["subtotal"], "?")
    return master, blocks


def _tokens(s):
    return set(re.findall(r"[a-z0-9]+", s.lower())) - {"training", "awareness", "level", "the", "and", "for"}


def _score(a, b):
    na, nb = norm(a), norm(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    ta, tb = _tokens(a), _tokens(b)
    sub = 0.0
    if ta and tb and (ta <= tb or tb <= ta):
        sub = 0.9
    return max(sub, difflib.SequenceMatcher(None, na, nb).ratio())


def _match_training(title, reg, settings, h2s):
    t_norm = title.lower()
    if re.search(r"h2s.*level.?1", t_norm):
        return h2s["l1"], 1.0, "H2S rule"
    if re.search(r"h2s.*level.?2", t_norm):
        return h2s["l2"], 1.0, "H2S rule"
    for key, needle in settings["checklist"]["block_overrides"].items():
        if key in t_norm:
            for t in reg.trainings:
                if needle in t.name.lower():
                    return t, 1.0, "override"
    best, best_s = None, 0.0
    for t in reg.trainings:
        s = _score(title, t.name)
        if s > best_s:
            best, best_s = t, s
    if best_s >= settings["checklist"]["min_match_score"]:
        return best, best_s, "name match"
    return None, best_s, "no match"


# ---------------------------------------------------------------- main entry
def fill_checklist(reg, settings, as_of, template_path, audit_no=None):
    soon = settings["expiry_soon_days"]
    keys = settings["key_trainings"]
    master, blocks = _read_layout(template_path)
    h2s = {"l1": reg.training_by_regex(keys["h2s_l1"]), "l2": reg.training_by_regex(keys["h2s_l2"])}
    warnings = []

    # headcount per category
    heads = {n: 0 for n in master}
    no_cat = 0
    for e in reg.employees:
        if e["category_no"] in heads:
            heads[e["category_no"]] += 1
        else:
            no_cat += 1
    if no_cat:
        warnings.append("%d employees have no valid Employee Category (1-56) and are not counted." % no_cat)
    if not reg.has_category:
        warnings.append("Register has no EMPLOYEE CATEGORY column - checklist cannot be filled correctly.")

    def valid(e, t):
        return status_of(e["t"][t.key], t.ttype, as_of, soon) in ("valid", "soon")

    def held_by_category(train):
        """People who have done this training at all (valid OR expired) per employee category."""
        d = {n: 0 for n in master}
        for e in reg.employees:
            if e["category_no"] in d and e["t"][train.key] is not None:
                d[e["category_no"]] += 1
        return d

    # attended per training per category
    cache = {}

    def attended(train, pool_rule=False):
        k = (train.key, pool_rule)
        if k in cache:
            return cache[k]
        d = {n: 0 for n in master}
        for e in reg.employees:
            n = e["category_no"]
            if n not in d or not valid(e, train):
                continue
            if pool_rule and h2s["l1"] is not None and valid(e, h2s["l1"]):
                continue                    # already counted in L1
            d[n] += 1
        cache[k] = d
        return d

    # ---- open template
    zin = zipfile.ZipFile(template_path)
    wbxml = etree.fromstring(zin.read("xl/workbook.xml"))
    rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    rid_target = {r.get("Id"): r.get("Target") for r in rels}
    sheet_path = {}
    for s in wbxml.find(Q("sheets")):
        sheet_path[s.get("name")] = "xl/" + rid_target[s.get("{%s}id" % RNS)].lstrip("/").replace("xl/", "")
    sheets = {n: Sheet(zin.read(p)) for n, p in sheet_path.items()}
    for sh in sheets.values():
        sh.unshare_formulas()
    main = sheets[MAIN_SHEET]

    # ---- header
    main.set_text("E5", settings.get("project_short", "BFN"))
    main.set_text("N6", settings["contractor"])
    serial = (as_of - dt.date(1899, 12, 30)).days
    main.set_number("E6", serial)
    if audit_no:
        main.set_text("E7", audit_no)

    # ---- master table
    for n, (row, col, name) in master.items():
        main.set_number("%s%d" % (col, row), heads[n])
    master_by_name = {norm(v[2]): n for n, v in master.items()}

    held_l1 = held_by_category(h2s["l1"]) if h2s["l1"] else {}

    # ---- blocks
    report = []
    l1_rows = {}                                       # category no -> row in L1 block
    for b in blocks:
        train, score, how = _match_training(b["title"], reg, settings, h2s)
        entry = {"title": b["title"], "side": b["side"], "training": train.name if train else None,
                 "score": round(score, 2), "how": how, "to_attend": 0, "attended": 0, "rows": len(b["rows"])}
        is_l1 = bool(re.search(r"h2s.*level.?1", b["title"].lower()))
        is_l2 = bool(re.search(r"h2s.*level.?2", b["title"].lower()))
        if train is None:
            warnings.append("Block '%s' has no matching column in the register (best score %.2f) - left unchanged."
                            % (b["title"], score))
            report.append(entry)
            continue
        if how == "name match" and score < 0.8:
            warnings.append("Block '%s' matched to '%s' with low confidence (%.2f) - please check."
                            % (b["title"], train.name, score))
        if is_l2 and settings.get("h2s_pool_rule", True):
            att = attended(train, pool_rule=True)
        else:
            att = attended(train)
        for row, name in b["rows"]:
            n = master_by_name.get(norm(name))
            if n is None:                              # fall back to closest master name
                cand = difflib.get_close_matches(norm(name), list(master_by_name), n=1, cutoff=0.8)
                n = master_by_name[cand[0]] if cand else None
            if n is None:
                warnings.append("Row %d '%s' (%s) not found in the master category table - skipped."
                                % (row, name, b["title"]))
                continue
            mrow, mcol, _ = master[n]
            head_ref = "$%s$%d" % (mcol, mrow)
            if is_l1:
                l1_rows[n] = row
            if is_l2 and settings.get("h2s_pool_rule", True) and n in l1_rows:
                main.set_formula("I%d" % row, "MAX(0,%s-J%d)" % (head_ref, l1_rows[n]))
            elif is_l1:
                # H2S L1: TO ATTEND = people who have done L1 (valid or expired), ATTENDED = those still valid
                main.set_number("I%d" % row, held_l1[n])
            else:
                main.set_formula("I%d" % row, head_ref)
            main.set_number("J%d" % row, att[n])
            main.set_formula("K%d" % row, 'IF(I%d=0,"",J%d/I%d)' % (row, row, row))
            entry["attended"] += att[n]
        report.append(entry)

    # python-side totals for the on-screen preview
    l1_att = attended(h2s["l1"]) if h2s["l1"] else {}
    for entry, b in zip(report, blocks):
        tot = 0
        is_l2 = bool(re.search(r"h2s.*level.?2", b["title"].lower())) and settings.get("h2s_pool_rule", True)
        is_l1_block = bool(re.search(r"h2s.*level.?1", b["title"].lower()))
        for row, name in b["rows"]:
            n = master_by_name.get(norm(name))
            if n is None:
                continue
            if is_l2:
                tot += max(0, heads[n] - l1_att.get(n, 0))
            elif is_l1_block:
                tot += held_l1.get(n, 0)
            else:
                tot += heads[n]
        entry["to_attend"] = tot
        entry["pct"] = round(100.0 * entry["attended"] / tot, 1) if tot and entry["training"] else None

    # plain A/B percentage formulas -> 0 instead of #DIV/0! when a block has nobody to train
    for name in (MAIN_SHEET, CALC_SHEET):
        for c in sheets[name].cells.values():
            f = c.find(Q("f"))
            if f is not None and f.text and re.match(r"^[A-Z]+\d+/[A-Z]+\d+$", f.text):
                f.text = "IFERROR(%s,0)" % f.text

    # ---- housekeeping: no stale cached numbers, force recalculation
    for sh in sheets.values():
        sh.strip_cached_values()
    cp = wbxml.find(Q("calcPr"))
    if cp is None:
        cp = etree.SubElement(wbxml, Q("calcPr"))
    cp.set("fullCalcOnLoad", "1")

    out = io.BytesIO()
    changed = {sheet_path[n]: sh.tobytes() for n, sh in sheets.items()}
    changed["xl/workbook.xml"] = etree.tostring(wbxml, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename.endswith("/"):
                continue
            zout.writestr(item.filename, changed.get(item.filename, zin.read(item.filename)))
    summary = {"3rd party": [0, 0], "In-house": [0, 0]}
    for e in report:
        if e["side"] in summary and e["training"]:
            summary[e["side"]][0] += e["attended"]
            summary[e["side"]][1] += e["to_attend"]
    totals = {k: {"attended": a, "to_attend": t, "pct": round(100.0 * a / t, 1) if t else None}
              for k, (a, t) in summary.items()}
    return out.getvalue(), {"blocks": report, "warnings": warnings, "totals": totals,
                            "headcount_total": sum(heads.values()), "excluded": len(reg.excluded)}
