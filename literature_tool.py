#!/usr/bin/env python3
"""
literature_tool.py - keep all literature workbooks and macro metadata in ONE master Excel file.

Standard library only (zipfile, xml.etree, csv). No openpyxl, no pandas, no pip install.

Commands
--------
  merge  : per-target files  -> master workbook
      python literature_tool.py merge --out literature_master.xlsx  FOLDER [FOLDER ...]
      python literature_tool.py merge --base literature_master.xlsx --out literature_master.xlsx  NEW_FOLDER
      Reads every  <target>_literature_workbook.xlsx  and  macro_metadata_<target>.csv  found in the
      folders (sub-folders included). With --base, targets already in the master are kept and any target
      found in the folders replaces its old rows.

  split  : master workbook -> per-target files the pipeline reads
      python literature_tool.py split literature_master.xlsx OUTDIR [--targets rwa,tier1_capital]
      Writes OUTDIR/config/<target>_literature_workbook.xlsx and OUTDIR/data/macro_metadata_<target>.csv

  list   : show the targets in a master workbook and what each one has
      python literature_tool.py list literature_master.xlsx

  check  : basic checks on every target in the master (one core model, lags 0-4, allowed transforms,
           model variables defined as concepts, sign agreement)
      python literature_tool.py check literature_master.xlsx

Edit the master in Excel (filter the 'target' column), then run split to regenerate the files.
"""
import csv
import io
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}

WB_SHEETS = ["models", "concepts", "settings", "references", "how to fill"]
MODEL_COLS = ["model", "source", "variable", "transform", "lag", "expected_sign", "core"]
CONCEPT_COLS = ["concept_id", "concept", "channel", "expected_sign", "effect_on_loan_growth", "why",
                "evidence", "transform", "theory_lags", "core", "long_run_form", "include",
                "moodys_name", "notes"]
SETTING_COLS = ["setting", "value", "meaning"]
REF_COLS = ["short", "reference", "link"]
META_COLS = ["mnemonic", "description", "geography", "theme", "transform", "expected_sign",
             "sign_reason", "changed_from_mortgage"]
TRANSFORMS = {"dlog", "yoy", "dlog2", "acc", "gap", "diff", "diff2", "diff4", "lvl", "rgap"}

WB_RE = re.compile(r"^(?:[0-9a-f]{8}-)?(?P<t>.+?)_literature_workbook\.xlsx$", re.I)
META_RE = re.compile(r"^(?:[0-9a-f]{8}-)?macro_metadata_(?P<t>.+?)\.csv$", re.I)


# ----------------------------------------------------------------------------------------------
# xlsx reading (stdlib)
# ----------------------------------------------------------------------------------------------
def _col_index(ref):
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _num(text):
    try:
        f = float(text)
    except ValueError:
        return text
    return int(f) if f.is_integer() and "e" not in text.lower() and "." not in text else f


def read_xlsx(path):
    """Return {sheet_name: [row_list, ...]} with values as str / int / float / None."""
    out = {}
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si", NS):
                shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target_of = {r.get("Id"): r.get("Target") for r in rels.findall("pr:Relationship", NS)}
        for sh in wb.find("m:sheets", NS).findall("m:sheet", NS):
            rid = sh.get("{%s}id" % NS["r"])
            tgt = target_of[rid].lstrip("/")
            tgt = tgt if tgt.startswith("xl/") else "xl/" + tgt
            root = ET.fromstring(z.read(tgt))
            rows = []
            for row in root.iter("{%s}row" % NS["m"]):
                r_idx = int(row.get("r")) - 1 if row.get("r") else len(rows)
                while len(rows) < r_idx:
                    rows.append([])
                vals = []
                for c in row.findall("m:c", NS):
                    ci = _col_index(c.get("r")) if c.get("r") else len(vals)
                    while len(vals) < ci:
                        vals.append(None)
                    t = c.get("t")
                    v = c.find("m:v", NS)
                    if t == "s":
                        val = shared[int(v.text)]
                    elif t == "inlineStr":
                        is_ = c.find("m:is", NS)
                        val = "" if is_ is None else "".join(x.text or "" for x in is_.iter("{%s}t" % NS["m"]))
                    elif t in ("str", "e"):
                        val = v.text if v is not None else ""
                    elif t == "b":
                        val = bool(int(v.text)) if v is not None else None
                    else:
                        val = _num(v.text) if v is not None and v.text is not None else None
                    vals.append(val)
                rows.append(vals)
            out[sh.get("name")] = rows
    return out


def table(rows, cols=None):
    """Sheet rows -> list of dicts keyed by the header row (empty rows dropped)."""
    if not rows:
        return []
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    recs = []
    for r in rows[1:]:
        if not any(v not in (None, "") for v in r):
            continue
        d = {h: (r[i] if i < len(r) else None) for i, h in enumerate(header) if h}
        recs.append(d)
    if cols:
        recs = [{c: d.get(c) for c in cols} for d in recs]
    return recs


# ----------------------------------------------------------------------------------------------
# xlsx writing (stdlib)
# ----------------------------------------------------------------------------------------------
def _col_letter(i):
    s = ""
    i += 1
    while i:
        i, rem = divmod(i - 1, 26)
        s = chr(65 + rem) + s
    return s


_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _cell(ref, v, style):
    st = ' s="%d"' % style if style else ""
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return '<c r="%s" t="b"%s><v>%d</v></c>' % (ref, st, int(v))
    if isinstance(v, (int, float)):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, st, repr(v) if isinstance(v, float) else v)
    text = escape(_BAD.sub("", str(v)))
    sp = ' xml:space="preserve"' if text != text.strip() or "\n" in text else ""
    return '<c r="%s" t="inlineStr"%s><is><t%s>%s</t></is></c>' % (ref, st, sp, text)


def write_xlsx(path, sheets, widths=None):
    """sheets: list of (name, rows). Row 1 is written bold and frozen."""
    widths = widths or {}
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   + "".join('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
                             for i in range(len(sheets)))
                   + "</Types>")
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   "</Relationships>")
        z.writestr("xl/workbook.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="%s" xmlns:r="%s"><sheets>' % (NS["m"], NS["r"])
                   + "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(n, {'"': "&quot;"}), i + 1, i + 1)
                             for i, (n, _) in enumerate(sheets))
                   + "</sheets></workbook>")
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1)
                             for i in range(len(sheets)))
                   + '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (len(sheets) + 1)
                   + "</Relationships>")
        z.writestr("xl/styles.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<styleSheet xmlns="%s">' % NS["m"]
                   + '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
                     '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
                     '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
                     '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill></fills>'
                     '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                     '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                     '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                     '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
                     '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
                     "</styleSheet>")
        for i, (name, rows) in enumerate(sheets):
            ncols = max((len(r) for r in rows), default=1)
            w = widths.get(name, {})
            cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (j + 1, j + 1, w.get(j, 16))
                           for j in range(ncols))
            body = []
            for ri, r in enumerate(rows):
                cells = "".join(_cell("%s%d" % (_col_letter(ci), ri + 1), v, 1 if ri == 0 else 0)
                                for ci, v in enumerate(r))
                body.append('<row r="%d">%s</row>' % (ri + 1, cells))
            pane = ('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
                    'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>') if len(rows) > 1 else ""
            z.writestr("xl/worksheets/sheet%d.xml" % (i + 1),
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<worksheet xmlns="%s">%s<cols>%s</cols><sheetData>%s</sheetData></worksheet>'
                       % (NS["m"], pane, cols, "".join(body)))


# ----------------------------------------------------------------------------------------------
# master model
# ----------------------------------------------------------------------------------------------
def _clean(v):
    return "" if v is None else v


def _csv_value(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _sign(v):
    if isinstance(v, str) and re.fullmatch(r"[+-]?\d+", v.strip()):
        return int(v)
    return v


class Master:
    def __init__(self):
        self.targets = {}     # name -> dict(label, notes)
        self.models, self.concepts, self.settings, self.howto = [], [], [], []
        self.refs = []        # dicts short, reference, link, used_by(set)
        self.meta = []        # dicts target + META_COLS

    # ---- load / save master ----
    @classmethod
    def load(cls, path):
        m = cls()
        s = read_xlsx(path)
        for d in table(s.get("targets", [])):
            m.targets[str(d["target"])] = {"label": _clean(d.get("label")), "notes": _clean(d.get("notes"))}
        m.models = table(s.get("models", []), ["target"] + MODEL_COLS)
        m.concepts = table(s.get("concepts", []), ["target"] + CONCEPT_COLS)
        m.settings = table(s.get("settings", []), ["target"] + SETTING_COLS)
        m.howto = table(s.get("how_to_fill", []), ["target", "line"])
        for d in table(s.get("references", []), REF_COLS + ["used_by"]):
            d["used_by"] = set(x.strip() for x in str(_clean(d["used_by"])).split(",") if x.strip())
            m.refs.append(d)
        m.meta = table(s.get("metadata", []), ["target"] + META_COLS)
        for lst in (m.models, m.concepts, m.settings, m.howto, m.meta):
            for d in lst:
                d["target"] = str(d["target"])
                m.targets.setdefault(d["target"], {"label": "", "notes": ""})
        return m

    def save(self, path):
        tnames = sorted(self.targets)
        trows = [["target", "label", "has_workbook", "has_metadata", "concepts", "models", "core_model", "notes"]]
        for t in tnames:
            core = sorted({str(d["model"]) for d in self.models if d["target"] == t and str(_clean(d["core"])).lower() == "yes"})
            trows.append([t, self.targets[t]["label"],
                          "yes" if any(d["target"] == t for d in self.models) else "no",
                          "yes" if any(d["target"] == t for d in self.meta) else "no",
                          sum(d["target"] == t for d in self.concepts),
                          len({d["model"] for d in self.models if d["target"] == t}),
                          "; ".join(core), self.targets[t]["notes"]])

        def rows(lst, cols):
            return [cols] + [[d.get(c) for c in cols] for d in lst]

        refs = [REF_COLS + ["used_by"]] + [[d["short"], d["reference"], d["link"], ", ".join(sorted(d["used_by"]))]
                                            for d in sorted(self.refs, key=lambda d: str(d["short"]).lower())]
        # sign matrix: one row per mnemonic, one column per target (view only)
        mt = sorted({d["target"] for d in self.meta})
        first, signs = {}, {}
        for d in self.meta:
            first.setdefault(d["mnemonic"], d)
            signs[(d["mnemonic"], d["target"])] = d["expected_sign"]
        matrix = [["mnemonic", "description", "transform"] + mt]
        for mn, d in first.items():
            matrix.append([mn, d["description"], d["transform"]] + [signs.get((mn, t)) for t in mt])
        readme = [["literature_master.xlsx - how to use"]] + [[l] for l in README_LINES]
        sheets = [("README", readme), ("targets", trows),
                  ("concepts", rows(self.concepts, ["target"] + CONCEPT_COLS)),
                  ("models", rows(self.models, ["target"] + MODEL_COLS)),
                  ("settings", rows(self.settings, ["target"] + SETTING_COLS)),
                  ("references", refs),
                  ("metadata", rows(self.meta, ["target"] + META_COLS)),
                  ("sign_matrix", matrix),
                  ("how_to_fill", rows(self.howto, ["target", "line"]))]
        widths = {"README": {0: 140}, "targets": {0: 26, 1: 70, 6: 40, 7: 60},
                  "concepts": {0: 24, 1: 22, 2: 30, 5: 50, 6: 40, 13: 40},
                  "models": {0: 24, 1: 30, 2: 40, 3: 22}, "settings": {0: 24, 1: 22, 2: 30, 3: 70},
                  "references": {0: 30, 1: 90, 2: 50, 3: 60},
                  "metadata": {0: 24, 1: 16, 2: 45, 7: 60, 8: 60}, "sign_matrix": {1: 45},
                  "how_to_fill": {0: 24, 1: 140}}
        write_xlsx(path, sheets, widths)

    # ---- per-target files ----
    def drop_target(self, t):
        for name in ("models", "concepts", "settings", "howto", "meta"):
            setattr(self, name, [d for d in getattr(self, name) if d["target"] != t])
        for d in self.refs:
            d["used_by"].discard(t)
        self.refs = [d for d in self.refs if d["used_by"]]

    def add_workbook(self, t, path):
        s = read_xlsx(path)
        missing = [n for n in WB_SHEETS if n not in s]
        if missing:
            raise ValueError("%s: missing sheet(s) %s" % (path, missing))
        for lst, sheet, cols in ((self.models, "models", MODEL_COLS), (self.concepts, "concepts", CONCEPT_COLS),
                                 (self.settings, "settings", SETTING_COLS)):
            for d in table(s[sheet], cols):
                lst.append(dict(target=t, **d))
        for d in table(s["references"], REF_COLS):
            key = (d["short"], d["reference"], d["link"])
            for r in self.refs:
                if (r["short"], r["reference"], r["link"]) == key:
                    r["used_by"].add(t)
                    break
            else:
                self.refs.append(dict(d, used_by={t}))
        label = ""
        for r in s["how to fill"]:
            line = r[0] if r else None
            self.howto.append({"target": t, "line": line})
            if isinstance(line, str) and line.startswith("Target:"):
                label = line[len("Target:"):].strip()
        info = self.targets.setdefault(t, {"label": "", "notes": ""})
        info["label"] = label or info["label"]

    def add_metadata(self, t, path, note=None):
        with open(path, newline="", encoding="utf-8-sig") as f:
            for d in csv.DictReader(f):
                rec = {"target": t}
                for c in META_COLS:
                    rec[c] = d.get(c, "")
                rec["expected_sign"] = _sign(rec["expected_sign"])
                self.meta.append(rec)
        info = self.targets.setdefault(t, {"label": "", "notes": ""})
        if note:
            info["notes"] = note

    def write_target(self, t, outdir):
        written = []
        if any(d["target"] == t for d in self.models + self.concepts):
            def rows(lst, cols):
                return [cols] + [[d.get(c) for c in cols] for d in lst if d["target"] == t]
            refs = [REF_COLS] + [[d["short"], d["reference"], d["link"]] for d in self.refs if t in d["used_by"]]
            how = [[d["line"]] for d in self.howto if d["target"] == t]
            p = os.path.join(outdir, "config", "%s_literature_workbook.xlsx" % t)
            write_xlsx(p, [("models", rows(self.models, MODEL_COLS)), ("concepts", rows(self.concepts, CONCEPT_COLS)),
                           ("settings", rows(self.settings, SETTING_COLS)), ("references", refs),
                           ("how to fill", how or [["How to fill this workbook"]])],
                       {"models": {0: 30, 1: 40, 2: 22}, "concepts": {1: 30, 5: 50, 6: 40, 13: 40},
                        "settings": {0: 24, 1: 30, 2: 70}, "references": {0: 30, 1: 90, 2: 50},
                        "how to fill": {0: 140}})
            written.append(p)
        meta = [d for d in self.meta if d["target"] == t]
        if meta:
            p = os.path.join(outdir, "data", "macro_metadata_%s.csv" % t)
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(META_COLS)
                for d in meta:
                    w.writerow([_csv_value(d[c]) for c in META_COLS])
            written.append(p)
        return written


README_LINES = [
    "One file for all literature workbooks and macro metadata. Every sheet has a 'target' column; filter it to work on one target.",
    "targets      : one row per target (label, what it has, core model). Summary only - it is rebuilt on every save.",
    "concepts     : the 'concepts' sheet of every literature workbook, with a target column.",
    "models       : the 'models' sheet of every literature workbook, with a target column. Exactly one core model per target.",
    "settings     : the 'settings' sheet of every workbook (ECM, channels, regime threshold, sign rule).",
    "references   : every study once; used_by lists the targets whose workbook cites it. Add a target name to used_by to cite it there.",
    "metadata     : the macro_metadata_<target>.csv files stacked (400 Moody's series per target). Edit expected_sign / sign_reason here.",
    "sign_matrix  : VIEW ONLY - expected_sign of every series for every target side by side. Edits here are ignored; edit 'metadata'.",
    "how_to_fill  : the 'how to fill' text of each workbook.",
    "",
    "Regenerate the files the pipeline reads (standard-library Python, no packages):",
    "    python literature_tool.py split literature_master.xlsx OUTDIR",
    "    -> OUTDIR/config/<target>_literature_workbook.xlsx and OUTDIR/data/macro_metadata_<target>.csv",
    "Add or replace targets from per-target files (e.g. the mortgage workbook and metadata):",
    "    python literature_tool.py merge --base literature_master.xlsx --out literature_master.xlsx FOLDER",
    "Check every target:  python literature_tool.py check literature_master.xlsx",
    "",
    "Rules (same as each workbook): transform in dlog, yoy, dlog2, acc, gap, diff, diff2, diff4, lvl, rgap; lag 0-4; "
    "expected_sign +1/-1 (0 in metadata keeps a series out of the data-driven screen).",
]


# ----------------------------------------------------------------------------------------------
# commands
# ----------------------------------------------------------------------------------------------
def find_inputs(folders):
    wbs, metas = {}, {}
    for folder in folders:
        for root, _, files in os.walk(folder):
            for fn in files:
                if fn.startswith("~$"):
                    continue
                mw, mm = WB_RE.match(fn), META_RE.match(fn)
                if mw:
                    wbs[mw.group("t")] = os.path.join(root, fn)
                elif mm:
                    metas[mm.group("t")] = os.path.join(root, fn)
    return wbs, metas


def cmd_merge(args):
    out, base, folders = None, None, []
    it = iter(args)
    for a in it:
        if a == "--out":
            out = next(it)
        elif a == "--base":
            base = next(it)
        else:
            folders.append(a)
    if not out or not folders:
        sys.exit("usage: merge --out MASTER.xlsx [--base MASTER.xlsx] FOLDER [FOLDER ...]")
    m = Master.load(base) if base else Master()
    wbs, metas = find_inputs(folders)
    for t in sorted(set(wbs) | set(metas)):
        if t in m.targets:
            m.drop_target(t)
        if t in wbs:
            m.add_workbook(t, wbs[t])
        if t in metas:
            m.add_metadata(t, metas[t])
        print("  %-28s workbook=%s metadata=%s" % (t, "yes" if t in wbs else "-", "yes" if t in metas else "-"))
    m.save(out)
    print("wrote %s (%d targets)" % (out, len(m.targets)))


def cmd_split(args):
    if len(args) < 2:
        sys.exit("usage: split MASTER.xlsx OUTDIR [--targets a,b]")
    m = Master.load(args[0])
    targets = sorted(m.targets)
    if "--targets" in args:
        targets = [x.strip() for x in args[args.index("--targets") + 1].split(",") if x.strip()]
    n = 0
    for t in targets:
        for p in m.write_target(t, args[1]):
            n += 1
            print("  " + p)
    print("wrote %d files" % n)


def cmd_list(args):
    m = Master.load(args[0])
    for t in sorted(m.targets):
        print("%-28s concepts=%-3d models=%-2d metadata_rows=%-4d %s" % (
            t, sum(d["target"] == t for d in m.concepts), len({d["model"] for d in m.models if d["target"] == t}),
            sum(d["target"] == t for d in m.meta), m.targets[t]["label"][:60]))


def cmd_check(args):
    m = Master.load(args[0])
    problems = 0
    for t in sorted(m.targets):
        models = [d for d in m.models if d["target"] == t]
        concepts = {str(d["concept_id"]): d for d in m.concepts if d["target"] == t}
        if not models:
            continue
        msgs = []
        cores = {d["model"] for d in models if str(_clean(d["core"])).lower() == "yes"}
        if len(cores) != 1:
            msgs.append("needs exactly one core model, found %d" % len(cores))
        for d in models:
            if str(d["transform"]) not in TRANSFORMS:
                msgs.append("model %s: transform %r not allowed" % (d["model"], d["transform"]))
            if not (isinstance(d["lag"], int) and 0 <= d["lag"] <= 4):
                msgs.append("model %s: lag %r outside 0-4" % (d["model"], d["lag"]))
            c = concepts.get(str(d["variable"]))
            if c is None:
                msgs.append("model %s: variable %s is not a concept (fine only if it is a series name)" % (d["model"], d["variable"]))
            elif d["expected_sign"] not in (None, "") and c["expected_sign"] not in (None, "", 0) \
                    and _sign(d["expected_sign"]) != _sign(c["expected_sign"]):
                msgs.append("model %s: sign of %s differs from the concept's" % (d["model"], d["variable"]))
        for c in concepts.values():
            if c["transform"] not in (None, "") and str(c["transform"]) not in TRANSFORMS:
                msgs.append("concept %s: transform %r not allowed" % (c["concept_id"], c["transform"]))
        print("%-28s %s" % (t, "ok" if not msgs else "%d issue(s)" % len(msgs)))
        for x in msgs:
            print("    - " + x)
        problems += len(msgs)
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in ("merge", "split", "list", "check"):
        print(__doc__)
        sys.exit(1)
    {"merge": cmd_merge, "split": cmd_split, "list": cmd_list, "check": cmd_check}[sys.argv[1]](sys.argv[2:])
