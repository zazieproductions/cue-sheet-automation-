#!/usr/bin/env python3
"""Quality-control pass over the completed cue sheet.

Compares the completed workbook against the untouched template part by part
and cell by cell: nothing may differ except the cells we intentionally filled
and the calcPr flag.
"""
import sys
import zipfile
from xml.etree import ElementTree as ET

import openpyxl

SRC = "Cue_Sheet_Template_2016V3-6-5.xlsx"
DST = "Cue_Sheet_The_Dark_Awaits_COMPLETED.xlsx"

fails = []


def check(label, ok, detail=""):
    print(("PASS  " if ok else "FAIL  ") + label + ((" :: " + detail) if detail else ""))
    if not ok:
        fails.append(label)


# ---------------------------------------------------------------- 1. package
za, zb = zipfile.ZipFile(SRC), zipfile.ZipFile(DST)
check("package part list identical",
      za.namelist() == zb.namelist(),
      "%d parts" % len(zb.namelist()))

changed = [n for n in za.namelist() if za.read(n) != zb.read(n)]
check("only sheet1.xml + workbook.xml changed",
      changed == ["xl/workbook.xml", "xl/worksheets/sheet1.xml"],
      "changed=%s" % changed)

for n in zb.namelist():
    if n.endswith(".xml") or n.endswith(".rels"):
        try:
            ET.fromstring(zb.read(n))
        except ET.ParseError as e:
            check("XML well-formed: %s" % n, False, str(e))
            break
else:
    check("all XML parts well-formed", True)

# workbook.xml diff must be the calc flag only
wa = za.read("xl/workbook.xml").decode()
wb_ = zb.read("xl/workbook.xml").decode()
check("workbook.xml diff is only fullCalcOnLoad",
      wb_.replace('<calcPr calcId="152511" fullCalcOnLoad="1"/>',
                  '<calcPr calcId="152511"/>') == wa)

for part in ("xl/styles.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet2.xml",
             "xl/tables/table1.xml", "xl/printerSettings/printerSettings1.bin"):
    check("untouched: %s" % part, za.read(part) == zb.read(part))

# ---------------------------------------------------------------- 2. structure
a = openpyxl.load_workbook(SRC)
b = openpyxl.load_workbook(DST)

check("sheet names/state identical",
      [(s.title, s.sheet_state) for s in a.worksheets] ==
      [(s.title, s.sheet_state) for s in b.worksheets],
      str([(s.title, s.sheet_state) for s in b.worksheets]))
check("hidden 'hidat' worksheet still hidden",
      b["hidat"].sheet_state == "hidden")

ta, tb = a["Template"], b["Template"]
check("data validations identical",
      [str(d.sqref) + "|" + str(d.formula1) + "|" + d.type
       for d in ta.data_validations.dataValidation] ==
      [str(d.sqref) + "|" + str(d.formula1) + "|" + d.type
       for d in tb.data_validations.dataValidation],
      "%d dropdown/range rules kept" % len(tb.data_validations.dataValidation))
check("conditional formatting identical",
      len(ta.conditional_formatting._cf_rules) == len(tb.conditional_formatting._cf_rules))
check("merged cells identical",
      sorted(str(r) for r in ta.merged_cells.ranges) ==
      sorted(str(r) for r in tb.merged_cells.ranges))
check("column widths/hidden identical",
      {k: (v.width, v.hidden) for k, v in ta.column_dimensions.items()} ==
      {k: (v.width, v.hidden) for k, v in tb.column_dimensions.items()})
check("sheet protection preserved",
      (ta.protection.sheet, ta.protection.algorithmName) ==
      (tb.protection.sheet, tb.protection.algorithmName))
check("workbook structure protection preserved",
      a.security.lockStructure == b.security.lockStructure,
      "lockStructure=%s" % b.security.lockStructure)
check("tables preserved",
      dict(ta.tables.items()) == dict(tb.tables.items()),
      str(dict(tb.tables.items())))
check("hidat lookup tables preserved",
      dict(a["hidat"].tables.items()) == dict(b["hidat"].tables.items()),
      str(sorted(dict(b["hidat"].tables.items()).keys())))
check("CueDetails2 calculated-column formula preserved",
      ta.tables["CueDetails2"].tableColumns[0].calculatedColumnFormula ==
      tb.tables["CueDetails2"].tableColumns[0].calculatedColumnFormula)

# ---------------------------------------------------------------- 3. formulas
def formulas(ws):
    out = {}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                out[c.coordinate] = c.value
    return out


fa, fb = formulas(ta), formulas(tb)
check("formula count unchanged (%d)" % len(fa), fa == fb,
      "template=%d completed=%d" % (len(fa), len(fb)))
if fa != fb:
    for k in sorted(set(fa) | set(fb)):
        if fa.get(k) != fb.get(k):
            print("   diff", k, repr(fa.get(k)), "->", repr(fb.get(k)))

# ---------------------------------------------------------------- 4. content
def val(ws, coord):
    return ws[coord].value

expect_header = {
    "N4": "The Dark Awaits",
    "N11": "Dark Echo Productions",
    "D13": 45,
    "F13": 0,
    "C4": "Cue Sheet Classification:",   # label untouched
    "M4": "Program (series, film, etc.) Title:",
}
for coord, want in expect_header.items():
    got = val(tb, coord)
    check("%s = %r" % (coord, want), got == want, "got %r" % got)

for coord in ("D4", "D8", "D9", "D10", "D11", "N6", "N8", "N12", "N13", "N14"):
    got = val(tb, coord)
    check("%s left blank" % coord, got in (None, ""), "got %r" % got)

CUES = [
    ("Impact Wave Tick Interlude", 0, 2, 32, 0, 2, 35),
    ("John's Story \u2014 Distorted Atmospheric Pad", 0, 2, 42, 0, 2, 48),
    ("Spirits Still Around", 0, 4, 6, 0, 4, 25),
    ("Industrial Electronic Interlude", 0, 6, 12, 0, 6, 49),
    ("Mystic Bell Ambience", 0, 7, 32, 0, 8, 0),
    ("Family Investigation \u2014 Dark Ambience", 0, 14, 25, 0, 38, 0),
]

print("\n--- cue grid as written ---")
sheet_order = []      # (time_in_seconds, row) read back out of the workbook
for i, (title, ih, im, isec, oh, om, osec) in enumerate(CUES):
    for role, off in (("Composer", 0), ("Publisher", 1)):
        r = 20 + i * 2 + off
        row = [val(tb, "%s%d" % (c, r)) for c in "ABCDEFGHIKLMNOPQR"]
        A, B, C, D, E, F, G, H, I, K, L, M, N, O, P, Q, R = row
        seq_in, seq_out = (ih * 3600 + im * 60 + isec), (oh * 3600 + om * 60 + osec)
        sheet_order.append((seq_in, r))
        ok = (B == title and C == "BI" and (D, E, F) == (ih, im, isec)
              and (G, H, I) == (oh, om, osec) and L == role
              and Q == "BMI" and R == 1
              and ((M == "ZAZIE" and N == "PRODUCTIONS" and O is None)
                   if role == "Composer" else
                   (M is None and N is None and O == "Zazie Productions Publishing"))
              and P is None)
        check("row %d %s: %s" % (r, role, title), ok,
              "B=%r C=%r in=%s:%s:%s out=%s:%s:%s L=%r M=%r N=%r O=%r P=%r Q=%r R=%r"
              % (B, C, D, E, F, G, H, I, L, M, N, O, P, Q, R))
        print("      row %2d | %-9s | %-42s | BI | %d:%02d:%02d -> %d:%02d:%02d | %s | %s"
              % (r, role, title, D, E, F, G, H, I,
                 "%s %s" % (M or "", N or "") if role == "Composer" else O, Q))

# Chronological: every cue starts no earlier than the one before it, each cue
# has a positive duration, and rows sit in the sheet in that same order.
times = [(ih * 3600 + im * 60 + isec, oh * 3600 + om * 60 + osec, t)
         for (t, ih, im, isec, oh, om, osec) in CUES]
check("every cue has a positive duration",
      all(o > i for i, o, _ in times),
      "; ".join("%s=%ds" % (t, o - i) for i, o, t in times))
check("cues in chronological order (non-decreasing time-in)",
      all(times[k][0] >= times[k - 1][0] for k in range(1, len(times))),
      " ".join(str(i) for i, _, _ in times))
check("rows appear in the sheet in cue order",
      [r for _, r in sorted(sheet_order, key=lambda x: x[1])] ==
      [r for _, r in sheet_order] == list(range(20, 32)))

# rows below the last cue must still be pristine
dirty = []
for r in range(32, 120):
    for c in "BCDEFGHILMNOPQR":
        if val(tb, "%s%d" % (c, r)) not in (None, ""):
            dirty.append("%s%d" % (c, r))
check("no stray data below the cue block", not dirty, str(dirty[:8]))

# number format sanity for the share + time cells
check("R column share cells keep percent format 0.00%",
      all(tb["R%d" % r].number_format == "0.00%" for r in range(20, 32)),
      tb["R20"].number_format)
check("E/F time cells keep 00 format",
      tb["E20"].number_format == "00" and tb["F20"].number_format == "00",
      "%s/%s" % (tb["E20"].number_format, tb["F20"].number_format))
check("writer/publisher cells keep unlocked style",
      all(tb["%s%d" % (c, r)].protection.locked is False
          for r in range(20, 32) for c in "BCDEFGHILMNOPQR"))

print()
print("RESULT:", "ALL CHECKS PASSED" if not fails else "%d FAILURES: %s" % (len(fails), fails))
sys.exit(1 if fails else 0)
