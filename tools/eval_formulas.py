#!/usr/bin/env python3
"""Evaluate the template's own formula strings against the filled-in data.

The engine used here (`formulas`) cannot parse Excel table structured
references or INDIRECT, so this script takes the *verbatim* formula text out
of the untouched template and resolves only those two reference forms to the
literal cells they denote for the CueDetails2 table (headerRowCount=0, ref
A20:R999, so [#This Row] is the row itself and Column2..Column9 are B..I).
Everything else -- IFERROR/IF/OR/ISTEXT/COUNTA/INT/MOD/TRUNC/SUM and all the
arithmetic -- is the template's own formula, evaluated as-is on the values
read back out of the completed workbook.
"""
import re
import warnings

import formulas
import openpyxl
from openpyxl import Workbook

warnings.filterwarnings("ignore")

SRC = "Cue_Sheet_Template_2016V3-6-5.xlsx"
DST = "Cue_Sheet_The_Dark_Awaits_COMPLETED.xlsx"
COLMAP = {i + 2: chr(ord("A") + i + 1) for i in range(18)}   # Column1 -> A ...


def resolve(f, row):
    """Structured refs / INDIRECT -> literal A1 refs for one cue row."""
    f = re.sub(r"CueDetails2\[\[#This Row\],\[Column(\d+)\]\]",
               lambda m: "%s%d" % (COLMAP[int(m.group(1))], row), f)
    f = f.replace('INDIRECT("B20")', "B20")
    f = f.replace('INDIRECT("A"&ROW()-1)', "A%d" % (row - 1))
    assert "INDIRECT" not in f and "[[#This Row]" not in f, f
    return f


tpl = openpyxl.load_workbook(SRC)["Template"]
done = openpyxl.load_workbook(DST)["Template"]

wb = Workbook()
ws = wb.active
ws.title = "Template"
ws["A19"] = "Seq. #"                      # the label the seq formula tests for
ws["A1"] = tpl["A1"].value                # verbatim title formulas + precedents
ws["A2"] = tpl["A2"].value
for c in ("N4", "N6", "D4"):
    ws[c] = done[c].value
ws["C13"] = tpl["C13"].value
ws["D13"] = done["D13"].value             # 45 (program duration, minutes)
ws["F13"] = done["F13"].value
for name, coord in (("D14", "D14"), ("F14", "F14")):
    ws[coord] = tpl[coord].value          # verbatim template total formulas

for r in range(20, 32):
    for c in "BCDEFGHI":                  # inputs read back from the deliverable
        ws["%s%d" % (c, r)] = done["%s%d" % (c, r)].value
    for c in "AKJ":
        ws["%s%d" % (c, r)] = resolve(tpl["%s%d" % (c, r)].value, r)

wb.save("/tmp/recalc.xlsx")
sol = formulas.ExcelModel().loads("/tmp/recalc.xlsx").finish().calculate()


def get(coord):
    for k, v in sol.items():
        if "!" in k and k.split("!")[1] == coord.upper():
            try:
                return v.value[0, 0]
            except Exception:
                return v.value
    return "<not calculated>"


print("Template formulas evaluated on the completed workbook's data:\n")
print("  %-5s %-42s %-9s %-9s %-9s" % ("row", "cue title", "usage", "dur m", "dur s"))
total = 0
for r in range(20, 32):
    m, s = get("J%d" % r), get("K%d" % r)
    total += int(m) * 60 + int(s)
    print("  %-5d %-42s %-9s %-9s %-9s  (seq %s, role %s)"
          % (r, str(done["B%d" % r].value)[:42], get("C%d" % r) or done["C%d" % r].value,
             m, s, get("A%d" % r), done["L%d" % r].value))

print("\n  D14 (Total Music Duration, min) = %s" % get("D14"))
print("  F14 (Total Music Duration, sec) = %s" % get("F14"))
print("  cross-check sum of J/K rows     = %ds = %dm %ds"
      % (total, total // 60, total % 60))
print("  unique-cue music duration       = %ds = %dm %ds (each cue is listed twice: "
      "Composer + Publisher row)" % (total // 2, total // 120, (total // 2) % 60))
print("  A1  (title formula)             = %r" % get("A1"))
print("  A2  (subtitle formula)          = %r" % get("A2"))
print("  H14 (auto-calc flag)            = %r" % get("H14"))
