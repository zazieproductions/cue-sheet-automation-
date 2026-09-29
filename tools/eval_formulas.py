#!/usr/bin/env python3
"""Evaluate the template's own formula strings against the filled-in data.

The engine used here (`formulas`) cannot parse Excel table structured
references or INDIRECT, so this module takes the *verbatim* formula text out
of the untouched template and resolves only those two reference forms to the
literal cells they denote for the CueDetails2 table (headerRowCount=0, ref
A20:R999, so [#This Row] is the row itself and Column2..Column9 are B..I).
Everything else -- IFERROR/IF/OR/ISTEXT/COUNTA/INT/MOD/TRUNC/SUM and all the
arithmetic -- is the template's own formula, evaluated as-is on the values
read back out of the completed workbook.

`evaluate()` returns {cell coordinate: value} and is also what the preview
server renders, so the preview shows these numbers and nothing else.
"""
import re
import sys
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


def evaluate(src=SRC, dst=DST, first_row=20, last_row=31,
             tmp="/tmp/recalc.xlsx", extra=()):
    """Return {coord: value} for the template's calculated cells."""
    tpl = openpyxl.load_workbook(src)["Template"]
    done = openpyxl.load_workbook(dst)["Template"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    ws["A19"] = "Seq. #"                       # the label the seq formula tests for
    ws["A1"] = tpl["A1"].value                 # verbatim title formulas + precedents
    ws["A2"] = tpl["A2"].value
    for c in ("N4", "N6", "D4", "C13", "D13", "F13", "D14", "E14", "F14", "G14", "H14"):
        ws[c] = done[c].value if done[c].value is not None else tpl[c].value
    for r in range(first_row, last_row + 1):
        for c in "BCDEFGHI":                   # inputs read back from the deliverable
            ws["%s%d" % (c, r)] = done["%s%d" % (c, r)].value
        for c in "AKJ":                        # template's own formulas, resolved
            ws["%s%d" % (c, r)] = resolve(tpl["%s%d" % (c, r)].value, r)
    for c, v in extra:
        ws[c] = v

    wb.save(tmp)
    sol = formulas.ExcelModel().loads(tmp).finish().calculate()

    out = {}
    for k, v in sol.items():
        if "!" not in k:
            continue
        coord = k.split("!")[1]
        if ":" in coord:                       # skip range nodes
            continue
        try:
            out[coord] = v.value[0, 0]
        except Exception:
            out[coord] = v.value
    return out


def _fmt(v):
    """Render an engine value the way the template's number formats would."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def main():
    vals = evaluate()
    done = openpyxl.load_workbook(DST)["Template"]

    print("Template formulas evaluated on the completed workbook's data:\n")
    print("  %-5s %-42s %-9s %-9s %-9s" % ("row", "cue title", "usage", "dur m", "dur s"))
    total = 0
    for r in range(20, 32):
        m, s = vals.get("J%d" % r), vals.get("K%d" % r)
        total += int(float(m)) * 60 + int(float(s))
        print("  %-5d %-42s %-9s %-9s %-9s  (seq %s, role %s)"
              % (r, str(done["B%d" % r].value)[:42], done["C%d" % r].value,
                 _fmt(m), _fmt(s), _fmt(vals.get("A%d" % r)), done["L%d" % r].value))

    print("\n  D14 (Total Music Duration, min) = %s" % _fmt(vals.get("D14")))
    print("  F14 (Total Music Duration, sec) = %s" % _fmt(vals.get("F14")))
    print("  cross-check sum of J/K rows     = %ds = %dm %ds"
          % (total, total // 60, total % 60))
    print("  unique-cue music duration       = %ds = %dm %ds (each cue is listed twice: "
          "Composer + Publisher row)" % (total // 2, total // 120, (total // 2) % 60))
    print("  A1  (title formula)             = %r" % vals.get("A1"))
    print("  A2  (subtitle formula)          = %r" % vals.get("A2"))
    print("  H14 (auto-calc flag)            = %r" % vals.get("H14"))


if __name__ == "__main__":
    sys.exit(main())
