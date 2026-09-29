#!/usr/bin/env python3
"""
Fill in the "Cue Sheet Template 2016 V3-6-5" workbook for
The Dark Awaits / Dark Echo Productions.

Design constraints (per instruction):
  * No structural change to the workbook: every part other than
    xl/worksheets/sheet1.xml and xl/workbook.xml is copied byte-for-byte.
  * No formula is written, replaced or deleted -- only previously empty
    input cells receive values, plus one workbook calc flag so Excel
    recalculates the untouched formulas on open.
  * Values are written as inline strings / plain numbers, so
    sharedStrings.xml and styles.xml are left completely alone and the
    original cell style indexes (s="..") are preserved.
"""

import re
import shutil
import zipfile

SRC = "Cue_Sheet_Template_2016V3-6-5.xlsx"
DST = "Cue_Sheet_The_Dark_Awaits_COMPLETED.xlsx"

PROGRAM_TITLE = "The Dark Awaits"
PRODUCTION_COMPANY = "Dark Echo Productions"
PROGRAM_DURATION_MIN = 45

WRITER_FIRST = "ZAZIE"            # from BMI registered name "PRODUCTIONS, ZAZIE"
WRITER_LAST = "PRODUCTIONS"
PUBLISHER = "Zazie Productions Publishing"
PUBLISHER_IPI = "01372389230"      # stored as text: column P is number-formatted
                                   # "0", which would drop the leading zero
PRO = "BMI"
SHARE = 1                          # 100% -- column R is number-formatted 0.00%

USAGE = "BI"

# (title, time_in (h,m,s), time_out (h,m,s)) in chronological order
CUES = [
    ("Impact Wave Tick Interlude",               (0, 2, 32),  (0, 2, 35)),
    ("John's Story \u2014 Distorted Atmospheric Pad", (0, 2, 42),  (0, 2, 48)),
    ("Spirits Still Around",                     (0, 4, 6),   (0, 4, 25)),
    ("Industrial Electronic Interlude",          (0, 6, 12),  (0, 6, 49)),
    ("Mystic Bell Ambience",                     (0, 7, 32),  (0, 8, 0)),
    ("Family Investigation \u2014 Dark Ambience",     (0, 14, 25), (0, 38, 0)),
]

FIRST_CUE_ROW = 20   # first data row of the CueDetails2 table


def esc(text):
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;"))


def fill_empty(xml, coord, inner, extra_attr=""):
    """Put content into an existing, still-empty cell, keeping its style."""
    pat = re.compile(r'<c r="%s"((?:\s[^>]*)?)/>' % re.escape(coord))
    matches = pat.findall(xml)
    if len(matches) != 1:
        raise SystemExit("ERROR: cell %s is not a single empty cell (%d matches)"
                         % (coord, len(matches)))
    return pat.sub(lambda m: '<c r="%s"%s%s>%s</c>'
                   % (coord, m.group(1), extra_attr, inner), xml, count=1)


def set_text(xml, coord, text):
    return fill_empty(xml, coord,
                      "<is><t>%s</t></is>" % esc(text),
                      extra_attr=' t="inlineStr"')


def set_number(xml, coord, number):
    return fill_empty(xml, coord, "<v>%s</v>" % number)


def replace_value(xml, coord, old, new):
    """Replace the cached value of an existing numeric cell (keeps style)."""
    pat = re.compile(r'(<c r="%s"(?:\s[^>]*)?>)<v>%s</v>(</c>)'
                     % (re.escape(coord), re.escape(old)))
    matches = pat.findall(xml)
    if len(matches) != 1:
        raise SystemExit("ERROR: cell %s does not hold the single value %r"
                         % (coord, old))
    return pat.sub(lambda m: "%s<v>%s</v>%s" % (m.group(1), new, m.group(2)),
                   xml, count=1)


def build_sheet(xml):
    # ---- program information -------------------------------------------
    xml = set_text(xml, "N4", PROGRAM_TITLE)        # Program (series, film, etc.) Title
    xml = set_text(xml, "N11", PRODUCTION_COMPANY)  # Production Company
    xml = replace_value(xml, "D13", "0", str(PROGRAM_DURATION_MIN))  # minutes
    # F13 (seconds) already 0 -> program duration reads 45 min. 0 sec.

    # ---- cues: one Composer row + one Publisher row per cue ------------
    for i, (title, t_in, t_out) in enumerate(CUES):
        for role, offset in (("Composer", 0), ("Publisher", 1)):
            r = FIRST_CUE_ROW + (i * 2) + offset
            xml = set_text(xml, "B%d" % r, title)     # Cue Title
            xml = set_text(xml, "C%d" % r, USAGE)     # Usage
            for col, val in zip("DEFGHI", t_in + t_out):
                xml = set_number(xml, "%s%d" % (col, r), val)
            # J/K (Duration) are formulas -> left untouched, they compute.
            xml = set_text(xml, "L%d" % r, role)      # Role
            if role == "Composer":
                xml = set_text(xml, "M%d" % r, WRITER_FIRST)
                xml = set_text(xml, "N%d" % r, WRITER_LAST)
                # O (Publisher Name) stays blank on composer rows: the
                # template strikes it out when Role = Composer/Arranger.
            else:
                # M/N (writer first/last) stay blank on publisher rows:
                # the template strikes them out when Role = Publisher.
                xml = set_text(xml, "O%d" % r, PUBLISHER)
                # P = Publisher IPI # (the template's only IPI column, hidden)
                xml = set_text(xml, "P%d" % r, PUBLISHER_IPI)
            xml = set_text(xml, "Q%d" % r, PRO)       # PRO / Affiliation
            xml = set_number(xml, "R%d" % r, SHARE)   # Shares % (0.00% format)
    return xml


def build_workbook(xml):
    old = '<calcPr calcId="152511"/>'
    new = '<calcPr calcId="152511" fullCalcOnLoad="1"/>'
    if xml.count(old) != 1:
        raise SystemExit("ERROR: calcPr not found verbatim in workbook.xml")
    return xml.replace(old, new)


def main():
    edits = {
        "xl/worksheets/sheet1.xml": build_sheet,
        "xl/workbook.xml": build_workbook,
    }
    with zipfile.ZipFile(SRC) as zin, \
         zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in edits:
                data = edits[item.filename](data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    print("wrote", DST)


if __name__ == "__main__":
    main()
