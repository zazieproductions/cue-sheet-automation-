#!/usr/bin/env python3
"""
Fill the "Cue Sheet Template 2016 V3-6-5" workbook for
The Dark Awaits / Dark Echo Productions.

Revision 2 -- why this build differs from the first completed file
-----------------------------------------------------------------
The first completed file was structurally valid, but it did not *work*
well outside a full desktop-Excel recalculation:

1. New text was written as OOXML inline strings (t="inlineStr").  Excel
   normally stores worksheet text in xl/sharedStrings.xml, and some
   importers/viewers are unreliable with inline strings inside an Excel
   table.  This build appends new unique strings to sharedStrings.xml and
   writes ordinary t="s" cells, exactly as Excel itself would.
2. Formula cells still carried the blank template's cached results, so
   tools/viewers that do not recalculate (data-only readers, previews,
   some imports) showed "Music Cue Sheet Template", no sequence numbers,
   no durations and no Total Music Duration.  This build refreshes only
   the cached <v> result of formula cells whose inputs changed.  Every
   <f>...</f> formula remains byte-for-byte identical; Excel is also told
   to recalculate on load (fullCalcOnLoad), so it recomputes them anyway.
3. The first build repeated the cue title/usage/timings on both the
   Composer row and the Publisher row.  In this standardized template a
   cue is a grouped block: cue title, usage and timings belong on the
   first party row only, with additional writers/publishers on following
   rows.  That is how the template's sample cue sheet ("Weekly World of
   Music, Rock and Roll") is filled, and it is the only layout under
   which the template's own Seq. # formula numbers cues (1..6, not party
   rows 1..12) and its Total Music Duration formula returns the actual
   music duration (25:08, not a double-counted 50:16 -- longer than this
   45-minute program).

Preservation rules
------------------
* No worksheet, table, style, validation, conditional-formatting, merged
  cell, protection, printer-setting, theme or hidden-sheet part is
  redesigned.  Only these parts change:
    - xl/worksheets/sheet1.xml  (previously-empty input cells filled;
      formula *cached results* refreshed where inputs changed)
    - xl/sharedStrings.xml       (new unique strings appended; count and
      uniqueCount updated to match actual cell usage)
    - xl/workbook.xml            (fullCalcOnLoad="1" only)
* Every other package part, including xl/calcChain.xml, is copied
  byte-for-byte.  The set of formula cells is unchanged, so the saved
  calculation chain still references exactly the same cells; Excel
  rebuilds/validates it during the forced full calculation on load.
* Unknown production fields are left blank.  Nothing is invented.
"""

import re
import zipfile
from xml.etree import ElementTree as ET

SRC = "Cue_Sheet_Template_2016V3-6-5.xlsx"
DST = "Cue_Sheet_The_Dark_Awaits_COMPLETED.xlsx"

PROGRAM_TITLE = "The Dark Awaits"
PRODUCTION_COMPANY = "Dark Echo Productions"
PROGRAM_DURATION_MIN = 45
PROGRAM_DURATION_SEC = 0  # F13 already holds 0 in the template

# BMI registered writer name is "PRODUCTIONS, ZAZIE"; the template splits it
# into First (and Middle) Name / Last Name.  There is no writer-IPI field in
# this template, so writer IPI 01280391067 is intentionally not placed in
# the worksheet (see the review notes' Missing Information section).
WRITER_FIRST = "ZAZIE"
WRITER_LAST = "PRODUCTIONS"

# Column P is headed "IPI #" under Publisher.  Enter the supplied Publisher
# IPI there -- not BMI publisher account 4380631, which has no field in
# this template.  Store the IPI as text because column P's number format is
# "0"; a numeric 1372389230 would display without its leading zero.
PUBLISHER = "Zazie Productions Publishing"
PUBLISHER_IPI = "01372389230"

PRO = "BMI"
# Repository default (Zazie_Productions_Publishing.md): one writer at 100%
# writer share and one publisher at 100% publisher share, absent a contract
# stating otherwise.  Column R is formatted 0.00%, so 100% is stored as 1.
SHARE = 1
USAGE = "BI"  # Background Instrumental, for every cue

# (title, time_in (h,m,s), time_out (h,m,s)) in chronological order.
CUES = [
    ("Impact Wave Tick Interlude",                (0, 2, 32),  (0, 2, 35)),
    ("John's Story \u2014 Distorted Atmospheric Pad", (0, 2, 42),  (0, 2, 48)),
    ("Spirits Still Around",                      (0, 4, 6),   (0, 4, 25)),
    ("Industrial Electronic Interlude",           (0, 6, 12),  (0, 6, 49)),
    ("Mystic Bell Ambience",                      (0, 7, 32),  (0, 8, 0)),
    ("Family Investigation \u2014 Dark Ambience",      (0, 14, 25), (0, 38, 0)),
]

FIRST_CUE_ROW = 20  # first data row of the CueDetails2 table (A20:R999)
SST_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def esc(text):
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;"))


def seconds(hms):
    h, m, s = hms
    return h * 3600 + m * 60 + s


class SharedStrings:
    """Append-only view of xl/sharedStrings.xml.

    Existing <si> entries and their indexes never move.  A new string is
    appended once and then reused; every cell written through this object
    increments the workbook-wide shared-string cell count.
    """

    def __init__(self, xml):
        root = ET.fromstring(xml)
        self.original_count = int(root.attrib["count"])
        self.original_unique = int(root.attrib["uniqueCount"])
        self.index = {}
        for i, si in enumerate(root.findall("{%s}si" % SST_NS)):
            text = "".join(node.text or "" for node in si.iter("{%s}t" % SST_NS))
            # Excel keeps the first index for a repeated string in this file.
            self.index.setdefault(text, i)
        if len(root.findall("{%s}si" % SST_NS)) != self.original_unique:
            raise SystemExit("ERROR: template sharedStrings uniqueCount does not "
                             "match its <si> entries")
        self.additions = []
        self.cells_written = 0

    def ref(self, text):
        if text != text.strip():
            # A leading/trailing-space shared string would need
            # xml:space="preserve"; none of this cue sheet's values has one.
            raise SystemExit("ERROR: refusing shared string with edge spaces: %r" % text)
        self.cells_written += 1
        if text in self.index:
            return self.index[text]
        idx = self.original_unique + len(self.additions)
        self.index[text] = idx
        self.additions.append(text)
        return idx

    def build(self, xml):
        new_count = self.original_count + self.cells_written
        new_unique = self.original_unique + len(self.additions)
        attrs = 'count="%d" uniqueCount="%d"' % (new_count, new_unique)
        xml, n = re.subn(r'count="\d+" uniqueCount="\d+"', attrs, xml, count=1)
        if n != 1 or xml.count("</sst>") != 1:
            raise SystemExit("ERROR: could not update sharedStrings counts/closing tag")
        appended = "".join("<si><t>%s</t></si>" % esc(s) for s in self.additions)
        return xml.replace("</sst>", appended + "</sst>")


def fill_empty(xml, coord, inner, extra_attr=""):
    """Put content into an existing, still-empty cell, keeping its style."""
    pat = re.compile(r'<c r="%s"((?:\s[^>]*)?)/>' % re.escape(coord))
    matches = pat.findall(xml)
    if len(matches) != 1:
        raise SystemExit("ERROR: cell %s is not a single empty cell (%d matches)"
                         % (coord, len(matches)))
    return pat.sub(lambda m: '<c r="%s"%s%s>%s</c>'
                   % (coord, m.group(1), extra_attr, inner), xml, count=1)


def set_shared_text(xml, coord, text, sst):
    return fill_empty(xml, coord, "<v>%d</v>" % sst.ref(text),
                      extra_attr=' t="s"')


def set_number(xml, coord, number):
    return fill_empty(xml, coord, "<v>%s</v>" % number)


def replace_value(xml, coord, old, new):
    """Replace the value of an existing numeric input cell (keeps style)."""
    pat = re.compile(r'(<c r="%s"(?:\s[^>]*)?>)<v>%s</v>(</c>)'
                     % (re.escape(coord), re.escape(old)))
    if len(pat.findall(xml)) != 1:
        raise SystemExit("ERROR: cell %s does not hold the single value %r"
                         % (coord, old))
    return pat.sub(lambda m: "%s<v>%s</v>%s" % (m.group(1), new, m.group(2)),
                   xml, count=1)


def update_formula_cache(xml, coord, cached, numeric):
    """Refresh only a formula cell's cached result; keep <f> verbatim.

    A formula cell's t attribute describes the type of its cached result:
    t="str" for a string result, no t attribute for a numeric result.  The
    template's duration/sequence formulas can return either "" or a number,
    so moving from the blank template's "" cache to a numeric cache also
    requires dropping t="str" -- precisely what Excel writes after it
    calculates and saves the workbook.
    """
    pat = re.compile(
        r'<c r="%s"((?:\s[^>]*)?)>'
        r'(<f(?:\s[^>]*)?>.*?</f>)'
        r'(?:<v/>|<v>.*?</v>)'
        r'</c>' % re.escape(coord), re.DOTALL)
    matches = list(pat.finditer(xml))
    if len(matches) != 1:
        raise SystemExit("ERROR: formula cell %s not found exactly once (%d)"
                         % (coord, len(matches)))
    match = matches[0]
    attrs, formula = match.group(1), match.group(2)
    if numeric:
        attrs = attrs.replace(' t="str"', "")
        if 't="' in attrs:
            raise SystemExit("ERROR: unexpected type on numeric cache cell %s: %s"
                             % (coord, attrs))
        cache_xml = "<v>%s</v>" % cached
    else:
        if 't="str"' not in attrs:
            raise SystemExit("ERROR: string cache cell %s lacks t=\"str\": %s"
                             % (coord, attrs))
        cache_xml = "<v>%s</v>" % esc(str(cached))
    replacement = '<c r="%s"%s>%s%s</c>' % (coord, attrs, formula, cache_xml)
    return xml[:match.start()] + replacement + xml[match.end():]


def build_sheet(xml, sst):
    starts = [seconds(t_in) for _, t_in, _ in CUES]
    if starts != sorted(starts):
        raise SystemExit("ERROR: cues are not in chronological time-in order")

    # ---- program information -------------------------------------------
    xml = set_shared_text(xml, "N4", PROGRAM_TITLE, sst)       # Program Title
    xml = set_shared_text(xml, "N11", PRODUCTION_COMPANY, sst)  # Production Co.
    xml = replace_value(xml, "D13", "0", str(PROGRAM_DURATION_MIN))
    # F13 (seconds) already holds 0, so duration reads 45 min. 0 sec.
    # Episode title/number, network, airdate, classification, category and
    # all other unsupplied header fields remain blank.

    # ---- cues ------------------------------------------------------------
    # One two-row block per cue:
    #   first row      cue title, usage, time in/out + Composer party
    #   following row  Publisher party only (B:K left blank)
    total_music_seconds = 0
    for i, (title, t_in, t_out) in enumerate(CUES):
        first = FIRST_CUE_ROW + (i * 2)
        publisher_row = first + 1
        duration = seconds(t_out) - seconds(t_in)
        if duration <= 0:
            raise SystemExit("ERROR: cue %r has a non-positive duration" % title)
        total_music_seconds += duration

        # Cue-identifying fields: first row of the block only.  A/J/K are
        # formula columns; their formulas are left alone and only their
        # calculated caches are refreshed below.
        xml = set_shared_text(xml, "B%d" % first, title, sst)
        xml = set_shared_text(xml, "C%d" % first, USAGE, sst)
        for col, val in zip("DEFGHI", t_in + t_out):
            xml = set_number(xml, "%s%d" % (col, first), val)

        # Composer party on the first row.  O stays blank: conditional
        # formatting strikes Publisher Name when Role is Composer.
        xml = set_shared_text(xml, "L%d" % first, "Composer", sst)
        xml = set_shared_text(xml, "M%d" % first, WRITER_FIRST, sst)
        xml = set_shared_text(xml, "N%d" % first, WRITER_LAST, sst)
        xml = set_shared_text(xml, "Q%d" % first, PRO, sst)
        xml = set_number(xml, "R%d" % first, SHARE)

        # Publisher party on the following row.  M/N stay blank: the sheet
        # strikes writer names when Role is Publisher.  B:K stay blank so
        # Seq. # numbers cues and Total Music Duration is not doubled.
        xml = set_shared_text(xml, "L%d" % publisher_row, "Publisher", sst)
        xml = set_shared_text(xml, "O%d" % publisher_row, PUBLISHER, sst)
        xml = set_shared_text(xml, "P%d" % publisher_row, PUBLISHER_IPI, sst)
        xml = set_shared_text(xml, "Q%d" % publisher_row, PRO, sst)
        xml = set_number(xml, "R%d" % publisher_row, SHARE)

        # Cached results of the template's own A/J/K formulas for this cue.
        # Publisher-row A/J/K formulas evaluate to "" (B and F/I are blank),
        # which is already their cached value, so those cells are untouched.
        xml = update_formula_cache(xml, "A%d" % first, i + 1, numeric=True)
        xml = update_formula_cache(xml, "J%d" % first, duration // 60, numeric=True)
        xml = update_formula_cache(xml, "K%d" % first,
                                   (duration % 3600) % 60, numeric=True)

    # ---- cached results in the header ------------------------------------
    # A1: N6 (Episode Title) is blank, so the formula returns N4 alone.
    xml = update_formula_cache(xml, "A1", PROGRAM_TITLE, numeric=False)
    # A2: a non-revision sheet for a named program.
    xml = update_formula_cache(xml, "A2", "Music Cue Sheet", numeric=False)
    # D14/F14 sum J:J and K:K.  Only the six first-rows contribute.
    xml = update_formula_cache(xml, "D14", total_music_seconds // 60, numeric=True)
    xml = update_formula_cache(xml, "E14", "min.", numeric=False)
    xml = update_formula_cache(xml, "F14", total_music_seconds % 60, numeric=True)
    xml = update_formula_cache(xml, "G14", "sec.", numeric=False)
    xml = update_formula_cache(xml, "H14", "(auto calc.)", numeric=False)
    return xml, total_music_seconds


def build_workbook(xml):
    old = '<calcPr calcId="152511"/>'
    new = '<calcPr calcId="152511" fullCalcOnLoad="1"/>'
    if xml.count(old) != 1:
        raise SystemExit("ERROR: calcPr not found verbatim in workbook.xml")
    return xml.replace(old, new)


def main():
    with zipfile.ZipFile(SRC) as zin:
        infos = zin.infolist()
        parts = {item.filename: zin.read(item.filename) for item in infos}

    sst = SharedStrings(parts["xl/sharedStrings.xml"].decode("utf-8"))
    sheet_xml, total_music = build_sheet(
        parts["xl/worksheets/sheet1.xml"].decode("utf-8"), sst)
    parts["xl/worksheets/sheet1.xml"] = sheet_xml.encode("utf-8")
    parts["xl/sharedStrings.xml"] = sst.build(
        parts["xl/sharedStrings.xml"].decode("utf-8")).encode("utf-8")
    parts["xl/workbook.xml"] = build_workbook(
        parts["xl/workbook.xml"].decode("utf-8")).encode("utf-8")

    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in infos:  # preserve the template's package part order
            info = zipfile.ZipInfo(item.filename, item.date_time)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = item.create_system
            info.external_attr = item.external_attr
            info.internal_attr = item.internal_attr
            zout.writestr(info, parts[item.filename])

    print("wrote %s" % DST)
    print("sharedStrings: %d new string cells, %d new unique strings "
          "(count %d, uniqueCount %d)"
          % (sst.cells_written, len(sst.additions),
             sst.original_count + sst.cells_written,
             sst.original_unique + len(sst.additions)))
    print("total music (template formula caches): %d sec = %d min. %d sec."
          % (total_music, total_music // 60, total_music % 60))


if __name__ == "__main__":
    main()
