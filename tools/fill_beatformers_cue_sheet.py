#!/usr/bin/env python3
"""
Fill the flat (non-AcroForm) PDF "Beatformers - Music Cue Sheet-2.pdf"
(Famous Studios, sent by Ricky Taitt on 2026-11-02) for Zazie Productions.

The source PDF has no form fields, so entries are overlaid as typed text
in DejaVuSerif (the typeface the form itself is set in), sitting on the
form's own printed underline in each "Details" cell.  No original text,
line or table is moved, covered or deleted; signature/date lines and
everything marked "(production fills)" are left blank.

Every value below is sourced -- see Cue_Sheet_Beatformers_REVIEW_NOTES.md
for the full provenance.  Short version:
  * Writer / publisher IPIs, BMI account numbers: Zazie_Productions_Publishing.md
    and the artist's BMI correspondence of 2026-09-29 / 2026-11-02.
  * Titles, ISWCs, BMI Work #, writer identity + IPI, Manuscript status:
    "BMI Catalog Detail.pdf" dated 2026-09-29 (Drive).
    - TIMELESS RETRO SPLURGE      ISWC T3372349996, BMI Work #77028290
      (NOT "TIMELESS RETRO SPLURGE EDGE", BMI Work #77273295, Slynth Music)
    - SOUND SYSTEM NEOCORTEX ACTIVATION  ISWC T3310780775, BMI Work #71873631
  * ISRCs: Timeless = QZHNB2278384 (DistroKid sales/streaming report, Drive;
    the ISRC database spreadsheet wrongly shows "--" because it only listed
    the Bandcamp issue), Sound System = QZDA52378036 ("BMI Sound System Info"
    doc + ISRC database + streaming report).
  * Lengths, album, release date, (P) line: Apple Music / Deezer release
    metadata for the two EPs (DistroKid UPCs 196864398824 / 197631604926).
  * Genre "Experimental Electronic" for both: artist instruction, 2026-11-02.
  * Usage (BI / BV), publisher named with publisher share left blank,
    city/state-only address, credit wording: artist decisions, 2026-11-02.
"""

import pymupdf

SRC = "Beatformers - Music Cue Sheet-2.pdf"
DST = "Beatformers - Music Cue Sheet-2 - Zazie Productions (filled, unsigned).pdf"
FONTFILE = "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"
FONTNAME = "DejaVuSerifFill"
INK = (0, 0, 0)

LEGAL_NAME = "Zazie Kanwar-Torge"
STAGE_NAME = "Zazie Productions"
ADDRESS_LINES = ["Asheville,", "North Carolina"]
PHONE = "+1 (864) 430-0606"
EMAIL = "zaziediya@gmail.com"
# BMI registered writer identity for these two works: PRODUCTIONS, ZAZIE
WRITER_BMI_NAME_LINES = ["PRODUCTIONS,", "ZAZIE"]
WRITER_IPI = "01280391067"
PRO = "BMI"
PUBLISHER_LINES = ["Zazie Productions", "Publishing"]
PUBLISHER_IPI = "01372389230"
MASTER_OWNER_LINES = ["Zazie", "Productions LLC"]
MASTER_OWNER = "Zazie Productions LLC"
GENRE = "Experimental Electronic"

SONG_1 = {  # page 2, top table
    "title": "Timeless Retro Splurge",
    "alt_title": "N/A",
    "iswc": "T-337.234.999-6",
    "isrc": "QZHNB2278384",
    "length": "02:16",
    "album": "Interference Archive 01010101 - EP",
    "release_date": "May 19, 2021",
    "usage": "Background Instrumental",
}
SONG_2 = {  # page 2, bottom table; splits + credit on page 3
    "title": "Sound System: Neocortex Activation",
    "alt_title": "N/A",
    "iswc": "T-331.078.077-5",
    "isrc": "QZDA52378036",
    "length": "02:16",
    "album": "Distant Bells For Surgical Minds - EP",
    "release_date": "February 10, 2023",
    "usage": "Background Vocal",
}

_font = pymupdf.Font(fontfile=FONTFILE)


def text_width(text, size):
    return _font.text_length(text, fontsize=size)


def put(page, x, baseline, text, size=7.5, max_width=None):
    """Insert one line, shrinking the size until it fits max_width."""
    s = size
    if max_width:
        while s > 4.8 and text_width(text, s) > max_width:
            s -= 0.2
        if text_width(text, s) > max_width:
            raise SystemExit("Does not fit (%.1fpt in %.1fpt): %r"
                             % (text_width(text, s), max_width, text))
    page.insert_text((x, baseline), text, fontname=FONTNAME,
                     fontfile=FONTFILE, fontsize=s, color=INK)
    return s


def put_stacked(page, x, underline_y, lines, size=5.6, max_width=None, leading=6.0):
    """Two short lines in a one-line-high cell; last line sits on the underline."""
    top_baseline = underline_y - leading * (len(lines) - 1)
    for i, line in enumerate(lines):
        put(page, x, top_baseline + i * leading, line, size=size, max_width=max_width)


def put_wrapped(page, x, first_baseline, text, size, max_width, leading):
    """Greedy word-wrap at a fixed size; returns the number of lines used."""
    words, lines, cur = text.split(), [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if cur and text_width(trial, size) > max_width:
            lines.append(cur)
            cur = w
        else:
            cur = trial
    if cur:
        lines.append(cur)
    for i, line in enumerate(lines):
        put(page, x, first_baseline + i * leading, line, size=size)
    return len(lines)


def song_table(page, song, underlines):
    """One Song N details table. Timecode/duration rows stay blank."""
    x, w = 310.2, 246.0
    rows = ["title", "alt_title", "iswc", "isrc", "length", None,
            "album", "release_date", None, "usage"]
    for key, y in zip(rows, underlines):
        if key is None and y == underlines[5]:
            put(page, x, y, GENRE, max_width=w)
        elif key is None:
            put(page, x, y, MASTER_OWNER, max_width=w)
        else:
            put(page, x, y, song[key], max_width=w)


def splits_row(page, row_top):
    """First (only) writer row of a Writers-and-splits table."""
    mid = row_top + 11.7                     # single-line baseline
    st1, st2 = row_top + 7.6, row_top + 13.6  # stacked baselines
    colw = 63.0
    put(page, 55.6, st1, WRITER_BMI_NAME_LINES[0], size=5.8, max_width=colw)
    put(page, 55.6, st2, WRITER_BMI_NAME_LINES[1], size=5.8, max_width=colw)
    put(page, 128.4, mid, "Composer", max_width=colw)
    put(page, 201.3, mid, WRITER_IPI, max_width=colw)
    put(page, 273.9, mid, PRO, max_width=colw)
    put(page, 346.5, mid, "100%", max_width=colw)
    put(page, 419.1, st1, PUBLISHER_LINES[0], size=5.8, max_width=colw)
    put(page, 419.1, st2, PUBLISHER_LINES[1], size=5.8, max_width=colw)
    # Publisher share % deliberately left blank (artist's choice, 2026-11-02,
    # while BMI's works-maintenance update is still pending).


def credit(page, x, first_baseline, title, size, max_width, leading):
    text = (u"\u201c%s\u201d written by %s, performed by %s, courtesy of %s"
            % (title, LEGAL_NAME, STAGE_NAME, MASTER_OWNER))
    return put_wrapped(page, x, first_baseline, text, size, max_width, leading)


def main():
    doc = pymupdf.open(SRC)
    if len(doc) != 3:
        raise SystemExit("Expected the 3-page Beatformers form, got %d pages" % len(doc))
    p1, p2, p3 = doc[0], doc[1], doc[2]

    # ---- page 1: Production Details (blank rows only) --------------------
    put(p1, 366.6, 284.9, "Ricky Taitt \u2014 taitt.film@gmail.com",
        max_width=189.6)
    # Film runtime / Release year: production's to complete -- left blank.

    # ---- page 1: Artist / Contributor Information ------------------------
    x, w = 462.3, 93.9
    put(p1, x, 413.9, LEGAL_NAME, max_width=w)
    put(p1, x, 430.6, STAGE_NAME, max_width=w)
    put_stacked(p1, x, 447.3, ADDRESS_LINES, max_width=w)
    put(p1, x, 464.0, PHONE, max_width=w)
    put(p1, x, 480.7, EMAIL, max_width=w)
    put(p1, x, 497.4, WRITER_IPI, max_width=w)
    put(p1, x, 514.1, PRO, max_width=w)
    put_stacked(p1, x, 530.8, PUBLISHER_LINES, max_width=w)
    put(p1, x, 547.5, PUBLISHER_IPI, max_width=w)
    put(p1, x, 564.2, PRO, max_width=w)
    put_stacked(p1, x, 580.9, MASTER_OWNER_LINES, max_width=w)

    # ---- page 2: Song 1 + splits + credit --------------------------------
    song_table(p2, SONG_1, [61.4, 78.1, 94.8, 111.5, 128.2,
                            144.9, 161.6, 178.3, 195.0, 211.7])
    splits_row(p2, 323.5)
    n = credit(p2, 310.2, 421.5, SONG_1["title"], size=6.0,
               max_width=246.0, leading=6.8)
    if n != 2:
        raise SystemExit("Song 1 credit wrapped to %d lines (cell holds 2)" % n)

    # ---- page 2: Song 2 details ------------------------------------------
    song_table(p2, SONG_2, [517.0, 533.7, 550.4, 567.1, 583.8,
                            600.5, 617.2, 633.9, 650.6, 667.3])

    # ---- page 3: Song 2 splits + credit, artist print name ---------------
    splits_row(p3, 71.2)
    n = credit(p3, 191.0, 169.2, SONG_2["title"], size=6.4,
               max_width=365.0, leading=6.8)
    if n != 2:
        raise SystemExit("Song 2 credit wrapped to %d lines (cell holds 2)" % n)
    put(p3, 105.1, 593.0, LEGAL_NAME, size=8.6, max_width=138.3)
    # Signature + Date (artist) and the whole Producer block: left blank.

    meta = dict(doc.metadata or {})
    meta["title"] = ("Beatformers \u2013 Music Cue Sheet \u2014 Zazie Productions "
                     "(filled, unsigned)")
    meta["author"] = "Zazie Productions"
    doc.set_metadata(meta)
    doc.save(DST, garbage=3, deflate=True)
    print("wrote", DST)


if __name__ == "__main__":
    main()
