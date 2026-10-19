# The Dark Awaits — Completed Music Cue Sheet (Review Notes)

**Deliverable:** `Cue_Sheet_The_Dark_Awaits_COMPLETED.xlsx`
**Source template:** `Cue_Sheet_Template_2016V3-6-5.xlsx` (left untouched)
**Prepared:** 2026-10-19 · **Revision:** 2 · Status: **ready for review, not yet submitted**

Template SHA-256: `2624e0184b806a22fe21998351116acfae6c281fc9a492317a61f4de0574d02d`
Completed SHA-256: `ec28fddedf8a884ce8f4d295819b1e54b4f2b1d799dc22138246366b300f549c`

---

## What was wrong with the previous completed file, and what changed

The previous file was parseable, but it did not behave like a finished Excel
workbook in important ways:

1. **Its formula-result caches were still the blank template's.** Excel
   workbooks store the last calculated result alongside each formula.
   The previous file changed the inputs but left those stored results alone,
   so any viewer/importer that displayed the file before recalculating it
   still showed **“Music Cue Sheet Template”** as the title and showed **no
   Seq. numbers, no cue durations and no Total Music Duration**. This build
   refreshes the cached result of every formula cell affected by the inputs,
   while leaving every formula itself byte-for-byte identical. It also keeps
   `fullCalcOnLoad="1"`, so desktop Excel recalculates the untouched formulas
   when the file opens.
2. **Its recalculated total was impossible.** Cue title/usage/timings were
   repeated on both the Composer and Publisher rows, so the template's
   `SUM(J:J)`/`SUM(K:K)` Total Music Duration double-counted all six cues:
   **50 min. 16 sec. of music in a 45-minute program**. This standardized
   BMI/RapidCue template uses grouped cue blocks — the official sample cue
   sheet for this exact template version, “Weekly World of Music, Rock and
   Roll,” puts a cue's title, usage and timings on its first party row and
   additional writers/publishers on the rows below. This build follows that
   layout, so Seq. # correctly numbers cues **1–6** and Total Music Duration
   is the actual **25 min. 8 sec.** BMI's current page links both this
   template version and the sample: <https://www.bmi.com/creators/what_is_a_cue_sheet>.
3. **Its text was stored as inline strings.** This build stores text the way
   Excel itself does, in `xl/sharedStrings.xml`, with exact count accounting.
   That removes a class of importer/table compatibility problems with no loss
   of values.

No structural redesign was made. Only 3 of the workbook's 24 package parts
differ from the blank template:

- `xl/worksheets/sheet1.xml` — previously empty input cells filled, plus the
  formula-result caches described above;
- `xl/sharedStrings.xml` — 12 new unique strings appended after the original
  290 (all original strings/order preserved; `count=378`,
  `uniqueCount=302`, matching actual cell usage);
- `xl/workbook.xml` — only `fullCalcOnLoad="1"` added to `calcPr`.

The other 21 parts — including styles, the hidden `hidat` worksheet, all 8
table definitions, printer settings, theme, relationships, content types and
`xl/calcChain.xml` — are byte-for-byte identical to the template.

---

## Program information

| Cell | Field | Value |
|---|---|---|
| N4 | Program (series, film, etc.) Title | The Dark Awaits |
| N11 | Production Company | Dark Echo Productions |
| D13 / F13 | Program/Show Duration | 45 min. 0 sec. |
| A1 / A2 | Auto titles (template formulas) | “The Dark Awaits” / “Music Cue Sheet” |
| D14–H14 | Total Music Duration (template formulas) | 25 min. 8 sec. (auto calc.) |

All fields not supplied, or expressly instructed to be left blank, remain
blank. See **Missing Information** at the end.

## Party information

Each cue occupies a two-row block in rows 20–31. The first row is the writer
party; the following row is the publisher party.

| Column(s) | Field | Value |
|---|---|---|
| L (first row) | Role | Composer |
| M / N (first row) | Composer First / Last Name | ZAZIE / PRODUCTIONS — i.e. BMI registered name **PRODUCTIONS, ZAZIE** |
| L (second row) | Role | Publisher |
| O (second row) | Publisher Name | Zazie Productions Publishing |
| P (second row, hidden column) | Publisher IPI # | 01372389230 |
| Q (both rows) | PRO / Affiliation | BMI |
| R (both rows) | Shares | 100.00% (stored as `1`; the column format is `0.00%`) |

The template's conditional formatting strikes Publisher Name on Composer
rows and writer names on Publisher rows; those inapplicable cells were left
blank rather than filled with struck-through text. Shares follow this
repository's recorded default ownership (single writer 100%, single
publisher 100%, absent a contrary contract) in
`Zazie_Productions_Publishing.md`.

### IPI / account-number handling

Column P is the template's only IPI column and is headed under **Publisher**,
so it contains the expressly supplied **Publisher IPI 01372389230**. The BMI
publisher **account number 4380631 was not entered as an IPI** and has no
separate field in this template. Publisher IPI is stored as text on purpose:
column P's number format is `0`, and a numeric entry would display
`1372389230`, losing the leading zero. Excel's “number stored as text”
marker on those cells is therefore expected. Column P remains hidden, as in
the template.

There is no writer-IPI field anywhere in this template, so writer IPI
**01280391067** is recorded here rather than forced into the publisher
column.

## Cues

Chronological, all Usage Code **BI** (Background Instrumental). Duration and
Seq. # below are the outputs/caches of the template's own formulas; no value
was typed into a formula cell's formula.

| Seq. # | Sheet rows | Cue Title | Usage | Time In | Time Out | Duration |
|---:|---:|---|---|---|---|---:|
| 1 | 20–21 | Impact Wave Tick Interlude | BI | 00:02:32 | 00:02:35 | 0:03 |
| 2 | 22–23 | John's Story — Distorted Atmospheric Pad | BI | 00:02:42 | 00:02:48 | 0:06 |
| 3 | 24–25 | Spirits Still Around | BI | 00:04:06 | 00:04:25 | 0:19 |
| 4 | 26–27 | Industrial Electronic Interlude | BI | 00:06:12 | 00:06:49 | 0:37 |
| 5 | 28–29 | Mystic Bell Ambience | BI | 00:07:32 | 00:08:00 | 0:28 |
| 6 | 30–31 | Family Investigation — Dark Ambience | BI | 00:14:25 | 00:38:00 | 23:35 |
|  |  | **Total Music Duration (template formula)** |  |  |  | **25:08** |

In each block, title/usage/timings are on the first row only; the Publisher
row contains only Role, Publisher Name, Publisher IPI, Affiliation and
Shares. That blank appearance in columns B:K of rows 21, 23, 25, 27, 29 and
31 is intentional and matches the sample layout for this template.

---

## Quality control actually run

Regenerate and verify with:

```bash
python3 -m venv /tmp/venv-cue
/tmp/venv-cue/bin/pip install -r tools/requirements.txt
/tmp/venv-cue/bin/python tools/fill_cue_sheet.py
/tmp/venv-cue/bin/python tools/qc_cue_sheet.py
/tmp/venv-cue/bin/python tools/eval_formulas.py
```

Results on the delivered file:

- `tools/qc_cue_sheet.py`: **105 checks passed, 0 failed**, exit code 0.
  It verifies ZIP integrity and identical 24-part list/order; only the 3
  expected parts changed; every XML/rels part is well-formed; styles,
  `hidat`, tables, printer settings, theme and calcChain are byte-identical;
  hidden `hidat` remains hidden; 9 validation rules/dropdowns, conditional
  formatting, merged cells, dimensions, column widths/hidden states and
  sheet/workbook protection are preserved; all **2,947 raw `<f>` formulas
  are identical** to the template and the calcChain references exactly that
  same set; every one of sheet1's 17,973 cells remains in the same XML order
  with the same style index; shared-string counts equal actual usage and no
  `inlineStr` cells remain; all header/cue values, cached formula results,
  chronology, BI usage, exact writer/publisher/PRO/IPI/share values and
  blanks match this document.
- `tools/eval_formulas.py`: independently evaluates the template's own
  Seq. #, Duration and Total Music Duration formulas (resolving only the
  structured references/`INDIRECT` forms that the `formulas` engine cannot
  parse) and obtains Seq. # **1–6**, durations **0:03 / 0:06 / 0:19 / 0:37 /
  0:28 / 23:35**, Total Music Duration **25 min. 8 sec.**, A1 = “The Dark
  Awaits”, A2 = “Music Cue Sheet” — exactly the caches stored in the file.
- Cross-parser smoke tests (run in the sandbox, not committed as dependencies):
  **SheetJS `xlsx` 0.18.5** and **ExcelJS 4.4.0** both opened the delivered
  file, saw the hidden sheet, read the unicode cue titles and text IPI, and
  read the cached title/total/sequence/duration results and underlying
  formulas correctly.
- Data-only reading (openpyxl `data_only=True`) now returns the finished
  sheet — A1 “The Dark Awaits”, A2 “Music Cue Sheet”, sequence 1–6, all cue
  durations and 25/8 total — instead of the blank template's cached view.

Not verified here: desktop Microsoft Excel and LibreOffice are not available
in this sandbox, and Debian package repositories are unreachable from it, so
the file was not opened in those applications. It was therefore built to be
correct both **with** recalculation (unchanged formulas plus
`fullCalcOnLoad`) and **without** it (correct Excel-style cached results),
and checked with the independent parsers/calculation engine above. A first
open in desktop Excel before submission is still recommended; if Excel shows
a Protected View banner because the file was downloaded, choose **Enable
Editing** — the template's own sheet protection remains on, but the input
cells retain their unlocked styles.

## Reviewer attention items (not missing data)

1. **“Family Investigation — Dark Ambience” runs 23 min. 35 sec.**
   (00:14:25 → 00:38:00), about 94% of all music on the sheet. It was
   entered exactly as supplied; spot-check it against the locked cut /
   spotting notes before submission.
2. **Shares** are set to writer 100% / publisher 100% from the repository's
   stated default ownership. Confirm no cue-specific split or co-party
   applies to this production.

---

## Missing Information

Nothing below was invented. Cells are left blank in the workbook unless a
note says the template simply has no field for supplied information.

1. **Episode Title (N6)** — instructed: leave blank.
2. **Episode Number (N8)** — left blank. The instruction text is contradictory
   here (“Episode Number: 2” followed by “Leave blank”); the explicit
   “Leave blank” was followed and **2 was not entered**. Confirm whether the
   episode number should in fact be 2.
3. **Network / Source (D11)** — instructed: leave blank.
4. **Initial Airdate (D8)** — instructed: leave blank.
5. **Cue Sheet Classification (D4)** — unknown; dropdown offers Original /
   Revision.
6. **Category (D9)** — unknown; dropdown includes Series, Movie, Broadcast
   Series, Reality, etc.
7. **Date Prepared (D5)** — not supplied.
8. **Version (D10)** — not supplied; dropdown includes Cable, Network,
   Edited for TV, US, Non US, etc.
9. **Production Number (N9)** — not supplied.
10. **Program Title AKA(s) (N5)** — not supplied.
11. **Episode Title AKA(s) (N7)** — not supplied.
12. **Production Company Mailing Address (N12)** — not supplied.
13. **Cue Sheet Prepared By (N13)** — not supplied.
14. **Email Address (N14)** — not supplied.
15. **Writer IPI / CAE 01280391067** — supplied and verified for the notes,
    but this template has no writer-IPI field; not entered in the worksheet.
16. **Public credit “Zazie Productions”** — supplied, but the template has no
    separate public-credit/AKA writer field. The BMI registered name
    PRODUCTIONS, ZAZIE was entered in the First/Last Name fields.
17. **BMI publisher account number 4380631** — supplied, but this template
    has no publisher-account-number field. It was deliberately not used in
    the Publisher IPI column; that column holds Publisher IPI 01372389230.
18. **Per-cue scene numbers/descriptions, ISWC/catalog identifiers and cue
    registration details** — not supplied; this template version has no
    columns for them on the cue grid.
