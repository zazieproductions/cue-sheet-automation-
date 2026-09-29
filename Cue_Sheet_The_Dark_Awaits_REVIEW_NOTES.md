# The Dark Awaits — Completed Music Cue Sheet (Review Notes)

**Deliverable:** `Cue_Sheet_The_Dark_Awaits_COMPLETED.xlsx`
**Source template:** `Cue_Sheet_Template_2016V3-6-5.xlsx` (left untouched)
**Prepared:** 2026-09-29 · Status: **ready for review, not yet submitted**

The workbook was filled by writing values only into cells that were empty in
the template. No formula, table, dropdown, hidden sheet, style or protection
setting was altered — 22 of the 24 package parts are byte-for-byte identical to
the template, and the only two that differ are `xl/worksheets/sheet1.xml`
(filled cells) and `xl/workbook.xml` (one `fullCalcOnLoad="1"` flag so Excel
recalculates the untouched formulas when the file opens).

---

## Program information

| Cell | Field | Value |
|---|---|---|
| N4 | Program (series, film, etc.) Title | The Dark Awaits |
| N11 | Production Company | Dark Echo Productions |
| D13 / F13 | Program/Show Duration | 45 min. 0 sec. |
| A1 / A2 | Auto titles (formulas) | "The Dark Awaits" / "Music Cue Sheet" |

## Party information (per cue row)

| Column | Field | Value |
|---|---|---|
| L | Role | `Composer` on odd cue rows, `Publisher` on even cue rows |
| M / N | Composer First / Last Name | ZAZIE / PRODUCTIONS (BMI registered name *PRODUCTIONS, ZAZIE*) |
| O | Publisher Name | Zazie Productions Publishing |
| P | Publisher IPI # | 01372389230 (Publisher rows only) |
| Q | PRO / Affiliation | BMI |
| R | Shares | 1 (column is formatted `0.00%`, so it displays **100.00%**) |

## Cues (rows 20–31, chronological)

| Seq | Row | Cue Title | Usage | Time In | Time Out | Duration | Role |
|---|---|---|---|---|---|---|---|
| 1 | 20 | Impact Wave Tick Interlude | BI | 00:02:32 | 00:02:35 | 0:03 | Composer |
| 2 | 21 | Impact Wave Tick Interlude | BI | 00:02:32 | 00:02:35 | 0:03 | Publisher |
| 3 | 22 | John's Story — Distorted Atmospheric Pad | BI | 00:02:42 | 00:02:48 | 0:06 | Composer |
| 4 | 23 | John's Story — Distorted Atmospheric Pad | BI | 00:02:42 | 00:02:48 | 0:06 | Publisher |
| 5 | 24 | Spirits Still Around | BI | 00:04:06 | 00:04:25 | 0:19 | Composer |
| 6 | 25 | Spirits Still Around | BI | 00:04:06 | 00:04:25 | 0:19 | Publisher |
| 7 | 26 | Industrial Electronic Interlude | BI | 00:06:12 | 00:06:49 | 0:37 | Composer |
| 8 | 27 | Industrial Electronic Interlude | BI | 00:06:12 | 00:06:49 | 0:37 | Publisher |
| 9 | 28 | Mystic Bell Ambience | BI | 00:07:32 | 00:08:00 | 0:28 | Composer |
| 10 | 29 | Mystic Bell Ambience | BI | 00:07:32 | 00:08:00 | 0:28 | Publisher |
| 11 | 30 | Family Investigation — Dark Ambience | BI | 00:14:25 | 00:38:00 | 23:35 | Composer |
| 12 | 31 | Family Investigation — Dark Ambience | BI | 00:14:25 | 00:38:00 | 23:35 | Publisher |

Duration (columns J/K) and Seq. # (column A) are the template's own formulas;
nothing was typed into them.

---

## Two things a reviewer should look at before submission

1. **"Total Music Duration" (D14/F14) reads 50 min. 16 sec., not 25 min. 8 sec.**
   That formula is `SUM(J:J)` + `SUM(K:K)` over the whole column, so it counts
   every listed row. Because each cue is deliberately listed twice (one Composer
   row, one Publisher row), the auto total is exactly double the music:
   6 cues × 1,508 s of unique music = 25:08, doubled = 50:16.
   If the sheet is being sent somewhere that reads that cell as program music
   time, clear D:I on the six **Publisher** rows (21, 23, 25, 27, 29, 31) — the
   duration formulas there then return blank and the total reads 25:08. Nothing
   was changed on your instruction to duplicate timing on both rows.

2. **"Family Investigation — Dark Ambience" runs 23 min 35 s** (00:14:25 →
   00:38:00) — 94% of all the music on this sheet. Entered exactly as supplied;
   worth a spot-check against the spotting notes in case a time-out was meant
   to be 00:18:00 or similar.

## Template mechanics worth knowing

* The sheet strikes through (hatches) column **O** on any row whose Role is
  Composer/Arranger, and columns **M:N** on any row whose Role is Publisher —
  i.e. it is designed for one party per row, which is why each cue has two rows.
  Those cells were left blank accordingly, so no struck-out text appears.
* Column **P** is the only IPI field in this template and its header sits under
  **Publisher**, so it asks for a *Publisher* IPI #. It now holds
  **01372389230** on the six Publisher rows, and is blank on the Composer rows.
  It is stored as text rather than as a number on purpose: the column's number
  format is `0`, so a numeric entry would display `1372389230` and silently drop
  the leading zero. Excel may therefore show its "number stored as text" marker
  on those cells — that is expected and is the safer option for an 11-digit IPI.
  Column P is hidden in the template and stays hidden; unhide it to see the IPIs.
* The **BMI publisher account number 4380631** has no field in this template
  (the only publisher identifier column is the IPI one), so it is recorded here
  and in `Zazie_Productions_Publishing.md` only.
* Column **R** is number-formatted `0.00%`. 100% is therefore stored as `1` and
  displays as `100.00%`; typing `100` would display `10000.00%`.
* There is **no writer-IPI field** anywhere in this template (columns are
  First Name / Last Name only), so writer IPI 01280391067 has nowhere to go.

## Quality control actually run

`python3 tools/qc_cue_sheet.py` — 57 checks, all passed, exit code 0: identical package part
list; only sheet1.xml + workbook.xml changed; all 24 XML parts well-formed;
styles.xml / sharedStrings.xml / hidat sheet / table definitions / printer
settings byte-identical; hidden `hidat` sheet still hidden; 9 data-validation
rules (the Usage, Affiliation, Role, Category, Classification and Version
dropdowns) intact; conditional formatting, merged cells, column widths and
sheet/workbook protection intact; **2,947 formulas in the completed file are
character-for-character the same as the 2,947 in the template**; cues in
chronological order; every cue uses BI; no stray data below row 31.

`python3 tools/eval_formulas.py` — the template's own Seq. #, Duration and
Total Music Duration formulas, evaluated by the `formulas` engine against the
values in the completed file (structured references and `INDIRECT` resolved to
their literal cells, since that engine cannot parse them): Seq. # 1–12,
durations 0:03 / 0:06 / 0:19 / 0:37 / 0:28 / 23:35, total 50 min 16 sec,
A1 = "The Dark Awaits", A2 = "Music Cue Sheet".

Not verified here: the file was **not** opened in Excel or LibreOffice — no
spreadsheet application could be installed in this sandbox (Debian package
repos are unreachable). Structural validity and formula behaviour were checked
programmatically as described above.

---

## Missing Information

Fields left blank because the information was not supplied (or was explicitly
requested to be left blank). Nothing was invented.

**Resolved in this revision:** Publisher IPI / CAE **01372389230** was supplied
and entered in column P on the six Publisher rows. (Earlier drafts left P blank
because only the BMI account number had been given.)

**Requested blank by instruction**
1. Episode Title (N6)
2. Episode Number (N8)
3. Network / Source (D11)
4. Initial Airdate (D8)
5. Cue Sheet Classification (D4) — dropdown offers `Original` / `Revision`; not confirmed
6. Category (D9) — dropdown (Series, Movie, Broadcast Series, Reality, …); not confirmed

**Not supplied at all**
7. Date Prepared (D5)
8. Version / edit version (D10) — dropdown (Original, Edited for TV, US, Non US, …)
9. Production Number (N9)
10. Program Title AKA(s) (N5) and Episode Title AKA(s) (N7)
11. Production Company mailing address (N12)
12. Cue Sheet Prepared By (N13) and Email Address (N14)
13. **Writer IPI / CAE 01280391067** — recorded here; this template has no writer-IPI field to hold it
14. BMI publisher account number **4380631** — no field for it in this template
15. Public credit "Zazie Productions" — the template only has First/Last name fields; the BMI registered name `PRODUCTIONS, ZAZIE` was used instead
16. Per-cue scene numbers / descriptions and any co-writer or co-publisher splits (none reported; single writer, single publisher at 100%/100% assumed per instruction)
