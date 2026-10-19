#!/usr/bin/env python3
"""Read-only browser preview of the completed cue sheet.

Nothing here writes to the workbook. The page is rendered from
`Cue_Sheet_The_Dark_Awaits_COMPLETED.xlsx` itself (openpyxl) and the values
shown for formula cells come from tools/eval_formulas.evaluate() -- the same
code path that produced the QC numbers, so the preview cannot drift from the
file.  Served on 0.0.0.0 so the Arena live preview can proxy it.
"""
import html
import os
import subprocess
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import eval_formulas  # noqa: E402

PORT = int(os.environ.get("PORT", "8765"))
XLSX = "Cue_Sheet_The_Dark_Awaits_COMPLETED.xlsx"
NOTES = "Cue_Sheet_The_Dark_Awaits_REVIEW_NOTES.md"
TEMPLATE = "Cue_Sheet_Template_2016V3-6-5.xlsx"

LEFT_ROWS = [4, 5, 8, 9, 10, 11, 13, 14]   # label in C, value in D
RIGHT_ROWS = [4, 5, 6, 7, 8, 9, 11, 12, 13, 14]  # label in M, value in N
CUE_FIRST, CUE_LAST = 20, 31
COLS = list("ABCDEFGHIJKLMNOPQR")

print("loading workbook ...", flush=True)
WB = openpyxl.load_workbook(XLSX)
WS = WB["Template"]
TPL = openpyxl.load_workbook(TEMPLATE)["Template"]
print("evaluating the template's formulas ...", flush=True)
CALC = eval_formulas.evaluate()
QC = subprocess.run([sys.executable, "tools/qc_cue_sheet.py"],
                    capture_output=True, text=True).stdout
QC_PASS = QC.count("\nPASS") + QC.startswith("PASS")
QC_FAIL = QC.count("\nFAIL") + QC.startswith("FAIL")
print("ready: %d calculated cells, QC %d pass / %d fail" % (len(CALC), QC_PASS, QC_FAIL),
      flush=True)


# ------------------------------------------------------------------ display
def is_formula(coord):
    v = TPL[coord].value
    return isinstance(v, str) and v.startswith("=")


def show(coord):
    """What Excel would display in this cell."""
    if is_formula(coord):
        v = CALC.get(coord)
        if v is None or v == "":
            return "", True
        v = int(v) if isinstance(v, float) and v.is_integer() else v
        return fmt(coord, v), True
    v = WS[coord].value
    if v is None or v == "":
        return "", False
    return fmt(coord, v), False


def fmt(coord, v):
    nf = WS[coord].number_format
    if isinstance(v, (int, float)):
        if nf == "0.00%":
            return "%.2f%%" % (float(v) * 100)
        if nf == "00":
            return "%02d" % int(v)
        if nf in ("0", "General"):
            return str(int(v)) if float(v).is_integer() else str(v)
    return str(v)


def cell_html(coord, css):
    text, calculated = show(coord)
    filled = (not calculated) and WS[coord].value not in (None, "")
    classes = [css]
    if calculated:
        classes.append("calc")
    elif filled:
        classes.append("filled")
    else:
        classes.append("empty")
    tip = coord
    if calculated:
        tip += " \u2014 template formula, auto-calculated"
    return '<td class="%s" title="%s">%s</td>' % (
        " ".join(classes), html.escape(tip), html.escape(text))


def width(col):
    d = WS.column_dimensions.get(col)
    w = d.width if (d and d.width) else 9.0
    return max(28, min(230, int(w * 7.2)))


def program_block():
    out = ['<div class="prog">']
    for label_col, value_col, rows in (("C", "D", LEFT_ROWS), ("M", "N", RIGHT_ROWS)):
        out.append("<table>")
        for r in rows:
            label = TPL["%s%d" % (label_col, r)].value or ""
            text, calc = show("%s%d" % (value_col, r))
            if r == 13 and label_col == "C":
                # "Program/Show Duration" spans D13 (min) + F13 (sec)
                mins, _ = show("D13")
                secs, _ = show("F13")
                text = "%s min. %s sec." % (mins, secs)
            if r == 14 and label_col == "C":
                m, _ = show("D14")
                s, _ = show("F14")
                text, calc = "%s min. %s sec." % (m, s), True
            state = "calc" if calc else ("filled" if text else "empty")
            out.append('<tr><td class="lbl">%s</td><td class="%s">%s</td></tr>'
                       % (html.escape(str(label)), state, html.escape(text)))
        out.append("</table>")
    out.append("</div>")
    return "\n".join(out)


def cue_grid():
    cols = "".join('<col style="width:%dpx">' % width(c) for c in COLS)
    head = ['<tr><th class="seq">Seq. #</th><th>Cue Title<br><span>(Song/Track Name)</span></th>'
            "<th>Usage</th>"
            '<th colspan="3">Time In</th><th colspan="3">Time Out</th>'
            '<th colspan="2">Duration</th>'
            "<th>Role</th><th>Composer/Writer<br><span>First (and Middle) Name</span></th>"
            "<th>Composer/Writer<br><span>Last Name</span></th>"
            "<th>Publisher<br><span>Name</span></th>"
            '<th class="hidden-col">Publisher<br><span>IPI # (hidden col.)</span></th>'
            "<th>PRO<br><span>Affiliation</span></th><th>%<br><span>Shares</span></th></tr>"]
    sub = ['<tr class="sub"><th></th><th></th><th></th><th>h</th><th>mm</th><th>ss</th>'
           "<th>h</th><th>mm</th><th>ss</th><th>min.</th><th>sec.</th>"
           "<th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"]
    body = []
    for r in range(CUE_FIRST, CUE_LAST + 1):
        cells = [cell_html("%s%d" % (c, r), "seq" if c == "A" else
                           ("hidden-col" if c == "P" else ""))
                 for c in COLS]
        role = WS["L%d" % r].value or ""
        body.append('<tr class="%s">%s</tr>' % (role.lower(), "".join(cells)))
    return ("<table class='grid'><colgroup>%s</colgroup>%s</table>"
            % (cols, "".join(head + sub + body)))


CSS = """
 body{font:13px/1.45 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:0;
      background:#eceff3;color:#1c2430}
 .wrap{max-width:1400px;margin:0 auto;padding:22px}
 h1{font-size:20px;margin:0 0 2px} h2{font-size:15px;margin:26px 0 8px}
 .sub2{color:#5b6675;margin:0 0 14px}
 .banner{background:#fff;border:1px solid #d4dae2;border-left:4px solid #2f6fd0;
         padding:10px 14px;border-radius:6px;margin:14px 0}
 .banner b{color:#1c2430}
 a.btn{display:inline-block;background:#2f6fd0;color:#fff;text-decoration:none;
       padding:7px 13px;border-radius:5px;font-weight:600;margin-right:8px}
 a.plain{color:#2f6fd0}
 .prog{display:flex;gap:26px;flex-wrap:wrap}
 .prog table{border-collapse:collapse;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,.07);
             min-width:430px}
 .prog td{padding:4px 9px;border-bottom:1px solid #eceff3;vertical-align:top}
 .prog td.lbl{color:#5b6675;white-space:nowrap;width:1%;padding-right:14px}
 table.grid{border-collapse:collapse;background:#fff;width:100%;
            box-shadow:0 1px 2px rgba(0,0,0,.07);table-layout:fixed}
 table.grid th,table.grid td{border:1px solid #dbe1e8;padding:4px 5px;overflow:hidden;
            text-overflow:ellipsis;white-space:nowrap;font-size:12px}
 table.grid th{background:#f3f5f8;text-align:left;font-weight:600;color:#3c4757}
 table.grid th span{font-weight:400;color:#7b8695;font-size:10.5px}
 tr.sub th{background:#fafbfc;font-size:10.5px;color:#7b8695;padding:2px 5px}
 td.filled{background:#e9f6ec}
 td.calc{background:#eaf1fb;color:#28518f}
 td.empty{background:#fbfbfc;color:#c3cad3}
 td.seq{text-align:right;color:#7b8695}
 th.seq{text-align:right}
 td.hidden-col,th.hidden-col{background:#fff8e6;color:#8a6d1a}
 tr.publisher td{border-bottom:2px solid #cfd6de}
 .legend span{display:inline-block;margin-right:16px;font-size:12px;color:#4b5665}
 .sw{display:inline-block;width:12px;height:12px;border:1px solid #c9d1da;
     vertical-align:-2px;margin-right:5px;border-radius:2px}
 .qc{background:#fff;border:1px solid #d4dae2;border-radius:6px;padding:12px 14px}
 .qc code{font-size:12px}
 .ok{color:#137a3d;font-weight:700} .bad{color:#b3261e;font-weight:700}
 .notes{background:#fff;border:1px solid #d4dae2;border-radius:6px;padding:4px 14px;
        max-height:520px;overflow:auto}
 .notes pre{white-space:pre-wrap;font:12px/1.5 ui-monospace,Menlo,Consolas,monospace}
"""


def page():
    a1, _ = show("A1")
    a2, _ = show("A2")
    return """<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>The Dark Awaits \u2014 Music Cue Sheet (preview)</title>
<style>%s</style></head><body><div class="wrap">
<h1>%s</h1><p class="sub2">%s &middot; Dark Echo Productions &middot; 45 min. program
&middot; 6 cues, all usage code BI</p>

<div class="banner"><b>Read-only preview of the completed workbook.</b>
The file you are looking at is rendered from
<code>Cue_Sheet_The_Dark_Awaits_COMPLETED.xlsx</code>; the spreadsheet itself was not
modified to make this page. Cells marked <span class="sw" style="background:#eaf1fb"></span>
are the template's own formulas \u2014 the numbers shown are what Excel calculates from
them (Seq. #, Duration, Total Music Duration). Cells marked
<span class="sw" style="background:#e9f6ec"></span> are the metadata that was entered.
<br><br><a class="btn" href="/download">Download the .xlsx</a>
<a class="plain" href="/notes">Review notes + Missing Information</a>
<a class="plain" href="/qc">QC log (%d checks)</a></div>

<h2>Program information</h2>%s

<h2>Cues (rows %d\u2013%d of the Template sheet)</h2>
<p class="legend"><span><span class="sw" style="background:#e9f6ec"></span>entered value</span>
<span><span class="sw" style="background:#eaf1fb"></span>template formula (auto-calculated)</span>
<span><span class="sw" style="background:#fff8e6"></span>column P is hidden in the workbook</span></p>
%s

<div class="banner" style="border-left-color:#c98a00"><b>Note:</b> Total Music Duration
reads <b>50 min. 16 sec.</b> because the template sums every listed row and each cue is
listed twice (one Composer row, one Publisher row). The unique music is
<b>25 min. 8 sec.</b> Clearing D:I on the six Publisher rows makes the auto-total read
25:08 \u2014 left duplicated as instructed.</div>

<h2>Verification</h2><div class="qc"><span class="ok">%d checks passed</span>,
<span class="%s">%d failed</span> \u2014 <code>tools/qc_cue_sheet.py</code>, run when this
preview started. Formulas, tables, dropdowns, the hidden <code>hidat</code> sheet, styles
and sheet/workbook protection are unchanged from the blank template.</div>
</div></body></html>""" % (
        CSS, html.escape(a1 or "Music Cue Sheet"), html.escape(a2 or ""),
        QC_PASS, program_block(), CUE_FIRST, CUE_LAST, cue_grid(),
        QC_PASS, "bad" if QC_FAIL else "ok", QC_FAIL)


class Handler(BaseHTTPRequestHandler):
    server_version = "CueSheetPreview/1.0"

    def _send(self, code, body, ctype, extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            self._send(200, page(), "text/html; charset=utf-8")
        elif path == "/download":
            with open(XLSX, "rb") as fh:
                self._send(200, fh.read(),
                           "application/vnd.openxmlformats-officedocument."
                           "spreadsheetml.sheet",
                           {"Content-Disposition":
                            'attachment; filename="%s"' % XLSX})
        elif path == "/notes":
            with open(NOTES, encoding="utf-8") as fh:
                self._send(200, "<!doctype html><meta charset='utf-8'><style>"
                           "body{font:12.5px/1.55 ui-monospace,Menlo,Consolas,monospace;"
                           "margin:24px;max-width:900px}a{color:#2f6fd0}</style>"
                           "<p><a href='/'>\u2190 back to the cue sheet</a></p><pre>%s</pre>"
                           % html.escape(fh.read()), "text/html; charset=utf-8")
        elif path == "/qc":
            self._send(200, "<!doctype html><meta charset='utf-8'><style>"
                       "body{font:12.5px/1.5 ui-monospace,Menlo,Consolas,monospace;"
                       "margin:24px}</style><p><a href='/'>\u2190 back</a></p><pre>%s</pre>"
                       % html.escape(QC), "text/html; charset=utf-8")
        elif path == "/healthz":
            self._send(200, "ok", "text/plain")
        else:
            self._send(404, "not found", "text/plain")

    def do_HEAD(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()

    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))


if __name__ == "__main__":
    print("serving preview on 0.0.0.0:%d" % PORT, flush=True)
    ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
