"""
═══════════════════════════════════════════════════════════════════
  DONOR TRACKING SERVER  —  tracker_server.py
  ─────────────────────────────────────────────
  Runs locally on http://localhost:5050

  Catches THREE events from every email sent:
    1. OPEN      — donor opened the email (invisible 1×1 pixel)
    2. CLICK      — donor clicked the donate link
    3. DONATION   — donor submitted the donation form

  Every event:
    • Appends to  tracking_log.json  (raw audit trail)
    • Instantly updates  donor_tracker.xlsx  with latest status

  HOW IT WORKS (no external server needed):
    • Email links point to  http://localhost:5050/click?id=XXX
    • Email pixel points to http://localhost:5050/open?id=XXX
    • Donation form at     http://localhost:5050/donate?id=XXX
    • After donation form submit → writes to Excel immediately

  RUN:
    python tracker_server.py
    (keep this running while your campaign is active)
═══════════════════════════════════════════════════════════════════
"""

import json, os, threading
from datetime  import datetime
from pathlib   import Path
from flask     import Flask, request, redirect, send_file, Response
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils  import get_column_letter

# ── Config ────────────────────────────────────────────────────────────────
TRACKING_LOG   = "tracking_log.json"
EXCEL_FILE     = "donor_tracker.xlsx"
DONATE_URL     = os.getenv("DONATE_URL", "https://www.yourwebsite.org/donate")
PORT           = 5050
LOCK           = threading.Lock()   # thread-safe file writes

app = Flask(__name__)

# ══════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════

def _load_log() -> list[dict]:
    p = Path(TRACKING_LOG)
    if not p.exists():
        return []
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return []


def _append_event(tracking_id: str, event: str, extra: dict = None) -> None:
    """Thread-safe: append one event to the JSON log."""
    entry = {
        "id":    tracking_id,
        "event": event,
        "ts":    datetime.now().isoformat(timespec="seconds"),
    }
    if extra:
        entry.update(extra)

    with LOCK:
        log_data = _load_log()
        log_data.append(entry)
        Path(TRACKING_LOG).write_text(
            json.dumps(log_data, indent=2), encoding="utf-8"
        )
    _rebuild_excel()


def _get_donor_meta(tracking_id: str) -> dict:
    """Pull the original 'sent' record for this id (has email, name, tier)."""
    for entry in _load_log():
        if entry["id"] == tracking_id and entry["event"] == "sent":
            return entry
    return {}


# ══════════════════════════════════════════════════════════════════════════
#  EXCEL BUILDER  — called after every event
# ══════════════════════════════════════════════════════════════════════════

# Colour palette
COL_HEADER   = "1B4332"   # dark green
COL_DONATED  = "D4EDDA"   # light green
COL_CLICKED  = "FFF3CD"   # amber
COL_OPENED   = "D1ECF1"   # blue tint
COL_NOACTION = "F8D7DA"   # light red
COL_WHITE    = "FFFFFF"

STATUS_ORDER = ["donated", "clicked_only", "opened_only", "no_action"]
STATUS_LABEL = {
    "donated":      "✅ Donated",
    "clicked_only": "🟡 Clicked — No Donation Yet",
    "opened_only":  "📧 Opened — No Click",
    "no_action":    "❌ No Response",
}
STATUS_COLOR = {
    "donated":      COL_DONATED,
    "clicked_only": COL_CLICKED,
    "opened_only":  COL_OPENED,
    "no_action":    COL_NOACTION,
}


def _derive_status(events: list[str]) -> str:
    if "donated" in events:  return "donated"
    if "click"   in events:  return "clicked_only"
    if "open"    in events:  return "opened_only"
    return "no_action"


def _rebuild_excel() -> None:
    """Rebuild the full Excel workbook from the current tracking_log.json."""
    with LOCK:
        log_data = _load_log()

    # ── Aggregate per donor ───────────────────────────────────────────────
    donors: dict[str, dict] = {}   # tracking_id → aggregated info

    for entry in log_data:
        tid = entry["id"]
        if tid not in donors:
            donors[tid] = {
                "tracking_id":     tid,
                "email":           entry.get("email", ""),
                "name":            entry.get("name",  ""),
                "tier":            entry.get("tier",  ""),
                "events":          [],
                "donation_amount": "",
                "donation_cause":  "",
                "donation_note":   "",
                "first_open_ts":   "",
                "first_click_ts":  "",
                "donation_ts":     "",
                "sent_ts":         "",
            }

        ev = entry["event"]
        donors[tid]["events"].append(ev)

        if ev == "sent"    and not donors[tid]["sent_ts"]:
            donors[tid]["sent_ts"]       = entry["ts"]
            donors[tid]["email"]         = entry.get("email", donors[tid]["email"])
            donors[tid]["name"]          = entry.get("name",  donors[tid]["name"])
            donors[tid]["tier"]          = entry.get("tier",  donors[tid]["tier"])
        if ev == "open"    and not donors[tid]["first_open_ts"]:
            donors[tid]["first_open_ts"] = entry["ts"]
        if ev == "click"   and not donors[tid]["first_click_ts"]:
            donors[tid]["first_click_ts"]= entry["ts"]
        if ev == "donated":
            donors[tid]["donation_ts"]     = entry["ts"]
            donors[tid]["donation_amount"] = entry.get("amount", "")
            donors[tid]["donation_cause"]  = entry.get("cause",  "")
            donors[tid]["donation_note"]   = entry.get("note",   "")

    # ── Sort: donated first, then by status ──────────────────────────────
    rows = list(donors.values())
    rows.sort(key=lambda r: STATUS_ORDER.index(_derive_status(r["events"])))

    # ── Build workbook ────────────────────────────────────────────────────
    wb = openpyxl.Workbook()

    # ── Sheet 1: Master Tracker ───────────────────────────────────────────
    ws = wb.active
    ws.title = "Donor Tracker"

    headers = [
        "Name", "Email", "Tier", "Status",
        "Donation Amount", "Cause", "Note",
        "Email Sent", "Opened At", "Clicked At", "Donated At",
    ]
    col_widths = [22, 32, 10, 28, 18, 14, 24, 20, 20, 20, 20]

    # Header row
    header_fill = PatternFill("solid", fgColor=COL_HEADER)
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin         = Side(style="thin", color="CCCCCC")
    border       = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"

    # Data rows
    for ri, r in enumerate(rows, start=2):
        status = _derive_status(r["events"])
        fill   = PatternFill("solid", fgColor=STATUS_COLOR[status])

        values = [
            r["name"]            or r["email"],
            r["email"],
            r["tier"].upper()    if r["tier"] else "",
            STATUS_LABEL[status],
            f"${r['donation_amount']}" if r["donation_amount"] else "",
            r["donation_cause"],
            r["donation_note"],
            r["sent_ts"],
            r["first_open_ts"],
            r["first_click_ts"],
            r["donation_ts"],
        ]

        for ci, val in enumerate(values, start=1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill      = fill
            cell.border    = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if ci == 1:
                cell.font = Font(bold=True)

        ws.row_dimensions[ri].height = 20

    # ── Sheet 2: Summary Dashboard ────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 16
    ws2.column_dimensions["C"].width = 16

    total     = len(rows)
    donated   = sum(1 for r in rows if _derive_status(r["events"]) == "donated")
    clicked   = sum(1 for r in rows if _derive_status(r["events"]) == "clicked_only")
    opened    = sum(1 for r in rows if _derive_status(r["events"]) == "opened_only")
    no_action = sum(1 for r in rows if _derive_status(r["events"]) == "no_action")

    total_raised = 0.0
    for r in rows:
        try:
            total_raised += float(r["donation_amount"]) if r["donation_amount"] else 0
        except ValueError:
            pass

    pct = lambda n: f"{n/total*100:.1f}%" if total else "0%"

    summary_data = [
        ("CAMPAIGN SUMMARY", "", ""),
        ("Last updated", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ""),
        ("", "", ""),
        ("Metric", "Count", "Rate"),
        ("Total emails sent",      total,     "100%"),
        ("✅ Donated",              donated,   pct(donated)),
        ("🟡 Clicked — no donation",clicked,   pct(clicked)),
        ("📧 Opened — no click",    opened,    pct(opened)),
        ("❌ No response",          no_action, pct(no_action)),
        ("", "", ""),
        ("💰 Total donations raised", f"${total_raised:,.2f}", ""),
    ]

    h_fill  = PatternFill("solid", fgColor=COL_HEADER)
    h_font  = Font(color="FFFFFF", bold=True, size=12)
    b_font  = Font(bold=True)

    for ri, (a, b, c) in enumerate(summary_data, start=1):
        ws2.cell(row=ri, column=1, value=a)
        ws2.cell(row=ri, column=2, value=b)
        ws2.cell(row=ri, column=3, value=c)
        if ri == 1:
            for ci in range(1, 4):
                ws2.cell(row=ri, column=ci).fill = h_fill
                ws2.cell(row=ri, column=ci).font = h_font
        elif ri == 4:
            for ci in range(1, 4):
                ws2.cell(row=ri, column=ci).font = b_font
        ws2.row_dimensions[ri].height = 22

    # ── Sheet 3: Donations Only ───────────────────────────────────────────
    ws3 = wb.create_sheet("Donations")
    ws3.column_dimensions["A"].width = 22
    ws3.column_dimensions["B"].width = 32
    ws3.column_dimensions["C"].width = 18
    ws3.column_dimensions["D"].width = 16
    ws3.column_dimensions["E"].width = 22

    don_headers = ["Name", "Email", "Amount", "Cause", "Donated At"]
    for ci, h in enumerate(don_headers, 1):
        cell = ws3.cell(row=1, column=ci, value=h)
        cell.fill = PatternFill("solid", fgColor="1B4332")
        cell.font = Font(color="FFFFFF", bold=True)

    don_rows = [r for r in rows if _derive_status(r["events"]) == "donated"]
    for ri, r in enumerate(don_rows, start=2):
        ws3.cell(row=ri, column=1, value=r["name"] or r["email"])
        ws3.cell(row=ri, column=2, value=r["email"])
        ws3.cell(row=ri, column=3, value=f"${r['donation_amount']}" if r["donation_amount"] else "")
        ws3.cell(row=ri, column=4, value=r["donation_cause"])
        ws3.cell(row=ri, column=5, value=r["donation_ts"])
        for ci in range(1, 6):
            ws3.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor=COL_DONATED)

    wb.save(EXCEL_FILE)


# ══════════════════════════════════════════════════════════════════════════
#  FLASK ROUTES
# ══════════════════════════════════════════════════════════════════════════

# 1×1 transparent GIF bytes
PIXEL = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\xff\xff\xff\x00\x00\x00"
    b"!\xf9\x04\x00\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01"
    b"\x00\x00\x02\x02D\x01\x00;"
)


@app.route("/open")
def track_open():
    """Called silently when donor opens the email (img src pixel)."""
    tid = request.args.get("id", "")
    if tid:
        _append_event(tid, "open")
    return Response(PIXEL, mimetype="image/gif")


@app.route("/click")
def track_click():
    """Called when donor clicks the donate button in the email."""
    tid      = request.args.get("id", "")
    dest_url = request.args.get("url", DONATE_URL)
    if tid:
        _append_event(tid, "click")
        # Redirect to the real donation page with the tracking id appended
        sep = "&" if "?" in dest_url else "?"
        dest_url = f"{dest_url}{sep}tid={tid}"
    return redirect(dest_url, code=302)


@app.route("/donate", methods=["GET", "POST"])
def donation_form():
    """
    Local donation form page.
    GET  → show the form (donor lands here after clicking the link)
    POST → record the donation, update Excel immediately
    """
    tid = request.args.get("tid", request.args.get("id", ""))
    meta = _get_donor_meta(tid) if tid else {}

    if request.method == "POST":
        tid    = request.form.get("tid", tid)
        amount = request.form.get("amount", "").strip()
        cause  = request.form.get("cause",  "general").strip()
        note   = request.form.get("note",   "").strip()
        name   = request.form.get("name",   meta.get("name", "")).strip()

        if tid:
            _append_event(tid, "donated", {
                "amount": amount,
                "cause":  cause,
                "note":   note,
                "name":   name,
                "email":  meta.get("email", ""),
            })

        return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<title>Thank you!</title>
<style>
  body{{font-family:Georgia,serif;background:#f0faf4;display:flex;
       align-items:center;justify-content:center;height:100vh;margin:0}}
  .box{{background:#fff;border-radius:12px;padding:48px 56px;
        text-align:center;box-shadow:0 4px 24px rgba(0,0,0,.08);max-width:480px}}
  h1{{color:#1b4332;font-size:28px;margin-bottom:12px}}
  p{{color:#444;line-height:1.7;font-size:16px}}
  .amount{{font-size:42px;font-weight:700;color:#52b788;margin:16px 0}}
</style></head><body>
<div class="box">
  <h1>Thank you{', ' + name if name else ''}! 🎉</h1>
  <div class="amount">${amount}</div>
  <p>Your donation toward <strong>{cause}</strong> has been recorded.<br>
  We'll be in touch with your impact update shortly.</p>
</div></body></html>"""

    # ── GET: show donation form ───────────────────────────────────────────
    prefill_name  = meta.get("name",  "")
    prefill_email = meta.get("email", "")
    prefill_cause = meta.get("cause", "general")

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Make a Donation</title>
<style>
  *{{box-sizing:border-box}}
  body{{font-family:Georgia,serif;background:#f9f6f1;margin:0;padding:24px}}
  .wrap{{max-width:520px;margin:40px auto}}
  .header{{background:#1b4332;color:#fff;padding:28px 32px;border-radius:10px 10px 0 0}}
  .header h1{{margin:0;font-size:22px;font-weight:400}}
  .header p{{margin:6px 0 0;font-size:13px;color:#a8d5b5}}
  .form-body{{background:#fff;padding:32px;border:1px solid #ddd;
              border-radius:0 0 10px 10px}}
  label{{display:block;font-size:13px;font-weight:bold;color:#333;
         margin-bottom:5px;margin-top:18px}}
  input,select,textarea{{width:100%;padding:10px 12px;border:1px solid #ccc;
    border-radius:5px;font-size:15px;font-family:inherit}}
  textarea{{height:80px;resize:vertical}}
  .amounts{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:8px}}
  .amt-btn{{flex:1;min-width:70px;padding:10px;border:2px solid #ddd;
            background:#fff;border-radius:5px;cursor:pointer;font-size:15px;
            font-family:inherit;transition:all .15s}}
  .amt-btn:hover{{border-color:#52b788;background:#f0faf4}}
  .submit{{width:100%;margin-top:24px;padding:14px;background:#1b4332;
           color:#fff;border:none;border-radius:5px;font-size:16px;
           font-family:inherit;cursor:pointer;letter-spacing:.02em}}
  .submit:hover{{background:#2d6a4f}}
</style>
<script>
function setAmt(v){{
  document.getElementById('amount').value=v;
  document.querySelectorAll('.amt-btn').forEach(b=>
    b.style.borderColor = b.dataset.val==v ? '#1b4332' : '#ddd');
}}
</script>
</head><body>
<div class="wrap">
  <div class="header">
    <h1>Make a Donation</h1>
    <p>Your generosity makes a real difference.</p>
  </div>
  <div class="form-body">
    <form method="POST">
      <input type="hidden" name="tid" value="{tid}">

      <label>Your Name</label>
      <input type="text" name="name" value="{prefill_name}" placeholder="Full name" required>

      <label>Email</label>
      <input type="email" name="email" value="{prefill_email}" placeholder="you@example.com" required>

      <label>Select a Cause</label>
      <select name="cause">
        <option value="education"   {'selected' if prefill_cause=='education'   else ''}>Education</option>
        <option value="health"      {'selected' if prefill_cause=='health'      else ''}>Health</option>
        <option value="environment" {'selected' if prefill_cause=='environment' else ''}>Environment</option>
        <option value="general"     {'selected' if prefill_cause=='general'     else ''}>General Fund</option>
      </select>

      <label>Donation Amount ($)</label>
      <div class="amounts">
        <button type="button" class="amt-btn" data-val="25"  onclick="setAmt('25')">$25</button>
        <button type="button" class="amt-btn" data-val="50"  onclick="setAmt('50')">$50</button>
        <button type="button" class="amt-btn" data-val="100" onclick="setAmt('100')">$100</button>
        <button type="button" class="amt-btn" data-val="250" onclick="setAmt('250')">$250</button>
        <button type="button" class="amt-btn" data-val="500" onclick="setAmt('500')">$500</button>
      </div>
      <input type="number" id="amount" name="amount" placeholder="Or enter custom amount" min="1" required>

      <label>Note (optional)</label>
      <textarea name="note" placeholder="Anything you'd like us to know…"></textarea>

      <button type="submit" class="submit">Complete My Donation →</button>
    </form>
  </div>
</div>
</body></html>"""


@app.route("/status")
def status_page():
    """Quick browser dashboard — open http://localhost:5050/status"""
    log_data = _load_log()
    donors: dict[str, dict] = {}
    for entry in log_data:
        tid = entry["id"]
        if tid not in donors:
            donors[tid] = {"email": entry.get("email",""), "name": entry.get("name",""),
                           "tier": entry.get("tier",""), "events": []}
        donors[tid]["events"].append(entry["event"])

    rows = list(donors.values())
    rows.sort(key=lambda r: STATUS_ORDER.index(_derive_status(r["events"])))

    tbl_rows = ""
    bg = {"donated":"#d4edda","clicked_only":"#fff3cd",
          "opened_only":"#d1ecf1","no_action":"#f8d7da"}
    for r in rows:
        st = _derive_status(r["events"])
        tbl_rows += (
            f"<tr style='background:{bg[st]}'>"
            f"<td>{r['name'] or r['email']}</td>"
            f"<td>{r['email']}</td>"
            f"<td>{r['tier'].upper()}</td>"
            f"<td>{STATUS_LABEL[st]}</td>"
            f"</tr>"
        )

    return f"""<!DOCTYPE html><html><head><meta charset='UTF-8'>
<title>Campaign Status</title>
<meta http-equiv='refresh' content='15'>
<style>
  body{{font-family:Georgia,serif;padding:24px;background:#f9f6f1}}
  h1{{color:#1b4332}} table{{width:100%;border-collapse:collapse;margin-top:16px}}
  th{{background:#1b4332;color:#fff;padding:10px 14px;text-align:left}}
  td{{padding:9px 14px;border-bottom:1px solid #e0e0e0}}
  .badge{{font-size:11px;background:#1b4332;color:#fff;
          padding:3px 8px;border-radius:99px;margin-left:8px}}
</style></head><body>
<h1>Campaign Live Status <span class="badge">Auto-refreshes every 15s</span></h1>
<p>Excel file: <strong>{EXCEL_FILE}</strong> — updated in real-time</p>
<table><thead><tr><th>Name</th><th>Email</th><th>Tier</th><th>Status</th></tr></thead>
<tbody>{tbl_rows}</tbody></table>
</body></html>"""


# ══════════════════════════════════════════════════════════════════════════
#  STARTUP
# ══════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    # Build initial Excel from whatever is already in the log
    if Path(TRACKING_LOG).exists():
        _rebuild_excel()
        print(f"  ✓ Loaded existing tracking log → rebuilt {EXCEL_FILE}")

    print(f"""
╔══════════════════════════════════════════════════════╗
║   DONOR TRACKING SERVER  —  running on port {PORT}     ║
╠══════════════════════════════════════════════════════╣
║  Open browser status  : http://localhost:{PORT}/status  ║
║  Excel updates live   : {EXCEL_FILE:<28}║
║  Press Ctrl+C to stop                                ║
╚══════════════════════════════════════════════════════╝
""")
    app.run(host="0.0.0.0", port=PORT, debug=False)
