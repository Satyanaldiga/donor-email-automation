# Donor Email Automation

A fully local Python pipeline that sends personalised donor outreach emails, tracks opens / clicks / donations in real-time, and writes every event to a live-updating Excel sheet — no external services, no Google Sheets, no paid tools required.

---

## What it does

| Step | Script | Output |
|------|--------|--------|
| Load donors from CSV / Excel | `donor_automation.py` | — |
| Score every donor (RFM model) → hot / warm / cold | `donor_automation.py` | tier + recommended ask |
| Send personalised emails via Gmail | `donor_automation.py` | emails delivered |
| Track opens (pixel), clicks (link), donations (form) | `tracker_server.py` | `tracking_log.json` |
| Live Excel dashboard, auto-updates on every event | `tracker_server.py` | `donor_tracker.xlsx` |

### Donor status categories tracked

- ✅ **Donated** — clicked the link and submitted the donation form
- 🟡 **Clicked — No Donation Yet** — clicked the donate link but hasn't given yet
- 📧 **Opened — No Click** — opened the email but didn't click
- ❌ **No Response** — no open or click recorded

---

## Files

```
donor-email-automation/
├── donor_automation.py   # Main script: load → score → send
├── tracker_server.py     # Local Flask server: catches opens, clicks, donations
├── .env.example          # Credentials template (copy to .env, never commit .env)
├── .gitignore            # Keeps secrets and donor data off GitHub
└── README.md
```

> **Your donor file (`donors.csv` / `donors.xlsx`) is listed in `.gitignore` and will never be pushed to GitHub.**

---

## Setup

### 1. Install dependencies

```bash
pip install pandas openpyxl python-dotenv flask
```

### 2. Configure credentials

```bash
cp .env.example .env
```

Edit `.env`:

```env
GMAIL_USER=you@gmail.com
GMAIL_PASSWORD=xxxx xxxx xxxx xxxx   # Gmail App Password (not your regular password)
FROM_NAME=Your Organization Name
DONATE_URL=http://localhost:5050/donate   # or your real website donation URL
```

> **Gmail App Password:** Google Account → Security → 2-Step Verification → App Passwords → Generate one for "Mail".

### 3. Prepare your donor file

Place your file in the same folder as the scripts. Supported formats: `.xlsx`, `.xls`, `.csv`.

Expected columns (edit `COLUMN_MAP` in `donor_automation.py` to match your headers):

| Column | Example |
|--------|---------|
| Email | alice@example.com |
| First Name | Alice |
| Last Name | Johnson |
| Lifetime Donations | 1500.00 |
| Last Donation Amount | 500.00 |
| Last Donation Date | 2025-02-10 |
| Donation Frequency | 6 |
| Cause | education |

---

## Running

Open **two terminals** in the project folder:

```bash
# Terminal 1 — start the tracker (keep running the whole time)
python tracker_server.py

# Terminal 2 — send the emails
python donor_automation.py
```

Then open:
- **Browser dashboard** → http://localhost:5050/status (auto-refreshes every 15s)
- **Excel file** → `donor_tracker.xlsx` (updates live on every open / click / donation)

---

## Test mode

`donor_automation.py` ships with `TEST_MODE = True` — all emails are redirected to your own Gmail address so you can verify formatting before going live.

When ready to send to real donors:
```python
# donor_automation.py  line ~74
TEST_MODE = False
```

---

## Donor Scoring (RFM)

Each donor is scored 0–100 across four signals:

| Signal | Max pts | Logic |
|--------|---------|-------|
| Recency | 30 | Gave in last 30 days = full points |
| Frequency | 25 | 10+ lifetime gifts = full points |
| Monetary | 30 | $2,000+ lifetime = full points |
| Trajectory | 15 | Last gift > average gift = trending up |

Tiers: **Hot ≥ 70 · Warm 40–69 · Cold < 40**

Each tier gets a different email template and a calculated recommended ask amount.

---

## Privacy & Security

- `.env` (credentials) is gitignored — never committed
- `donors.csv` / `donors.xlsx` are gitignored — never committed
- `tracking_log.json` and `donor_tracker.xlsx` are gitignored — stay local only
- All tracking runs on `localhost` — no data leaves your machine

---

## Requirements

- Python 3.10+
- Gmail account with [App Password](https://myaccount.google.com/apppasswords) enabled
- `pandas`, `openpyxl`, `python-dotenv`, `flask`
