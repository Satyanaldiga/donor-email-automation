"""
═══════════════════════════════════════════════════════════════════
  DONOR EMAIL AUTOMATION  —  donor_automation.py
  ─────────────────────────────────────────────────
  1. Load donors from your Excel / CSV file
  2. Score every donor (RFM model) → hot / warm / cold tier
  3. Auto-draft the right personalised email per tier
  4. Embed open-pixel + tracked donate link per donor
  5. Send via Gmail (STARTTLS / App Password)
  6. Write every send to tracking_log.json
  7. Print a post-send summary

  REQUIRES tracker_server.py to be running FIRST so that
  opens, clicks and donations are caught in real-time.

  SETUP:
    pip install pandas openpyxl python-dotenv flask

  STEPS:
    1. python tracker_server.py      ← keep running in one terminal
    2. python donor_automation.py    ← run in a second terminal
    3. Open donor_tracker.xlsx       ← updates live as events come in
═══════════════════════════════════════════════════════════════════
"""

import smtplib, ssl, json, time, logging, os, uuid
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from datetime             import datetime, date
from dataclasses          import dataclass, field
from pathlib              import Path

import pandas as pd

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Logging ───────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("automation.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════
#  ❶  CONFIGURATION  — edit here or put values in a .env file
# ══════════════════════════════════════════════════════════════════════════

GMAIL_USER     = os.getenv("GMAIL_USER",     "you@gmail.com")
GMAIL_PASSWORD = os.getenv("GMAIL_PASSWORD", "xxxx xxxx xxxx xxxx")
FROM_NAME      = os.getenv("FROM_NAME",      "Your Organization")

# Where the tracking server is running (default: local)
TRACKER_HOST   = os.getenv("TRACKER_HOST",   "http://localhost:5050")

# Your real donation page (tracker will redirect here after click)
DONATE_URL     = os.getenv("DONATE_URL",     "http://localhost:5050/donate")

# Files
DONOR_FILE         = "donors.xlsx"       # or donors.csv
TRACKING_LOG       = "tracking_log.json"
SEND_DELAY_SECONDS = 1.5                 # pause between sends

# Set to True to send only to yourself (safe testing)
TEST_MODE          = True
# Set to False once you're ready to send to real donors


# ══════════════════════════════════════════════════════════════════════════
#  ❷  DONOR MODEL
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class Donor:
    email:                str
    first_name:           str
    last_name:            str   = ""
    lifetime_donations:   float = 0.0
    last_donation_amount: float = 0.0
    last_donation_date:   str   = ""
    donation_frequency:   int   = 0
    cause:                str   = "general"

    # scored fields
    potential_score:  float = 0.0
    tier:             str   = "cold"
    recommended_ask:  float = 0.0

    # unique per-send tracking id
    tracking_id: str = field(default_factory=lambda: uuid.uuid4().hex[:12])

    @property
    def full_name(self) -> str:
        return f"{self.first_name} {self.last_name}".strip()

    @property
    def days_since_last_gift(self) -> int:
        if not self.last_donation_date:
            return 9999
        try:
            last = datetime.strptime(self.last_donation_date[:10], "%Y-%m-%d").date()
            return (date.today() - last).days
        except ValueError:
            return 9999


# ══════════════════════════════════════════════════════════════════════════
#  ❸  LOAD DONORS
#     Edit the left-hand keys to match YOUR column headers exactly.
# ══════════════════════════════════════════════════════════════════════════

COLUMN_MAP = {
    "Email":               "email",
    "First Name":          "first_name",
    "Last Name":           "last_name",
    "Lifetime Donations":  "lifetime_donations",
    "Last Donation Amount":"last_donation_amount",
    "Last Donation Date":  "last_donation_date",
    "Donation Frequency":  "donation_frequency",
    "Cause":               "cause",
}

def load_donors(filepath: str = DONOR_FILE) -> list[Donor]:
    path = Path(filepath)
    if not path.exists():
        log.warning("'%s' not found — using built-in demo donors.", filepath)
        return _demo_donors()

    df = (pd.read_excel(path) if path.suffix.lower() in (".xlsx", ".xls")
          else pd.read_csv(path))

    rename = {k: v for k, v in COLUMN_MAP.items() if k in df.columns}
    df.rename(columns=rename, inplace=True)

    defaults = {
        "last_name": "", "lifetime_donations": 0, "last_donation_amount": 0,
        "last_donation_date": "", "donation_frequency": 0, "cause": "general",
    }
    for col, val in defaults.items():
        if col not in df.columns:
            df[col] = val

    donors = []
    for _, row in df.iterrows():
        try:
            donors.append(Donor(
                email               = str(row["email"]).strip(),
                first_name          = str(row["first_name"]).strip(),
                last_name           = str(row.get("last_name", "")).strip(),
                lifetime_donations  = float(row.get("lifetime_donations",  0) or 0),
                last_donation_amount= float(row.get("last_donation_amount",0) or 0),
                last_donation_date  = str(row.get("last_donation_date", ""))[:10],
                donation_frequency  = int(  row.get("donation_frequency",   0) or 0),
                cause               = str(row.get("cause", "general")).strip().lower(),
            ))
        except Exception as e:
            log.warning("Skipping row — %s: %s", row.get("email", "?"), e)

    log.info("Loaded %d donors from '%s'", len(donors), filepath)
    return donors


def _demo_donors() -> list[Donor]:
    return [
        Donor("alice@example.com", "Alice", "Johnson",
              lifetime_donations=1500, last_donation_amount=500,
              last_donation_date="2025-02-10", donation_frequency=6, cause="education"),
        Donor("bob@example.com", "Bob", "Martinez",
              lifetime_donations=320, last_donation_amount=60,
              last_donation_date="2025-04-20", donation_frequency=8, cause="health"),
        Donor("carol@example.com", "Carol", "Lee",
              lifetime_donations=850, last_donation_amount=300,
              last_donation_date="2024-08-05", donation_frequency=3, cause="environment"),
        Donor("david@example.com", "David", "Kim",
              lifetime_donations=2800, last_donation_amount=900,
              last_donation_date="2025-05-01", donation_frequency=10, cause="health"),
        Donor("emily@example.com", "Emily", "Chen",
              lifetime_donations=90, last_donation_amount=25,
              last_donation_date="2025-03-15", donation_frequency=2, cause="education"),
    ]


# ══════════════════════════════════════════════════════════════════════════
#  ❹  SCORING  (RFM — no ML needed)
# ══════════════════════════════════════════════════════════════════════════

def score_donors(donors: list[Donor]) -> list[Donor]:
    def recency(days):
        return 30 if days<=30 else 25 if days<=90 else 18 if days<=180 \
               else 10 if days<=365 else 4 if days<=730 else 0
    def frequency(f):
        return 25 if f>=10 else 20 if f>=6 else 14 if f>=3 else 7 if f>=1 else 0
    def monetary(m):
        return 30 if m>=2000 else 24 if m>=1000 else 17 if m>=500 \
               else 10 if m>=200 else 4 if m>=50 else 0
    def trajectory(last, freq, total):
        if freq==0 or total==0: return 0
        ratio = last / (total/freq)
        return 15 if ratio>=1.5 else 10 if ratio>=1.0 else 5 if ratio>=0.5 else 0

    for d in donors:
        d.potential_score = round(
            recency(d.days_since_last_gift) +
            frequency(d.donation_frequency) +
            monetary(d.lifetime_donations) +
            trajectory(d.last_donation_amount, d.donation_frequency, d.lifetime_donations), 1
        )
        d.tier = "hot" if d.potential_score>=70 else "warm" if d.potential_score>=40 else "cold"

        avg = (d.lifetime_donations/d.donation_frequency) if d.donation_frequency else d.last_donation_amount
        if d.tier == "hot":
            d.recommended_ask = round(max(avg*1.25, d.last_donation_amount*1.1), -1)
        elif d.tier == "warm":
            d.recommended_ask = round(d.last_donation_amount*1.15, -1)
        else:
            d.recommended_ask = round(max(d.last_donation_amount or 25, 25), -1)

    donors.sort(key=lambda x: x.potential_score, reverse=True)
    log.info("Scored — hot:%d  warm:%d  cold:%d",
             sum(1 for d in donors if d.tier=="hot"),
             sum(1 for d in donors if d.tier=="warm"),
             sum(1 for d in donors if d.tier=="cold"))
    return donors


# ══════════════════════════════════════════════════════════════════════════
#  ❺  EMAIL TEMPLATES
#     Each email contains:
#       • 1×1 open-tracking pixel  → tracker_server /open
#       • Donate button link       → tracker_server /click → your site
# ══════════════════════════════════════════════════════════════════════════

def _pixel(d: Donor) -> str:
    return (f'<img src="{TRACKER_HOST}/open?id={d.tracking_id}" '
            f'width="1" height="1" style="display:none" alt="">')

def _donate_link(d: Donor) -> str:
    dest = f"{DONATE_URL}?tid={d.tracking_id}"
    click_url = f"{TRACKER_HOST}/click?id={d.tracking_id}&url={dest}"
    return click_url

def _subject(d: Donor) -> str:
    return {
        "hot":  f"{d.first_name}, you're one of our most impactful supporters",
        "warm": f"Your generosity is working, {d.first_name} — a personal update",
        "cold": f"{d.first_name}, it's been a while — we'd love to reconnect",
    }[d.tier]

def _plain(d: Donor) -> str:
    ask = f"${d.recommended_ask:,.0f}"
    body = {
        "hot":  f"As one of our top supporters (${d.lifetime_donations:,.0f} lifetime), "
                f"we'd love for you to consider a gift of {ask} this quarter.",
        "warm": f"Your recent gift of ${d.last_donation_amount:,.0f} is making a real "
                f"difference. Consider {ask} this quarter to keep it going.",
        "cold": f"It's been a while. A gift of just {ask} would help us continue "
                f"the {d.cause.title()} work you once supported.",
    }[d.tier]
    return f"Hi {d.first_name},\n\n{body}\n\nDonate: {_donate_link(d)}\n\nThank you,\n{FROM_NAME}"

def _html(d: Donor) -> str:
    ask   = f"${d.recommended_ask:,.0f}"
    link  = _donate_link(d)
    pixel = _pixel(d)

    header_color = {"hot": "#1b4332", "warm": "#2d6a4f", "cold": "#444444"}[d.tier]

    intros = {
        "hot": (
            f"As one of our most dedicated supporters — with <strong>${d.lifetime_donations:,.0f}</strong> "
            f"given across <strong>{d.donation_frequency} gifts</strong> — we wanted to reach out personally "
            f"with an impact update and a special ask."
        ),
        "warm": (
            f"Your recent gift of <strong>${d.last_donation_amount:,.0f}</strong> is actively funding our "
            f"<strong>{d.cause.title()}</strong> programs right now. Here's a quick update on what "
            f"your support is achieving — and how you can build on it."
        ),
        "cold": (
            f"It's been a while since we last connected, and we wanted to reach out personally. "
            f"The <strong>{d.cause.title()}</strong> programs you once supported have grown significantly, "
            f"and we'd love to have you back."
        ),
    }

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>A message from {FROM_NAME}</title>
</head><body style="margin:0;padding:0;background:#f9f6f1;font-family:Georgia,serif;color:#1a1a1a;">

<div style="max-width:580px;margin:32px auto;">

  <!-- Header -->
  <div style="background:{header_color};padding:28px 36px;border-radius:10px 10px 0 0;">
    <p style="color:#a8d5b5;margin:0;font-size:12px;letter-spacing:.08em;
              text-transform:uppercase;">{FROM_NAME}</p>
    <h1 style="color:#fff;margin:8px 0 0;font-size:21px;font-weight:400;line-height:1.4;">
      {_subject(d)}
    </h1>
  </div>

  <!-- Body -->
  <div style="background:#fff;padding:32px 36px;border:1px solid #e0e0e0;">
    <p style="font-size:16px;">Dear {d.first_name},</p>
    <p style="font-size:15px;line-height:1.8;">{intros[d.tier]}</p>

    <!-- Impact box -->
    <div style="background:#f0faf4;border-left:4px solid #52b788;padding:16px 20px;
                margin:22px 0;border-radius:4px;">
      <p style="margin:0;font-size:14px;color:#1b4332;">
        <strong>Your {d.cause.title()} impact at a glance:</strong><br>
        Lifetime giving: <strong>${d.lifetime_donations:,.0f}</strong> across
        {d.donation_frequency} gift{"s" if d.donation_frequency!=1 else ""}.
        Thank you for being part of this.
      </p>
    </div>

    <p style="font-size:15px;line-height:1.8;">
      Based on your giving history, we believe a gift of <strong>{ask}</strong> this
      quarter would make a meaningful difference to our {d.cause.title()} work.
    </p>

    <!-- CTA Button -->
    <div style="text-align:center;margin:28px 0;">
      <a href="{link}"
         style="background:{header_color};color:#fff;text-decoration:none;
                padding:14px 36px;border-radius:5px;font-size:15px;
                display:inline-block;letter-spacing:.02em;">
        Give {ask} Now →
      </a>
    </div>

    <p style="font-size:14px;color:#555;line-height:1.7;">
      Any amount helps — the button above will take you to our secure donation page.
      If you have any questions, simply reply to this email.
    </p>

    <p style="font-size:14px;margin-top:28px;">
      With gratitude,<br>
      <strong>The {FROM_NAME} Team</strong>
    </p>
  </div>

  <!-- Footer -->
  <div style="background:#f0f0f0;padding:14px 36px;font-size:11px;color:#999;
              text-align:center;border-radius:0 0 10px 10px;">
    {FROM_NAME} ·
    <a href="https://www.yourorg.org/unsubscribe" style="color:#999;">Unsubscribe</a>
  </div>

</div>
{pixel}
</body></html>"""


# ══════════════════════════════════════════════════════════════════════════
#  ❻  TRACKING LOG  — record every "sent" event before sending
# ══════════════════════════════════════════════════════════════════════════

def log_sent(donors: list[Donor]) -> None:
    p = Path(TRACKING_LOG)
    existing = []
    if p.exists():
        try:
            existing = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass

    known_ids = {e["id"] for e in existing}
    new = [
        {
            "id":    d.tracking_id,
            "event": "sent",
            "email": d.email,
            "name":  d.full_name,
            "tier":  d.tier,
            "cause": d.cause,
            "score": d.potential_score,
            "ask":   d.recommended_ask,
            "ts":    datetime.now().isoformat(timespec="seconds"),
        }
        for d in donors if d.tracking_id not in known_ids
    ]
    p.write_text(json.dumps(existing + new, indent=2), encoding="utf-8")
    log.info("Tracking log updated — %d new entries.", len(new))


# ══════════════════════════════════════════════════════════════════════════
#  ❼  SEND VIA GMAIL
# ══════════════════════════════════════════════════════════════════════════

def send_all(donors: list[Donor]) -> list[dict]:
    results = []
    ctx = ssl.create_default_context()
    ts  = lambda: datetime.now().isoformat(timespec="seconds")

    if TEST_MODE:
        log.warning("⚠  TEST MODE — all emails go to %s", GMAIL_USER)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.ehlo()
        server.starttls(context=ctx)
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        log.info("Gmail authenticated ✓  sending %d emails …", len(donors))

        for d in donors:
            to = GMAIL_USER if TEST_MODE else d.email
            try:
                msg = MIMEMultipart("alternative")
                msg["Subject"] = _subject(d)
                msg["From"]    = f"{FROM_NAME} <{GMAIL_USER}>"
                msg["To"]      = to
                if TEST_MODE:
                    msg["X-Original-To"] = d.email
                msg.attach(MIMEText(_plain(d), "plain", "utf-8"))
                msg.attach(MIMEText(_html(d),  "html",  "utf-8"))

                server.sendmail(GMAIL_USER, to, msg.as_string())
                log.info("✓  %-36s  tier=%-4s  score=%5.1f",
                         d.email, d.tier, d.potential_score)
                results.append({"email": d.email, "status": "sent",   "ts": ts()})

            except Exception as e:
                log.error("✗  %s — %s", d.email, e)
                results.append({"email": d.email, "status": "error",  "ts": ts(), "error": str(e)})

            time.sleep(SEND_DELAY_SECONDS)

    return results


# ══════════════════════════════════════════════════════════════════════════
#  ❽  SUMMARY
# ══════════════════════════════════════════════════════════════════════════

def print_summary(donors: list[Donor], results: list[dict]) -> None:
    sent   = sum(1 for r in results if r["status"] == "sent")
    errors = len(results) - sent
    print("\n" + "═"*56)
    print("  SEND SUMMARY  —", datetime.now().strftime("%Y-%m-%d %H:%M"))
    print("═"*56)
    print(f"  Emails sent    : {sent} / {len(results)}")
    print(f"  Errors         : {errors}")
    print(f"  Tier breakdown : hot={sum(1 for d in donors if d.tier=='hot')}  "
          f"warm={sum(1 for d in donors if d.tier=='warm')}  "
          f"cold={sum(1 for d in donors if d.tier=='cold')}")
    print()
    print("  ── NEXT STEPS ──────────────────────────────────")
    print("  1. Keep tracker_server.py running")
    print(f"  2. Open donor_tracker.xlsx — it updates live")
    print(f"  3. Browser dashboard → http://localhost:5050/status")
    print("═"*56 + "\n")


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════

def main():
    log.info("═══ Donor Automation Pipeline ═══")

    donors  = load_donors(DONOR_FILE)
    donors  = score_donors(donors)

    # ── Optional: filter by tier ─────────────────────────────────────────
    # donors = [d for d in donors if d.tier in ("hot", "warm")]

    log_sent(donors)          # write to tracking_log.json BEFORE sending
    results = send_all(donors)
    print_summary(donors, results)


if __name__ == "__main__":
    main()
