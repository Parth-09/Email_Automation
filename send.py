import os, time, base64, email.message, csv, random
from datetime import datetime, timedelta, timezone
from dateutil import parser as dtparse

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape
from jinja2.exceptions import TemplateNotFound
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.date import DateTrigger

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv
load_dotenv()
# -------- Config via env --------
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
]
CSV_PATH = os.getenv("CSV_PATH", "contacts.csv")
GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS")  # optional; "me" inferred if omitted

# CAP: 0 / "" / None => no cap
_raw_cap = os.getenv("DAILY_CAP", "")
DAILY_CAP = int(_raw_cap) if str(_raw_cap).strip().isdigit() else 0

RATE_PER_SECOND = float(os.getenv("RATE_PER_SECOND", "1.0"))  # throttle ~1/sec
DRY_RUN = os.getenv("DRY_RUN", "0") == "1"
DRY_RUN_TO = os.getenv("DRY_RUN_TO", "")
SKIP_WEEKENDS = os.getenv("SKIP_WEEKENDS", "0") == "1"
BUSINESS_HOURS = os.getenv("BUSINESS_HOURS", "09:00-17:00")  # advisory only

# Staggering when CSV has no explicit send time
SPACING = int(os.getenv("SPACING_SECONDS", "20"))  # seconds between sends
JITTER_MAX = int(os.getenv("JITTER_MAX_SECONDS", "4"))
MISFIRE_GRACE = int(os.getenv("MISFIRE_GRACE_SECONDS", "600"))  # 10 min

ATTACH_RESUME = os.getenv("ATTACH_RESUME", "")      # default resume for all
ATTACH_COVER  = os.getenv("ATTACH_COVER", "")       # default cover letter for all

env = Environment(
    loader=FileSystemLoader("templates"),
    autoescape=select_autoescape(["html", "xml"])
)

def gmail_service():
    """Authenticate and return a Gmail API service (Windows/PowerShell friendly)."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            from google.auth.transport.requests import Request
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            try:
                creds = flow.run_local_server(port=0, open_browser=True)
            except Exception:
                auth_url, _ = flow.authorization_url(
                    prompt="consent",
                    access_type="offline",
                    include_granted_scopes="true",
                )
                print("\n== Manual authentication ==")
                print("1) Open this URL in your browser, allow access:")
                print(auth_url)
                redirected = input(
                    "\n2) After allowing, copy the FULL redirected URL from the address bar\n"
                    "   and paste it here, then press Enter:\n> "
                )
                from urllib.parse import urlparse, parse_qs
                code = parse_qs(urlparse(redirected).query).get("code", [None])[0]
                if not code:
                    raise RuntimeError("No 'code' parameter found in the redirected URL.")
                flow.fetch_token(code=code)
                creds = flow.credentials

        with open("token.json", "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)

def as_raw_email(msg_from, to, subject, html, attachments=None):
    """Build RFC822 email with optional attachments; return Gmail API 'raw' payload."""
    m = email.message.EmailMessage()
    m["From"] = msg_from
    m["To"] = to
    m["Subject"] = subject

    m.set_content("HTML email. Please view in an HTML-capable client.")
    m.add_alternative(html, subtype="html")

    attachments = attachments or []
    for path in attachments:
        if not path:
            continue
        try:
            import mimetypes, os as _os
            ctype, _ = mimetypes.guess_type(path)
            maintype, subtype = (ctype.split("/", 1) if ctype else ("application", "octet-stream"))
            filename = _os.path.basename(path)
            with open(path, "rb") as f:
                data = f.read()
            m.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
        except Exception as e:
            print(f"[warn] Could not attach {path}: {e}")

    raw = base64.urlsafe_b64encode(m.as_bytes()).decode()
    return {"raw": raw}

def render_html(template_id, ctx):
    chosen = (template_id or "intro_v1").strip()
    try:
        tpl = env.get_template(f"{chosen}.html.j2")
    except TemplateNotFound:
        tpl = env.get_template("intro_v1.html.j2")
    return tpl.render(**ctx)

def within_business_hours(dt_local):
    if SKIP_WEEKENDS and dt_local.weekday() >= 5:
        return False
    try:
        start_s, end_s = BUSINESS_HOURS.split("-")
        s_h, s_m = map(int, start_s.split(":"))
        e_h, e_m = map(int, end_s.split(":"))
        mins = dt_local.hour * 60 + dt_local.minute
        return (s_h * 60 + s_m) <= mins <= (e_h * 60 + e_m)
    except Exception:
        return True

def send_one(service, row, last_sent_ts, scheduler):
    """Send one email (rate-limited), log it, then auto-exit after last job."""
    # Throttle
    now = time.time()
    need_gap = 1.0 / max(RATE_PER_SECOND, 0.1)
    elapsed = now - last_sent_ts[0]
    if elapsed < need_gap:
        time.sleep(need_gap - elapsed)

    # Render email
    html = render_html(row.get("template_id", "intro_v1"), {
        "first_name": row.get("first_name", ""),
        "title": row.get("title", ""),
        "company": row.get("company", ""),
        "position": row.get("position", ""),
    })

    # Subject
    first_name = (row.get("first_name", "") or "").strip()
    company = (row.get("company", "") or "").strip()
    position = (row.get("position", "") or "").strip()

    if position and company:
        subject = f"{first_name}, interest in contributing to {company} as a {position}"
    elif company:
        subject = f"{first_name}, opportunities at {company}"
    else:
        subject = f"{first_name}, quick note"

    # Recipient + sender
    to_addr = (DRY_RUN_TO or row["email"]) if DRY_RUN else row["email"]
    frm = GMAIL_ADDRESS or "me"

    # Attachments
    attach_paths = []
    csv_resume = (row.get("resume_path", "") or "").strip()
    csv_cover  = (row.get("cover_letter_path", "") or "").strip()

    if csv_resume:
        attach_paths.append(csv_resume)
    elif ATTACH_RESUME:
        attach_paths.append(ATTACH_RESUME)

    if csv_cover:
        attach_paths.append(csv_cover)
    elif ATTACH_COVER:
        attach_paths.append(ATTACH_COVER)

    # Send
    body = as_raw_email(frm, to_addr, subject, html, attachments=attach_paths)
    resp = service.users().messages().send(userId="me", body=body).execute()

    # Log
    last_sent_ts[0] = time.time()
    os.makedirs("logs", exist_ok=True)
    log_path = os.path.join("logs", "sent.csv")
    newfile = not os.path.exists(log_path)

    with open(log_path, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if newfile:
            w.writerow([
                "ts_utc","email","campaign_id","template_id","subject","dry_run",
                "gmail_message_id","gmail_thread_id"
            ])
        w.writerow([
            datetime.utcnow().isoformat(), row["email"], row.get("campaign_id",""),
            row.get("template_id","intro_v1"), subject, int(DRY_RUN),
            resp.get("id",""), resp.get("threadId","")
        ])

    print(f"Sent → {to_addr}  (thread {resp.get('threadId','?')})")

    # ✅ Auto-exit after the last scheduled job runs
    # APScheduler's BlockingScheduler keeps running even when all jobs are done,
    # so we explicitly shutdown after the final job completes.
    try:
        remaining = scheduler.get_jobs()
        if len(remaining) <= 1:
            scheduler.shutdown(wait=False)
    except Exception:
        pass

def load_jobs_from_csv():
    df = pd.read_csv(CSV_PATH, dtype=str).fillna("")
    if DAILY_CAP > 0:
        df = df.head(DAILY_CAP)

    jobs = []
    for _, r in df.iterrows():
        send_at_s = (r.get("send_time_iso", "") or "").strip()
        when = dtparse.isoparse(send_at_s).astimezone(timezone.utc) if send_at_s else None
        jobs.append((when, r.to_dict()))
    return jobs

def main():
    service = gmail_service()
    last_sent_ts = [0.0]

    jobs = load_jobs_from_csv()
    if not jobs:
        print("No rows found in CSV.")
        return

    sched = BlockingScheduler(
        job_defaults={
            "coalesce": True,
            "max_instances": 1,
            "misfire_grace_time": MISFIRE_GRACE,
        },
        timezone="UTC",
    )

    base_start = datetime.now(timezone.utc) + timedelta(seconds=5)

    for i, (when_utc, row) in enumerate(jobs):
        if when_utc is None:
            jitter = random.randint(0, max(JITTER_MAX, 0))
            when_utc = base_start + timedelta(seconds=i * max(SPACING, 0) + jitter)

        try:
            local_guess = when_utc.astimezone()
            if not within_business_hours(local_guess):
                print(f"[warn] {row['email']} scheduled outside BUSINESS_HOURS: {local_guess}")
        except Exception:
            pass

        trig = DateTrigger(run_date=when_utc)
        sched.add_job(
            send_one,
            trigger=trig,
            args=[service, row, last_sent_ts, sched],  # ✅ pass scheduler
            id=f"send_{i}",
            replace_existing=True,
        )

    print(f"Scheduled {len(jobs)} job(s). Running…")
    try:
        sched.start()
    except (KeyboardInterrupt, SystemExit):
        print("Stopped.")

if __name__ == "__main__":
    main()
