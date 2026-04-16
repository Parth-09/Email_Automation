import os, base64, email.message, time, csv
from typing import Optional, Tuple
import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape
from jinja2.exceptions import TemplateNotFound

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv
load_dotenv()
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
]
CSV_PATH = os.getenv("CSV_PATH", "contacts.csv")
FOLLOWUP_TEMPLATE = os.getenv("FOLLOWUP_TEMPLATE", "followup_v1")
GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS")
RATE_PER_SECOND = float(os.getenv("RATE_PER_SECOND", "1.0"))

ATTACH_RESUME = os.getenv("ATTACH_RESUME", "")
ATTACH_COVER  = os.getenv("ATTACH_COVER", "")

env = Environment(
    loader=FileSystemLoader("templates"),
    autoescape=select_autoescape(["html","xml"])
)

def gmail_service():
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
                auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline",
                                                     include_granted_scopes="true")
                print("\n== Manual authentication ==")
                print("Open this URL in your browser and paste the FULL redirected URL here:\n", auth_url)
                redirected = input("> ")
                from urllib.parse import urlparse, parse_qs
                code = parse_qs(urlparse(redirected).query).get("code", [None])[0]
                if not code:
                    raise RuntimeError("No 'code' parameter found in the redirected URL.")
                flow.fetch_token(code=code)
                creds = flow.credentials
        with open("token.json","w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return build("gmail","v1",credentials=creds)

def render_html(template_id: str, ctx: dict) -> str:
    try:
        tpl = env.get_template(f"{template_id}.html.j2")
    except TemplateNotFound:
        tpl = env.get_template("followup_v1.html.j2")
    return tpl.render(**ctx)

def as_raw_reply(msg_from, to, subject, html,
                 in_reply_to: Optional[str],
                 references: Optional[str],
                 attachments=None):
    m = email.message.EmailMessage()
    m["From"] = msg_from
    m["To"] = to
    m["Subject"] = subject
    if in_reply_to:
        m["In-Reply-To"] = in_reply_to
    if references:
        m["References"] = references

    m.set_content("HTML email. Please view in an HTML-capable client.")
    m.add_alternative(html, subtype="html")

    attachments = attachments or []
    import mimetypes, os as _os
    for path in attachments:
        if not path:
            continue
        try:
            ctype,_ = mimetypes.guess_type(path)
            maintype, subtype = (ctype.split("/",1) if ctype else ("application","octet-stream"))
            with open(path,"rb") as f:
                data = f.read()
            m.add_attachment(data, maintype=maintype, subtype=subtype,
                             filename=_os.path.basename(path))
        except Exception as e:
            print(f"[warn] Could not attach {path}: {e}")

    raw = base64.urlsafe_b64encode(m.as_bytes()).decode()
    return {"raw": raw}

def lookup_thread_from_log(email_addr: str) -> Tuple[Optional[str], Optional[str]]:
    path = os.path.join("logs","sent.csv")
    if not os.path.exists(path):
        return None, None
    last_row = None
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("email","").strip().lower() == email_addr.strip().lower():
                last_row = row
    if not last_row:
        return None, None
    return last_row.get("gmail_thread_id") or None, last_row.get("gmail_message_id") or None

def find_last_outbound_to(service, to_email: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    q = f'to:{to_email} from:me -in:drafts'
    res = service.users().messages().list(userId="me", q=q, maxResults=1).execute()
    msgs = res.get("messages", [])
    if not msgs:
        return None, None, None
    msg = service.users().messages().get(
        userId="me",
        id=msgs[0]["id"],
        format="metadata",
        metadataHeaders=["Subject","Message-Id"]
    ).execute()
    thread_id = msg.get("threadId")
    headers = {h["name"].lower(): h["value"] for h in msg.get("payload",{}).get("headers",[])}
    message_id = headers.get("message-id")
    subject = headers.get("subject","")
    return thread_id, message_id, subject

def main():
    service = gmail_service()
    df = pd.read_csv(CSV_PATH, dtype=str).fillna("")
    last_sent_ts = [0.0]

    for _, r in df.iterrows():
        to_addr = r.get("email","").strip()
        if not to_addr:
            continue

        thread_id, _ = lookup_thread_from_log(to_addr)
        last_subject = ""
        thread_id, last_msg_id, last_subject = find_last_outbound_to(service, to_addr)
        if not thread_id:
            print(f"[skip] No prior thread found to {to_addr}; send an initial email first.")
            continue

        subject = f"Re: {last_subject or 'Following up'}"

        html = render_html(FOLLOWUP_TEMPLATE, {
            "first_name": r.get("first_name",""),
            "title": r.get("title",""),
            "company": r.get("company",""),
            "position": r.get("position",""),
        })

        attach_paths = []
        csv_resume = r.get("resume_path","").strip()
        csv_cover  = r.get("cover_letter_path","").strip()
        if csv_resume:
            attach_paths.append(csv_resume)
        elif ATTACH_RESUME:
            attach_paths.append(ATTACH_RESUME)
        if csv_cover:
            attach_paths.append(csv_cover)
        elif ATTACH_COVER:
            attach_paths.append(ATTACH_COVER)

        frm = GMAIL_ADDRESS or "me"
        body = as_raw_reply(frm, to_addr, subject, html, last_msg_id, last_msg_id, attachments=attach_paths)

        need_gap = 1.0 / max(RATE_PER_SECOND, 0.1)
        elapsed = time.time() - last_sent_ts[0]
        if elapsed < need_gap:
            time.sleep(need_gap - elapsed)

        service.users().messages().send(
            userId="me",
            body={**body, "threadId": thread_id}
        ).execute()

        last_sent_ts[0] = time.time()
        print(f"Follow-up sent → {to_addr} (thread {thread_id})")

if __name__ == "__main__":
    main()
