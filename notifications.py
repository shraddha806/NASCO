"""
notifications.py — Gmail Inbox Watcher for InsureBot Claims Dashboard
Polls GMAIL_USER via IMAP, stores insurance claim emails in-memory,
exposes get_claims_dashboard() and get_missed_email_by_id() for the API.
"""
import os
import re
import imaplib
import email as _email_lib
import threading
import time
import logging

from email.header import decode_header as _decode_header
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("insurebot.notifications")

#  Configuration 
GMAIL_USER          = os.environ.get("EMAIL_ADDRESS",       "priyankashah8324@gmail.com")
GMAIL_PASSWORD      = os.environ.get("EMAIL_PASSWORD",      "")
INBOX_POLL_SEC      = int(os.environ.get("INBOX_POLL_SEC",      "60"))
INBOX_MAX_AGE_HOURS = int(os.environ.get("INBOX_MAX_AGE_HOURS", "48"))

#  In-memory store (keyed by URL-safe Message-ID) 
_missed_emails: dict   = {}
_missed_emails_lock    = threading.Lock()
_inbox_watcher_started = False
_inbox_watcher_lock    = threading.Lock()


#  Header / body helpers 

def _decode_str(value) -> str:
    """Safely decode an email header value to a plain string."""
    if value is None:
        return ""
    parts = _decode_header(value)
    result = []
    for fragment, charset in parts:
        if isinstance(fragment, bytes):
            result.append(fragment.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(str(fragment))
    return " ".join(result).strip()


def _safe_id(msg_id: str) -> str:
    """Strip angle brackets and slashes so the ID is safe for use in a URL path."""
    return msg_id.strip("<>").replace("/", "_").replace(" ", "_")


def _extract_person_name(name: str) -> str:
    """
    Extract only the person's name (first word/name).
    - If name has multiple words, take first word only
    - If it looks like an email address, extract handle before @
    - Remove common suffixes and clean the name
    """
    if not name:
        return "Unknown"
    
    name = name.strip()
    
    # If it's an email address, extract the handle before @
    if "@" in name:
        email_handle = name.split("@")[0]
        # Clean up email handles with dots/underscores - take the first part
        name = email_handle.split(".")[0].split("_")[0]
    else:
        # Take only the first word (first name)
        name = name.split()[0] if name else "Unknown"
    
    # Remove common non-name suffixes and clean up
    name = name.strip().title() if name else "Unknown"
    
    return name


def _get_text_body(msg) -> str:
    """Extract plain-text body from an email.message.Message object."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                payload = part.get_payload(decode=True)
                if payload:
                    body = payload.decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    )
                    break
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            body = payload.decode(
                msg.get_content_charset() or "utf-8", errors="replace"
            )
    return body.strip()


#  Insurance-domain signal classifier 

_INSURANCE_SIGNALS = [
    "policy number", "policy no.", "policy no:", "policy no ",
    "claim reference", "claim ref",
    "insured vehicle", "insured amount",
    "insurance claim", "motor insurance", "health insurance",
    "claims team", "claims department", "claims processing",
    "claims settlement", "claims decision",
    "claim intimated", "claim approved", "claim rejected",
    "claim under review", "claim has been",
    "sum insured", "deductible applied",
    "date of intimation", "date of incident",
    "approved amount", "sanctioned amount",
    "reason for rejection", "reason for decline",
    "cashless", "reimbursement claim",
    "policy portal", "claims portal",
]


def _parse_claim_email(subject: str, body: str) -> dict:
    """
    Classify and extract structured fields from an insurance claim email.

    Returns a dict with at least:
      claim_status       – one of: intimated | under_review | approved | partial | rejected | unknown
      claim_status_label – human-readable emoji label

    An email is only classified as a claim if >= 2 insurance-domain signals
    are detected in the combined subject+body text.
    """
    STATUS_LABELS = {
        "intimated":    "Claim Intimated",
        "under_review": "Claim Under Review",
        "approved":     "Claim Approved",
        "partial":      "Claim Partially Approved",
        "rejected":     "Claim Rejected",
    }

    _UNKNOWN = {
        "claim_status": "unknown", "claim_status_label": None,
        "customer_name": None, "policy_number": None, "claim_ref": None,
        "vehicle": None, "date_of_intimation": None, "date_of_incident": None,
        "total_claimed": None, "approved_amount": None, "deductible": None,
        "payment_mode": None, "expected_credit_date": None,
        "rejection_reason": None, "hospital": None, "location": None,
    }

    combined = (subject + " " + body).lower()

    # Gate: require at least 2 insurance-specific signals
    signal_count = sum(1 for s in _INSURANCE_SIGNALS if s in combined)
    if signal_count < 2:
        return _UNKNOWN

    # Classify
    if "partially approved" in combined or "partial approval" in combined or "partially sanctioned" in combined:
        status = "partial"
    elif ("claim has been declined" in combined or "claim declined" in combined
          or "reason for rejection" in combined or "reason for decline" in combined
          or ("rejected" in combined and "claim" in combined)
          or ("declined" in combined and "claim" in combined)):
        status = "rejected"
    elif ("claim has been approved" in combined or "claim approved" in combined
          or ("approved" in combined and ("claim" in combined or "policy" in combined))):
        status = "approved"
    elif ("under review" in combined or "under assessment" in combined
          or "currently under review" in combined or "claim under review" in combined):
        status = "under_review"
    elif ("claim intimated" in combined or "claim has been successfully registered" in combined
          or "claim is now registered" in combined or "date of intimation" in combined
          or ("intimated" in combined and "claim" in combined)
          or ("successfully registered" in combined and "claim" in combined)):
        status = "intimated"
    else:
        status = "unknown"

    if status == "unknown":
        return _UNKNOWN

    def _rx(pattern):
        m = re.search(pattern, body, re.IGNORECASE)
        return m.group(1).strip() if m else None

    customer_name = _rx(r"Dear\s+(?:Mr\.|Ms\.|Mrs\.|Dr\.)?\s*([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+){0,3})")
    policy_number = _rx(r"Policy\s*(?:Number|No\.?|#)\s*[:\-]?\s*([\w\-]+)")
    claim_ref     = _rx(r"Claim\s*Reference\s*(?:Number|No\.?)?\s*[:\-]\s*([\w\-]+)")
    if not claim_ref:
        claim_ref = _rx(r"\bRef(?:erence)?\s*[:\-]\s*(CLM[\-\w]+)")
    if not claim_ref:
        claim_ref = _rx(r"\b(CLM-\d+)\b")

    vehicle              = _rx(r"Insured\s+Vehicle\s*[:\-]\s*(.+)")
    date_of_intimation   = _rx(r"Date\s+of\s+Intimation\s*[:\-]\s*([^;)\r\n]+)")
    date_of_incident     = _rx(r"Date\s+of\s+(?:Incident|Accident|Loss)\s*[:\-]\s*([^;)\r\n]+)")
    total_claimed        = _rx(r"Total\s+Claimed\s+Amount\s*[:\-]\s*([^\r\n]+)")
    approved_amount      = _rx(r"Approved\s+Amount\s*[:\-]\s*([^\r\n]+)")
    deductible           = _rx(r"Deductible\s+(?:Applied)?\s*[:\-]\s*([^\r\n]+)")
    payment_mode         = _rx(r"Payment\s+Mode\s*[:\-]\s*(.+)")
    expected_credit_date = _rx(r"Expected\s+Credit\s+(?:Date|Timeline)\s*[:\-]\s*(.+)")

    _rej_m = re.search(
        r"Reason\s+for\s+(?:Rejection|Decline)\s*[:\-]\s*(.+?)(?=\r?\n\r?\n|\r?\nAs per|\r?\nIf you|$)",
        body, re.IGNORECASE | re.DOTALL
    )
    rejection_reason = " ".join(_rej_m.group(1).split()) if _rej_m else None
    hospital = _rx(r"Hospital\s*(?:Name)?\s*[:\-]\s*(.+)")
    location = _rx(r"Location\s+of\s+(?:Incident|Accident|Loss)\s*[:\-]\s*(.+)")

    raw = {
        "claim_status":         status,
        "claim_status_label":   STATUS_LABELS.get(status, "Insurance Update"),
        "customer_name":        customer_name,
        "policy_number":        policy_number,
        "claim_ref":            claim_ref,
        "vehicle":              vehicle,
        "date_of_intimation":   date_of_intimation,
        "date_of_incident":     date_of_incident,
        "total_claimed":        total_claimed,
        "approved_amount":      approved_amount,
        "deductible":           deductible,
        "payment_mode":         payment_mode,
        "expected_credit_date": expected_credit_date,
        "rejection_reason":     rejection_reason,
        "hospital":             hospital,
        "location":             location,
    }
    return {k: v for k, v in raw.items()
            if v is not None or k in ("claim_status", "claim_status_label")}


#  IMAP poll 

def _poll_inbox_once(hours_back: int = 2) -> list:
    """
    Connect to Gmail IMAP, scan UNSEEN emails within the last `hours_back` hours,
    parse each one for claim content, store results in-memory.
    Non-claim emails are silently ignored.
    Returns list of newly stored claim entries.
    """
    if not GMAIL_PASSWORD:
        log.warning("[inbox] EMAIL_PASSWORD not set — cannot connect to Gmail IMAP")
        return []

    new_entries = []
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        mail.login(GMAIL_USER, GMAIL_PASSWORD)
        mail.select("INBOX")

        cutoff     = datetime.now(timezone.utc) - timedelta(hours=hours_back)
        since_date = cutoff.strftime("%d-%b-%Y")
        status, data = mail.search(None, f'UNSEEN SINCE "{since_date}"')
        if status != "OK" or not data[0].strip():
            mail.logout()
            return []

        seq_ids = data[0].split()
        log.info(f"[inbox] {len(seq_ids)} UNSEEN email(s) since {since_date} — scanning for claims")

        for seq in seq_ids:
            try:
                status2, msg_data = mail.fetch(seq, "(RFC822)")
                if status2 != "OK":
                    continue
                raw = msg_data[0][1]
                msg = _email_lib.message_from_bytes(raw)

                msg_id_raw = msg.get("Message-ID", "").strip()
                msg_id     = msg_id_raw if msg_id_raw else f"seq-{seq.decode()}"
                safe_id    = _safe_id(msg_id)

                with _missed_emails_lock:
                    if safe_id in _missed_emails:
                        continue

                from_raw = msg.get("From", "")
                _, from_addr = _email_lib.utils.parseaddr(from_raw)
                subject  = _decode_str(msg.get("Subject", "(no subject)"))
                date_str = msg.get("Date", "")

                try:
                    received_at = _email_lib.utils.parsedate_to_datetime(date_str)
                    if received_at.tzinfo is None:
                        received_at = received_at.replace(tzinfo=timezone.utc)
                    received_str = received_at.strftime("%d %b %Y  %H:%M UTC")
                except Exception:
                    received_at  = datetime.now(timezone.utc)
                    received_str = received_at.strftime("%d %b %Y  %H:%M UTC")

                if received_at < cutoff:
                    continue

                full_body    = _get_text_body(msg)
                body_preview = full_body[:300]

                claim_info = _parse_claim_email(subject, full_body)
                if claim_info["claim_status"] == "unknown":
                    log.debug(f"[inbox] Skipping non-claim email: {subject!r}")
                    continue

                entry = {
                    "safe_id":      safe_id,
                    "message_id":   msg_id,
                    "from_addr":    from_addr,
                    "subject":      subject,
                    "received_str": received_str,
                    "body_preview": body_preview,
                    "full_body":    full_body,
                    "claim_info":   claim_info,
                }

                with _missed_emails_lock:
                    _missed_emails[safe_id] = entry

                sep = "=" * 60
                ci  = claim_info
                print(sep)
                print(f"  CLAIM EMAIL from {from_addr}")
                print(f"  Status  : {ci.get('claim_status_label', ci.get('claim_status'))}")
                print(f"  Subject : {subject}")
                print(f"  Received: {received_str}")
                if ci.get("customer_name"):
                    print(f"  Customer: {ci['customer_name']}")
                if ci.get("policy_number"):
                    print(f"  Policy #: {ci['policy_number']}")
                if ci.get("claim_ref"):
                    print(f"  Claim # : {ci['claim_ref']}")
                print(sep)

                new_entries.append(entry)

            except Exception as exc:
                log.warning(f"[inbox] Could not parse email seq {seq}: {exc}")

        mail.logout()

    except imaplib.IMAP4.error as exc:
        log.error(f"[inbox] IMAP error: {exc}")
    except Exception as exc:
        log.error(f"[inbox] Unexpected error: {exc}")

    return new_entries


#  Public API 

def get_missed_email_by_id(safe_id: str):
    """Return a single email entry (including full_body) or None."""
    with _missed_emails_lock:
        return _missed_emails.get(safe_id)


def get_claims_dashboard(
    status_filter: str = "all",
    search: str = "",
    page: int = 1,
    per_page: int = 10,
    include_handled: bool = True,
) -> dict:
    """
    Return structured data for the Claims Review Dashboard.

    status_filter : "all" | "approved" | "rejected" | "pending"
    search        : free-text filter
    page          : 1-based page number
    per_page      : items per page
    include_handled: include all emails (no handled flag anymore, kept for compat)

    Returns { summary, claims, pagination }
    """
    with _missed_emails_lock:
        all_emails = list(_missed_emails.values())

    def _build_record(e: dict) -> dict:
        ci         = e.get("claim_info") or {}
        raw_status = ci.get("claim_status", "unknown")

        if raw_status in ("approved", "partial"):
            dash_status  = "approved"
            status_label = "Approved"
        elif raw_status == "rejected":
            dash_status  = "rejected"
            status_label = "Rejected"
        else:
            dash_status  = "pending"
            status_label = "Pending"

        if raw_status == "rejected":
            reason = ci.get("rejection_reason") or "Claim rejected"
        elif raw_status == "approved":
            reason = (
                f"Approved – payment issued ({ci['approved_amount']})"
                if ci.get("approved_amount")
                else "Claim meets all requirements"
            )
        elif raw_status == "partial":
            reason = (
                f"Partially approved: {ci['approved_amount']}"
                if ci.get("approved_amount") else "Partially approved"
            )
        elif raw_status == "under_review":
            reason = "Claim under review"
        elif raw_status == "intimated":
            reason = "Claim intimated / registered"
        else:
            reason = ci.get("claim_status_label") or "Under review"

        claim_id     = ci.get("claim_ref") or e.get("safe_id", "")[:10].upper()
        # Extract only the person's name (first word) for submitted_by
        raw_name = ci.get("customer_name") or e.get("from_addr", "Unknown")
        submitted_by = _extract_person_name(raw_name)
        date_submitted = ci.get("date_of_intimation") or e.get("received_str", "")

        return {
            "claim_id":       claim_id,
            "safe_id":        e.get("safe_id", ""),
            "submitted_by":   submitted_by,
            "status":         dash_status,
            "status_label":   status_label,
            "reason":         reason,
            "date_submitted": date_submitted,
            "policy_number":  ci.get("policy_number"),
            "from_addr":      e.get("from_addr", ""),
            "subject":        e.get("subject", ""),
            "body_preview":   e.get("body_preview") or e.get("full_body", "")[:300],
        }

    all_claims = [_build_record(e) for e in all_emails]

    summary = {
        "total":    len(all_claims),
        "approved": sum(1 for c in all_claims if c["status"] == "approved"),
        "rejected": sum(1 for c in all_claims if c["status"] == "rejected"),
        "pending":  sum(1 for c in all_claims if c["status"] == "pending"),
    }

    if status_filter and status_filter != "all":
        filtered = [c for c in all_claims if c["status"] == status_filter]
    else:
        filtered = list(all_claims)

    if search:
        q = search.lower()
        filtered = [
            c for c in filtered
            if (
                q in c["claim_id"].lower()
                or q in c["submitted_by"].lower()
                or q in c["reason"].lower()
                or q in c["from_addr"].lower()
                or q in c["subject"].lower()
            )
        ]

    filtered.sort(key=lambda x: x.get("date_submitted", ""), reverse=True)

    total_items = len(filtered)
    total_pages = max(1, (total_items + per_page - 1) // per_page)
    page        = max(1, min(page, total_pages))
    start       = (page - 1) * per_page
    page_items  = filtered[start: start + per_page]

    return {
        "summary": summary,
        "claims":  page_items,
        "pagination": {
            "page":        page,
            "per_page":    per_page,
            "total_items": total_items,
            "total_pages": total_pages,
        },
    }


#  Background watcher 

def _inbox_watcher_loop():
    log.info(
        f"[inbox] Watcher started — polling every {INBOX_POLL_SEC}s, Gmail: {GMAIL_USER}"
    )

    # One-time startup historical back-fill
    log.info(f"[inbox] Startup historical scan for last {INBOX_MAX_AGE_HOURS}h ...")
    try:
        historic = _poll_inbox_once(hours_back=INBOX_MAX_AGE_HOURS)
        if historic:
            log.info(f"[inbox] Historical scan: {len(historic)} claim email(s) loaded")
        else:
            log.info("[inbox] Historical scan complete — inbox is empty")
    except Exception as exc:
        log.warning(f"[inbox] Historical scan error: {exc}")

    while True:
        time.sleep(INBOX_POLL_SEC)
        try:
            new_entries = _poll_inbox_once(hours_back=2)
            if new_entries:
                log.info(f"[inbox] {len(new_entries)} new claim email(s) stored")
        except Exception as exc:
            log.warning(f"[inbox] watcher loop error: {exc}")


def start_inbox_watcher():
    """Start the Gmail inbox monitor daemon thread. Safe to call multiple times."""
    global _inbox_watcher_started
    with _inbox_watcher_lock:
        if _inbox_watcher_started:
            return
        _inbox_watcher_started = True

    t = threading.Thread(target=_inbox_watcher_loop, name="inbox-watcher", daemon=True)
    t.start()
    log.info("[inbox] Gmail inbox watcher thread launched")
