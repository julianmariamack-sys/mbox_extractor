import mailbox
import os
import email
import argparse
from email import policy
from email.utils import parsedate_to_datetime
import re
import hashlib


# -------------------------
# Helpers
# -------------------------

def safe_name(name: str) -> str:
    return "".join(c if c.isalnum() or c in "._-@" else "_" for c in name)


def get_year_month(msg):
    try:
        dt = parsedate_to_datetime(msg.get("date"))
        return str(dt.year), dt.strftime("%B")
    except:
        return None, None


def get_html(msg):
    if msg.is_multipart():
        for p in msg.walk():
            if p.get_content_type() == "text/html":
                return p.get_payload(decode=True).decode(errors="ignore")
        for p in msg.walk():
            if p.get_content_type() == "text/plain":
                text = p.get_payload(decode=True).decode(errors="ignore")
                return f"<pre>{text}</pre>"
    else:
        if msg.get_content_type() in ["text/html", "text/plain"]:
            return msg.get_payload(decode=True).decode(errors="ignore")
    return None


def build_base(output, year, month, sender=None, by_sender=False):
    base = output

    if year and month:
        base = os.path.join(base, year, month)
    else:
        base = os.path.join(base, "UnknownDate")

    if by_sender and sender:
        base = os.path.join(base, sender)

    return base


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def smart_filename(sender, subject, index):
    base = f"{sender}_{subject}_{index}"
    base = base[:80]  # truncate aggressively

    h = hashlib.md5(base.encode()).hexdigest()[:8]

    return f"{base}_{h}.html"

def unique_path(path):
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(path):
        path = f"{base}_{i}{ext}"
        i += 1
    return path


# -------------------------
# Main
# -------------------------

def main():
    parser = argparse.ArgumentParser(description="MBOX Extractor (clean refactor)")

    parser.add_argument("mbox")
    parser.add_argument("-o", "--output", default="output")

    parser.add_argument("--types", nargs="+")
    parser.add_argument("--by-sender", action="store_true")

    parser.add_argument("--html", action="store_true")
    parser.add_argument("--html-all", action="store_true")

    args = parser.parse_args()

    mbox = mailbox.mbox(
        args.mbox,
        factory=lambda f: email.message_from_binary_file(f, policy=policy.default)
    )

    email_count = 0
    attach_count = 0

    for i, msg in enumerate(mbox):

        # -------------------------
        # Metadata
        # -------------------------
        year, month = get_year_month(msg)

        raw_sender = msg.get("from", "unknown")
        sender_safe = safe_name(re.sub(r"<.*?>", "", raw_sender).strip())

        subject = safe_name(msg.get("subject", f"email_{i}"))

        base = build_base(
            args.output,
            year,
            month,
            sender_safe,
            args.by_sender
        )

        wrote_something = False

        # -------------------------
        # HTML EXPORT
        # -------------------------
        export_html = False

        if args.html:
            export_html = any(
                p.get_content_disposition() == "attachment"
                for p in msg.walk()
            )

        if args.html_all:
            export_html = True

        if export_html:
            html = get_html(msg)

            if html:
                ensure_dir(base)

                path = os.path.join(base, f"{sender_safe}_{subject}_{i}.html")
                path = unique_path(path)

                with open(path, "w", encoding="utf-8") as f:
                    f.write(html)

                print("Saved email:", path)
                email_count += 1
                wrote_something = True

        # -------------------------
        # ATTACHMENTS
        # -------------------------
        for part in msg.walk():
            if part.get_content_disposition() != "attachment":
                continue

            filename = part.get_filename() or f"attachment_{i}"
            filename = safe_name(filename)

            # type filtering
            if args.types:
                allowed = tuple(f".{t.lower().lstrip('.')}" for t in args.types)
                if not filename.lower().endswith(allowed):
                    continue

            ensure_dir(base)

            #path = os.path.join(base, filename)
            #path = unique_path(path)
            
            filename = smart_filename(sender_safe, subject, i)
            path = os.path.join(base, filename)
            path = unique_path(path)

            #with open(path, "wb") as f:
                #f.write(part.get_payload(decode=True))
                
            ##FIX1
            payload = part.get_payload(decode=True)

            # --- HARD GUARD ---
            if payload is None:
                raw = part.get_payload()

                if isinstance(raw, str):
                    payload = raw.encode("utf-8", errors="ignore")
                elif isinstance(raw, bytes):
                    payload = raw
                else:
                    print(f"Skipped invalid attachment (no usable data): {filename}")
                    continue

            # --- FINAL SAFETY CHECK ---
            if not isinstance(payload, (bytes, bytearray)):
                print(f"Skipped non-bytes attachment: {filename}")
                continue

            with open(path, "wb") as f:
                f.write(payload)
            ##FIX1

            print("Saved attachment:", path)
            attach_count += 1
            wrote_something = True

    print("\nDone.")
    print(f"Emails exported: {email_count}")
    print(f"Attachments saved: {attach_count}")


if __name__ == "__main__":
    main()
