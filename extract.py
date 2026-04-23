import mailbox
import os
import email
import argparse
from email import policy
from email.utils import parsedate_to_datetime

def safe_name(name):
    return "".join(c if c.isalnum() or c in "._-@" else "_" for c in name)

def get_year_month(message):
    try:
        dt = parsedate_to_datetime(message.get("date"))
        return str(dt.year), dt.strftime("%B")
    except:
        return None, None  # IMPORTANT: unknown dates won't create folders

def get_email_body_as_html(message):
    if message.is_multipart():
        for part in message.walk():
            if part.get_content_type() == "text/html":
                return part.get_payload(decode=True).decode(errors="ignore")
        for part in message.walk():
            if part.get_content_type() == "text/plain":
                text = part.get_payload(decode=True).decode(errors="ignore")
                return f"<pre>{text}</pre>"
    else:
        if message.get_content_type() == "text/html":
            return message.get_payload(decode=True).decode(errors="ignore")
        elif message.get_content_type() == "text/plain":
            text = message.get_payload(decode=True).decode(errors="ignore")
            return f"<pre>{text}</pre>"
    return None

def ensure_dir(path):
    """Create directory only when actually needed"""
    os.makedirs(path, exist_ok=True)

def main():
    parser = argparse.ArgumentParser()
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

    for i, message in enumerate(mbox):

        year, month = get_year_month(message)
        sender = safe_name(message.get("from", "unknown"))
        subject = safe_name(message.get("subject", f"email_{i}"))

        # Precompute folder path (BUT DO NOT CREATE YET)
        base_dir = args.output

        if year and month:
            base_dir = os.path.join(base_dir, year, month)
        else:
            base_dir = os.path.join(base_dir, "UnknownDate")

        if args.by_sender:
            base_dir = os.path.join(base_dir, sender)

        email_saved = False

        # ---------------- HTML ----------------
        export_html = False

        if args.html:
            export_html = any(
                part.get_content_disposition() == "attachment"
                for part in message.walk()
            )

        if args.html_all:
            export_html = True

        if export_html:
            html = get_email_body_as_html(message)

            if html:  # only create folder if content exists
                ensure_dir(base_dir)

                path = os.path.join(base_dir, f"{subject}_{i}.html")

                base, ext = os.path.splitext(path)
                c = 1
                while os.path.exists(path):
                    path = f"{base}_{c}{ext}"
                    c += 1

                with open(path, "w", encoding="utf-8") as f:
                    f.write(html)

                print(f"Saved email: {path}")
                email_saved = True

        # ---------------- ATTACHMENTS ----------------
        for part in message.walk():
            if part.get_content_disposition() != "attachment":
                continue

            filename = part.get_filename() or f"attachment_{i}"
            filename = safe_name(filename)

            filepath = os.path.join(base_dir, filename)

            base, ext = os.path.splitext(filepath)
            c = 1
            while os.path.exists(filepath):
                filepath = f"{base}_{c}{ext}"
                c += 1

            # ONLY create folder if we are actually saving something
            ensure_dir(base_dir)

            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))

            print(f"Saved attachment: {filepath}")
            email_saved = True

        # Optional: skip empty emails entirely
        # (no HTML + no attachments → no folder ever created)

    print("\nDone.")

if __name__ == "__main__":
    main()
