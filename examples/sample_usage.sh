# Basic extraction
python extract.py mailbox.mbox

# Full HTML export
python extract.py mailbox.mbox --html-all

# Attachments only (PDFs)
python extract.py mailbox.mbox --types pdf

# Organized archive
python extract.py mailbox.mbox --html-all --by-sender
