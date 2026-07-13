#!/usr/bin/env python3
"""
process_garamut.py  —  Gmail → PDF → GitHub pipeline for The Garamut website.

Usage:
  python process_garamut.py            # normal run (push to GitHub)
  python process_garamut.py --dry-run  # connect & parse only, no push
  python process_garamut.py --help
"""

import argparse
import base64
import email
import imaplib
import json
import os
import re
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

# ── Configuration ────────────────────────────────────────────────────────────

IMAP_HOST     = "imap.gmail.com"
GMAIL_USER    = "norm.reed@gmail.com"
# Set GARAMUT_APP_PASSWORD in your environment (or in a .env file).
# Never hard-code credentials here.
GMAIL_PASS    = os.environ.get("GARAMUT_APP_PASSWORD", "")

SENDER_FILTER = "bethelcentrepng@gmail.com"
SUBJECT_HINT  = "GARAMUT"          # case-insensitive substring present in every subject

REPO_ROOT   = Path(__file__).parent.resolve()
PDF_DIR     = REPO_ROOT / "public" / "pdfs"
INDEX_HTML  = REPO_ROOT / "public" / "index.html"
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")

# Bible book names for scripture-reference extraction
BIBLE_BOOKS = [
    "Genesis","Exodus","Leviticus","Numbers","Deuteronomy","Joshua","Judges","Ruth",
    "1 Samuel","2 Samuel","1 Kings","2 Kings","1 Chronicles","2 Chronicles",
    "Ezra","Nehemiah","Esther","Job","Psalm","Psalms","Proverbs","Ecclesiastes",
    "Song of Solomon","Song of Songs","Isaiah","Jeremiah","Lamentations","Ezekiel",
    "Daniel","Hosea","Joel","Amos","Obadiah","Jonah","Micah","Nahum","Habakkuk",
    "Zephaniah","Haggai","Zechariah","Malachi",
    "Matthew","Mark","Luke","John","Acts","Romans",
    "1 Corinthians","2 Corinthians","Galatians","Ephesians","Philippians","Colossians",
    "1 Thessalonians","2 Thessalonians","1 Timothy","2 Timothy","Titus","Philemon",
    "Hebrews","James","1 Peter","2 Peter","1 John","2 John","3 John","Jude","Revelation",
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def clean_filename(topic: str, year: int, issue: int) -> str:
    """Return e.g. 2026_No19_Be_Still_Quietening_Our_Hearts.pdf"""
    safe = re.sub(r"[^\w\s-]", "", topic)
    safe = re.sub(r"\s+", "_", safe.strip())
    safe = safe[:60]  # truncate so paths stay sane
    return f"{year}_No{issue:02d}_{safe}.pdf"


def parse_subject(subject: str):
    """
    Extract (issue_number, topic) from subjects like:
      '18. THE GARAMUT PS BARRY'S REFLECTION - BE STILL...'
      '19. THE GARAMUT - ANOTHER TOPIC'
    Returns (int, str) or (None, None) on failure.
    """
    # Collapse email line-folding (\r\n + whitespace -> single space)
    subject = re.sub(r"\r?\n\s*", " ", subject).strip()
    # Look for leading number
    m = re.match(r"(\d+)[.\s]+", subject.strip())
    issue = int(m.group(1)) if m else None

    # Extract topic after "REFLECTION - " or after the last " - "
    topic = subject
    for marker in ["REFLECTION - ", "REFLECTION- ", " - "]:
        idx = subject.upper().find(marker.upper())
        if idx != -1:
            topic = subject[idx + len(marker):].strip()
            break
    # Strip trailing "No.XX for ..." / ". No.XX for ..." suffixes added by the sender
    topic = re.sub(r"[.,]?\s*[Nn]o\.?\s*\d+\s*(for\s+.+)?$", "", topic).strip()
    # Strip trailing ellipsis or punctuation
    topic = topic.rstrip(".,").strip()
    # Collapse whitespace
    topic = re.sub(r"\s+", " ", topic)

    return issue, topic


def extract_text_from_pdf(pdf_path: Path) -> str:
    """Use pdfminer.six to extract text. Falls back to empty string."""
    try:
        from pdfminer.high_level import extract_text as pm_extract
        return pm_extract(str(pdf_path))
    except ImportError:
        log("  WARNING: pdfminer.six not installed -- scripture extraction skipped.")
        log("           Run: pip install pdfminer.six")
        return ""
    except Exception as exc:
        log(f"  WARNING: PDF text extraction failed: {exc}")
        return ""


def extract_scripture_refs(text: str) -> str:
    """
    Find Bible references like 'John 3:16', 'Psalm 46:10-11', 'Matthew 5:3,6'.
    Returns a semicolon-separated string.
    """
    # Sort longest-name books first so "1 Corinthians" matches before "Corinthians"
    books_sorted = sorted(BIBLE_BOOKS, key=len, reverse=True)
    pattern = (
        r"(?<!\w)("
        + "|".join(re.escape(b) for b in books_sorted)
        + r")\s+(\d+(?::\d+(?:[–\-]\d+)?(?:,\s*\d+(?:[–\-]\d+)?)*)?)"
    )
    found = []
    seen = set()
    for m in re.finditer(pattern, text, re.IGNORECASE):
        book  = m.group(1).title()
        verse = m.group(2).strip()
        ref   = f"{book} {verse}"
        if ref not in seen:
            seen.add(ref)
            found.append(ref)
    return "; ".join(found)


def today_as_str() -> str:
    """Return date like '6 June 2026'."""
    return datetime.now().strftime("%-d %B %Y") if sys.platform != "win32" \
        else datetime.now().strftime("%#d %B %Y")


# ── Gmail IMAP ────────────────────────────────────────────────────────────────

def connect_gmail() -> imaplib.IMAP4_SSL:
    if not GMAIL_PASS:
        sys.exit(
            "ERROR: GARAMUT_APP_PASSWORD environment variable is not set.\n"
            "Set it to your Gmail App Password before running this script.\n"
            "See the README section 'Gmail App Password' for instructions."
        )
    log(f"Connecting to Gmail IMAP as {GMAIL_USER}...")
    imap = imaplib.IMAP4_SSL(IMAP_HOST)
    imap.login(GMAIL_USER, GMAIL_PASS)
    log("  Connected.")
    return imap


def fetch_unprocessed_emails(imap: imaplib.IMAP4_SSL):
    """
    Search INBOX for unread emails from SENDER_FILTER that contain SUBJECT_HINT.
    Returns list of (uid_bytes, email.message.Message).
    """
    imap.select("INBOX")
    # Search for UNSEEN messages from the sender
    _, data = imap.search(None, f'(UNSEEN FROM "{SENDER_FILTER}")')
    uids = data[0].split()
    log(f"  Found {len(uids)} unread message(s) from {SENDER_FILTER}.")

    results = []
    for uid in uids:
        _, msg_data = imap.fetch(uid, "(BODY.PEEK[])")
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        subject = msg.get("Subject", "")
        if SUBJECT_HINT.upper() not in subject.upper():
            log(f"  Skipping uid={uid.decode()} -- subject does not match: {subject!r}")
            continue
        results.append((uid, msg))
    return results


def download_pdf_attachment(msg) -> tuple[bytes, str] | tuple[None, None]:
    """Return (pdf_bytes, original_filename) or (None, None) if no PDF found."""
    for part in msg.walk():
        if part.get_content_type() == "application/pdf":
            fname = part.get_filename() or "attachment.pdf"
            return part.get_payload(decode=True), fname
        # Some mailers send PDFs as application/octet-stream
        if part.get_content_type() == "application/octet-stream":
            fname = part.get_filename() or ""
            if fname.lower().endswith(".pdf"):
                return part.get_payload(decode=True), fname
    return None, None


def mark_as_read(imap: imaplib.IMAP4_SSL, uid: bytes) -> None:
    imap.store(uid, "+FLAGS", "\\Seen")


# ── CATALOGUE update ──────────────────────────────────────────────────────────

def load_catalogue(html: str) -> list[dict]:
    m = re.search(r"const CATALOGUE = (\[.*?\]);", html, re.DOTALL)
    if not m:
        sys.exit("ERROR: Could not locate CATALOGUE array in index.html")
    return json.loads(m.group(1))


def entry_exists(catalogue: list[dict], year: int, issue: int) -> bool:
    return any(e["year"] == year and e["issue"] == issue for e in catalogue)


def insert_entry(catalogue: list[dict], entry: dict) -> list[dict]:
    """Insert in sorted order by (year, issue)."""
    catalogue.append(entry)
    catalogue.sort(key=lambda e: (e["year"], e["issue"]))
    return catalogue


def save_catalogue(html_path: Path, catalogue: list[dict]) -> None:
    html = html_path.read_text(encoding="utf-8")
    new_json = json.dumps(catalogue, ensure_ascii=False, separators=(", ", ": "))
    # Reformat: one object per logical line keeps diffs readable
    new_json = re.sub(r"\}, \{", "},\n{", new_json)
    new_block = f"const CATALOGUE = {new_json};"
    html = re.sub(
        r"const CATALOGUE = \[.*?\];",
        new_block,
        html,
        flags=re.DOTALL,
    )
    html_path.write_text(html, encoding="utf-8")


# ── Git / GitHub ──────────────────────────────────────────────────────────────

def git_run(*args, check=True):
    return subprocess.run(
        ["git", *args],
        cwd=str(REPO_ROOT),
        capture_output=True,
        text=True,
        check=check,
    )


def push_to_github(topic: str, filename: str) -> None:
    log("Staging changes...")
    git_run("add", f"public/pdfs/{filename}", "public/index.html")

    log("Committing...")
    msg = f"Add Garamut reflection: {topic}"
    git_run("commit", "-m", msg)

    # Embed token into remote URL for authentication
    result = git_run("remote", "get-url", "origin")
    remote_url = result.stdout.strip()
    # Insert token: https://TOKEN@github.com/...
    auth_url = re.sub(
        r"https://",
        f"https://{GITHUB_TOKEN}@",
        remote_url,
    )
    log("Pushing to GitHub...")
    subprocess.run(
        ["git", "push", auth_url, "HEAD:main"],
        cwd=str(REPO_ROOT),
        capture_output=True,
        text=True,
        check=True,
    )
    log("  Pushed successfully.")


# ── Main pipeline ─────────────────────────────────────────────────────────────

def process_message(uid, msg, dry_run: bool) -> bool:
    """Process one email. Returns True if a new reflection was published."""
    subject = msg.get("Subject", "")
    log(f"\nProcessing: {subject!r}")

    issue, topic = parse_subject(subject)
    if issue is None:
        log("  Could not parse issue number from subject -- skipping.")
        return False

    pdf_bytes, orig_fname = download_pdf_attachment(msg)
    if pdf_bytes is None:
        log("  No PDF attachment found -- skipping.")
        return False

    year = datetime.now().year
    filename = clean_filename(topic, year, issue)
    pdf_path = PDF_DIR / filename

    log(f"  Issue: No.{issue}  |  Topic: {topic}")
    log(f"  File:  {filename}  ({len(pdf_bytes):,} bytes)")

    # Check catalogue for duplicate
    html_text = INDEX_HTML.read_text(encoding="utf-8")
    catalogue  = load_catalogue(html_text)
    if entry_exists(catalogue, year, issue):
        log(f"  Issue {year}/No.{issue} already in CATALOGUE -- skipping.")
        return False

    # Save PDF to disk
    if not dry_run:
        PDF_DIR.mkdir(parents=True, exist_ok=True)
        pdf_path.write_bytes(pdf_bytes)
        log(f"  Saved PDF -> {pdf_path.relative_to(REPO_ROOT)}")
    else:
        log(f"  [DRY RUN] Would save PDF -> {pdf_path.relative_to(REPO_ROOT)}")

    # Extract scripture refs
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = Path(tmp.name)
    try:
        text = extract_text_from_pdf(tmp_path)
        refs = extract_scripture_refs(text)
    finally:
        tmp_path.unlink(missing_ok=True)

    if refs:
        log(f"  Scripture refs: {refs[:120]}{'...' if len(refs) > 120 else ''}")
    else:
        log("  No scripture refs detected.")

    # Build catalogue entry
    entry = {
        "year": year,
        "issue": issue,
        "date": today_as_str(),
        "topic": topic,
        "scripture_refs": refs,
        "file_id": "",
        "drive_url": f"pdfs/{filename}",
    }

    if dry_run:
        log(f"  [DRY RUN] Would add CATALOGUE entry: {json.dumps(entry, ensure_ascii=False)}")
        return False

    # Update index.html
    insert_entry(catalogue, entry)
    save_catalogue(INDEX_HTML, catalogue)
    log("  Updated CATALOGUE in index.html.")

    # Push to GitHub
    push_to_github(topic, filename)
    return True


def main():
    parser = argparse.ArgumentParser(description="Garamut Gmail -> GitHub pipeline")
    parser.add_argument("--dry-run", action="store_true",
                        help="Connect to Gmail and parse emails but do not write files or push")
    args = parser.parse_args()

    if args.dry_run:
        log("=== DRY RUN -- no files will be written, no git push ===")

    imap = connect_gmail()
    try:
        messages = fetch_unprocessed_emails(imap)
        if not messages:
            log("No new Garamut reflections found.")
            return

        published = 0
        for uid, msg in messages:
            ok = process_message(uid, msg, dry_run=args.dry_run)
            if ok:
                published += 1
                mark_as_read(imap, uid)

        log(f"\nDone. {published} reflection(s) published.")
    finally:
        imap.logout()


if __name__ == "__main__":
    main()
