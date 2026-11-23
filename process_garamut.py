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
import traceback
from datetime import datetime, timedelta
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

LOOKBACK_DAYS = 60   # how far back to look for reflections that were never published
NET_TIMEOUT   = 60   # seconds before a stalled Gmail or GitHub call gives up

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


def decode_subject(msg) -> str:
    """Decode a (possibly MIME-encoded, e.g. '=?UTF-8?Q?...?=') Subject header."""
    from email.header import decode_header
    raw = msg.get("Subject", "") or ""
    parts = decode_header(raw)
    decoded = "".join(
        chunk.decode(charset or "utf-8", errors="replace") if isinstance(chunk, bytes) else chunk
        for chunk, charset in parts
    )
    return decoded


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

    # Extract topic after known markers (tried in order)
    topic = subject
    for marker in ["TOPIC: ", "TOPIC:", "THEME: ", "THEME:", "REFLECTION - ", "REFLECTION- ", " - "]:
        idx = subject.upper().find(marker.upper())
        if idx != -1:
            topic = subject[idx + len(marker):].strip()
            break
    # Strip surrounding quotes
    topic = topic.strip("'\"")
    # Strip trailing "No.XX for ..." / ". No.XX for ..." / "being for ..." suffixes
    # (case-insensitive: subjects are sometimes fully upper-cased)
    topic = re.sub(r"[.,]?\s*(being\s+)?for\s+.+$", "", topic, flags=re.IGNORECASE).strip()
    topic = re.sub(r"[.,]?\s*no\.?\s*\d+\s*(for\s+.+)?$", "", topic, flags=re.IGNORECASE).strip()
    # Strip surrounding quotes again now that a trailing suffix may have exposed one
    topic = topic.strip("'\"")
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


def email_date(msg) -> datetime:
    """Return the email's Date header as a datetime, falling back to now()."""
    from email.utils import parsedate_to_datetime
    raw = msg.get("Date", "")
    try:
        return parsedate_to_datetime(raw)
    except (TypeError, ValueError):
        return datetime.now()


def date_as_str(dt: datetime) -> str:
    """Return date like '6 June 2026'."""
    return dt.strftime("%-d %B %Y") if sys.platform != "win32" \
        else dt.strftime("%#d %B %Y")


# ── Gmail IMAP ────────────────────────────────────────────────────────────────

def connect_gmail() -> imaplib.IMAP4_SSL:
    if not GMAIL_PASS:
        sys.exit(
            "ERROR: GARAMUT_APP_PASSWORD environment variable is not set.\n"
            "Set it to your Gmail App Password before running this script.\n"
            "See the README section 'Gmail App Password' for instructions."
        )
    log(f"Connecting to Gmail IMAP as {GMAIL_USER}...")
    imap = imaplib.IMAP4_SSL(IMAP_HOST, timeout=NET_TIMEOUT)
    imap.login(GMAIL_USER, GMAIL_PASS)
    log("  Connected.")
    return imap


def imap_since(days: int) -> str:
    """IMAP date like '25-Sep-2026' (month names must be English whatever the locale)."""
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    d = datetime.now() - timedelta(days=days)
    return f"{d.day:02d}-{months[d.month - 1]}-{d.year}"


def select_all_mail(imap: imaplib.IMAP4_SSL) -> None:
    """Select Gmail's All Mail folder so archived emails are seen too; INBOX if it can't be found."""
    _, boxes = imap.list()
    for raw in boxes or []:
        line = raw.decode(errors="replace")
        # The folder is found by its \All flag because its name varies by region
        if "\\All" in line.split(")")[0]:
            name = line.rsplit(' "/" ', 1)[-1]
            if imap.select(name)[0] == "OK":
                return
    imap.select("INBOX")


def fetch_unpublished_emails(imap: imaplib.IMAP4_SSL):
    """
    Look at every recent Garamut email from SENDER_FILTER, whether read, unread
    or archived, and return the ones whose issue is not in the published catalogue.
    Returns list of (uid_bytes, email.message.Message).
    """
    published = committed_catalogue()
    select_all_mail(imap)
    _, data = imap.search(
        None, f'(FROM "{SENDER_FILTER}" SINCE {imap_since(LOOKBACK_DAYS)})'
    )
    uids = data[0].split()

    results = []
    for uid in uids:
        # Headers first, so already-published emails are never downloaded in full
        _, hdr_data = imap.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])")
        hdr = email.message_from_bytes(hdr_data[0][1])
        subject = decode_subject(hdr)
        if SUBJECT_HINT.upper() not in subject.upper():
            continue
        issue, _topic = parse_subject(subject)
        if issue is not None and entry_exists(published, email_date(hdr).year, issue):
            continue
        _, msg_data = imap.fetch(uid, "(BODY.PEEK[])")
        results.append((uid, email.message_from_bytes(msg_data[0][1])))

    log(f"  Checked {len(uids)} email(s) from the last {LOOKBACK_DAYS} days; "
        f"{len(results)} not yet published.")
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


def committed_catalogue() -> list[dict]:
    """The catalogue as of the last commit, i.e. what has really been published."""
    return load_catalogue(git_run("show", "HEAD:public/index.html").stdout)


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
    # Write beside the real file then swap, so a run killed mid-write
    # can never leave a truncated index.html behind.
    tmp_path = html_path.with_name(html_path.name + ".tmp")
    tmp_path.write_text(html, encoding="utf-8")
    os.replace(tmp_path, html_path)


# ── Git / GitHub ──────────────────────────────────────────────────────────────

def git_run(*args, check=True, timeout=NET_TIMEOUT, config=()):
    """Run git. `config` is a list of 'key=value' settings applied to this call only."""
    cmd = ["git"]
    for setting in config:
        cmd += ["-c", setting]
    cmd += args
    try:
        result = subprocess.run(
            cmd,
            cwd=str(REPO_ROOT),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        # Raised afresh so the command line (which may hold the token) is not logged
        raise RuntimeError(f"git {args[0]} timed out after {timeout}s") from None
    if check and result.returncode != 0:
        detail = (result.stderr or result.stdout).strip()
        if GITHUB_TOKEN:
            detail = detail.replace(GITHUB_TOKEN, "***")
        raise RuntimeError(f"git {args[0]} failed: {detail}")
    return result


def unpushed_commits() -> int:
    """How many local commits GitHub does not have yet, as far as git knows."""
    result = git_run("rev-list", "--count", "origin/main..HEAD", check=False)
    return int(result.stdout.strip() or 0) if result.returncode == 0 else 0


def git_push() -> None:
    log("Pushing to GitHub...")
    # Pushing to 'origin' by name (not a bare URL) keeps git's record of what
    # has been pushed up to date; the token is supplied for this call only.
    config = []
    if GITHUB_TOKEN:
        config.append(
            f"url.https://{GITHUB_TOKEN}@github.com/.insteadOf=https://github.com/"
        )
    git_run("push", "origin", "HEAD:main", config=config, timeout=NET_TIMEOUT * 2)
    log("  Pushed successfully.")


def push_to_github(topic: str, filename: str) -> None:
    log("Staging changes...")
    git_run("add", f"public/pdfs/{filename}", "public/index.html")

    if git_run("diff", "--cached", "--quiet", check=False).returncode != 0:
        log("Committing...")
        git_run("commit", "-m", f"Add Garamut reflection: {topic}")

    git_push()


# ── Main pipeline ─────────────────────────────────────────────────────────────

def process_message(uid, msg, dry_run: bool) -> bool:
    """Process one email. Returns True if a new reflection was published."""
    subject = decode_subject(msg)
    log(f"\nProcessing: {subject!r}")

    issue, topic = parse_subject(subject)
    if issue is None:
        log("  Could not parse issue number from subject -- skipping.")
        return False

    pdf_bytes, orig_fname = download_pdf_attachment(msg)
    if pdf_bytes is None:
        log("  No PDF attachment found -- skipping.")
        return False

    sent = email_date(msg)
    year = sent.year

    # Published already (e.g. the same issue was emailed twice)
    if entry_exists(committed_catalogue(), year, issue):
        log(f"  Issue {year}/No.{issue} already published -- skipping.")
        return False

    catalogue = load_catalogue(INDEX_HTML.read_text(encoding="utf-8"))
    leftover = next(
        (e for e in catalogue if e["year"] == year and e["issue"] == issue), None
    )
    if leftover:
        # An earlier run wrote this entry into index.html but was stopped before
        # committing it. Keep the entry as it stands and finish the job.
        filename = Path(leftover["drive_url"]).name
        pdf_path = PDF_DIR / filename
        log(f"  Issue: No.{issue}  |  finishing a publish that was interrupted")
        if dry_run:
            log("  [DRY RUN] Would commit and push the existing entry.")
            return False
        if not pdf_path.exists():
            pdf_path.write_bytes(pdf_bytes)
            log(f"  Saved PDF -> {pdf_path.relative_to(REPO_ROOT)}")
        push_to_github(leftover["topic"], filename)
        return True

    filename = clean_filename(topic, year, issue)
    pdf_path = PDF_DIR / filename

    log(f"  Issue: No.{issue}  |  Topic: {topic}")
    log(f"  File:  {filename}  ({len(pdf_bytes):,} bytes)")

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
        "date": date_as_str(sent),
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

    failures = 0

    # A run stopped between commit and push leaves work that never reached GitHub
    if not args.dry_run and unpushed_commits():
        log("Found commit(s) from an earlier run that were never pushed.")
        try:
            git_push()
        except RuntimeError as exc:
            failures += 1
            log(f"  ERROR: {exc}")

    imap = connect_gmail()
    try:
        messages = fetch_unpublished_emails(imap)
        if not messages:
            log("No new Garamut reflections found.")

        published = 0
        for uid, msg in messages:
            try:
                ok = process_message(uid, msg, dry_run=args.dry_run)
            except Exception:
                # One bad email must not stop the others; the next run retries it
                failures += 1
                log("  ERROR: this reflection was not published.")
                traceback.print_exc()
                continue
            if ok:
                published += 1
                mark_as_read(imap, uid)

        if messages:
            log(f"\nDone. {published} reflection(s) published.")
    finally:
        try:
            imap.logout()
        except Exception:
            pass

    if failures:
        sys.exit(1)


if __name__ == "__main__":
    main()
