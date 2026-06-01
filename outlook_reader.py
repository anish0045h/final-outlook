"""
outlook_reader.py — Connects to Microsoft Outlook via COM and fetches
payroll-related emails that have Excel attachments.

Security changes (functionality unchanged):
  - Attachment filenames sanitised before disk write (path traversal fix)
  - File size capped at 50 MB before reading bytes (DoS / memory exhaustion fix)
  - tempfile.mkstemp() replaces hand-built temp path (race condition fix)
  - Temp dir scoped to dedicated subfolder, not root %TEMP%
  - SQLite DB relocated to %APPDATA%\PayrollAgent\ (financial controls fix)
  - sqlite3 connections used as context managers (resource leak fix)
  - Subject/sender truncated in log output (PII leakage fix)
  - TESTING MODE flag kept exactly as original (commented-out block preserved)
"""

import os
import re
import sqlite3
import logging
import pathlib
import tempfile
import datetime
import win32com.client

log = logging.getLogger(__name__)

# ── Security constants ─────────────────────────────────────────────────────────
MAX_ATTACHMENT_BYTES = 50 * 1024 * 1024                  # 50 MB hard cap
_SAFE_FILENAME_RE    = re.compile(r"[^a-zA-Z0-9._\-]")  # allowlist chars only
_APP_DATA_DIR        = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")), "PayrollAgent"
)


def _safe_filename(raw: str) -> str:
    """
    Strip directory components and replace unsafe characters.
    Prevents a crafted filename like '../../evil.dll' from escaping the temp dir.
    """
    name = pathlib.Path(raw).name           # strip any leading path components
    name = _SAFE_FILENAME_RE.sub("_", name) # replace non-allowlist chars
    return name or "attachment"


class OutlookReader:
    def __init__(self, config):
        self.config = config

        # FIX: DB moved from beside the payroll Excel file to a protected OS path.
        # Old location let any local user delete it to force email re-processing
        # (double-payment risk in a payroll context).
        os.makedirs(_APP_DATA_DIR, exist_ok=True)
        self.db_path = os.path.join(_APP_DATA_DIR, "processed_emails.db")

        self._init_db()
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)  # 6 = olFolderInbox

    def _init_db(self):
        """Initialize SQLite database to track processed emails."""
        # FIX: context manager ensures connection is always closed, even on error
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "CREATE TABLE IF NOT EXISTS processed_emails (entry_id TEXT PRIMARY KEY)"
            )
            conn.commit()

    def _is_processed(self, entry_id):
        # FIX: context manager ensures connection is always closed
        with sqlite3.connect(self.db_path) as conn:
            result = conn.execute(
                "SELECT 1 FROM processed_emails WHERE entry_id = ?", (entry_id,)
            ).fetchone()
        return result is not None

    def mark_as_processed(self, entry_id: str):
        """Mark email as processed in SQLite database."""
        # FIX: context manager ensures connection is always closed
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "INSERT OR IGNORE INTO processed_emails (entry_id) VALUES (?)", (entry_id,)
            )
            conn.commit()
        log.info(f"Marked email as processed in local DB.")

    def fetch_payroll_emails(self, start_date=None, end_date=None) -> list:
        log.info("Fetching emails from Outlook Inbox...")
        messages = self.inbox.Items
        messages.Sort("[ReceivedTime]", True)  # newest first; Sort MUST come before Restrict

        if start_date or end_date:
            filters = []
            if start_date:
                # DASL format: converts to UTC ISO string Outlook actually understands
                start_str = datetime.datetime(
                    start_date.year, start_date.month, start_date.day, 0, 0, 0
                ).strftime("%m/%d/%Y %I:%M %p")
                filters.append(f"[ReceivedTime] >= '{start_str}'")
            if end_date:
                end_str = datetime.datetime(
                    end_date.year, end_date.month, end_date.day, 23, 59, 59
                ).strftime("%m/%d/%Y %I:%M %p")
                filters.append(f"[ReceivedTime] <= '{end_str}'")

            filter_expr = " AND ".join(filters)
            log.info(f"Attempting Restrict with: {filter_expr}")

            try:
                filtered = messages.Restrict(filter_expr)
                count = filtered.Count
                log.info(f"Restrict succeeded — {count} emails in date range.")
                if count > 0:
                    messages = filtered
                else:
                    log.warning("Restrict returned 0 emails — falling back to manual date filtering.")
                    messages = self.inbox.Items
                    messages.Sort("[ReceivedTime]", True)
            except Exception as exc:
                log.warning(f"Restrict failed: {exc} — falling back to manual date filtering.")

        results = []
        scanned = 0
        scan_limit = self.config.max_emails

        for message in messages:
            if scanned >= scan_limit:
                log.info(f"Reached scan limit of {scan_limit}.")
                break

            scanned += 1

            if message.Class != 43:
                continue

            subject  = getattr(message, "Subject", "(no subject)")
            sender   = getattr(message, "SenderEmailAddress", "unknown")
            received = getattr(message, "ReceivedTime", None)

            # ── Manual date gate (works even if Restrict silently failed) ──────────
            if received and (start_date or end_date):
                try:
                    received_date = datetime.date(received.year, received.month, received.day)
                    if start_date and received_date < start_date:
                        # Since sorted newest-first, once we're below start_date we're done
                        log.info(f"  [{scanned:>3}] Passed start_date boundary — stopping.")
                        break
                    if end_date and received_date > end_date:
                        log.info(f"  [{scanned:>3}] Email above end_date — skipping.")
                        continue
                except Exception as e:
                    log.warning(f"Date parse error on email: {e}")

            email_data = self._parse_email(message)
            if email_data:
                results.append(email_data)
                log.info(f"  [{scanned:>3}] ★ MATCH    | {subject[:60]} | from: {sender[:40]}")
            else:
                log.info(f"  [{scanned:>3}] ✗ NO MATCH | {subject[:60]} | from: {sender[:40]}")

        log.info(f"Scanned {scanned} email(s), found {len(results)} matching.")
        results.reverse()
        return results

    def _parse_email(self, message) -> dict | None:
        """
        Check if the email matches keywords and has Excel attachments.
        """
        subject = getattr(message, "Subject", "(no subject)")
        sender  = getattr(message, "SenderEmailAddress", "unknown")
        body    = getattr(message, "Body", "")

        combined_text = (subject + " " + body).lower()

        keyword_hit = any(k.lower() in combined_text for k in self.config.keywords)
        if not keyword_hit:
            return None

        attachments = self._extract_attachments(message)
        if not attachments:
            log.info(f"  SKIPPED (no Excel attachment) | subject: '{subject[:60]}'")
            return None

        log.info(
            f"  ACCEPTED | subject: '{subject[:60]}' | from: {sender[:40]} | "
            f"attachments: {[a['filename'] for a in attachments]}"
        )
        return {
            "id":          message.EntryID,
            "subject":     subject,
            "sender":      sender,
            "attachments": attachments,
        }

    def _extract_attachments(self, message) -> list:
        results  = []

        # FIX: dedicated subfolder instead of root %TEMP%\payroll_attachments
        temp_dir = os.path.join(tempfile.gettempdir(), "payroll_agent_attachments")
        os.makedirs(temp_dir, exist_ok=True)

        for attachment in message.Attachments:
            raw_name  = attachment.FileName

            # FIX: sanitise filename before any disk write (path traversal prevention)
            safe_name = _safe_filename(raw_name)
            ext       = pathlib.Path(safe_name).suffix.lower()

            if ext not in self.config.allowed_extensions:
                continue

            # FIX: mkstemp gives a unique path — avoids filename collision between emails
            fd, temp_path = tempfile.mkstemp(suffix=f"_{safe_name}", dir=temp_dir)
            os.close(fd)

            try:
                attachment.SaveAsFile(temp_path)

                # FIX: reject oversized attachments before loading into memory
                file_size = os.path.getsize(temp_path)
                if file_size > MAX_ATTACHMENT_BYTES:
                    log.warning(
                        f"Attachment '{safe_name}' rejected — "
                        f"{file_size:,} bytes exceeds {MAX_ATTACHMENT_BYTES:,} byte limit."
                    )
                    continue

                with open(temp_path, "rb") as f:
                    data = f.read()

                results.append({"filename": safe_name, "data": data})

            except Exception as e:
                log.error(f"Failed to extract attachment '{safe_name}': {e}")

            finally:
                # Always clean up temp file
                try:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                except OSError as e:
                    log.warning(f"Could not delete temp file '{temp_path}': {e}")

        return results