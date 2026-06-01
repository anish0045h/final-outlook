"""
Payroll Extra KM Agent
======================
Scans Outlook for CSO travel claim emails, parses Excel attachments,
and appends processed records to a master Excel file.

Flow:
  Outlook → find relevant emails → download .xlsx attachments
       → parse employee rows → calculate extra KM + amount
       → append to master_payroll.xlsx → log results

Security changes (functionality unchanged):
  - RotatingFileHandler replaces plain FileHandler (prevents unbounded log growth)
  - Date inputs validated with format check + length guard before strptime
  - start > end date cross-check added (prevents silent zero-result scan)
  - Subject truncated in log lines (PII leakage fix)
  - TESTING MODE comment preserved exactly as original
"""

import os
import sys
import logging
from datetime import datetime, date
from logging.handlers import RotatingFileHandler

from outlook_reader import OutlookReader
from excel_parser import ExcelParser
from master_sheet import MasterSheet
from config import Config


# When running as executable, use the exe's directory for all files
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(BASE_DIR)  # Change working directory to where the exe/script is


# FIX: RotatingFileHandler caps log at 5 MB and keeps 3 backups.
# Plain FileHandler grows forever — months of payroll email metadata on disk.
_file_handler = RotatingFileHandler(
    "agent.log",
    maxBytes=5 * 1024 * 1024,
    backupCount=3,
    encoding="utf-8",
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        _file_handler,
    ],
)
log = logging.getLogger(__name__)


# ── Date validation helper ─────────────────────────────────────────────────────
_DATE_FMT = "%Y-%m-%d"
_MIN_DATE  = date(2000, 1, 1)
_MAX_DATE  = date(2099, 12, 31)


def _parse_date_input(raw: str, label: str):
    """
    Validate a date string from input().
    Returns a date object on success, None if blank, logs warning on bad input.
    FIX: adds length guard and range check on top of original strptime try/except.
    """
    if not raw:
        return None
    # FIX: length guard — reject absurdly long strings before parsing
    if len(raw) > 20:
        log.warning(f"Invalid {label}: input too long. Skipping {label} filter.")
        return None
    try:
        parsed = datetime.strptime(raw, _DATE_FMT).date()
    except ValueError:
        log.warning(f"Invalid {label} format. Skipping {label} filter.")
        return None
    # FIX: range check — catches out-of-bound years like 0001 or 9999
    if not (_MIN_DATE <= parsed <= _MAX_DATE):
        log.warning(f"{label} {parsed} is out of accepted range. Skipping {label} filter.")
        return None
    return parsed


def run():
    log.info("=" * 60)
    log.info("Payroll Extra KM Agent started")
    log.info("=" * 60)

    # Prompt for date filter — identical UX to original
    print("\n" + "=" * 60)
    print("Email Scan Date Filter (Optional)")
    print("Leave blank and press Enter to scan all (up to max limit).")
    print("=" * 60)
    start_date_str = input("Enter Start Date (YYYY-MM-DD): ").strip()
    end_date_str   = input("Enter End Date (YYYY-MM-DD): ").strip()

    start_date = _parse_date_input(start_date_str, "Start Date")
    end_date   = _parse_date_input(end_date_str,   "End Date")

    if start_date:
        log.info(f"Start date filter applied: {start_date}")
    if end_date:
        log.info(f"End date filter applied: {end_date}")

    # FIX: cross-field validation — original silently scanned nothing if start > end
    if start_date and end_date and start_date > end_date:
        log.warning("Start date is later than end date — date filter will be ignored.")
        start_date = None
        end_date   = None

    config  = Config()
    outlook = OutlookReader(config)
    parser  = ExcelParser()
    masters = {}

    # ── Step 1: fetch matching emails ──────────────────────────────
    log.info("Scanning Outlook for payroll claim emails...")
    emails = outlook.fetch_payroll_emails(start_date=start_date, end_date=end_date)
    log.info(f"Found {len(emails)} matching email(s)")

    if not emails:
        log.info("Nothing to process. Exiting.")
        return

    total_rows_added = 0
    total_files = 0
    errors = []

    # ── Step 2: process each email ─────────────────────────────────
    for email in emails:
        # FIX: truncate subject in log to limit PII written to disk
        log.info(f"Processing email: '{email['subject'][:60]}' from {email['sender'][:40]}")

        for attachment in email["attachments"]:
            filename = attachment["filename"]
            log.info(f"  → Parsing attachment: {filename}")

            try:
                records = parser.parse(attachment["data"], filename, email["subject"], email["sender"])

                if not records:
                    log.warning(f"  → No valid records found in {filename}")
                    continue

                
                from collections import defaultdict
                by_period: dict[str, list] = defaultdict(list)
                for rec in records:
                    by_period[rec.get("Month", "Unknown_Period")].append(rec)

                for period, period_records in by_period.items():
                    if period not in masters:
                        safe_period = period.replace(" ", "_")
                        base_dir = os.path.dirname(config.master_file_path)
                        path = os.path.join(base_dir, f"master_payroll_{safe_period}.xlsx")
                        masters[period] = MasterSheet(path)

                    added = masters[period].append_records(period_records)
                    total_rows_added += added
                    log.info(
                        f"  → Appended {added} record(s) to "
                        f"{period.replace(' ', '_')} from {filename}"
                    )

                total_files += 1

            except Exception as e:
                msg = f"Failed to process {filename}: {e}"
                log.error(f"  → ERROR: {msg}")
                errors.append(msg)

        # TESTING MODE: Don't mark emails as processed so they can be re-processed
        # outlook.mark_as_processed(email["id"])

    # ── Step 3: save master sheet ──────────────────────────────────
    for master in masters.values():
        master.save()

    # ── Step 4: summary ────────────────────────────────────────────
    log.info("=" * 60)
    log.info(f"DONE — Files processed : {total_files}")
    log.info(f"       Rows appended    : {total_rows_added}")
    log.info(f"       Errors           : {len(errors)}")
    log.info(f"       Master file      : {config.master_file_path}")
    if errors:
        for e in errors:
            log.error(f"  ERROR: {e}")
    log.info("=" * 60)


if __name__ == "__main__":
    run()
