"""
config.py — All settings in one place.

Security hardening applied:
  - Settings can be overridden via a config.json file placed next to the EXE,
    removing the need to rebuild the EXE every time a value changes and
    eliminating hardcoded paths from the binary (reduces SAST findings).
  - master_file_path still defaults to BASE_DIR for backwards compatibility
    if no config.json is present.
  - No credentials, tokens, or secrets are stored here.
"""

import os
import sys
import json
import logging

log = logging.getLogger(__name__)

# ── Base directory ─────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

_CONFIG_FILE = os.path.join(BASE_DIR, "config.json")


def _load_external_config() -> dict:
    """
    Load optional config.json from BASE_DIR.
    Returns an empty dict if the file is absent or malformed.
    Any key present in the JSON overrides the class default.
    """
    if not os.path.exists(_CONFIG_FILE):
        return {}
    try:
        with open(_CONFIG_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if not isinstance(data, dict):
            raise ValueError("config.json must be a JSON object")
        log.info(f"Loaded external config from {_CONFIG_FILE}")
        return data
    except Exception as exc:
        log.warning(f"Could not read config.json — using defaults. Reason: {exc}")
        return {}


class Config:

    def __init__(self):
        ext = _load_external_config()

        # ── Email filtering ────────────────────────────────────────────────────
        self.keywords: list = ext.get("keywords", [
            "extra km",
            "extra kms",
            "Extra travelled KM",
            "extra travelled km",
            "extra Travelled KM",
            "Extra KM and Approval",
            "Extra KM and approval",
            "exception km",
            "travelled km",
            "travel claim",
            "CSO extra KM",
            "cso extra km",
            "extra Km expenses",
            "Extra km expenses",
            "Extra Kms claim",
            "extra Kms claim",
            "Extra Kms",
            "extra Kms",
            "extra km claim",
            "Extra km claim",
            "KMs claim",
            
        ])

        self.allowed_extensions: list = ext.get(
            "allowed_extensions", [".xlsx", ".xls"]
        )

        self.max_emails: int = int(ext.get("max_emails", 9999))

        # ── Excel parsing ──────────────────────────────────────────────────────
        self.sheet_name: str        = ext.get("sheet_name", "Present absent")
        self.data_start_row: int    = int(ext.get("data_start_row", 6))
        self.col_employee_id: int   = int(ext.get("col_employee_id", 1))
        self.col_employee_name: int = int(ext.get("col_employee_name", 2))
        self.col_designation: int   = int(ext.get("col_designation", 3))
        self.col_ase_manager: int   = int(ext.get("col_ase_manager", 4))
        self.col_asm_manager: int   = int(ext.get("col_asm_manager", 5))
        self.col_daily_km_start: int = int(ext.get("col_daily_km_start", 6))
        self.col_daily_km_end: int  = int(ext.get("col_daily_km_end", 36))
        self.col_extra_km: int      = int(ext.get("col_extra_km", 37))
        self.col_sm_approval: int   = int(ext.get("col_sm_approval", 38))
        self.col_amount: int        = int(ext.get("col_amount", 39))
        self.approval_value: str    = ext.get("approval_value", "Yes")
        self.rate_per_km: int       = int(ext.get("rate_per_km", 3))
        self.daily_limit: int       = int(ext.get("daily_limit", 50))

        # ── Output ────────────────────────────────────────────────────────────
        default_master = os.path.join(BASE_DIR, "master_payroll.xlsx")
        self.master_file_path: str  = ext.get("master_file_path", default_master)
        self.master_sheet_name: str = ext.get("master_sheet_name", "Payroll Records")

        self.master_columns: list = ext.get("master_columns", [
            "Employee ID",
            "Employee Name",
            "Designation",
            "ASE Manager",
            "ASM Manager",
            "Month",
            "Total Extra KM",
            "Amount INR",
            "Source File",
            "Email Subject",
            "Processed Date",
        ])
