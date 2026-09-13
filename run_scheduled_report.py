#!/usr/bin/env python3
"""run_scheduled_report.py — entry point for the twice-monthly scheduled run
(a Windows Task Scheduler job locally, or a Claude Code routine in the
cloud). Guards on compute_period's reporting-day check, then runs
generate_report.py's real pipeline for every office in offices.json, using
the computed period window. Which Bubble environment that hits is whatever
BUBBLE_BASE_URL resolves to (see generate_report.py's resolve_base_url) —
this script doesn't hard-code production.

Every generated PDF is then handed to drive_upload.upload_reports() and
email_report.send_report_email(), since a cloud routine's filesystem does not
persist between runs — a PDF that only lands in ./reports would otherwise be
silently lost. Both are no-ops (not errors) when their credentials aren't
configured, so local testing of this same entry point (as used throughout
development) still works with no Drive or email access at all."""
import json
import os
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from compute_period import compute_period
from generate_report import generate_all_reports
from drive_upload import upload_reports
from email_report import send_report_email


def period_dict(window: dict) -> dict:
    from_date = datetime.strptime(window["period_start"], "%Y-%m-%d")
    to_date = datetime.strptime(window["period_end"], "%Y-%m-%d")
    return {
        "label": "FORTNIGHTLY ACTIVITY REPORT",
        "from_date": window["period_start"],
        "to_date": window["period_end"],
        "prev_from_date": window["prev_period_start"],
        "prev_to_date": window["prev_period_end"],
        "range_label": f"{from_date.strftime('%d %B').upper()} - {to_date.strftime('%d %B %Y').upper()}",
        "matched_listings_label": f"{from_date.strftime('%d/%m/%y')} - {to_date.strftime('%d/%m/%y')}",
        "file_suffix": to_date.strftime("%Y%m%d"),
    }


def main() -> int:
    today = datetime.now()
    print(f"[{today.isoformat()}] run_scheduled_report starting")

    window = compute_period(today)
    if window is None:
        print(f"Not a reporting day (today is the {today.day}); skipping.")
        return 0

    with open(os.path.join(SCRIPT_DIR, "config.json")) as f:
        config = json.load(f)
    with open(os.path.join(SCRIPT_DIR, "offices.json")) as f:
        offices = json.load(f)

    period = period_dict(window)
    summary = generate_all_reports(offices, period, config, mock=False, use_test=False)

    pdf_paths = [s["path"] for s in summary.get("successes", [])]

    upload_results = upload_reports(pdf_paths, config.get("drive_reports_folder_id"))
    upload_failures = [r for r in upload_results if r["status"] == "failed"]

    to_addrs = recipients_from_config(config)
    email_result = send_report_email(pdf_paths, to_addrs)
    email_failed = email_result["status"] == "failed"

    # A PDF that generated but wasn't actually delivered anywhere is not
    # meaningfully "done" — on a cloud routine it will vanish along with the
    # session's filesystem — so delivery failures count the same as a
    # generation error would.
    total_errors = summary["errors"] + len(upload_failures) + (1 if email_failed else 0)
    if upload_failures:
        print(f"{len(upload_failures)} report(s) generated but failed to upload to Drive")
    if email_failed:
        print("report email failed to send")
    return 0 if total_errors == 0 else 1


def recipients_from_config(config: dict) -> list:
    """REPORT_EMAIL_TO (comma-separated) overrides config.json's
    report_email_recipients — a quick way to redirect a test run to a
    different inbox without editing the file."""
    env_override = (os.environ.get("REPORT_EMAIL_TO") or "").strip()
    if env_override:
        return [addr.strip() for addr in env_override.split(",") if addr.strip()]
    return config.get("report_email_recipients", [])


if __name__ == "__main__":
    sys.exit(main())
