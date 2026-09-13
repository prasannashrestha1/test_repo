#!/usr/bin/env python3
"""run_scheduled_report.py — entry point for the twice-monthly Windows Task
Scheduler job (1st and 15th). Guards on compute_period's reporting-day check,
then runs generate_report.py's real pipeline against PRODUCTION Bubble for
every office in offices.json, using the computed period window."""
import json
import os
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from compute_period import compute_period
from generate_report import generate_all_reports


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
    return 0 if summary["errors"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
