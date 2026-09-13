#!/usr/bin/env python3
"""
compute_period.py — the date guard + window logic for the Quiet List reports.

Exit code 1 (and prints a one-line reason) on any day that isn't the 1st or
15th — nothing else in the skill should run past that. On a valid day, prints
a JSON object with the four ISO dates generate_report.py needs.

Window logic:
  run on the 1st  -> period = previous month's 16th through its last day
                     (EOM); comparison period = previous month's 1st-15th
  run on the 15th -> period = this month's 1st through 15th;
                     comparison period = previous month's 16th-EOM
"""
import json
import sys
from datetime import datetime, timedelta


def compute_period(today: datetime) -> dict:
    if today.day == 1:
        first_of_this_month = today.replace(day=1)
        last_day_prev_month = first_of_this_month - timedelta(days=1)
        period_start = last_day_prev_month.replace(day=16)
        period_end = last_day_prev_month
        prev_period_start = last_day_prev_month.replace(day=1)
        prev_period_end = last_day_prev_month.replace(day=15)
    elif today.day == 15:
        period_start = today.replace(day=1)
        period_end = today.replace(day=15)
        first_of_this_month = today.replace(day=1)
        last_day_prev_month = first_of_this_month - timedelta(days=1)
        prev_period_start = last_day_prev_month.replace(day=16)
        prev_period_end = last_day_prev_month
    else:
        return None

    return {
        "period_start": period_start.strftime("%Y-%m-%d"),
        "period_end": period_end.strftime("%Y-%m-%d"),
        "prev_period_start": prev_period_start.strftime("%Y-%m-%d"),
        "prev_period_end": prev_period_end.strftime("%Y-%m-%d"),
    }


if __name__ == "__main__":
    today = datetime.now()
    result = compute_period(today)
    if result is None:
        print(f"Not a reporting day (today is the {today.day}); skipping.")
        sys.exit(1)
    print(json.dumps(result))
