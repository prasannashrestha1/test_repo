# Production Bubble access is intentionally disabled

**As of 2026-09-11**, `config.json` no longer contains `bubble_base_url` (the
production/live Bubble root). This was removed on purpose, not by accident —
this pipeline is still actively being reworked (the Listings-snapshot object
is being replaced by a per-match join against `property`'s
`malcolm_listing_state_option_os_malcolm_listing_state` field, still being
validated), and until that settles, nothing in this folder should be able to
touch the live Quiet List data.

## What this means in practice

- `generate_report.py` **without** `--use-test-version` and **without**
  `--mock` will now fail immediately with:
  ```
  config.json is missing "bubble_base_url".
  ```
  instead of silently hitting production. This is deliberate fail-safe
  behavior, not a bug.
- `python generate_report.py --use-test-version ...` still works exactly as
  before — it only ever reads `bubble_base_url_test`, which is untouched.
- `python generate_report.py --mock ...` still works exactly as before too —
  it never reads either URL.
- **The two Windows scheduled tasks** (`QuietList_Report_Day1`,
  `QuietList_Report_Day15`, via `run_scheduled_report.py`) call the pipeline
  in production mode (`use_test=False`). They will still fire on the 1st/15th
  as scheduled, but will now just log the "missing bubble_base_url" error to
  `logs/scheduler.log` (or `logs/scheduled_run.log` for the local `.bat`
  version) and exit — **they cannot reach production while this file is
  missing the key.** No PDF gets generated on those days until this is
  restored.

## How to restore production access

Once the pipeline is confirmed working end-to-end against
`--use-test-version` and everyone is ready for real scheduled runs, add the
key back to `config.json`:

```json
"bubble_base_url": "https://app.quietlist.com.au/api/1.1/obj",
```

(Right alongside the existing `"bubble_base_url_test"` line.) That's the only
change needed — no code changes required, since `generate_report.py` already
reads this key normally when present.
