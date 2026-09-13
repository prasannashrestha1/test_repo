# Quiet List Exchange Activity Report

Generates the Quiet List Exchange Activity Report — one PDF per office —
from Bubble's Data API, on the 1st and 15th of each month.

Designed to run as a **Claude Code routine** against a cloud environment that
has the Bubble domain allowlisted. It also runs fine locally for development
and testing.

---

## Quick start (local)

```bash
pip install -r requirements.txt
playwright install chromium

cp .env.example .env        # then paste in the real Bubble API token

# Synthetic data, no network, no token — proves the PDF pipeline works
python generate_report.py --mock --only-office 1

# Real data against the TEST/staging Bubble version
python generate_report.py --use-test-version --only-office 1 \
    --period-start 2026-08-19 --period-end 2026-09-06 \
    --prev-period-start 2026-07-01 --prev-period-end 2026-07-31
```

PDFs land in `reports/`.

---

## Running as a Claude Code routine

### Environment prerequisites

The cloud environment needs **both** of these on its allowed-domains list:

| Domain | Why |
|---|---|
| `app.quietlist.com.au` | The Bubble Data API itself |
| `cdn.playwright.dev` | Chromium download for `playwright install` |

> Package registries (PyPI) are allowlisted by default, so `pip install`
> works without configuration. The Playwright **browser binary** comes from a
> separate CDN host — miss that one and setup fails before it ever reaches
> Bubble.

### The token

`BUBBLE_API_TOKEN` is supplied **only** as an environment variable. It is
never stored in the repo, in `config.json`, or in the routine's prompt.

Set it as an environment variable / secret on the cloud environment. A real
environment variable always takes precedence over `.env`, so no `.env` file
needs to exist in the cloud at all — `.env` is purely a local-development
convenience.

### Connectors

Enable the **Google Drive** connector on the routine, so it can upload the
finished PDFs — cloud sessions are ephemeral and the filesystem does not
persist.

Connector traffic is proxied through Anthropic's servers rather than the
session's network, so it works regardless of the domain allowlist.

### What the routine should do

```bash
pip install -r requirements.txt
playwright install chromium
python run_scheduled_report.py
```

(with `BUBBLE_API_TOKEN` already present in the environment)

`run_scheduled_report.py` is the scheduled entry point. It:

1. Checks whether today is the 1st or 15th via `compute_period.py`, exiting
   cleanly if not — so an off-cycle run can't produce a wrong-period report
2. Computes the correct reporting window and its comparison period
3. Generates one PDF per office in `offices.json`, into `reports/`

The routine should then upload everything in `reports/` to the Drive folder
in `config.json` (`drive_reports_folder_id`), because the session's own
filesystem does not persist.

A run that isn't on a reporting day logs
`Not a reporting day (today is the N); skipping.` and exits 0. That's the
date guard working, not a failure.

---

## ⚠️ Production access is disabled

`config.json` deliberately has **no `bubble_base_url`**, so a real run fails
fast rather than silently reaching live data. Only `--use-test-version` and
`--mock` work as shipped.

See [PRODUCTION_ACCESS_DISABLED.md](PRODUCTION_ACCESS_DISABLED.md) for the
reasoning and the one-line change that restores it. Do not re-enable it until
a production dry-run has been checked against expected numbers.

---

## How the data fits together

Two Bubble objects, both read-only:

**`Property_Matches`** — one row per buyer-brief ↔ listing match. Carries its
own address, suburb, budget range, and buyer's-agent details, so most of the
report comes from here alone. Filtered by `office_id_text` and
`matched_date1_date`.

**`property`** — the listing records. Joined per match on
`property_listing_id_text` → `listing_id_text`, purely to read
`malcolm_listing_state_option_os_malcolm_listing_state` (the current listing
state, e.g. `CURRENT`).

### An important caveat about "Total Current Status Listings"

That metric — and the Performance Overview "Listings" column — count
**distinct current-status listings among those that were matched** in the
period. They are *not* total office inventory.

A listing that received zero matches never appears in `Property_Matches` at
all, so it cannot be counted. Total inventory independent of match activity
is not derivable from these two objects and would need a separate source.

This was a deliberate scope decision. The number's meaning on a
client-facing report changed accordingly.

---

## Files

| File | Purpose |
|---|---|
| `generate_report.py` | Main pipeline + CLI: fetch → compute → template → PDF |
| `run_scheduled_report.py` | Scheduled entry point (date guard + period window) |
| `compute_period.py` | 1st/15th guard and reporting-window arithmetic |
| `commentary.py` | Generates the narrative bullets |
| `template.html` | Jinja2 template, rendered to PDF via Playwright |
| `config.json` | Endpoints, object names, field maps |
| `offices.json` | Which offices get reports |
| `mock_data.py` | Synthetic data for `--mock` runs |
| `*_test_data.csv` | Synthetic rows loaded into Bubble's test version, for reconciliation |

## Status

- ✅ Validated end-to-end against TEST/staging: counts reconcile exactly
  against the test data (3 listings, 10 matches; Apartments 1/3, Houses 2/7)
- ⚠️ **Production has never been exercised**
- ⚠️ `CURRENT` is the only `malcolm_listing_state` value observed — filtering
  against other states is unproven
- ⚠️ `offices.json` has 1 of 8 offices; the rest need their IDs confirmed
