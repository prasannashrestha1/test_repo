# Quiet List Exchange Activity Report

Generates the Quiet List Exchange Activity Report — one PDF per office —
from Bubble's Data API, on the 1st and 15th of each month.

Designed to run as a **Claude Code routine** against a cloud environment that
has the Bubble domain allowlisted. It also runs fine locally for development
and testing.

---

## Getting this onto GitHub and running as a routine

**1. Push to GitHub.** This repo has no remote configured yet. Create an
empty **private** repo on GitHub first (no README/license, so it doesn't
conflict with what's already here), then:

```bash
git remote add origin https://github.com/YOUR_ORG/YOUR_REPO.git
git push -u origin main
```

**2. Allowlist two domains** on the Claude Code cloud environment the routine
will use (Admin settings → Cloud environments → the environment → Network
access → Custom → Allowed domains):

| Domain | Why |
|---|---|
| `app.quietlist.com.au` | The Bubble Data API itself |
| `cdn.playwright.dev` | Chromium download for `playwright install` |

PyPI is allowlisted by default, so `pip install` needs no configuration — but
Playwright's *browser binary* comes from a separate host, and missing it
means setup fails before ever reaching Bubble.

**3. Set environment variables / secrets on the routine.** At minimum:

| Variable | Required? |
|---|---|
| `BUBBLE_API_TOKEN` | Yes |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Only if delivering via Drive |
| `SMTP_USERNAME` / `SMTP_PASSWORD` | Only if delivering via email (verified working — see Status) |
| `BUBBLE_ALLOW_PRODUCTION` + `BUBBLE_BASE_URL` | Only once ready for real (non-test) data — see [PRODUCTION_ACCESS_DISABLED.md](PRODUCTION_ACCESS_DISABLED.md) |

**4. Create the routine**, pointed at this repo, scheduled for the 1st and
15th, running:

```bash
pip install -r requirements.txt
playwright install chromium
python run_scheduled_report.py
```

That's the entire routine prompt — `run_scheduled_report.py` already handles
the date guard, report generation, and delivery (Drive + email) itself. See
"Running as a Claude Code routine" below for the full detail on each piece.

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

### Environment variables

Set `BUBBLE_API_TOKEN` on the cloud environment. It's never stored in the
repo, in `config.json`, or in the routine's prompt.

| Variable | Purpose |
|---|---|
| `BUBBLE_API_TOKEN` | The Bubble API token |

Real environment variables take precedence over `.env`, so no `.env` file
needs to exist in the cloud at all — that file is purely a local-development
convenience.

**Production is currently double-gated and off by default — every run,
including the routine, targets test/staging until this is deliberately
changed.** Two separate variables, both required together, are what unlock
production:

| Variable | Purpose |
|---|---|
| `BUBBLE_ALLOW_PRODUCTION` | Must be a truthy value (`true`/`1`/`yes`/`on`) or production is never reachable, full stop |
| `BUBBLE_BASE_URL` | Which Bubble environment to target, e.g. `https://app.quietlist.com.au/api/1.1/obj` — ignored unless the above is also set |

See [PRODUCTION_ACCESS_DISABLED.md](PRODUCTION_ACCESS_DISABLED.md) for the
full reasoning. Keeping the URL here rather than in `config.json` means the
production URL never has to be committed; requiring the second variable means
a stray `BUBBLE_BASE_URL` left over from something else can never silently
send a run to production by itself. Must be `https://` — the token is sent as
a Bearer header and plaintext HTTP is refused.

### What the routine should do

```bash
pip install -r requirements.txt
playwright install chromium
python run_scheduled_report.py
```

(with `BUBBLE_API_TOKEN` and `GOOGLE_SERVICE_ACCOUNT_JSON` already present in
the environment — this runs against test/staging until `BUBBLE_ALLOW_PRODUCTION`
and `BUBBLE_BASE_URL` are both also set)

`run_scheduled_report.py` is the scheduled entry point. It:

1. Checks whether today is the 1st or 15th via `compute_period.py`, exiting
   cleanly if not — so an off-cycle run can't produce a wrong-period report
2. Computes the correct reporting window and its comparison period
3. Generates one PDF per office in `offices.json`, into `reports/`
4. **Uploads each PDF to Drive** (`drive_upload.py`) **and emails it**
   (`email_report.py`), since the session's own filesystem does not persist
   between runs

A run that isn't on a reporting day logs
`Not a reporting day (today is the N); skipping.` and exits 0. That's the
date guard working, not a failure.

The routine's own prompt does not need to handle delivery itself — it's code,
not agent behavior, so it happens the same way on every run rather than
depending on the routine correctly remembering to do it each time.

## Uploading PDFs to Drive

Upload is a **Google Cloud service account**, not the Drive connector Claude
sessions normally use — a plain Python subprocess has no access to that
connector; it needs its own credential.

**One-time setup:**

1. In Google Cloud Console, create a project (or reuse one) and enable the
   **Google Drive API**.
2. Create a **service account**, then generate a **JSON key** for it.
3. Open the target Drive folder (`drive_reports_folder_id` in `config.json`)
   and **share** it with the service account's email address — it looks like
   `xxx@your-project.iam.gserviceaccount.com` — with **Editor** access.
4. Set `GOOGLE_SERVICE_ACCOUNT_JSON` to that key file's **raw JSON content**
   (not a file path) as an environment variable / secret on the routine.

Leaving this unset is fine — `generate_report.py` behaves exactly as before
(PDFs only in `reports/`), and `run_scheduled_report.py` logs a clear skip
message rather than failing. Nothing here is needed for `--mock` or
`--use-test-version` runs.

For a manual CLI run, upload is opt-in via `--upload-to-drive` (off by
default, so a `--use-test-version` reconciliation pass doesn't push synthetic
test PDFs into the client's real Drive folder). `run_scheduled_report.py`
always uploads and does not use this flag.

## Emailing PDFs

No third-party package and no paid service — `email_report.py` uses only
Python's own standard-library `email`/`smtplib` modules. It reads each PDF
and hands the bytes to those libraries, which handle MIME/base64 encoding
internally as part of building the email; that's normal, invisible plumbing
every email client does, not something that has to pass through an LLM's
context to happen (the way, say, a chat-based Drive upload would).

**One-time setup — a Gmail App Password, not your real password:**

1. Turn on **2-Step Verification** on the sending Google account, if not
   already on (required before App Passwords are available).
2. Google Account → Security → **App Passwords** → create one.
3. Set `SMTP_USERNAME` to that Gmail address and `SMTP_PASSWORD` to the
   generated App Password.
4. Set recipients: `report_email_recipients` (a list) in `config.json`, or
   `REPORT_EMAIL_TO` (comma-separated) as a quick override without editing
   that file.

Leaving `SMTP_USERNAME`/`SMTP_PASSWORD` unset is fine — `run_scheduled_report.py`
logs a clear skip message rather than failing. Nothing here is needed for
`--mock` or `--use-test-version` runs.

For a manual CLI run, email is opt-in via `--email-report` (with an optional
`--email-to` override), off by default for the same reason `--upload-to-drive`
is. `run_scheduled_report.py` always attempts to send and does not use either
flag.

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
| `run_scheduled_report.py` | Scheduled entry point (date guard + period window + delivery) |
| `drive_upload.py` | Uploads generated PDFs to Drive via a service account |
| `email_report.py` | Emails generated PDFs via SMTP (standard library only) |
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
- ✅ **Email delivery verified with a real send** — `email_report.py`
  successfully delivered a generated PDF via Gmail SMTP
- ⚠️ **Drive delivery is code-complete but not yet verified with a real
  upload** — every skip/failure path is tested, but no Google Cloud service
  account has been created yet to prove an actual upload
- ⚠️ **Production has never been exercised**
- ⚠️ `CURRENT` is the only `malcolm_listing_state` value observed — filtering
  against other states is unproven
- ⚠️ `offices.json` has 1 of 8 offices; the rest need their IDs confirmed
