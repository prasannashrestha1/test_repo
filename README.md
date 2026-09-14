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

**2. Allowlist three domains** on the Claude Code cloud environment the routine
will use (Admin settings → Cloud environments → the environment → Network
access → Custom → Allowed domains):

| Domain | Why |
|---|---|
| `app.quietlist.com.au` | The Bubble Data API itself |
| `cdn.playwright.dev` | Chromium download for `playwright install` |
| `api.sendgrid.com` | Sending the report email (see Emailing reports below) |

PyPI is allowlisted by default, so `pip install` needs no configuration — but
Playwright's *browser binary* comes from a separate host, and missing it
means setup fails before ever reaching Bubble. Email uses SendGrid's HTTPS
API specifically *because* the routine's network sandbox is a domain
allowlist over HTTPS/443 only — plain SMTP (port 587) is unreachable from it
no matter what, so raw SMTP (e.g. a Gmail app password) will never work here.

**3. Set environment variables / secrets on the routine.** At minimum:

| Variable | Required? |
|---|---|
| `BUBBLE_API_TOKEN` | Yes |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Only if delivering via Drive |
| `SENDGRID_API_KEY` / `SENDGRID_FROM_EMAIL` | Only if delivering via email — see Emailing reports below |
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

The cloud environment needs **all three** of these on its allowed-domains list:

| Domain | Why |
|---|---|
| `app.quietlist.com.au` | The Bubble Data API itself |
| `cdn.playwright.dev` | Chromium download for `playwright install` |
| `api.sendgrid.com` | Sending the report email |

> Package registries (PyPI) are allowlisted by default, so `pip install`
> works without configuration. The Playwright **browser binary** comes from a
> separate CDN host — miss that one and setup fails before it ever reaches
> Bubble. Email specifically needs its own domain because the routine's
> sandbox only permits outbound HTTPS to allowlisted hosts, not arbitrary TCP
> — SMTP (port 587) never gets through, however it's configured. See
> "Emailing reports" below.

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
3. Generates one PDF **and one .docx** per office in `offices.json` (every
   office in the file, in one run — not just one), into `reports/`
4. **Uploads every file to Drive** (`drive_upload.py`) **and emails all of
   them in one message** (`email_report.py`), since the session's own
   filesystem does not persist between runs

A run that isn't on a reporting day logs
`Not a reporting day (today is the N); skipping.` and exits 0. That's the
date guard working, not a failure.

The routine's own prompt does not need to handle delivery itself — it's code,
not agent behavior, so it happens the same way on every run rather than
depending on the routine correctly remembering to do it each time.

### Testing the routine on a day that isn't the 1st or 15th

Set `FORCE_RUN_REPORT=true` as an environment variable / secret on the
routine to make it run anyway — this is an explicit, temporary opt-in for
testing only; it does not weaken the guard itself, which stays in place
and unset for every real scheduled run. With it set, the command is
unchanged (`python run_scheduled_report.py`) and behaves exactly like a real
run would (all offices, real Drive/email delivery), just using the period
window this month's 15th would compute. Remove it once you've confirmed the
routine works, so the next real 1st/15th run isn't skipped by mistake — it
wouldn't be (the guard still fires normally), but there's no reason to leave
a test-only override set permanently.

This still safely targets Bubble's test/staging root, same as any other run
— `FORCE_RUN_REPORT` only bypasses the date check, not the separate
`BUBBLE_ALLOW_PRODUCTION` + `BUBBLE_BASE_URL` gate production requires.

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

## Emailing reports

`email_report.py` sends via the **SendGrid** Web API (a plain HTTPS POST to
`api.sendgrid.com/v3/mail/send`, using `requests` — no new dependency), one
email per run with every generated file (all offices' PDFs and .docx files)
attached together.

**This replaced Gmail SMTP.** A Claude Code cloud routine's network sandbox
only allows outbound HTTPS to allowlisted domains, not arbitrary TCP — SMTP's
port 587 is unreachable from it no matter how it's configured (this was
confirmed directly: the SMTP host resolved and connected fine over IPv4,
authentication worked, and it still timed out at the port level from inside
the routine). It worked from a local machine because a local machine has no
such restriction. SendGrid's API needs only a domain on the allowlist
(`api.sendgrid.com`, see above), so the exact same call works in both places.

**One-time setup:**

1. Create a SendGrid account (sendgrid.com) — the free tier (100 emails/day)
   is enough for this.
2. Verify a sender under **Settings → Sender Authentication**: either
   **Single Sender Verification** (fastest — verify one address by clicking
   a link SendGrid emails to it) or full domain authentication (better
   deliverability, needs a few DNS records at your registrar). Whichever
   address ends up verified is what `SENDGRID_FROM_EMAIL` must be set to —
   SendGrid rejects sends from an unverified address.
3. Create an API key under **Settings → API Keys → Create API Key**.
   Restricted Access with only **Mail Send** enabled is enough; Full Access
   isn't needed.
4. Set `SENDGRID_API_KEY` (the key — a secret) and `SENDGRID_FROM_EMAIL`
   (the verified address) as environment variables / secrets — never in
   `config.json`, never committed.
5. Set recipients: `report_email_recipients` (a list) in `config.json`, or
   `REPORT_EMAIL_TO` (comma-separated) as a quick override without editing
   that file.

Leaving `SENDGRID_API_KEY`/`SENDGRID_FROM_EMAIL` unset is fine —
`run_scheduled_report.py` logs a clear skip message rather than failing.
Nothing here is needed for `--mock` or `--use-test-version` runs.

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
state, e.g. `CURRENT`). The fetch itself is also constrained to properties
with `status_option_os_property_status` = `Available` and
`backend_price_number` > `property_min_backend_price` (config.json, default
50000) — a matched listing whose property fails either isn't part of the
matching pool at all, so it's excluded from "Total Current Status Listings"
and the Performance Overview "Listings" column the same way a listing with no
match at all would be (see `fetch_property_status` in generate_report.py).

### An important caveat about "Total Current Status Listings"

That metric — and the Performance Overview "Listings" column — count
**distinct current-status listings among those that were matched** in the
period, and whose property record is `Available` and priced above the
configured floor. They are *not* total office inventory.

A listing that received zero matches never appears in `Property_Matches` at
all, so it cannot be counted. Total inventory independent of match activity
is not derivable from these two objects and would need a separate source.

This was a deliberate scope decision. The number's meaning on a
client-facing report changed accordingly.

---

## Files

| File | Purpose |
|---|---|
| `generate_report.py` | Main pipeline + CLI: fetch → compute → template → PDF + .docx |
| `run_scheduled_report.py` | Scheduled entry point (date guard + period window + delivery) |
| `drive_upload.py` | Uploads generated report files to Drive via a service account |
| `email_report.py` | Emails generated report files via the SendGrid API |
| `docx_report.py` | Renders the .docx report alongside the PDF |
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
- ✅ Both the PDF and .docx are generated per office and delivered together
  (Drive + email); confirmed the office loop runs all of `offices.json` in
  one pass, not just one office
- ⚠️ **Email delivery moved from Gmail SMTP to the SendGrid API** (SMTP's
  port 587 is unreachable from a Claude Code routine's network sandbox, so
  it could never work there) — code-complete and skip/failure paths tested,
  but no real SendGrid account has been set up yet to prove an actual send;
  the previous SMTP-based send *was* verified working, but only locally,
  which is exactly what the routine can't do
- ⚠️ **Drive delivery is code-complete but not yet verified with a real
  upload** — every skip/failure path is tested, but no Google Cloud service
  account has been created yet to prove an actual upload
- ⚠️ **Production has never been exercised**
- ⚠️ `CURRENT` is the only `malcolm_listing_state` value observed — filtering
  against other states is unproven
- ⚠️ `offices.json` has 1 of 8 offices; the rest need their IDs confirmed
