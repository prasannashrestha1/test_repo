Quiet List Office Reports — local test package
================================================
Generated 2026-09-11

WHAT THIS IS
--------------
Quiet List (quietlist.com.au) is an Australian real estate matching platform,
built on Bubble, that connects buyer's agents with property listings. This
package is the automation that generates the fortnightly "Quiet List
Exchange Activity Report" — a per-office PDF summarizing buyer-brief matches
and current listing inventory — pulled from Bubble's Data API and saved to
Google Drive.

The end goal is a Claude scheduled task that runs automatically on the 1st
and 15th of each month, generating one report per office. It is currently a
SINGLE-OFFICE PILOT: only Sydney Sooth (Bubble office_id "1") is configured
in offices.json. The remaining 7 offices get added once their office_id and
Malcolm ID are confirmed.

The master/source-of-truth version of this automation lives as a Claude
skill ("quiet-list-office-reports") maintained in an ongoing Claude session,
not in this zip. This package is a point-in-time snapshot pulled out for
local testing only — if anything is changed or fixed while testing locally,
report it back rather than editing this copy as if it were canonical, so the
fix can be folded into the actual skill.

WHY THIS IS BEING RUN LOCALLY RIGHT NOW
------------------------------------------
The Claude cloud workspace this pipeline was built in cannot reach
app.quietlist.com.au directly — its outbound network is restricted to a
small allowlist (package registries, Anthropic's own services), and Bubble's
app isn't on it. Every attempt from that workspace is rejected before it
even reaches Bubble:

    HTTP/1.1 403 Forbidden
    request blocked: no rule or allowlist entry allows host "app.quietlist.com.au"

The real fix is an org Owner adding app.quietlist.com.au to that Claude
environment's Network access allowlist (Admin settings -> Cloud environments
-> the environment -> Network access -> Custom -> Allowed domains) — that's
in progress separately. Until it's confirmed working, running this on a
machine with normal internet access (this one) is how the real API call
actually gets tested end to end.

WHAT'S BEING TESTED RIGHT NOW
--------------------------------
A single-office reconciliation test: does the pipeline correctly fetch real
Property Matches rows from Bubble for office_id "1" and compute the right
report metrics from them? 50 synthetic Property Matches rows have already
been added directly into Bubble for this (property_matches_test_data.csv is
the exact data that was entered — kept here for reference/cross-checking,
not for re-uploading).

FILES IN THIS PACKAGE
-----------------------
generate_report.py                        - the report generator / CLI (the main script)
config.json                               - all endpoints, object names, field maps
commentary.py                             - narrative bullet generation (pure functions)
mock_data.py                              - synthetic data used only by --mock runs
template.html                             - Jinja2 template, rendered to PDF via Playwright
offices.json                              - office list (currently just Sydney Sooth, office_id "1")
compute_period.py                         - date-guard / period-window logic (1st & 15th only)
property_matches_schema_api_spec_v0.5.md  - schema/API notes: what's confirmed, what isn't
property_matches_test_data.csv            - the 50 synthetic rows already added to Bubble,
                                             for reference/cross-checking during this test

SETUP (one time)
-----------------
1. Python 3.9+.
2. pip install jinja2 playwright requests
   playwright install chromium
3. Get the real Bubble API token from the private Google Doc named
   "quiet-list-bubble-api-token" (Drive root) and set it as an environment
   variable in your terminal session — never hard-code it into any file,
   commit it, or paste it into Slack/tickets:
     Windows (PowerShell):  $env:BUBBLE_API_TOKEN = "paste-token-here"
     Windows (cmd.exe):      set BUBBLE_API_TOKEN=paste-token-here
     macOS/Linux:            export BUBBLE_API_TOKEN=paste-token-here

RUNNING THE RECONCILIATION TEST (office 1, TEST/staging Bubble version)
--------------------------------------------------------------------------
Matches the 50 test rows in property_matches_test_data.csv, dated
2026-08-20 through 2026-09-04:

    python generate_report.py --use-test-version --only-office 1 \
        --period-start 2026-08-19 --period-end 2026-09-06 \
        --prev-period-start 2026-07-01 --prev-period-end 2026-07-31

--use-test-version hits the TEST/staging Bubble root
(https://app.quietlist.com.au/version-03jes/api/1.1/obj) — never production.
Only drop that flag for a deliberate real production run once everything
here is confirmed working. Output PDF lands in
./reports/Sydney_Sooth_<date>.pdf, plus a console summary line:
"N offices, M reports, K errors" (and a FAILED line per office if any office
errors out — the whole batch doesn't stop on one failure).

MOCK RUN (no token, no network — sanity-checks the PDF pipeline only)
------------------------------------------------------------------------
    python generate_report.py --mock --only-office 1

KNOWN OPEN ITEM
-----------------
The Listings-snapshot object (a separate Bubble table, used only for "Total
Current Status Listings" and inventory counts — independent of Property
Matches) still uses GUESSED object/field names ("Listings"; office_id,
listing_id, address, suburb, property_types, price, status, snapshot_date).
None of this has been confirmed against real Bubble data yet. If this run
fails on that step rather than on the Property Matches fetch, that is the
likely reason — not a token problem and not a Property Matches problem.
Whoever runs this should report back exactly which step failed and the
error text, rather than guessing at a fix for the Listings object.

NEVER commit BUBBLE_API_TOKEN to any file, script, or version control.
