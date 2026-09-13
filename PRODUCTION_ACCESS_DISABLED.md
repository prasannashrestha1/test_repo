# Production Bubble access is intentionally disabled

**As of 2026-09-11**, `config.json` no longer contains `bubble_base_url` (the
production/live Bubble root). **As of 2026-09-13**, that's backed by a second,
independent gate: `BUBBLE_ALLOW_PRODUCTION`, which must be explicitly set
before a production URL is even considered.

This is on purpose, not by accident — this pipeline is still actively being
reworked (the Listings-snapshot object is being replaced by a per-match join
against `property`'s `malcolm_listing_state_option_os_malcolm_listing_state`
field, still being validated), and until that settles, nothing in this folder
should be able to touch the live Quiet List data.

## What this means in practice

- **Every real run goes to test/staging, unconditionally**, until
  `BUBBLE_ALLOW_PRODUCTION` is explicitly set to a truthy value
  (`1`/`true`/`yes`/`on`) — regardless of whether `--use-test-version` was
  passed, and regardless of whether `BUBBLE_BASE_URL` happens to be set in
  someone's environment. A single stray `BUBBLE_BASE_URL` left over from
  testing something else can no longer send a run to production by itself.
- `python generate_report.py --mock ...` still works exactly as before —
  it never reads any URL at all.
- **The scheduled entry point** (`run_scheduled_report.py`, whether via the
  local Windows Task Scheduler jobs or a Claude Code routine) calls the
  pipeline with `use_test=False` — but that no longer matters while
  `BUBBLE_ALLOW_PRODUCTION` is unset, since the double-gate applies
  regardless of that argument. It will keep firing on the 1st/15th and
  successfully generating reports — against test/staging, not production —
  until production is deliberately enabled below.

## How to restore production access

Two separate, deliberate actions are required — not one:

```
BUBBLE_ALLOW_PRODUCTION=true
BUBBLE_BASE_URL=https://app.quietlist.com.au/api/1.1/obj
```

Set both as environment variables on whatever runs it. **Prefer this over
editing `config.json`.** Committing the production URL there puts it in git
history permanently and destroys the fail-safe for every future checkout —
setting environment variables instead keeps production reachable only where
someone has explicitly configured both of them, and leaves the repo itself
unable to touch live data on its own.

`--use-test-version` still forces the test root even when both of the above
are set, so a deliberate staging run always stays on staging no matter how
the surrounding environment is configured.
