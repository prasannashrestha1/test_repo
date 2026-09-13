# Property Matches Reporting — Schema & API Spec (v0.5)

Builds on v0.4. This version closes the `bubble_base_url` open item using the
real API endpoints the user provided on 2026-09-06, and confirms
`property_matches_object`'s exact real name/casing.

---

## 1. Real endpoints provided

- **Live/production:** `https://app.quietlist.com.au/api/1.1/obj/Property_Matches`
- **Test/staging version:** `https://app.quietlist.com.au/version-03jes/api/1.1/obj/Property_Matches`

Both point at the same object, `Property_Matches` (capitalized, underscore —
this also confirms `property_matches_object`'s exact casing, previously a
lowercase guess `"propertymatches"`).

The `version-03jes` segment is Bubble's named-version mechanism for a
development/test copy of the app, separate from the live production version.

## 2. How this pipeline uses the two roots

Per the user's explicit instruction: **production is the default/priority**,
and the **test version is only used for manual ad hoc runs**, never for the
scheduled task.

- `config.json` now has both:
  ```json
  "bubble_base_url": "https://app.quietlist.com.au/api/1.1/obj",
  "bubble_base_url_test": "https://app.quietlist.com.au/version-03jes/api/1.1/obj",
  "property_matches_object": "Property_Matches",
  ```
- `generate_report.py` gained a `--use-test-version` flag (default off). Off
  (the default — what the scheduled task always uses) resolves to
  `bubble_base_url` (production). Passing it resolves to
  `bubble_base_url_test` instead — for a manual run only.
- `BubbleClient` itself is unaffected — it just receives whichever base URL
  was resolved and builds `{base_url}/{object_name}` as before.

## 3. What's still open

Only the Listings-snapshot object remains, unchanged from v0.4:
- Its real Bubble internal object name (config.json currently guesses
  `"Listings"` — an updated guess following the naming convention
  `Property_Matches` confirmed, but NOT itself confirmed by a real sample or
  endpoint).
- A real sample response for it, showing field names for: office_id, a
  listing identifier matching `property_listing_id_text` (e.g. "1212"),
  address, suburb, property/dwelling type, price, status (the value meaning
  "current"), and whether it's a dated snapshot or a live mutable table.

Also still open, unchanged from v0.4:
- Confirm Sydney Sooth's Malcolm ID vs. Bubble `office_id` (both assumed "1").
- "Prepared for" contact name per office, and whether the Helpful Tip text is
  static across all offices or per-office.

## 4. `BUBBLE_API_TOKEN` — still not provided

Resolving the base URL does not unblock a live/test run by itself — the
actual Bubble API token has not been provided yet, so `config.json`'s
`bubble_token_drive_file_id` is still `null` and no Drive file holding it has
been created. See SKILL.md section 0 for the runtime-fetch mechanism this
will use once the token is supplied.
