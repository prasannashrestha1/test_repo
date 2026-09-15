#!/usr/bin/env python3
"""
generate_report.py — Quiet List Exchange Activity Report generator.

For each office in offices.json: fetch Property Matches from Bubble's Data
API (paginated), then a single per-office join against the `property` object
for each matched listing's current status, compute every metric on the
report, generate commentary dynamically (commentary.py), fill template.html,
and render to PDF. One office failing does not stop the batch — errors are
collected and reported in a final summary line.

STATUS / OPEN ITEMS (see property_matches_schema_api_spec_v0.5.md):
  - Bubble internal object name + ALL field names for Property Matches are now
    confirmed against a real, fully-populated sample response (2026-09-06):
    office_id_text, property_listing_id_text, property_address_text,
    property_type_option_os_property_type, matched_date1_date, brief_title_text,
    suburbs_list_text (a list field), brief_min_price_number,
    brief_max_price_number, property_listing_price_number,
    buyers_agent_name_text, buyers_agency_name_text — see MATCH_FIELD_MAP below.
    Every canonical field the report needs now has a real source on this
    object directly; no more guessed/None mappings for Property Matches.
  - "Total Current Status Listings" and the Performance Overview listings
    column no longer come from a separate Listings-snapshot object (that
    object's name was never confirmed and turned out to 404). Instead,
    fetch_property_status/attach_listing_status (below) join each match's
    listing_id against the real `property` object's
    malcolm_listing_state_option_os_malcolm_listing_state field — confirmed
    2026-09-11 via a live GET against the TEST/staging root. This is a
    deliberate scope narrowing: these two numbers now mean "current status
    among listings that were matched," not "total office inventory
    regardless of match activity" — the latter is no longer derivable at
    all, since a listing with zero matches never appears here to begin with.
  - Which Bubble environment a run targets is resolved by resolve_base_url(),
    and production requires TWO deliberate signals, not one: BUBBLE_ALLOW_
    PRODUCTION must be truthy AND a production URL must be configured (via
    BUBBLE_BASE_URL or config.json's "bubble_base_url"). Until BUBBLE_ALLOW_
    PRODUCTION is set, every run — with or without --use-test-version —
    always uses the test/staging root, so a stray BUBBLE_BASE_URL left in an
    environment can never send a run to production by itself, and nobody has
    to remember to pass --use-test-version for that guarantee to hold.
    Non-HTTPS roots are refused outright, since the token travels as a
    Bearer header.
  - PRODUCTION ACCESS IS DISABLED BY DEFAULT (see PRODUCTION_ACCESS_DISABLED.md)
    — every run uses test/staging until BUBBLE_ALLOW_PRODUCTION is explicitly
    set, at which point BUBBLE_BASE_URL/config.json's "bubble_base_url" is
    used instead.
  - Performance Overview shows one row per raw property_type value that
    actually has a match or current listing this period (Bubble's
    property_type option set has ~19 possible values), rather than a fixed
    set of categories — a type with zero activity simply doesn't get a row.
  - Budget-range bucket width defaults to $200,000 — tune during the
    single-office reconciliation test if it doesn't land on Will's numbers.
    Note: price_min/price_max now come from the buyer's BRIEF budget range
    (brief_min_price_number/brief_max_price_number), not the listing price —
    property_listing_price_number is mapped separately as "listing_price" and
    used as the fallback when a brief has no min/max set.

USAGE
  Real run — currently ALWAYS goes to the TEST/staging root, whether or not
  --use-test-version is passed, because BUBBLE_ALLOW_PRODUCTION is unset. See
  PRODUCTION_ACCESS_DISABLED.md.
    export BUBBLE_API_TOKEN=xxxxx
    python3 generate_report.py \\
        --config config.json --offices-file offices.json \\
        --period-start 2026-07-01 --period-end 2026-07-31 \\
        --prev-period-start 2026-06-01 --prev-period-end 2026-06-30

  Same thing, spelled out explicitly with --use-test-version — behaves
  identically right now, but keeps working the same way once production is
  eventually enabled elsewhere:
    export BUBBLE_API_TOKEN=xxxxx
    python3 generate_report.py --use-test-version \\
        --config config.json --offices-file offices.json \\
        --period-start 2026-07-01 --period-end 2026-07-31 \\
        --prev-period-start 2026-06-01 --prev-period-end 2026-06-30

  Dry run with synthetic data (no Bubble access needed) — proves the whole
  pipeline end to end:
    python3 generate_report.py --mock

  Single-office test (e.g. the Bresic reconciliation pass):
    python3 generate_report.py --mock --only-office 5
"""

import argparse
import asyncio
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime, timedelta

import requests
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader
from playwright.async_api import async_playwright

from commentary import build_commentary
from docx_report import render_docx

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# BUBBLE_API_TOKEN is read from the environment. A real environment variable
# always wins over .env, which is what lets a cloud routine or CI job inject it
# as a secret with no file on disk, while local development keeps using .env.
# Loaded at import time so library callers (run_scheduled_report.py) get it too;
# --env-file can point at a different file, and overrides this.
DEFAULT_ENV_FILE = os.path.join(SCRIPT_DIR, ".env")
load_dotenv(DEFAULT_ENV_FILE)

# Maps our canonical field names (used by every compute_* function below) to
# the REAL Bubble field names on the property_matches object. ALL of these are
# now confirmed from a real, fully-populated sample API response on 2026-09-06:
#   {"_id": "...", "brief_title_text": "Brand New Day",
#    "brief_max_price_number": 700000, "brief_min_price_number": 500000,
#    "buyers_agency_name_text": "Unicorn Buyers", "buyers_agent_name_text": "Dan Solo",
#    "Created By": "...", "Created Date": "...",
#    "matched_date1_date": "2025-12-11T18:27:00.000Z", "Modified Date": "...",
#    "office_id_text": "1", "property_address_text": "29 Gwayne Street",
#    "property_listing_id_text": "1212", "property_listing_price_number": 600000,
#    "property_type_option_os_property_type": "House",
#    "suburbs_list_text": ["SURRY HILLS NSW 2010"]}
#
# Notes:
#  - suburbs_list_text is itself a Bubble list field (not a single string) —
#    normalize_match_record below handles both shapes via _as_list().
#  - price_min/price_max map to the BUYER'S BRIEF budget range
#    (brief_min/max_price_number), not the listing's own price. The listing's
#    actual price (property_listing_price_number) is mapped separately as
#    "listing_price" and used by compute_budget_range() as a fallback only
#    when a brief has no min/max set.
MATCH_FIELD_MAP = {
    "match_id": "_id",
    "office_id": "office_id_text",
    "listing_id": "property_listing_id_text",
    "address": "property_address_text",
    "property_type": "property_type_option_os_property_type",
    "match_date": "matched_date1_date",
    "brief_title": "brief_title_text",
    "suburb": "suburbs_list_text",             # CONFIRMED 2026-09-06 — list field
    "price_min": "brief_min_price_number",      # CONFIRMED 2026-09-06 — buyer brief budget
    "price_max": "brief_max_price_number",      # CONFIRMED 2026-09-06 — buyer brief budget
    "listing_price": "property_listing_price_number",  # CONFIRMED 2026-09-06 — actual listing price
    "ba_name": "buyers_agent_name_text",        # CONFIRMED 2026-09-06
    "ba_agency": "buyers_agency_name_text",     # CONFIRMED 2026-09-06
}

# Maps canonical names to real Bubble field names on the `property` object.
# Confirmed 2026-09-11 via a live GET against the TEST/staging root — this
# replaces the old, never-confirmed separate Listings-snapshot object.
# "Total Current Status Listings" and the Performance Overview listings
# column are now computed from a per-match join against this object (see
# fetch_property_status/attach_listing_status below), keyed on the same
# listing_id every match record already carries.
PROPERTY_FIELD_MAP = {
    "listing_id": "listing_id_text",
    "malcolm_listing_state": "malcolm_listing_state_option_os_malcolm_listing_state",
    "status": "status_option_os_property_status",
    "backend_price": "backend_price_number",
}


# --------------------------------------------------------------------------
# Bubble Data API client
# --------------------------------------------------------------------------

class BubbleClient:
    def __init__(self, base_url: str, api_token: str, page_size: int = 100, timeout: int = 30):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({"Authorization": f"Bearer {api_token}"})
        self.page_size = page_size
        self.timeout = timeout

    def fetch_all(self, object_name: str, constraints: list) -> list:
        """Fetch every row matching constraints, paging via cursor until exhausted."""
        results = []
        cursor = 0
        while True:
            params = {
                "constraints": json.dumps(constraints),
                "cursor": cursor,
                "limit": self.page_size,
            }
            resp = self.session.get(
                f"{self.base_url}/{object_name}", params=params, timeout=self.timeout
            )
            resp.raise_for_status()
            payload = resp.json()["response"]
            results.extend(payload.get("results", []))
            remaining = payload.get("remaining", 0)
            if remaining <= 0:
                break
            cursor += len(payload.get("results", []))
        return results


def _as_list(val):
    """Normalize a raw Bubble value that may already be a list (e.g. a Bubble
    "list of texts" field like suburbs_list_text) or a single scalar into a
    plain list, dropping empties. Used for fields where Bubble's own field
    type varies by object."""
    if isinstance(val, list):
        return [v for v in val if v]
    return [val] if val else []


def normalize_match_record(raw: dict, field_map: dict = MATCH_FIELD_MAP) -> dict:
    """Translate one raw Bubble property_matches row into the canonical shape
    every compute_* function below expects, using field_map. A canonical key
    whose field_map value is None (not yet confirmed against real Bubble
    field names) comes through as None/empty — every consumer below already
    tolerates that (renders N/A, "—", or skips the group) rather than crashing."""
    def get(canonical):
        raw_key = field_map.get(canonical)
        return raw.get(raw_key) if raw_key else None

    property_type = get("property_type")
    return {
        "match_id": get("match_id"),
        "office_id": get("office_id"),
        "listing_id": get("listing_id"),
        "address": get("address"),
        "property_types": [property_type] if property_type else [],
        "match_date": get("match_date"),
        "brief_title": get("brief_title"),
        "suburbs": _as_list(get("suburb")),
        "price_min": get("price_min"),
        "price_max": get("price_max"),
        "listing_price": get("listing_price"),
        "ba_name": get("ba_name"),
        "ba_agency": get("ba_agency"),
    }


def fetch_matches(client: BubbleClient, object_name: str, office_id: str,
                   field_map: dict, from_date: str, to_date: str) -> list:
    office_key = field_map["office_id"]
    date_key = field_map["match_date"]
    constraints = [
        {"key": office_key, "constraint_type": "equals", "value": office_id},
        {"key": date_key, "constraint_type": "greater than", "value": from_date},
        {"key": date_key, "constraint_type": "less than", "value": to_date},
    ]
    raw_rows = client.fetch_all(object_name, constraints)
    return [normalize_match_record(r, field_map) for r in raw_rows]


def fetch_property_status(client: BubbleClient, object_name: str, listing_ids: list,
                           field_map: dict = PROPERTY_FIELD_MAP,
                           required_status: str = "Available",
                           min_backend_price: float = 50_000) -> dict:
    """One GET against the `property` object for every listing_id appearing in
    this office's matches (Bubble's "in" constraint), returning
    {listing_id: malcolm_listing_state}. This is the confirmed replacement for
    the old separate Listings-snapshot fetch — status is now looked up per
    match rather than pulled from an independent inventory snapshot.

    Also constrains the fetch itself to properties with status "Available"
    and a backend_price_number above min_backend_price -- only properties
    meeting both belong in the matching pool at all. A matched listing whose
    property fails either constraint (e.g. sold, or a placeholder/test
    listing priced under $50k) simply won't come back here, so it's absent
    from status_lookup and attach_listing_status leaves its listing_status
    as None -- is_current_listing_status already treats that as "not
    current", so no separate exclusion logic is needed downstream.
    """
    unique_ids = sorted({lid for lid in listing_ids if lid})
    if not unique_ids:
        return {}
    listing_id_key = field_map["listing_id"]
    state_key = field_map["malcolm_listing_state"]
    constraints = [
        {"key": listing_id_key, "constraint_type": "in", "value": unique_ids},
        {"key": field_map["status"], "constraint_type": "equals", "value": required_status},
        {"key": field_map["backend_price"], "constraint_type": "greater than", "value": min_backend_price},
    ]
    raw_rows = client.fetch_all(object_name, constraints)
    return {row.get(listing_id_key): row.get(state_key) for row in raw_rows}


def attach_listing_status(matches: list, status_lookup: dict) -> None:
    for m in matches:
        m["listing_status"] = status_lookup.get(m.get("listing_id"))


# --------------------------------------------------------------------------
# Metric computation — pure functions, all testable without any network call
# --------------------------------------------------------------------------

def pct_change(current, previous):
    if not previous:
        return 0.0
    return (current - previous) / previous * 100


def first_property_type(match: dict) -> str:
    """property_types may be a single string or a list — normalize to one representative value."""
    val = match.get("property_types")
    if isinstance(val, list):
        return val[0] if val else ""
    return val or ""


def is_current_listing_status(status, current_status_value="Current") -> bool:
    return bool(status) and str(status).strip().lower() == str(current_status_value).strip().lower()


def count_current_status_listings(matches: list, current_status_value="Current") -> int:
    """Distinct listings among these matches whose linked `property` record
    (via fetch_property_status/attach_listing_status) is still in the
    "current" state. Replaces the old separate Listings-snapshot count —
    inherently limited to listings that appear in at least one match, which
    is the intentional scope now (see PRODUCTION_ACCESS_DISABLED.md/chat
    history: total office inventory independent of matches is out of scope)."""
    return len({
        m["listing_id"] for m in matches
        if m.get("listing_id") and is_current_listing_status(m.get("listing_status"), current_status_value)
    })


def compute_exec_snapshot(matches_current, matches_previous,
                           current_status_value="Current", unique_agents_field="ba_name"):
    listings_current = count_current_status_listings(matches_current, current_status_value)
    listings_previous = count_current_status_listings(matches_previous, current_status_value)
    matches_current_n = len(matches_current)
    matches_previous_n = len(matches_previous)

    unique_agents = len({m.get(unique_agents_field) for m in matches_current if m.get(unique_agents_field)})
    unique_agents_prev = len({m.get(unique_agents_field) for m in matches_previous if m.get(unique_agents_field)})

    listing_counts = Counter(m.get("listing_id") for m in matches_current if m.get("listing_id"))
    highest_count = max(listing_counts.values()) if listing_counts else 0
    listing_counts_prev = Counter(m.get("listing_id") for m in matches_previous if m.get("listing_id"))
    highest_count_prev = max(listing_counts_prev.values()) if listing_counts_prev else 0

    def row(label, cur, prev):
        change = pct_change(cur, prev)
        return {
            "label": label,
            "result": str(cur),
            "change": f"{'+' if change >= 0 else ''}{change:.1f}%",
            "positive": change >= 0,
        }

    return [
        row("Total Current Status Listings", listings_current, listings_previous),
        row("Total Buyer Brief Matches", matches_current_n, matches_previous_n),
        row("Unique Buyer Agents Active", unique_agents, unique_agents_prev),
        row("Highest Matched Listing", highest_count, highest_count_prev),
    ], listing_counts


def compute_performance_overview(matches_current, current_status_value="Current"):
    current_listings = {}
    for m in matches_current:
        lid = m.get("listing_id")
        if lid and lid not in current_listings and is_current_listing_status(m.get("listing_status"), current_status_value):
            current_listings[lid] = m
    listings_by_type = Counter(
        first_property_type(l) for l in current_listings.values() if first_property_type(l)
    )
    matches_by_type = Counter(
        first_property_type(m) for m in matches_current if first_property_type(m)
    )

    # One row per raw property_type value that actually has a match or
    # current listing this period -- not a fixed Apartments/Houses pair.
    # Bubble's property_type option set has ~19 possible values (Commercial,
    # Land, Retirement, Studio, Warehouse, ...); most won't appear in any
    # given period, so a type with zero activity simply gets no row rather
    # than showing up as an empty 0/0 line. Sorted by match count so the
    # table reads as a ranking.
    property_types = sorted(
        set(listings_by_type) | set(matches_by_type),
        key=lambda t: matches_by_type.get(t, 0),
        reverse=True,
    )
    return [
        {
            "property_type": t,
            "listings": str(listings_by_type.get(t, 0)),
            "matches": str(matches_by_type.get(t, 0)),
        }
        for t in property_types
    ]


def compute_top_suburbs(matches_current, top_n=5):
    def suburb_of(m):
        val = m.get("suburbs")
        if isinstance(val, list):
            return val[0] if val else "Unknown"
        return val or "Unknown"

    counts = Counter(suburb_of(m) for m in matches_current)
    return counts.most_common(top_n)


def compute_budget_range(matches_current, bucket_width=200_000):
    points = []
    for m in matches_current:
        pmin = m.get("price_min")
        pmax = m.get("price_max")
        if pmin is not None and pmax is not None:
            points.append((pmin + pmax) / 2)
        elif m.get("listing_price") is not None:
            points.append(m["listing_price"])

    if not points:
        return "N/A", 0.0

    buckets = Counter(int(p // bucket_width) for p in points)
    top_bucket, top_count = buckets.most_common(1)[0]
    low = top_bucket * bucket_width
    high = low + bucket_width
    share_pct = top_count / len(points) * 100

    def fmt(v):
        return f"${v/1_000_000:.1f}M" if v >= 1_000_000 else f"${v:,.0f}"

    return f"{fmt(low)}-{fmt(high)}", share_pct


def compute_dwelling_type_display(matches_current):
    counts = Counter(first_property_type(m) for m in matches_current if first_property_type(m))
    if not counts:
        return "N/A"
    return counts.most_common(1)[0][0]


def compute_top_names(matches_current, field, top_n=3):
    counts = Counter(m.get(field) for m in matches_current if m.get(field))
    return [name for name, _ in counts.most_common(top_n)]


def compute_featured_listing(matches_current, listing_counts):
    if not listing_counts:
        return None, 0, 0, 0
    ranked = listing_counts.most_common()
    top_listing_id, top_count = ranked[0]
    address = next(
        (m.get("address") for m in matches_current if m.get("listing_id") == top_listing_id),
        top_listing_id,
    )
    second_tier = [c for _, c in ranked[1:6]]  # next up-to-5 listings
    if second_tier:
        return address, top_count, min(second_tier), max(second_tier)
    return address, top_count, top_count, top_count


def compute_listing_pages(matches_current, rows_per_col=15):
    seen = {}
    for m in matches_current:
        lid = m.get("listing_id")
        if lid and lid not in seen:
            seen[lid] = {
                "address": m.get("address", ""),
                "suburb": (m.get("suburbs")[0] if isinstance(m.get("suburbs"), list) and m.get("suburbs") else m.get("suburbs") or ""),
            }
    addresses = [v["address"] for v in sorted(seen.values(), key=lambda r: (r["suburb"], r["address"]))]

    per_page = rows_per_col * 2
    pages = []
    for i in range(0, len(addresses), per_page):
        chunk = addresses[i:i + per_page]
        left = chunk[:rows_per_col]
        right = chunk[rows_per_col:]
        pages.append({"left": left, "right": right})
    return pages


# --------------------------------------------------------------------------
# Report assembly for a single office
# --------------------------------------------------------------------------

def build_report_context(office: dict, period: dict, matches_current: list, matches_previous: list,
                          current_status_value: str = "Current") -> dict:
    exec_snapshot, listing_counts = compute_exec_snapshot(
        matches_current, matches_previous, current_status_value
    )
    performance_overview = compute_performance_overview(matches_current, current_status_value)
    top_suburbs = compute_top_suburbs(matches_current)
    budget_range_display, budget_share_pct = compute_budget_range(matches_current)
    dwelling_type_display = compute_dwelling_type_display(matches_current)
    top_agencies = compute_top_names(matches_current, "ba_agency")
    top_operatives = compute_top_names(matches_current, "ba_name")
    featured_address, featured_count, second_min, second_max = compute_featured_listing(
        matches_current, listing_counts
    )
    listing_pages = compute_listing_pages(matches_current)

    listings_current_n = count_current_status_listings(matches_current, current_status_value)
    listings_previous_n = count_current_status_listings(matches_previous, current_status_value)

    # Only gate on what build_commentary's 5 bullets actually consume — none of
    # them use top_agencies/top_operatives, so a period with real matches but
    # no known buyer-agent data (a genuine gap right now, see MATCH_FIELD_MAP)
    # must still get real commentary, not a false "no matches" message.
    if matches_current and top_suburbs and featured_address:
        commentary = build_commentary({
            "office_name": office["office_name"],
            "listings_current": listings_current_n,
            "listings_previous": listings_previous_n,
            "matches_current": len(matches_current),
            "matches_previous": len(matches_previous),
            "top_suburbs": top_suburbs,
            "budget_range_display": budget_range_display,
            "budget_range_share_pct": budget_share_pct,
            "dwelling_type_display": dwelling_type_display,
            "featured_listing_address": featured_address,
            "featured_listing_count": featured_count,
            "second_tier_min": second_min,
            "second_tier_max": second_max,
            "performance_overview": performance_overview,
        })
    else:
        # zero-match office: render clean rather than crash on empty groupings
        commentary = [f"No buyer brief matches were recorded for {office['office_name']} this period."]

    top_suburb_display = top_suburbs[0][0].upper() if top_suburbs else "N/A"

    return {
        "office_name": office["office_name"],
        "reporting_period_label": period.get("label", "FORTNIGHTLY ACTIVITY REPORT"),
        "period_range_label": period["range_label"],
        "prepared_for": office.get("prepared_for", ""),
        "report_date_label": datetime.now().strftime("%d %B %Y").upper(),
        "helpful_tip": office.get(
            "helpful_tip",
            "Buyer's Agents must complete AML checks before a brief can enter the exchange.",
        ),
        "exec_snapshot": exec_snapshot,
        "performance_overview": performance_overview,
        "featured_listing_address": featured_address or "N/A",
        "featured_listing_caption": "Highest number of matches." if featured_address else "",
        "top_suburb_display": top_suburb_display,
        "budget_range_display": budget_range_display,
        "dwelling_type_display": dwelling_type_display,
        "top_agencies": (top_agencies + ["—"] * 3)[:3],
        "top_operatives": (top_operatives + ["—"] * 3)[:3],
        "commentary": commentary,
        "listing_pages": listing_pages,
        "matched_listings_period_label": period["matched_listings_label"],
        "footer_contact": "quietlist.com.au | info@quietlist.com.au",
        "logo_text": "QUIET LIST",
    }


async def render_pdf(context: dict, output_path: str):
    env = Environment(loader=FileSystemLoader(SCRIPT_DIR))
    tpl = env.get_template("template.html")
    html_out = tpl.render(**context)

    rendered_html_path = os.path.abspath(output_path.replace(".pdf", ".rendered.html"))
    with open(rendered_html_path, "w") as f:
        f.write(html_out)

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(f"file://{rendered_html_path}")
        await page.pdf(path=output_path, format="A4", print_background=True)
        await browser.close()


# --------------------------------------------------------------------------
# Per-office pipeline with error tolerance
# --------------------------------------------------------------------------

_FILENAME_UNSAFE_RE = re.compile(r'[<>:"/\\|?*]')


def _safe_filename(name: str) -> str:
    """Turn an office name into something safe to use in a file path on any
    OS. Spaces become underscores; Windows' reserved filename characters
    (< > : " / \\ | ? *) are stripped outright rather than just spaces --
    an office name containing one of these (e.g. "McGrath Collaroy | Dee
    Why") previously crashed report generation on Windows with
    "[Errno 22] Invalid argument", since only spaces were being replaced."""
    return _FILENAME_UNSAFE_RE.sub("", name).replace(" ", "_")


def generate_report_for_office(office: dict, period: dict, config: dict,
                                client: BubbleClient = None, mock: bool = False) -> dict:
    """Returns {"pdf_path": ..., "docx_path": ...}. Raises on failure — caller decides tolerance."""
    if mock:
        from mock_data import get_mock_data
        matches_current, matches_previous = get_mock_data(office["office_id"])
    else:
        field_map = config.get("match_field_map", MATCH_FIELD_MAP)
        matches_current = fetch_matches(
            client, config["property_matches_object"], office["office_id"],
            field_map, period["from_date"], period["to_date"],
        )
        matches_previous = fetch_matches(
            client, config["property_matches_object"], office["office_id"],
            field_map, period["prev_from_date"], period["prev_to_date"],
        )
        property_field_map = config.get("property_field_map", PROPERTY_FIELD_MAP)
        all_listing_ids = [m.get("listing_id") for m in matches_current + matches_previous]
        status_lookup = fetch_property_status(
            client, config["property_object"], all_listing_ids, property_field_map,
            required_status=config.get("property_required_status", "Available"),
            min_backend_price=config.get("property_min_backend_price", 50_000),
        )
        attach_listing_status(matches_current, status_lookup)
        attach_listing_status(matches_previous, status_lookup)

    context = build_report_context(
        office, period, matches_current, matches_previous,
        config.get("current_status_value", "Current"),
    )

    os.makedirs(config["output_dir"], exist_ok=True)
    safe_name = _safe_filename(office["office_name"])
    pdf_path = os.path.join(config["output_dir"], f"{safe_name}_{period['file_suffix']}.pdf")
    docx_path = os.path.join(config["output_dir"], f"{safe_name}_{period['file_suffix']}.docx")

    # PDF stays exactly as before — kept for future use, not being replaced.
    asyncio.run(render_pdf(context, pdf_path))
    # Word version, generated alongside from the same context dict, so the
    # two formats can never disagree on the numbers — only how they're laid out.
    render_docx(context, docx_path)

    return {"pdf_path": pdf_path, "docx_path": docx_path}


def _is_truthy_env(name: str) -> bool:
    return (os.environ.get(name) or "").strip().lower() in ("1", "true", "yes", "on")


def resolve_base_url(config: dict, use_test: bool) -> str:
    """Pick which Bubble API root to talk to.

    Production requires TWO deliberate, separate signals, not one:
      1. BUBBLE_ALLOW_PRODUCTION set to a truthy value
      2. a production URL, from BUBBLE_BASE_URL or config.json's "bubble_base_url"

    Until BUBBLE_ALLOW_PRODUCTION is set, this ALWAYS resolves to the test
    root — regardless of --use-test-version, regardless of whether
    BUBBLE_BASE_URL happens to be set in someone's environment. That is the
    point: a single stray env var can no longer send a run to production, and
    no flag needs to be remembered for every invocation to stay safely on
    test. --use-test-version still works as an explicit, self-documenting way
    to say "use test" in a script, but it is no longer the only thing
    standing between a run and production.

    See PRODUCTION_ACCESS_DISABLED.md for how to actually enable production
    when ready.
    """
    allow_production = _is_truthy_env("BUBBLE_ALLOW_PRODUCTION")

    if use_test or not allow_production:
        base_url = config.get("bubble_base_url_test")
        if not base_url:
            raise SystemExit('config.json has no "bubble_base_url_test" — nothing to run against.')
    else:
        base_url = (os.environ.get("BUBBLE_BASE_URL") or "").strip() or config.get("bubble_base_url")
        if not base_url:
            raise SystemExit(
                "BUBBLE_ALLOW_PRODUCTION is set, but no production URL is configured. Either:\n"
                "  - set BUBBLE_BASE_URL, or\n"
                '  - add "bubble_base_url" to config.json.\n'
                "See PRODUCTION_ACCESS_DISABLED.md."
            )

    # The token rides on every request as a Bearer header; plaintext HTTP would
    # put it on the wire in the clear.
    if not base_url.startswith("https://"):
        raise SystemExit(f"Refusing to send the API token over a non-HTTPS connection: {base_url}")

    return base_url


def generate_all_reports(offices: list, period: dict, config: dict, mock: bool = False,
                          only_office: str = None, use_test: bool = False) -> dict:
    if only_office:
        offices = [o for o in offices if str(o["office_id"]) == str(only_office)]

    client = None
    if not mock:
        # .strip() because a token pasted into .env or a secrets UI very often
        # picks up a trailing newline, which Bubble rejects as a bad token.
        token = (os.environ.get("BUBBLE_API_TOKEN") or "").strip()
        if not token:
            raise SystemExit(
                "BUBBLE_API_TOKEN is not set.\n"
                "  - Cloud routine / CI: set it as an environment variable or secret.\n"
                "  - Local: put BUBBLE_API_TOKEN=... in a .env file beside this "
                "script (see .env.example), or pass --env-file to use another file."
            )
        base_url = resolve_base_url(config, use_test)
        client = BubbleClient(base_url, token, config.get("page_size", 100))

    successes = []
    errors = []
    for office in offices:
        try:
            paths = generate_report_for_office(office, period, config, client=client, mock=mock)
            # "path" stays the PDF path for backward compatibility — drive_upload.py
            # and email_report.py already read successes[i]["path"] for delivery.
            successes.append({
                "office_name": office["office_name"],
                "path": paths["pdf_path"],
                "docx_path": paths["docx_path"],
            })
        except Exception as e:  # noqa: BLE001 — intentional: one office must not stop the batch
            errors.append({"office_name": office["office_name"], "error": str(e)})

    summary = {
        "total": len(offices),
        "reports_generated": len(successes),
        "errors": len(errors),
        "successes": successes,
        "error_details": errors,
    }
    print(f"{summary['total']} offices, {summary['reports_generated']} reports, {summary['errors']} errors")
    for e in errors:
        print(f"  FAILED: {e['office_name']} — {e['error']}")
    return summary


# --------------------------------------------------------------------------
# Period helpers (manual-testing convenience only — SKILL.md owns the real
# date-guard + window logic in production and should pass explicit dates)
# --------------------------------------------------------------------------

def default_period_for_today() -> dict:
    today = datetime.now()
    if today.day >= 16:
        from_date = today.replace(day=1)
        to_date = today
        prev_to = from_date - timedelta(days=1)
        prev_from = prev_to.replace(day=16)
    else:
        from_date = today.replace(day=1)
        to_date = today
        prev_month_end = from_date - timedelta(days=1)
        prev_from = prev_month_end.replace(day=1)
        prev_to = prev_month_end

    return {
        "label": "FORTNIGHTLY ACTIVITY REPORT",
        "from_date": from_date.strftime("%Y-%m-%d"),
        "to_date": to_date.strftime("%Y-%m-%d"),
        "prev_from_date": prev_from.strftime("%Y-%m-%d"),
        "prev_to_date": prev_to.strftime("%Y-%m-%d"),
        "range_label": f"{from_date.strftime('%d %B').upper()} - {to_date.strftime('%d %B %Y').upper()}",
        "matched_listings_label": f"{from_date.strftime('%d/%m/%y')} - {to_date.strftime('%d/%m/%y')}",
        "file_suffix": to_date.strftime("%Y%m%d"),
    }


def period_from_args(args) -> dict:
    if not args.period_start:
        return default_period_for_today()
    from_date = datetime.strptime(args.period_start, "%Y-%m-%d")
    to_date = datetime.strptime(args.period_end, "%Y-%m-%d")
    prev_from = datetime.strptime(args.prev_period_start, "%Y-%m-%d")
    prev_to = datetime.strptime(args.prev_period_end, "%Y-%m-%d")
    return {
        "label": "FORTNIGHTLY ACTIVITY REPORT",
        "from_date": from_date.strftime("%Y-%m-%d"),
        "to_date": to_date.strftime("%Y-%m-%d"),
        "prev_from_date": prev_from.strftime("%Y-%m-%d"),
        "prev_to_date": prev_to.strftime("%Y-%m-%d"),
        "range_label": f"{from_date.strftime('%d %B').upper()} - {to_date.strftime('%d %B %Y').upper()}",
        "matched_listings_label": f"{from_date.strftime('%d/%m/%y')} - {to_date.strftime('%d/%m/%y')}",
        "file_suffix": to_date.strftime("%Y%m%d"),
    }


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Generate Quiet List Exchange Activity Reports.")
    parser.add_argument("--config", default=os.path.join(SCRIPT_DIR, "config.json"))
    parser.add_argument("--offices-file", default=os.path.join(SCRIPT_DIR, "offices.json"))
    parser.add_argument("--period-start")
    parser.add_argument("--period-end")
    parser.add_argument("--prev-period-start")
    parser.add_argument("--prev-period-end")
    parser.add_argument("--only-office", default=None, help="office_id to restrict to a single office")
    parser.add_argument("--env-file", default=None,
                         help="path to an env file supplying BUBBLE_API_TOKEN, instead of the "
                              "default .env beside this script. Ignored if the variable is "
                              "already set in the real environment.")
    parser.add_argument("--mock", action="store_true", help="use synthetic data, no Bubble access required")
    parser.add_argument("--use-test-version", action="store_true",
                         help="explicitly use config.json's bubble_base_url_test. Currently the default "
                              "for every run regardless of this flag, since BUBBLE_ALLOW_PRODUCTION is "
                              "unset — pass it anyway to self-document intent in scripts.")
    parser.add_argument("--upload-to-drive", action="store_true",
                         help="upload generated PDFs to Drive (config.json's drive_reports_folder_id) "
                              "via GOOGLE_SERVICE_ACCOUNT_JSON. Off by default for manual CLI runs, so a "
                              "--use-test-version reconciliation pass doesn't push synthetic test PDFs "
                              "into the client's real Drive folder. run_scheduled_report.py (the routine "
                              "entry point) always uploads and doesn't use this flag.")
    parser.add_argument("--email-report", action="store_true",
                         help="email generated PDFs and .docx files via the SendGrid API "
                              "(SENDGRID_API_KEY/SENDGRID_FROM_EMAIL) to config.json's "
                              "report_email_recipients, or --email-to if given. Off by default for "
                              "manual CLI runs, for the same reason as --upload-to-drive.")
    parser.add_argument("--email-to", default=None,
                         help="comma-separated recipient override for --email-report, instead of "
                              "config.json's report_email_recipients.")
    args = parser.parse_args()

    if args.env_file:
        if not os.path.exists(args.env_file):
            raise SystemExit(f"--env-file not found: {args.env_file}")
        load_dotenv(args.env_file, override=True)

    with open(args.config) as f:
        config = json.load(f)
    with open(args.offices_file) as f:
        offices = json.load(f)

    period = period_from_args(args)

    summary = generate_all_reports(offices, period, config, mock=args.mock,
                                    only_office=args.only_office, use_test=args.use_test_version)

    total_errors = summary["errors"]
    successes = summary.get("successes", [])
    all_paths = [s["path"] for s in successes] + [s["docx_path"] for s in successes]

    if args.upload_to_drive:
        from drive_upload import upload_reports
        upload_failures = [r for r in upload_reports(all_paths, config.get("drive_reports_folder_id"))
                            if r["status"] == "failed"]
        total_errors += len(upload_failures)

    if args.email_report:
        from email_report import send_report_email
        to_addrs = ([a.strip() for a in args.email_to.split(",") if a.strip()] if args.email_to
                    else config.get("report_email_recipients", []))
        if send_report_email(all_paths, to_addrs)["status"] == "failed":
            total_errors += 1

    sys.exit(0 if total_errors == 0 else 1)


if __name__ == "__main__":
    main()
