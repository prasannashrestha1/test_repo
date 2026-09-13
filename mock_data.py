"""
mock_data.py — synthetic Property Matches rows, shaped like the Bubble Data
API response described in property_matches_schema_api_spec_v0.5.md, plus a
`listing_status` field standing in for the per-match `property` join
(fetch_property_status/attach_listing_status in generate_report.py).
Lets generate_report.py be run end-to-end (--mock) with no Bubble
credentials, to prove the fetch -> compute -> template -> PDF pipeline
actually works.
"""

SUBURBS = ["Paddington", "Surry Hills", "Darlinghurst", "Balmain", "Woollahra",
           "Redfern", "Waterloo", "Rozelle", "Glebe", "Newtown"]

DWELLING_TYPES = ["House", "Semi", "Terrace", "Apartment", "Apartment", "Unit"]

AGENCIES = ["Unicorn Buyers Agents", "Aurum Advisory", "Advantage Buyer's Agents",
            "Sterling Property Buyers", "Coastal Buyer Co"]

OPERATIVES = ["Dan Sofo", "Leon Jacques", "Sam Green", "Priya Nair", "Marcus Webb"]


def _match(i, listing_id, suburb, dwelling, price_min, price_max, agency, operative,
           listing_status="CURRENT"):
    return {
        "_id": f"match_{i}",
        "office_id": "1",
        "match_date": "2026-07-15T00:00:00.000Z",
        "listing_id": listing_id,
        "contact_id": f"contact_{i}",
        "address": f"{i} {suburb} Street, {suburb} NSW 2000",
        "listing_price": (price_min + price_max) // 2,
        "dwelling_type": dwelling,
        "brief_ref": f"BRIEF-{i:04d}",
        "price_min": price_min,
        "price_max": price_max,
        "property_types": [dwelling],
        "beds": 3,
        "baths": 2,
        "parking": 1,
        "suburbs": [suburb],
        "ba_name": operative,
        "ba_agency": agency,
        "ba_email": operative.lower().replace(" ", ".") + "@example.com",
        "source": "quiet_list",
        "listing_status": listing_status,
    }


def _generate_matches(n, seed_offset=0):
    matches = []
    # one clear standout listing (id "listing_0") to exercise the featured-listing logic
    for i in range(9):
        matches.append(_match(
            seed_offset + i, "listing_0", "Paddington", "Terrace",
            2_400_000, 2_800_000, AGENCIES[i % len(AGENCIES)], OPERATIVES[i % len(OPERATIVES)],
        ))
    remaining = n - 9
    for i in range(remaining):
        suburb = SUBURBS[i % len(SUBURBS)]
        dwelling = DWELLING_TYPES[i % len(DWELLING_TYPES)]
        listing_id = f"listing_{1 + (i % 20)}"
        price_min = 1_000_000 + (i % 10) * 200_000
        price_max = price_min + 400_000
        # every 7th distinct listing is a withdrawn/no-longer-current one, so
        # the mock run also exercises the "not current" filtering path.
        status = "WITHDRAWN" if (1 + (i % 20)) % 7 == 0 else "CURRENT"
        matches.append(_match(
            seed_offset + 9 + i, listing_id, suburb, dwelling,
            price_min, price_max, AGENCIES[i % len(AGENCIES)], OPERATIVES[i % len(OPERATIVES)],
            listing_status=status,
        ))
    return matches


def get_mock_data(office_id: str):
    matches_current = _generate_matches(181)
    matches_previous = _generate_matches(91, seed_offset=1000)
    return matches_current, matches_previous
