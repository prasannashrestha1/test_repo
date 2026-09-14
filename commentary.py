"""
commentary.py — generates the 5 narrative bullet points for the Quiet List
Activity Report purely from computed metrics. No manual editing, no LLM call
at run time: deterministic and reconcilable, so the same inputs always
produce the same commentary (important for the Bresic line-by-line test).

Import build_commentary() from generate_report.py once metrics are computed.
"""


def _pct_change(current: float, previous: float) -> float:
    if not previous:
        return 0.0
    return (current - previous) / previous * 100


def _direction_word(pct: float, up_word: str, down_word: str) -> str:
    return up_word if pct >= 0 else down_word


def bullet_stock_vs_matches(office_name, listings_current, listings_previous,
                             matches_current, matches_previous) -> str:
    listings_pct = _pct_change(listings_current, listings_previous)
    matches_pct = _pct_change(matches_current, matches_previous)
    listings_dir = _direction_word(listings_pct, "an increase", "a reduction")
    matches_verb = _direction_word(matches_pct, "increased", "decreased")

    opener = "Following" if listings_pct >= 0 else "Despite"

    if listings_pct < 0 and matches_pct >= 0:
        closing = "aligning with buyer demand"
    elif listings_pct >= 0 and matches_pct >= 0:
        closing = "reflecting sustained engagement"
    elif listings_pct < 0 and matches_pct < 0:
        closing = "lowering buyer engagement"
    else:
        closing = "engagement eased slightly"

    return (
        f"{opener} {listings_dir} of {abs(listings_pct):.0f}% in stock "
        f"({listings_current} vs {listings_previous}), matches {matches_verb} from "
        f"{matches_previous} to {matches_current} — {closing}."
    )


def bullet_top_suburbs(office_name, top_suburbs, matches_current) -> str:
    # top_suburbs: list of (suburb_display, count), ranked desc, needs >= 2 entries
    top = top_suburbs[:5]
    lead_suburb, lead_count = top[0]
    rest = top[1:]
    rest_str = ", ".join(f"{s} ({c})" for s, c in rest[:-1])
    if len(rest) > 1:
        rest_str += f" and {rest[-1][0]} ({rest[-1][1]})"
    elif len(rest) == 1:
        rest_str = f"{rest[0][0]} ({rest[0][1]})"

    return f"{lead_suburb} leads with {lead_count} matches, followed by {rest_str}."


def bullet_budget_and_dwelling(budget_range_display, budget_range_share_pct,
                                dwelling_type_display) -> str:
    return (
        f"Demand was strongest in the {budget_range_display} segment, led by "
        f"{dwelling_type_display} listings ({budget_range_share_pct:.0f}% of matches)."
    )


def bullet_featured_listing(featured_address, featured_count,
                             second_tier_min, second_tier_max) -> str:
    if second_tier_min == second_tier_max:
        second_tier_clause = f"{second_tier_min} matches each"
    else:
        second_tier_clause = f"between {second_tier_min} and {second_tier_max} matches"

    return (
        f"{featured_address} led interest with {featured_count} matches; the next-highest "
        f"listings drew {second_tier_clause}."
    )


def bullet_property_type_split(performance_overview) -> str:
    # performance_overview: [{property_type, listings, matches}, ...] -- one row
    # per property type actually active this period, so this can be any
    # length (1 if only one type had matches, several if many did), not
    # always exactly 2.
    if not performance_overview:
        return "No matches were recorded against any property type this period."

    enriched = []
    for row in performance_overview:
        listings = float(row["listings"])
        matches = float(row["matches"])
        ratio = matches / listings if listings else 0
        enriched.append({**row, "ratio": ratio})
    enriched.sort(key=lambda r: r["ratio"], reverse=True)
    a = enriched[0]

    if len(enriched) == 1:
        return (
            f"{a['property_type']} was the only active property type this period, "
            f"generating {a['matches']} matches from {a['listings']} listings."
        )

    b = enriched[1]
    others_clause = ""
    if len(enriched) > 2:
        others_clause = f", with {len(enriched) - 2} other property type(s) also active"

    return (
        f"{a['property_type']} led with {a['matches']} matches from {a['listings']} listings "
        f"vs {b['matches']} from {b['listings']} {b['property_type'].lower()}{others_clause} — "
        f"strongest buyer-brief relevance this period."
    )


def build_commentary(metrics: dict) -> list:
    """
    metrics keys required:
      office_name, listings_current, listings_previous, matches_current, matches_previous,
      top_suburbs (list of (suburb, count)), budget_range_display, budget_range_share_pct,
      dwelling_type_display, featured_listing_address, featured_listing_count,
      second_tier_min, second_tier_max, performance_overview
    """
    return [
        bullet_stock_vs_matches(
            metrics["office_name"], metrics["listings_current"], metrics["listings_previous"],
            metrics["matches_current"], metrics["matches_previous"],
        ),
        bullet_top_suburbs(
            metrics["office_name"], metrics["top_suburbs"], metrics["matches_current"],
        ),
        bullet_budget_and_dwelling(
            metrics["budget_range_display"], metrics["budget_range_share_pct"],
            metrics["dwelling_type_display"],
        ),
        bullet_featured_listing(
            metrics["featured_listing_address"], metrics["featured_listing_count"],
            metrics["second_tier_min"], metrics["second_tier_max"],
        ),
        bullet_property_type_split(metrics["performance_overview"]),
    ]
