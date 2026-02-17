import hashlib
import json
import os

from config import get_client, DASHBOARD_SHEET_ID
from weekly_purchase import (
    fetch_raw_data,
    aggregate_weekly,
    build_dashboard as build_weekly_purchase,
)
from doc_price import (
    fetch_price_data,
    aggregate_weekly_prices,
    build_dashboard as build_doc_price,
)
from competitor_selling import (
    fetch_competitor_data,
    aggregate_by_date_location,
    build_dashboard as build_competitor_selling,
)
from competitor_buying import (
    fetch_buying_data,
    build_dashboard as build_competitor_buying,
)

HASH_FILE = os.path.join(os.path.dirname(__file__), "data_hashes.json")
IS_CI = os.getenv("GITHUB_ACTIONS") == "true"


def compute_hash(data):
    return hashlib.sha256(
        json.dumps(data, default=str, sort_keys=True).encode()
    ).hexdigest()


def load_hashes():
    if os.path.exists(HASH_FILE):
        with open(HASH_FILE) as f:
            return json.load(f)
    return {}


def save_hashes(hashes):
    with open(HASH_FILE, "w") as f:
        json.dump(hashes, f, indent=2)


def data_changed(key, data, prev_hashes, new_hashes):
    """Return True if data differs from previous run. Always True locally."""
    h = compute_hash(data)
    new_hashes[key] = h
    if not IS_CI:
        return True
    if h != prev_hashes.get(key):
        return True
    print("  No changes detected, skipping")
    return False


def main():
    print("Connecting to Google Sheets...")
    client = get_client()
    dash_sh = client.open_by_key(DASHBOARD_SHEET_ID)

    prev_hashes = load_hashes() if IS_CI else {}
    new_hashes = {}
    updated = []

    # --- Sheet 1: Weekly Purchase ---
    print("\n[Weekly Purchase]")
    print("  Fetching purchase data...")
    raw = fetch_raw_data(client)
    if data_changed("weekly_purchase", raw, prev_hashes, new_hashes):
        print("  Aggregating weekly data...")
        weeks = aggregate_weekly(raw)
        if not weeks:
            print("  No current-year data found, skipping")
        else:
            print(f"  Found {len(weeks)} weeks of data")
            print("  Building dashboard...")
            build_weekly_purchase(dash_sh, weeks)
            updated.append("Weekly Purchase")

    # --- Sheet 2: DOC Price Trends ---
    print("\n[DOC Price Trends]")
    print("  Fetching price data...")
    price_data, _ = fetch_price_data(client)
    if data_changed("doc_price", price_data, prev_hashes, new_hashes):
        print("  Aggregating weekly prices...")
        price_weeks = aggregate_weekly_prices(price_data)
        if not price_weeks:
            print("  No current-year data found, skipping")
        else:
            print(f"  Found {len(price_weeks)} weeks of data")
            print("  Building dashboard...")
            build_doc_price(dash_sh, price_weeks)
            updated.append("DOC Price Trends")

    # --- Sheet 3: Competitor Selling Prices ---
    print("\n[Competitor Selling Prices]")
    print("  Fetching competitor data...")
    comp_records = fetch_competitor_data(client)
    if data_changed("competitor_selling", comp_records, prev_hashes, new_hashes):
        print("  Aggregating by date and location...")
        comp_aggregated = aggregate_by_date_location(comp_records)
        if not comp_aggregated:
            print("  No current-year data found, skipping")
        else:
            print(f"  Found {len(comp_aggregated)} entries")
            print("  Building dashboard...")
            build_competitor_selling(dash_sh, comp_aggregated)
            updated.append("Competitor Selling Prices")

    # --- Sheet 4: Competitor Buying Prices ---
    print("\n[Competitor Buying Prices]")
    print("  Fetching buying price data...")
    buying_records = fetch_buying_data(client)
    if data_changed("competitor_buying", buying_records, prev_hashes, new_hashes):
        if not buying_records:
            print("  No current-year data found, skipping")
        else:
            print("  Building dashboard...")
            build_competitor_buying(dash_sh, buying_records)
            updated.append("Competitor Buying Prices")

    if IS_CI:
        save_hashes(new_hashes)

    if updated:
        print(f"\nUpdated: {', '.join(updated)}")
    else:
        print("\nNo changes detected, dashboard untouched.")


if __name__ == "__main__":
    main()
