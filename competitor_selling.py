from datetime import datetime
from collections import defaultdict

from config import (
    COMPETITOR_SELLING_PRICE_SHEET_ID, COMPETITOR_SELLING_PRICE_TABS,
    NAVY, TEAL, GREEN, RED, LIGHT_BG, WHITE, DARK_TEXT,
    CARD_BLUE, CARD_ORANGE, CARD_PURPLE, hex_to_rgb,
)
from helpers import (
    parse_date, safe_float,
    grid_range, cell_fmt, fmt_request, merge_request,
    col_width_request, row_height_request, border_request,
    conditional_format_request, clear_sheet,
)

# Location display names (strip _Entry suffix)
LOCATION_NAMES = {
    "Abuja_Entry": "Abuja",
    "Kaduna_Entry": "Kaduna",
    "Kano_Entry": "Kano",
}

# Products to track (column header -> short name)
PRODUCTS = ["Whole Chicken", "Gizzard"]


def parse_price(val):
    """Parse price values, handling formats like '3900/4000' by averaging."""
    if not val or not val.strip():
        return None
    val = val.strip()
    if "/" in val:
        parts = [safe_float(p) for p in val.split("/") if safe_float(p) > 0]
        return round(sum(parts) / len(parts), 0) if parts else None
    p = safe_float(val)
    return p if p > 0 else None


# ---------- Read & Aggregate ----------
def fetch_competitor_data(client):
    """Fetch data from all location tabs and return structured records."""
    sh = client.open_by_key(COMPETITOR_SELLING_PRICE_SHEET_ID)
    all_records = []

    for tab_name in COMPETITOR_SELLING_PRICE_TABS:
        tab_name = tab_name.strip()
        if not tab_name:
            continue
        location = LOCATION_NAMES.get(tab_name, tab_name)
        ws = sh.worksheet(tab_name)
        rows = ws.get_all_values()
        header = rows[0]
        data = rows[1:]

        # Build column index map
        col_map = {h.strip(): i for i, h in enumerate(header)}

        current_year = datetime.now().year
        kept = 0
        for row in data:
            dt = parse_date(row[0])
            if dt is None or dt.year != current_year:
                continue
            brand = row[1].strip() if len(row) > 1 else ""
            if not brand:
                continue

            kept += 1
            record = {
                "date": dt,
                "location": location,
                "brand": brand,
            }
            for product_name in PRODUCTS:
                idx = col_map.get(product_name)
                if idx is not None and idx < len(row):
                    record[product_name] = parse_price(row[idx])
                else:
                    record[product_name] = None

            all_records.append(record)

        print(f"  {location}: {kept} of {len(data)} rows kept")

    return all_records


def aggregate_by_date_location(records):
    """For each (date, location): compute Pullus price and competitor average per product."""
    groups = defaultdict(lambda: {"pullus": defaultdict(list), "competitors": defaultdict(list), "comp_brands": set()})

    for r in records:
        key = (r["date"], r["location"])
        is_pullus = r["brand"].lower() == "pullus"

        for product_name in PRODUCTS:
            price = r.get(product_name)
            if price is None:
                continue
            if is_pullus:
                groups[key]["pullus"][product_name].append(price)
            else:
                groups[key]["competitors"][product_name].append(price)
                groups[key]["comp_brands"].add(r["brand"])

    result = []
    for (dt, loc), vals in sorted(groups.items()):
        row = {"date": dt, "location": loc, "comp_brands": len(vals["comp_brands"])}
        for product_name in PRODUCTS:
            pullus_prices = vals["pullus"].get(product_name, [])
            pullus_price = round(sum(pullus_prices) / len(pullus_prices), 0) if pullus_prices else None
            comp_prices = vals["competitors"].get(product_name, [])
            comp_avg = round(sum(comp_prices) / len(comp_prices), 0) if comp_prices else None

            row[f"{product_name}_pullus"] = pullus_price
            row[f"{product_name}_comp_avg"] = comp_avg

            if pullus_price and comp_avg and comp_avg > 0:
                row[f"{product_name}_diff_pct"] = round(
                    (pullus_price - comp_avg) / comp_avg * 100, 1
                )
            else:
                row[f"{product_name}_diff_pct"] = None

        result.append(row)

    return result


# ---------- Chart ----------
def build_competitor_chart_request(sheet_id, data_count, table_start_row, axis_min=0):
    """Line chart: Pullus WC vs Competitor Avg WC, using table columns."""
    return {
        "addChart": {
            "chart": {
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": sheet_id,
                            "rowIndex": 3,
                            "columnIndex": 9,
                        },
                        "offsetXPixels": 20,
                        "offsetYPixels": 0,
                        "widthPixels": 780,
                        "heightPixels": 420,
                    }
                },
                "spec": {
                    "title": "Pullus vs Competitor Avg — Whole Chicken",
                    "titleTextFormat": {"fontSize": 12, "bold": True, "foregroundColor": DARK_TEXT},
                    "basicChart": {
                        "chartType": "COMBO",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {
                                "position": "BOTTOM_AXIS",
                                "title": "Date / Location",
                                "format": {"fontSize": 8, "foregroundColor": DARK_TEXT},
                            },
                            {
                                "position": "LEFT_AXIS",
                                "title": "Price (₦)",
                                "format": {"fontSize": 9, "foregroundColor": DARK_TEXT},
                                "viewWindowOptions": {
                                    "viewWindowMin": axis_min,
                                },
                            },
                        ],
                        "domains": [
                            {
                                "domain": {
                                    "sourceRange": {
                                        "sources": [
                                            grid_range(sheet_id, table_start_row, table_start_row + data_count + 1, 0, 1)
                                        ]
                                    }
                                }
                            }
                        ],
                        "series": [
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [
                                            grid_range(sheet_id, table_start_row, table_start_row + data_count + 1, 2, 3)
                                        ]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "type": "LINE",
                                "color": CARD_BLUE,
                            },
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [
                                            grid_range(sheet_id, table_start_row, table_start_row + data_count + 1, 3, 4)
                                        ]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "type": "LINE",
                                "color": CARD_ORANGE,
                                "lineStyle": {"type": "MEDIUM_DASHED"},
                            },
                        ],
                        "headerCount": 1,
                    },
                },
            }
        }
    }


# ---------- Dashboard Writing ----------
def get_or_create_sheet(dash_sh, title, index=2):
    for ws in dash_sh.worksheets():
        if ws.title == title:
            return ws
    ws = dash_sh.add_worksheet(title=title, rows=500, cols=20)
    dash_sh.batch_update({
        "requests": [{
            "updateSheetProperties": {
                "properties": {"sheetId": ws.id, "index": index},
                "fields": "index",
            }
        }]
    })
    return ws


def build_dashboard(dash_sh, aggregated):
    ws = get_or_create_sheet(dash_sh, "Competitor Selling Prices", index=2)
    sid = ws.id

    clear_sheet(dash_sh, ws, sid)

    # ---- Prepare summary data ----
    # Latest entry per location
    latest_by_loc = {}
    for row in aggregated:
        loc = row["location"]
        if loc not in latest_by_loc or row["date"] > latest_by_loc[loc]["date"]:
            latest_by_loc[loc] = row

    # Average Pullus WC across locations (latest)
    pullus_wc_prices = [r["Whole Chicken_pullus"] for r in latest_by_loc.values() if r.get("Whole Chicken_pullus")]
    comp_wc_prices = [r["Whole Chicken_comp_avg"] for r in latest_by_loc.values() if r.get("Whole Chicken_comp_avg")]

    latest_pullus_wc = round(sum(pullus_wc_prices) / len(pullus_wc_prices), 0) if pullus_wc_prices else 0
    latest_comp_wc = round(sum(comp_wc_prices) / len(comp_wc_prices), 0) if comp_wc_prices else 0
    premium_pct = round((latest_pullus_wc - latest_comp_wc) / latest_comp_wc * 100, 1) if latest_comp_wc else 0

    def pct_str(v):
        if v is None:
            return "N/A"
        sign = "+" if v > 0 else ""
        return f"{sign}{v}%"

    # ---- Write cell values ----
    all_values = []
    NUM_COLS = 9

    # Row 1: Title
    all_values.append(["PULLUS COMPETITOR SELLING PRICES"] + [""] * (NUM_COLS - 1))
    # Row 2: Subtitle
    all_dates = [r["date"] for r in aggregated]
    first_date = min(all_dates).strftime("%d %b %Y")
    last_date = max(all_dates).strftime("%d %b %Y")
    now_ts = datetime.now().strftime("%d %b %Y %I:%M %p")
    all_values.append([f"Data range: {first_date} - {last_date}  |  Last updated: {now_ts}"] + [""] * (NUM_COLS - 1))
    # Row 3: Spacer
    all_values.append([""] * NUM_COLS)

    # Rows 4-8: Summary Cards
    # Card 1: Pullus Whole Chicken (latest avg across locations)
    # Card 2: Competitor Avg Whole Chicken
    # Card 3: Premium/Discount %
    above_or_below = "ABOVE" if premium_pct > 0 else "BELOW"
    all_values.append(["PULLUS WHOLE CHICKEN", "", "", "COMPETITOR AVG WC", "", "", f"PULLUS vs MARKET", "", ""])
    all_values.append([latest_pullus_wc, "", "", latest_comp_wc, "", "", pct_str(premium_pct), "", ""])
    all_values.append(["Avg Across Locations", "", "", "Avg Across Locations", "", "", f"Priced {above_or_below.lower()} competitors", "", ""])

    # Sub-value row: per-location latest prices
    loc_pullus_parts = []
    loc_comp_parts = []
    for loc in ["Abuja", "Kaduna", "Kano"]:
        lr = latest_by_loc.get(loc)
        if lr and lr.get("Whole Chicken_pullus"):
            loc_pullus_parts.append(f"{loc}: ₦{int(lr['Whole Chicken_pullus']):,}")
        if lr and lr.get("Whole Chicken_comp_avg"):
            loc_comp_parts.append(f"{loc}: ₦{int(lr['Whole Chicken_comp_avg']):,}")

    all_values.append([
        " | ".join(loc_pullus_parts) if loc_pullus_parts else "",
        "", "",
        " | ".join(loc_comp_parts) if loc_comp_parts else "",
        "", "",
        f"{'+'if (latest_pullus_wc - latest_comp_wc) >= 0 else '-'}₦{abs(int(latest_pullus_wc - latest_comp_wc)):,}" if latest_comp_wc else "",
        "", "",
    ])

    # Latest survey dates
    latest_dates = [f"{loc}: {r['date'].strftime('%d %b')}" for loc, r in sorted(latest_by_loc.items())]
    all_values.append([" | ".join(latest_dates)] + [""] * (NUM_COLS - 1))

    # Row 9: Spacer
    all_values.append([""] * NUM_COLS)

    # Row 10: Explainer note
    all_values.append([
        "Shows Pullus selling prices vs competitor averages per location and survey date. "
        "Diff % = (Pullus - Competitor Avg) / Competitor Avg. Positive = Pullus priced higher."
    ] + [""] * (NUM_COLS - 1))

    # Row 11+: Table
    TABLE_START = 10
    all_values.append([
        "Date", "Location",
        "Pullus WC", "Comp Avg WC", "WC Diff %",
        "Pullus Gzd", "Comp Avg Gzd", "Gzd Diff %",
        "Competitors",
    ])

    for row in aggregated:
        dt_str = row["date"].strftime("%d %b %Y")

        wc_pullus = int(row["Whole Chicken_pullus"]) if row.get("Whole Chicken_pullus") else ""
        wc_comp = int(row["Whole Chicken_comp_avg"]) if row.get("Whole Chicken_comp_avg") else ""
        wc_diff = row.get("Whole Chicken_diff_pct")
        wc_diff = round(wc_diff, 1) if wc_diff is not None else ""

        gzd_pullus = int(row["Gizzard_pullus"]) if row.get("Gizzard_pullus") else ""
        gzd_comp = int(row["Gizzard_comp_avg"]) if row.get("Gizzard_comp_avg") else ""
        gzd_diff = row.get("Gizzard_diff_pct")
        gzd_diff = round(gzd_diff, 1) if gzd_diff is not None else ""

        all_values.append([
            dt_str,
            row["location"],
            wc_pullus, wc_comp, wc_diff,
            gzd_pullus, gzd_comp, gzd_diff,
            row.get("comp_brands", ""),
        ])

    ws.update(all_values, "A1")
    print(f"  Wrote {len(all_values)} rows of data")

    # ---- Formatting requests ----
    reqs = []

    col_widths = [100, 85, 95, 110, 85, 95, 110, 85, 90]
    for i, w in enumerate(col_widths):
        reqs.append(col_width_request(sid, i, w))

    # Row heights
    reqs.append(row_height_request(sid, 0, 50))
    reqs.append(row_height_request(sid, 1, 30))
    reqs.append(row_height_request(sid, 2, 10))
    for r in range(3, 8):
        reqs.append(row_height_request(sid, r, 32))
    reqs.append(row_height_request(sid, 8, 10))
    reqs.append(row_height_request(sid, 9, 30))
    reqs.append(row_height_request(sid, 10, 36))

    # Title row
    reqs.append(merge_request(sid, 0, 1, 0, NUM_COLS))
    fmt = cell_fmt(bg=NAVY, fg=WHITE, bold=True, size=18, halign="CENTER", valign="MIDDLE")
    reqs.append(fmt_request(sid, 0, 1, 0, NUM_COLS, fmt))

    # Subtitle row
    reqs.append(merge_request(sid, 1, 2, 0, NUM_COLS))
    fmt = cell_fmt(bg=NAVY, fg={"red": 0.75, "green": 0.78, "blue": 0.82}, bold=False, size=10, halign="CENTER")
    reqs.append(fmt_request(sid, 1, 2, 0, NUM_COLS, fmt))

    # Spacer row 3
    fmt = cell_fmt(bg=WHITE)
    reqs.append(fmt_request(sid, 2, 3, 0, NUM_COLS, fmt))

    # Summary Cards
    card_defs = [
        (0, 3, CARD_BLUE),
        (3, 6, CARD_ORANGE),
        (6, 9, CARD_PURPLE),
    ]

    for c1, c2, accent in card_defs:
        reqs.append(merge_request(sid, 3, 4, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg=DARK_TEXT, bold=True, size=9, halign="CENTER")
        reqs.append(fmt_request(sid, 3, 4, c1, c2, fmt))

        reqs.append(merge_request(sid, 4, 5, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg=NAVY, bold=True, size=22, halign="CENTER",
                           pattern={"type": "NUMBER", "pattern": "₦#,##0"})
        reqs.append(fmt_request(sid, 4, 5, c1, c2, fmt))

        reqs.append(merge_request(sid, 5, 6, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg={"red": 0.6, "green": 0.6, "blue": 0.6}, bold=False, size=8, halign="CENTER")
        reqs.append(fmt_request(sid, 5, 6, c1, c2, fmt))

        reqs.append(merge_request(sid, 6, 7, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg=DARK_TEXT, bold=True, size=9, halign="CENTER", wrap=True)
        reqs.append(fmt_request(sid, 6, 7, c1, c2, fmt))

        reqs.append(merge_request(sid, 7, 8, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg={"red": 0.6, "green": 0.6, "blue": 0.6}, bold=False, size=7, halign="CENTER", wrap=True)
        reqs.append(fmt_request(sid, 7, 8, c1, c2, fmt))

        reqs.append({
            "updateBorders": {
                "range": grid_range(sid, 3, 8, c1, c2),
                "top": {"style": "SOLID_MEDIUM", "color": accent, "width": 3},
                "bottom": {"style": "SOLID", "color": hex_to_rgb("#DEE2E6")},
                "left": {"style": "SOLID", "color": hex_to_rgb("#DEE2E6")},
                "right": {"style": "SOLID", "color": hex_to_rgb("#DEE2E6")},
            }
        })

    # Premium card color (card 3, row 5 big number is text so color the sub-value row)
    premium_color = RED if premium_pct > 0 else GREEN
    fmt = cell_fmt(bg=WHITE, fg=premium_color, bold=True, size=22, halign="CENTER")
    reqs.append(fmt_request(sid, 4, 5, 6, 9, fmt))

    # Spacer row 9
    fmt = cell_fmt(bg=WHITE)
    reqs.append(fmt_request(sid, 8, 9, 0, NUM_COLS, fmt))

    # Explainer note row
    reqs.append(merge_request(sid, 9, 10, 0, NUM_COLS))
    fmt = cell_fmt(bg=WHITE, fg={"red": 0.5, "green": 0.5, "blue": 0.5}, bold=False, size=8, halign="LEFT", valign="MIDDLE", wrap=True)
    reqs.append(fmt_request(sid, 9, 10, 0, NUM_COLS, fmt))

    # Table Header
    fmt = cell_fmt(bg=TEAL, fg=WHITE, bold=True, size=10, halign="CENTER", wrap=True)
    reqs.append(fmt_request(sid, TABLE_START, TABLE_START + 1, 0, NUM_COLS, fmt))

    # Table Data Rows
    data_start = TABLE_START + 1
    data_end = data_start + len(aggregated)

    for i in range(len(aggregated)):
        row_idx = data_start + i
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        fmt = cell_fmt(bg=bg, fg=DARK_TEXT, size=10, halign="CENTER", valign="MIDDLE")
        reqs.append(fmt_request(sid, row_idx, row_idx + 1, 0, NUM_COLS, fmt))

    # Number formatting: Price columns (indices 2,3,5,6) - integer with commas
    for col in [2, 3, 5, 6]:
        fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "₦#,##0"})
        reqs.append(fmt_request(sid, data_start, data_end, col, col + 1, fmt))

    # Diff % columns (indices 4, 7)
    for col in [4, 7]:
        fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "+#,##0.0;-#,##0.0;\"--\""})
        reqs.append(fmt_request(sid, data_start, data_end, col, col + 1, fmt))

    # Table borders
    reqs.append(border_request(sid, TABLE_START, data_end, 0, NUM_COLS, hex_to_rgb("#DEE2E6")))

    # Conditional formatting for Diff % columns (indices 4, 7)
    # Note: For selling prices, higher than competitors = premium (red for concern, green for competitive)
    # Using red for positive (Pullus more expensive) and green for negative (Pullus cheaper)
    reqs.extend(conditional_format_request(sid, data_start, data_end, 4, 5, RED, GREEN))
    reqs.extend(conditional_format_request(sid, data_start, data_end, 7, 8, RED, GREEN))

    # Freeze header
    reqs.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sid,
                "gridProperties": {"frozenRowCount": TABLE_START + 1},
            },
            "fields": "gridProperties.frozenRowCount",
        }
    })

    dash_sh.batch_update({"requests": reqs})
    print("  Applied formatting")

    # Chart
    all_prices = []
    for row in aggregated:
        for key in ["Whole Chicken_pullus", "Whole Chicken_comp_avg"]:
            if row.get(key):
                all_prices.append(row[key])
    chart_axis_min = int(min(all_prices) // 100 * 100 - 200) if all_prices else 0
    chart_req = build_competitor_chart_request(sid, len(aggregated), TABLE_START, axis_min=chart_axis_min)
    dash_sh.batch_update({"requests": [chart_req]})
    print("  Created chart")

    print(f"  Competitor Selling Prices dashboard built: {len(aggregated)} entries across {len(latest_by_loc)} locations")
