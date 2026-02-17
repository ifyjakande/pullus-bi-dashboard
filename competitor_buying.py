from datetime import datetime

from config import (
    COMPETITOR_BUYING_PRICE_SHEET_ID, COMPETITOR_BUYING_PRICE_TAB,
    NAVY, TEAL, GREEN, RED, LIGHT_BG, WHITE, DARK_TEXT,
    CARD_BLUE, CARD_ORANGE, CARD_PURPLE, hex_to_rgb,
)
from helpers import (
    parse_date, safe_float,
    grid_range, cell_fmt, fmt_request, merge_request,
    col_width_request, row_height_request, border_request,
    conditional_format_request, clear_sheet,
)


def parse_buying_price(val):
    """Parse buying price values, handling ranges like '3,300 - 3,500' by averaging."""
    if not val or not val.strip():
        return None
    val = val.strip().replace(",", "")
    if " - " in val:
        parts = [safe_float(p) for p in val.split(" - ") if safe_float(p) > 0]
        return round(sum(parts) / len(parts), 0) if parts else None
    p = safe_float(val)
    return p if p > 0 else None


# ---------- Read & Process ----------
def fetch_buying_data(client):
    sh = client.open_by_key(COMPETITOR_BUYING_PRICE_SHEET_ID)
    ws = sh.worksheet(COMPETITOR_BUYING_PRICE_TAB)
    rows = ws.get_all_values()

    # Find the actual header row (contains "Entry ID")
    header_idx = None
    for i, row in enumerate(rows):
        if any("Entry ID" in cell for cell in row):
            header_idx = i
            break

    if header_idx is None:
        print("  Could not find header row")
        return []

    header = rows[header_idx]
    data = rows[header_idx + 1:]
    col_map = {h.strip(): i for i, h in enumerate(header)}

    current_year = datetime.now().year
    records = []
    for row in data:
        dt = parse_date(row[col_map.get("Date", 1)])
        if dt is None or dt.year != current_year:
            continue

        competitor = row[col_map.get("Competitor Name", 5)].strip()
        product_type = row[col_map.get("Product Type", 6)].strip().title()
        comp_price = parse_buying_price(row[col_map.get("Competitor Price (N)", 7)])
        pullus_price = parse_buying_price(row[col_map.get("Pullus Price (N)", 8)])
        notes = row[col_map.get("Notes", 9)].strip() if col_map.get("Notes", 9) < len(row) else ""
        state = row[col_map.get("State", 4)].strip()

        if not competitor:
            continue

        records.append({
            "date": dt,
            "state": state,
            "competitor": competitor,
            "product_type": product_type,
            "comp_price": comp_price,
            "pullus_price": pullus_price,
            "notes": notes,
        })

    print(f"  Fetched {len(records)} of {len(data)} rows (current year)")
    return records


def compute_summary(records):
    """Compute summary stats from the records."""
    dressed = [r for r in records if r["product_type"] == "Dressed Birds"]
    live = [r for r in records if r["product_type"] == "Live Birds"]

    pullus_prices = [r["pullus_price"] for r in dressed if r.get("pullus_price")]
    comp_prices = [r["comp_price"] for r in dressed if r.get("comp_price")]

    avg_pullus = round(sum(pullus_prices) / len(pullus_prices), 0) if pullus_prices else 0
    avg_comp = round(sum(comp_prices) / len(comp_prices), 0) if comp_prices else 0

    diff_pct = round((avg_pullus - avg_comp) / avg_comp * 100, 1) if avg_comp else 0
    diff_abs = round(avg_pullus - avg_comp, 0)

    live_comp_prices = [r["comp_price"] for r in live if r.get("comp_price")]

    return {
        "avg_pullus": avg_pullus,
        "avg_comp": avg_comp,
        "diff_pct": diff_pct,
        "diff_abs": diff_abs,
        "total_entries": len(records),
        "dressed_entries": len(dressed),
        "live_entries": len(live),
        "live_comp_prices": live_comp_prices,
    }


# ---------- Chart ----------
def build_buying_chart_request(sheet_id, dressed_count, table_start_row, axis_min=0):
    """Chart references table columns directly (dressed birds rows only)."""
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
                        "widthPixels": 720,
                        "heightPixels": 420,
                    }
                },
                "spec": {
                    "title": "Pullus vs Competitor Buying Prices (Dressed Birds)",
                    "titleTextFormat": {"fontSize": 12, "bold": True, "foregroundColor": DARK_TEXT},
                    "basicChart": {
                        "chartType": "COMBO",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {
                                "position": "BOTTOM_AXIS",
                                "title": "",
                                "format": {"fontSize": 7, "foregroundColor": DARK_TEXT},
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
                                            grid_range(sheet_id, table_start_row, table_start_row + dressed_count + 1, 0, 1)
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
                                            grid_range(sheet_id, table_start_row, table_start_row + dressed_count + 1, 4, 5)
                                        ]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "type": "COLUMN",
                                "color": CARD_ORANGE,
                            },
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [
                                            grid_range(sheet_id, table_start_row, table_start_row + dressed_count + 1, 5, 6)
                                        ]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "type": "LINE",
                                "color": CARD_BLUE,
                            },
                        ],
                        "headerCount": 1,
                    },
                },
            }
        }
    }


# ---------- Dashboard Writing ----------
def get_or_create_sheet(dash_sh, title, index=3):
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


def build_dashboard(dash_sh, records):
    ws = get_or_create_sheet(dash_sh, "Competitor Buying Prices", index=3)
    sid = ws.id

    clear_sheet(dash_sh, ws, sid)

    summary = compute_summary(records)

    def pct_str(v):
        if v is None:
            return "N/A"
        sign = "+" if v > 0 else ""
        return f"{sign}{v}%"

    # Split and sort records
    sorted_records = sorted(records, key=lambda r: r["date"])
    dressed_records = [r for r in sorted_records if r["product_type"] == "Dressed Birds"]
    live_records = [r for r in sorted_records if r["product_type"] == "Live Birds"]

    # ---- Write cell values ----
    all_values = []
    NUM_COLS = 9

    # Row 0: Title
    all_values.append(["PULLUS COMPETITOR BUYING PRICES"] + [""] * (NUM_COLS - 1))
    # Row 1: Subtitle
    all_dates = [r["date"] for r in records]
    first_date = min(all_dates).strftime("%d %b %Y")
    last_date = max(all_dates).strftime("%d %b %Y")
    now_ts = datetime.now().strftime("%d %b %Y %I:%M %p")
    all_values.append([f"Data range: {first_date} - {last_date}  |  Last updated: {now_ts}"] + [""] * (NUM_COLS - 1))
    # Row 2: Spacer
    all_values.append([""] * NUM_COLS)

    # Rows 3-7: Summary Cards (Dressed Birds)
    # Paying LESS = bad (competitors attract farmers), paying MORE = good
    if summary["diff_pct"] < 0:
        market_note = "Competitors pay more to farmers"
    else:
        market_note = "Pullus pays more to farmers"
    dressed_label = f"{summary['dressed_entries']} of {summary['total_entries']} entries (Dressed Birds)"

    all_values.append(["PULLUS BUYING PRICE", "", "", "COMPETITOR AVG PRICE", "", "", "PULLUS vs MARKET", "", ""])
    all_values.append([summary["avg_pullus"], "", "", summary["avg_comp"], "", "", pct_str(summary["diff_pct"]), "", ""])
    all_values.append(["Dressed Birds Avg", "", "", "Dressed Birds Avg", "", "", market_note, "", ""])
    all_values.append([
        dressed_label,
        "", "",
        dressed_label,
        "", "",
        f"{'+'if summary['diff_abs'] >= 0 else '-'}₦{abs(int(summary['diff_abs'])):,}",
        "", "",
    ])
    all_values.append([
        last_date,
        "", "",
        last_date,
        "", "",
        "",
        "", "",
    ])

    # Row 8: Spacer
    all_values.append([""] * NUM_COLS)

    # Row 9: Live Birds insight
    live_comp_prices = summary["live_comp_prices"]
    if live_comp_prices:
        live_min = int(min(live_comp_prices))
        live_max = int(max(live_comp_prices))
        if live_min == live_max:
            price_str = f"₦{live_min:,}"
        else:
            price_str = f"₦{live_min:,} - ₦{live_max:,}"
        all_values.append([
            "LIVE BIRDS MARKET", "", "",
            f"Competitor prices: {price_str}", "", "",
            f"{summary['live_entries']} entries  |  No Pullus data", "", "",
        ])
    else:
        all_values.append(["No Live Birds data available"] + [""] * (NUM_COLS - 1))

    # Row 10: Spacer
    all_values.append([""] * NUM_COLS)

    # Row 11: Explainer note
    all_values.append([
        "Shows what Pullus pays farmers for birds vs competitor buying prices. "
        "Diff % = (Pullus - Competitor) / Competitor. "
        "Negative = Pullus pays less (competitors have sourcing advantage)."
    ] + [""] * (NUM_COLS - 1))

    # Row 12: Table header
    TABLE_START = 12
    all_values.append([
        "Date", "Location", "Competitor", "Product Type",
        "Comp Price", "Pullus Price", "Diff %",
        "Diff (₦)", "Notes",
    ])

    # Table data: Dressed Birds first, then Live Birds
    combined_records = dressed_records + live_records
    for r in combined_records:
        comp_price = int(r["comp_price"]) if r.get("comp_price") else ""
        pullus_price = int(r["pullus_price"]) if r.get("pullus_price") else ""

        if r.get("comp_price") and r.get("pullus_price") and r["comp_price"] > 0:
            diff_pct = round((r["pullus_price"] - r["comp_price"]) / r["comp_price"] * 100, 1)
            diff_abs = int(r["pullus_price"] - r["comp_price"])
        else:
            diff_pct = ""
            diff_abs = ""

        all_values.append([
            r["date"].strftime("%d %b %Y"),
            r["state"],
            r["competitor"],
            r["product_type"],
            comp_price,
            pullus_price,
            diff_pct,
            diff_abs,
            r.get("notes", ""),
        ])

    ws.update(all_values, "A1")
    print(f"  Wrote {len(all_values)} rows of data")

    # ---- Formatting requests ----
    reqs = []

    col_widths = [100, 80, 170, 105, 95, 95, 75, 85, 155]
    for i, w in enumerate(col_widths):
        reqs.append(col_width_request(sid, i, w))

    # Row heights
    reqs.append(row_height_request(sid, 0, 50))
    reqs.append(row_height_request(sid, 1, 30))
    reqs.append(row_height_request(sid, 2, 10))
    for r in range(3, 8):
        reqs.append(row_height_request(sid, r, 32))
    reqs.append(row_height_request(sid, 8, 10))
    reqs.append(row_height_request(sid, 9, 32))
    reqs.append(row_height_request(sid, 10, 10))
    reqs.append(row_height_request(sid, 11, 30))
    reqs.append(row_height_request(sid, TABLE_START, 36))

    # Title row
    reqs.append(merge_request(sid, 0, 1, 0, NUM_COLS))
    fmt = cell_fmt(bg=NAVY, fg=WHITE, bold=True, size=18, halign="CENTER", valign="MIDDLE")
    reqs.append(fmt_request(sid, 0, 1, 0, NUM_COLS, fmt))

    # Subtitle row
    reqs.append(merge_request(sid, 1, 2, 0, NUM_COLS))
    fmt = cell_fmt(bg=NAVY, fg={"red": 0.75, "green": 0.78, "blue": 0.82}, bold=False, size=10, halign="CENTER")
    reqs.append(fmt_request(sid, 1, 2, 0, NUM_COLS, fmt))

    # Spacer row 2
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

    # Buying: paying LESS = bad (RED), paying MORE = good (GREEN)
    diff_color = RED if summary["diff_pct"] < 0 else GREEN
    fmt = cell_fmt(bg=WHITE, fg=diff_color, bold=True, size=22, halign="CENTER")
    reqs.append(fmt_request(sid, 4, 5, 6, 9, fmt))

    # Spacer row 8
    fmt = cell_fmt(bg=WHITE)
    reqs.append(fmt_request(sid, 8, 9, 0, NUM_COLS, fmt))

    # Live Birds insight row (row 9) - warm yellow banner
    YELLOW_BG = hex_to_rgb("#FFF3CD")
    YELLOW_BORDER = hex_to_rgb("#FFECB5")
    reqs.append(merge_request(sid, 9, 10, 0, 3))
    reqs.append(merge_request(sid, 9, 10, 3, 6))
    reqs.append(merge_request(sid, 9, 10, 6, 9))
    fmt = cell_fmt(bg=YELLOW_BG, fg=DARK_TEXT, bold=True, size=10, halign="LEFT", valign="MIDDLE")
    reqs.append(fmt_request(sid, 9, 10, 0, 3, fmt))
    fmt = cell_fmt(bg=YELLOW_BG, fg=DARK_TEXT, bold=False, size=10, halign="CENTER", valign="MIDDLE")
    reqs.append(fmt_request(sid, 9, 10, 3, 6, fmt))
    fmt = cell_fmt(bg=YELLOW_BG, fg={"red": 0.6, "green": 0.6, "blue": 0.6}, bold=False, size=9, halign="CENTER", valign="MIDDLE")
    reqs.append(fmt_request(sid, 9, 10, 6, 9, fmt))
    reqs.append({
        "updateBorders": {
            "range": grid_range(sid, 9, 10, 0, NUM_COLS),
            "top": {"style": "SOLID", "color": YELLOW_BORDER},
            "bottom": {"style": "SOLID", "color": YELLOW_BORDER},
            "left": {"style": "SOLID", "color": YELLOW_BORDER},
            "right": {"style": "SOLID", "color": YELLOW_BORDER},
        }
    })

    # Spacer row 10
    fmt = cell_fmt(bg=WHITE)
    reqs.append(fmt_request(sid, 10, 11, 0, NUM_COLS, fmt))

    # Explainer note row 11
    reqs.append(merge_request(sid, 11, 12, 0, NUM_COLS))
    fmt = cell_fmt(bg=WHITE, fg={"red": 0.5, "green": 0.5, "blue": 0.5}, bold=False, size=8, halign="LEFT", valign="MIDDLE", wrap=True)
    reqs.append(fmt_request(sid, 11, 12, 0, NUM_COLS, fmt))

    # Table Header
    fmt = cell_fmt(bg=TEAL, fg=WHITE, bold=True, size=10, halign="CENTER", wrap=True)
    reqs.append(fmt_request(sid, TABLE_START, TABLE_START + 1, 0, NUM_COLS, fmt))

    # Table Data Rows
    data_start = TABLE_START + 1
    data_end = data_start + len(combined_records)

    for i in range(len(combined_records)):
        row_idx = data_start + i
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        fmt = cell_fmt(bg=bg, fg=DARK_TEXT, size=10, halign="CENTER", valign="MIDDLE")
        reqs.append(fmt_request(sid, row_idx, row_idx + 1, 0, NUM_COLS, fmt))

    # Number formatting: Price columns (indices 4, 5)
    for col in [4, 5]:
        fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "₦#,##0"})
        reqs.append(fmt_request(sid, data_start, data_end, col, col + 1, fmt))

    # Diff % column (index 6)
    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "+#,##0.0;-#,##0.0;\"--\""})
    reqs.append(fmt_request(sid, data_start, data_end, 6, 7, fmt))

    # Diff ₦ column (index 7)
    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "+₦#,##0;-₦#,##0;\"--\""})
    reqs.append(fmt_request(sid, data_start, data_end, 7, 8, fmt))

    # Table borders
    reqs.append(border_request(sid, TABLE_START, data_end, 0, NUM_COLS, hex_to_rgb("#DEE2E6")))

    # Conditional formatting: paying LESS = bad (RED), paying MORE = good (GREEN)
    # Swapped vs selling sheet: here positive diff = green, negative diff = red
    reqs.extend(conditional_format_request(sid, data_start, data_end, 6, 7, GREEN, RED))
    reqs.extend(conditional_format_request(sid, data_start, data_end, 7, 8, GREEN, RED))

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

    # Chart (dressed birds only, references table directly)
    dressed_prices = []
    for r in dressed_records:
        if r.get("comp_price"):
            dressed_prices.append(r["comp_price"])
        if r.get("pullus_price"):
            dressed_prices.append(r["pullus_price"])
    chart_axis_min = int(min(dressed_prices) // 500 * 500 - 500) if dressed_prices else 0
    chart_axis_min = max(chart_axis_min, 0)

    chart_req = build_buying_chart_request(sid, len(dressed_records), TABLE_START, axis_min=chart_axis_min)
    dash_sh.batch_update({"requests": [chart_req]})
    print("  Created chart")

    print(f"  Competitor Buying Prices dashboard built: {len(records)} entries")
    print(f"  Dressed Birds: Pullus avg {int(summary['avg_pullus']):,} vs Comp avg {int(summary['avg_comp']):,} ({pct_str(summary['diff_pct'])})")
    if summary["live_entries"] > 0:
        lp = summary["live_comp_prices"]
        print(f"  Live Birds: {summary['live_entries']} entries, comp prices {int(min(lp)):,} - {int(max(lp)):,}")
