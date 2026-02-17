from datetime import datetime
from collections import defaultdict

from config import (
    DOC_PRICE_SHEET_ID, DOC_PRICE_SHEET_NAME,
    NAVY, TEAL, GREEN, RED, LIGHT_BG, WHITE, DARK_TEXT,
    CARD_BLUE, CARD_ORANGE, CARD_PURPLE, hex_to_rgb,
)
from helpers import (
    parse_date, safe_float,
    grid_range, cell_fmt, fmt_request, merge_request,
    col_width_request, row_height_request, border_request,
    conditional_format_request, clear_sheet,
)


# ---------- Read & Aggregate ----------
def fetch_price_data(client):
    sh = client.open_by_key(DOC_PRICE_SHEET_ID)
    ws = sh.worksheet(DOC_PRICE_SHEET_NAME)
    rows = ws.get_all_values()
    header = rows[0]
    data = rows[1:]
    suppliers = header[1:]
    print(f"  Fetched {len(data)} price entries, {len(suppliers)} suppliers")
    return data, suppliers


def aggregate_weekly_prices(data):
    """Aggregate daily supplier prices into weekly averages."""
    weeks = defaultdict(lambda: {"prices": [], "dates": []})

    current_year = datetime.now().year
    kept = 0
    for row in data:
        dt = parse_date(row[0])
        if dt is None or dt.year != current_year:
            continue

        # Collect all non-empty supplier prices for this date
        day_prices = []
        for val in row[1:]:
            p = safe_float(val)
            if p > 0:
                day_prices.append(p)

        if not day_prices:
            continue

        kept += 1
        iso_year, iso_week, _ = dt.isocalendar()
        key = (iso_year, iso_week)
        weeks[key]["dates"].append(dt)
        # Store the daily average across suppliers
        weeks[key]["prices"].append({
            "avg": sum(day_prices) / len(day_prices),
            "min": min(day_prices),
            "max": max(day_prices),
        })

    print(f"  Filtered {kept} of {len(data)} rows (current year with prices)")

    result = []
    for (yr, wk), vals in sorted(weeks.items()):
        mn_date = min(vals["dates"])
        mx_date = max(vals["dates"])
        entries = len(vals["prices"])

        all_avgs = [p["avg"] for p in vals["prices"]]
        all_mins = [p["min"] for p in vals["prices"]]
        all_maxs = [p["max"] for p in vals["prices"]]

        week_avg = round(sum(all_avgs) / len(all_avgs), 0)
        week_min = min(all_mins)
        week_max = max(all_maxs)

        result.append({
            "year": yr,
            "week": wk,
            "start": mn_date,
            "end": mx_date,
            "entries": entries,
            "avg_price": week_avg,
            "min_price": week_min,
            "max_price": week_max,
            "spread": round(week_max - week_min, 0),
        })

    # WoW change on average price
    for i, w in enumerate(result):
        if i == 0:
            w["wow_pct"] = None
            w["wow_abs"] = None
        else:
            prev = result[i - 1]
            if prev["avg_price"]:
                w["wow_pct"] = round((w["avg_price"] - prev["avg_price"]) / prev["avg_price"] * 100, 1)
                w["wow_abs"] = round(w["avg_price"] - prev["avg_price"], 0)
            else:
                w["wow_pct"] = None
                w["wow_abs"] = None

    return result


# ---------- Chart ----------
def build_price_chart_request(sheet_id, weeks_count, table_start_row, axis_min=0):
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
                    "title": "DOC Price Trend (Weekly Avg)",
                    "titleTextFormat": {"fontSize": 12, "bold": True, "foregroundColor": DARK_TEXT},
                    "basicChart": {
                        "chartType": "COMBO",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {
                                "position": "BOTTOM_AXIS",
                                "title": "Week",
                                "format": {"fontSize": 9, "foregroundColor": DARK_TEXT},
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
                                            grid_range(sheet_id, table_start_row, table_start_row + weeks_count + 1, 0, 1)
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
                                            grid_range(sheet_id, table_start_row, table_start_row + weeks_count + 1, 3, 4)
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
                                            grid_range(sheet_id, table_start_row, table_start_row + weeks_count + 1, 4, 5)
                                        ]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "type": "LINE",
                                "color": hex_to_rgb("#BDC3C7"),
                                "lineStyle": {"type": "MEDIUM_DASHED"},
                            },
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [
                                            grid_range(sheet_id, table_start_row, table_start_row + weeks_count + 1, 5, 6)
                                        ]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "type": "LINE",
                                "color": hex_to_rgb("#BDC3C7"),
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
def get_or_create_sheet(dash_sh, title, index=1):
    """Get existing sheet by title or create a new one."""
    for ws in dash_sh.worksheets():
        if ws.title == title:
            return ws
    ws = dash_sh.add_worksheet(title=title, rows=500, cols=20)
    # Move to desired position
    dash_sh.batch_update({
        "requests": [{
            "updateSheetProperties": {
                "properties": {"sheetId": ws.id, "index": index},
                "fields": "index",
            }
        }]
    })
    return ws


def build_dashboard(dash_sh, weeks):
    ws = get_or_create_sheet(dash_sh, "DOC Price Trends", index=1)
    sid = ws.id

    clear_sheet(dash_sh, ws, sid)

    # ---- Prepare summary data ----
    latest = weeks[-1]
    prev = weeks[-2] if len(weeks) > 1 else None

    # YTD stats
    ytd_high = max(w["max_price"] for w in weeks)
    ytd_low = min(w["min_price"] for w in weeks)
    ytd_high_week = next(w for w in weeks if w["max_price"] == ytd_high)
    ytd_low_week = next(w for w in weeks if w["min_price"] == ytd_low)

    def wow_str(v):
        if v is None:
            return "N/A"
        sign = "+" if v > 0 else ""
        return f"{sign}{v}%"

    def abs_wow_str(v):
        if v is None:
            return "N/A"
        sign = "+" if v > 0 else "-" if v < 0 else ""
        return f"{sign}₦{abs(int(v)):,}"

    # ---- Write cell values ----
    all_values = []
    NUM_COLS = 9

    # Row 1: Title
    all_values.append(["PULLUS DOC PRICE TRENDS"] + [""] * (NUM_COLS - 1))
    # Row 2: Subtitle
    first_date = weeks[0]["start"].strftime("%d %b %Y")
    last_date = weeks[-1]["end"].strftime("%d %b %Y")
    now_ts = datetime.now().strftime("%d %b %Y %I:%M %p")
    all_values.append([f"Data range: {first_date} - {last_date}  |  Last updated: {now_ts}"] + [""] * (NUM_COLS - 1))
    # Row 3: Spacer
    all_values.append([""] * NUM_COLS)

    # Rows 4-8: Summary Cards
    # Card 1: Latest Avg Price + WoW
    # Card 2: YTD High
    # Card 3: YTD Low
    all_values.append(["LATEST AVG PRICE", "", "", "YTD HIGH", "", "", "YTD LOW", "", ""])
    all_values.append([latest["avg_price"], "", "", ytd_high, "", "", ytd_low, "", ""])
    all_values.append(["WoW Change", "", "", "2026 Maximum", "", "", "2026 Minimum", "", ""])
    all_values.append([
        f"{abs_wow_str(latest['wow_abs'])} ({wow_str(latest['wow_pct'])})",
        "", "",
        f"Spread: ₦{int(ytd_high - ytd_low):,}",
        "", "",
        f"W{ytd_low_week['week']} ({ytd_low_week['start'].strftime('%d %b')})",
        "", "",
    ])
    latest_label = f"Week {latest['week']}, {latest['year']}"
    high_label = f"Week {ytd_high_week['week']}, {ytd_high_week['year']}"
    low_label = f"Week {ytd_low_week['week']}, {ytd_low_week['year']}"
    all_values.append([latest_label, "", "", high_label, "", "", low_label, "", ""])

    # Row 9: Spacer
    all_values.append([""] * NUM_COLS)

    # Row 10: Explainer note
    all_values.append([
        "Prices are averaged across all reporting suppliers per date, then aggregated weekly. "
        "Spread = Max - Min supplier price. WoW = week-over-week change in average price."
    ] + [""] * (NUM_COLS - 1))

    # Row 11+: Table
    TABLE_START = 10
    all_values.append([
        "Week", "Date Range", "Days",
        "Avg Price", "Min Price", "Max Price",
        "Spread", "WoW Change", "WoW %",
    ])

    for w in weeks:
        dr = f"{w['start'].strftime('%d %b')} - {w['end'].strftime('%d %b %Y')}"
        wow_abs = w["wow_abs"] if w["wow_abs"] is not None else ""
        wow_pct = w["wow_pct"] if w["wow_pct"] is not None else ""
        if isinstance(wow_abs, float):
            wow_abs = int(wow_abs)
        if isinstance(wow_pct, (int, float)):
            wow_pct = round(wow_pct, 1)
        all_values.append([
            f"W{w['week']}",
            dr,
            w["entries"],
            int(w["avg_price"]),
            int(w["min_price"]),
            int(w["max_price"]),
            int(w["spread"]),
            wow_abs,
            wow_pct,
        ])

    ws.update(all_values, "A1")
    print(f"  Wrote {len(all_values)} rows of data")

    # ---- Formatting requests ----
    reqs = []

    col_widths = [70, 180, 60, 95, 95, 95, 85, 100, 85]
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
        # Card label row
        reqs.append(merge_request(sid, 3, 4, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg=DARK_TEXT, bold=True, size=9, halign="CENTER")
        reqs.append(fmt_request(sid, 3, 4, c1, c2, fmt))

        # Big number row
        reqs.append(merge_request(sid, 4, 5, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg=NAVY, bold=True, size=22, halign="CENTER",
                           pattern={"type": "NUMBER", "pattern": "₦#,##0"})
        reqs.append(fmt_request(sid, 4, 5, c1, c2, fmt))

        # Sub-label row
        reqs.append(merge_request(sid, 5, 6, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg={"red": 0.6, "green": 0.6, "blue": 0.6}, bold=False, size=8, halign="CENTER")
        reqs.append(fmt_request(sid, 5, 6, c1, c2, fmt))

        # Sub-value row
        reqs.append(merge_request(sid, 6, 7, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg=DARK_TEXT, bold=True, size=12, halign="CENTER")
        reqs.append(fmt_request(sid, 6, 7, c1, c2, fmt))

        # Week label row
        reqs.append(merge_request(sid, 7, 8, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg={"red": 0.6, "green": 0.6, "blue": 0.6}, bold=False, size=8, halign="CENTER")
        reqs.append(fmt_request(sid, 7, 8, c1, c2, fmt))

        # Top accent border
        reqs.append({
            "updateBorders": {
                "range": grid_range(sid, 3, 8, c1, c2),
                "top": {"style": "SOLID_MEDIUM", "color": accent, "width": 3},
                "bottom": {"style": "SOLID", "color": hex_to_rgb("#DEE2E6")},
                "left": {"style": "SOLID", "color": hex_to_rgb("#DEE2E6")},
                "right": {"style": "SOLID", "color": hex_to_rgb("#DEE2E6")},
            }
        })

    # Color the WoW text on card 1 (row 7, cols 0-3)
    wow_color = GREEN if latest.get("wow_pct") is not None and latest["wow_pct"] > 0 else RED
    fmt = cell_fmt(bg=WHITE, fg=wow_color, bold=True, size=12, halign="CENTER")
    reqs.append(fmt_request(sid, 6, 7, 0, 3, fmt))

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
    data_end = data_start + len(weeks)

    for i in range(len(weeks)):
        row_idx = data_start + i
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        fmt = cell_fmt(bg=bg, fg=DARK_TEXT, size=10, halign="CENTER", valign="MIDDLE")
        reqs.append(fmt_request(sid, row_idx, row_idx + 1, 0, NUM_COLS, fmt))

    # Number formatting: Entries column (col C, index 2) - integer
    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "#,##0"})
    reqs.append(fmt_request(sid, data_start, data_end, 2, 3, fmt))

    # Price columns (cols D-G, index 3-6) - integer with commas
    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "₦#,##0"})
    reqs.append(fmt_request(sid, data_start, data_end, 3, 7, fmt))

    # WoW Change column (col H, index 7) - signed integer
    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "+₦#,##0;-₦#,##0;\"--\""})
    reqs.append(fmt_request(sid, data_start, data_end, 7, 8, fmt))

    # WoW % column (col I, index 8)
    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "+#,##0.0;-#,##0.0;\"--\""})
    reqs.append(fmt_request(sid, data_start, data_end, 8, 9, fmt))

    # Table borders
    reqs.append(border_request(sid, TABLE_START, data_end, 0, NUM_COLS, hex_to_rgb("#DEE2E6")))

    # Conditional formatting for WoW columns (cols H-I, index 7-8)
    reqs.extend(conditional_format_request(sid, data_start, data_end, 7, NUM_COLS, GREEN, RED))

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
    # Set chart axis min to round down from the lowest price, giving breathing room
    chart_axis_min = int(min(w["min_price"] for w in weeks) // 100 * 100 - 100)
    chart_req = build_price_chart_request(sid, len(weeks), TABLE_START, axis_min=chart_axis_min)
    dash_sh.batch_update({"requests": [chart_req]})
    print("  Created chart")

    print(f"  DOC Price Trends dashboard built: {len(weeks)} weeks")
    print(f"  Latest: W{latest['week']} avg price {int(latest['avg_price'])}")
