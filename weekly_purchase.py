from datetime import datetime
from collections import defaultdict

from config import (
    PURCHASE_SHEET_ID, PURCHASE_TAB,
    NAVY, TEAL, GREEN, RED, LIGHT_BG, WHITE, DARK_TEXT,
    CARD_BLUE, CARD_ORANGE, CARD_PURPLE, hex_to_rgb,
)
from helpers import (
    parse_date, safe_float,
    grid_range, cell_fmt, fmt_request, merge_request,
    col_width_request, row_height_request, border_request,
    conditional_format_request, clear_sheet,
)


# Source sheet column indices (update if sheet layout changes)
COL_DATE = 0
COL_BIRDS = 5
COL_CHICKEN_WT = 8
COL_GIZZARD_WT = 9


# ---------- Read & Aggregate ----------
def fetch_raw_data(client):
    sh = client.open_by_key(PURCHASE_SHEET_ID)
    ws = sh.worksheet(PURCHASE_TAB)
    rows = ws.get_all_values()
    header = rows[0]
    data = rows[1:]
    print(f"  Fetched {len(data)} rows, columns: {header}")
    return data


def aggregate_weekly(data):
    weeks = defaultdict(lambda: {"birds": 0, "chicken_wt": 0.0, "gizzard_wt": 0.0, "dates": []})

    current_year = datetime.now().year
    kept = 0
    for row in data:
        dt = parse_date(row[COL_DATE])
        if dt is None or dt.year != current_year:
            continue
        kept += 1
        iso_year, iso_week, _ = dt.isocalendar()
        key = (iso_year, iso_week)

        weeks[key]["birds"] += int(safe_float(row[COL_BIRDS])) if row[COL_BIRDS] else 0
        weeks[key]["chicken_wt"] += safe_float(row[COL_CHICKEN_WT])
        weeks[key]["gizzard_wt"] += safe_float(row[COL_GIZZARD_WT])
        weeks[key]["dates"].append(dt)
    print(f"  Filtered {kept} of {len(data)} rows (current year)")

    result = []
    for (yr, wk), vals in sorted(weeks.items()):
        mn = min(vals["dates"])
        mx = max(vals["dates"])
        purchase_days = len(set(d.date() for d in vals["dates"]))
        birds = vals["birds"]
        chicken_wt = round(vals["chicken_wt"], 2)
        gizzard_wt = round(vals["gizzard_wt"], 2)
        total_wt = round(vals["chicken_wt"] + vals["gizzard_wt"], 2)
        result.append({
            "year": yr,
            "week": wk,
            "start": mn,
            "end": mx,
            "purchase_days": purchase_days,
            "birds": birds,
            "chicken_wt": chicken_wt,
            "gizzard_wt": gizzard_wt,
            "total_wt": total_wt,
            "avg_birds_day": round(birds / purchase_days, 1),
            "avg_wt_day": round(total_wt / purchase_days, 2),
        })

    for i, w in enumerate(result):
        if i == 0:
            w["birds_wow"] = None
            w["wt_wow"] = None
        else:
            prev = result[i - 1]
            w["birds_wow"] = (
                round((w["avg_birds_day"] - prev["avg_birds_day"]) / prev["avg_birds_day"] * 100, 1)
                if prev["avg_birds_day"] else None
            )
            w["wt_wow"] = (
                round((w["avg_wt_day"] - prev["avg_wt_day"]) / prev["avg_wt_day"] * 100, 1)
                if prev["avg_wt_day"] else None
            )

    return result


# ---------- Chart ----------
def build_chart_request(sheet_id, weeks_count, table_start_row):
    return {
        "addChart": {
            "chart": {
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": sheet_id,
                            "rowIndex": 3,
                            "columnIndex": 10,
                        },
                        "offsetXPixels": 20,
                        "offsetYPixels": 0,
                        "widthPixels": 720,
                        "heightPixels": 420,
                    }
                },
                "spec": {
                    "title": "Weekly Birds & Chicken Weight",
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
                                "title": "Birds",
                                "format": {"fontSize": 9, "foregroundColor": DARK_TEXT},
                            },
                            {
                                "position": "RIGHT_AXIS",
                                "title": "Chicken Weight (kg)",
                                "format": {"fontSize": 9, "foregroundColor": DARK_TEXT},
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
                                "type": "COLUMN",
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
                                "targetAxis": "RIGHT_AXIS",
                                "type": "LINE",
                                "color": CARD_ORANGE,
                            },
                        ],
                        "headerCount": 1,
                    },
                },
            }
        }
    }


# ---------- Dashboard Writing ----------
def build_dashboard(dash_sh, weeks):
    ws = dash_sh.sheet1
    ws.update_title("Weekly Purchase")
    sid = ws.id

    clear_sheet(dash_sh, ws, sid)

    # ---- Prepare data ----
    latest = weeks[-1]
    prev = weeks[-2] if len(weeks) > 1 else None

    def avg_wow_pct(cur, prv, total_key):
        if prv is None:
            return "N/A"
        cur_avg = cur[total_key] / cur["purchase_days"]
        prv_avg = prv[total_key] / prv["purchase_days"]
        if prv_avg == 0:
            return "N/A"
        return round((cur_avg - prv_avg) / prv_avg * 100, 1)

    birds_wow = avg_wow_pct(latest, prev, "birds")
    chicken_wow = avg_wow_pct(latest, prev, "chicken_wt")
    gizzard_wow = avg_wow_pct(latest, prev, "gizzard_wt")

    def wow_str(v):
        if v == "N/A":
            return "N/A"
        sign = "+" if v > 0 else ""
        return f"{sign}{v}%"

    # ---- Write cell values ----
    all_values = []

    NUM_COLS = 10

    # Row 1: Title
    all_values.append(["PULLUS WEEKLY PURCHASE DASHBOARD"] + [""] * (NUM_COLS - 1))
    # Row 2: Subtitle
    first_date = weeks[0]["start"].strftime("%d %b %Y")
    last_date = weeks[-1]["end"].strftime("%d %b %Y")
    now_ts = datetime.now().strftime("%d %b %Y %I:%M %p")
    all_values.append([f"Data range: {first_date} - {last_date}  |  Last updated: {now_ts}"] + [""] * (NUM_COLS - 1))
    # Row 3: Spacer
    all_values.append([""] * NUM_COLS)

    # YTD totals
    ytd_birds = sum(w["birds"] for w in weeks)
    ytd_chicken = round(sum(w["chicken_wt"] for w in weeks), 2)
    ytd_gizzard = round(sum(w["gizzard_wt"] for w in weeks), 2)

    # Rows 4-8: Summary Cards
    all_values.append(["TOTAL BIRDS (YTD)", "", "", "TOTAL CHICKEN WT (YTD)", "", "", "TOTAL GIZZARD WT (YTD)", "", "", ""])
    all_values.append([ytd_birds, "", "", ytd_chicken, "", "", ytd_gizzard, "", "", ""])
    all_values.append(["Avg/Day WoW", "", "", "Avg/Day WoW", "", "", "Avg/Day WoW", "", "", ""])
    all_values.append([wow_str(birds_wow), "", "", wow_str(chicken_wow), "", "", wow_str(gizzard_wow), "", "", ""])
    latest_label = f"Week {latest['week']}, {latest['year']} ({latest['purchase_days']} days)"
    all_values.append([latest_label, "", "", latest_label, "", "", latest_label, "", "", ""])

    # Row 9: Spacer
    all_values.append([""] * NUM_COLS)

    # Row 10: Explainer note
    all_values.append([
        "WoW % = Week-over-Week change based on daily averages (total / purchase days), not raw totals. "
        "Weight WoW % uses combined chicken + gizzard weight."
    ] + [""] * (NUM_COLS - 1))

    # Row 11: Table header
    TABLE_START = 10
    all_values.append([
        "Week", "Date Range", "Purchase Days", "Total Birds",
        "Chicken Wt (kg)", "Gizzard Wt (kg)", "Total Wt (kg)",
        "Avg Birds/Day", "Birds Avg/Day WoW %", "Weight Avg/Day WoW %",
    ])

    # Data rows
    for w in weeks:
        dr = f"{w['start'].strftime('%d %b')} - {w['end'].strftime('%d %b %Y')}"
        b_wow = w["birds_wow"] if w["birds_wow"] is not None else ""
        w_wow = w["wt_wow"] if w["wt_wow"] is not None else ""
        if isinstance(b_wow, (int, float)):
            b_wow = round(b_wow, 1)
        if isinstance(w_wow, (int, float)):
            w_wow = round(w_wow, 1)
        all_values.append([
            f"W{w['week']}",
            dr,
            w["purchase_days"],
            w["birds"],
            w["chicken_wt"],
            w["gizzard_wt"],
            w["total_wt"],
            w["avg_birds_day"],
            b_wow,
            w_wow,
        ])

    ws.update(all_values, "A1")
    print(f"  Wrote {len(all_values)} rows of data")

    # ---- Formatting requests ----
    reqs = []

    col_widths = [70, 180, 70, 90, 120, 120, 110, 105, 120, 130]
    for i, w in enumerate(col_widths):
        reqs.append(col_width_request(sid, i, w))

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
        (0, 3, CARD_BLUE, birds_wow, "#,##0"),
        (3, 6, CARD_ORANGE, chicken_wow, "#,##0.00\" kg\""),
        (6, 9, CARD_PURPLE, gizzard_wow, "#,##0.00\" kg\""),
    ]

    for c1, c2, accent, card_wow, num_pattern in card_defs:
        reqs.append(merge_request(sid, 3, 4, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg=DARK_TEXT, bold=True, size=9, halign="CENTER")
        reqs.append(fmt_request(sid, 3, 4, c1, c2, fmt))

        reqs.append(merge_request(sid, 4, 5, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg=NAVY, bold=True, size=22, halign="CENTER",
                           pattern={"type": "NUMBER", "pattern": num_pattern})
        reqs.append(fmt_request(sid, 4, 5, c1, c2, fmt))

        reqs.append(merge_request(sid, 5, 6, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg={"red": 0.6, "green": 0.6, "blue": 0.6}, bold=False, size=8, halign="CENTER")
        reqs.append(fmt_request(sid, 5, 6, c1, c2, fmt))

        reqs.append(merge_request(sid, 6, 7, c1, c2))
        wow_color = GREEN if isinstance(card_wow, (int, float)) and card_wow > 0 else RED
        fmt = cell_fmt(bg=WHITE, fg=wow_color, bold=True, size=14, halign="CENTER")
        reqs.append(fmt_request(sid, 6, 7, c1, c2, fmt))

        reqs.append(merge_request(sid, 7, 8, c1, c2))
        fmt = cell_fmt(bg=WHITE, fg={"red": 0.6, "green": 0.6, "blue": 0.6}, bold=False, size=8, halign="CENTER")
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

    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "#,##0"})
    reqs.append(fmt_request(sid, data_start, data_end, 2, 3, fmt))

    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "#,##0"})
    reqs.append(fmt_request(sid, data_start, data_end, 3, 4, fmt))

    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "#,##0.00"})
    reqs.append(fmt_request(sid, data_start, data_end, 4, 7, fmt))

    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "#,##0.0"})
    reqs.append(fmt_request(sid, data_start, data_end, 7, 8, fmt))

    fmt = cell_fmt(pattern={"type": "NUMBER", "pattern": "+#,##0.0;-#,##0.0;\"--\""})
    reqs.append(fmt_request(sid, data_start, data_end, 8, NUM_COLS, fmt))

    reqs.append(border_request(sid, TABLE_START, data_end, 0, NUM_COLS, hex_to_rgb("#DEE2E6")))

    reqs.extend(conditional_format_request(sid, data_start, data_end, 8, NUM_COLS, GREEN, RED))

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

    chart_req = build_chart_request(sid, len(weeks), TABLE_START)
    dash_sh.batch_update({"requests": [chart_req]})
    print("  Created chart")

    print(f"  Weekly Purchase dashboard built: {len(weeks)} weeks")
    print(f"  Latest: W{latest['week']} ({latest['start'].strftime('%d %b')} - {latest['end'].strftime('%d %b %Y')})")
