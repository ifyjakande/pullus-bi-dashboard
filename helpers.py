from datetime import datetime

from config import hex_to_rgb


def parse_date(val):
    for fmt in (
        "%d-%b-%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y",
        "%d %b %Y", "%b %d, %Y", "%Y/%m/%d", "%d-%m-%Y",
        "%d %B %Y", "%B %d, %Y",
    ):
        try:
            return datetime.strptime(val.strip(), fmt)
        except (ValueError, AttributeError):
            continue
    return None


def safe_float(val):
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


# ---------- Sheets API Formatting Helpers ----------

def grid_range(sheet_id, r1, r2, c1, c2):
    return {
        "sheetId": sheet_id,
        "startRowIndex": r1,
        "endRowIndex": r2,
        "startColumnIndex": c1,
        "endColumnIndex": c2,
    }


def cell_fmt(bg=None, fg=None, bold=False, size=10, halign="LEFT", valign="MIDDLE", pattern=None, wrap=None):
    fmt = {}
    if bg:
        fmt["backgroundColor"] = bg
    if bold or size != 10:
        tf = {"bold": bold, "fontSize": size}
        if fg:
            tf["foregroundColor"] = fg
        fmt["textFormat"] = tf
    elif fg:
        fmt["textFormat"] = {"foregroundColor": fg}
    if halign != "LEFT":
        fmt["horizontalAlignment"] = halign
    if valign != "MIDDLE":
        fmt["verticalAlignment"] = valign
    if pattern:
        fmt["numberFormat"] = pattern
    if wrap is not None:
        fmt["wrapStrategy"] = "WRAP" if wrap else "CLIP"
    return fmt


def fmt_request(sheet_id, r1, r2, c1, c2, fmt):
    fields = ",".join(f"userEnteredFormat.{k}" for k in fmt)
    return {
        "repeatCell": {
            "range": grid_range(sheet_id, r1, r2, c1, c2),
            "cell": {"userEnteredFormat": fmt},
            "fields": fields,
        }
    }


def merge_request(sheet_id, r1, r2, c1, c2, merge_type="MERGE_ALL"):
    return {
        "mergeCells": {
            "range": grid_range(sheet_id, r1, r2, c1, c2),
            "mergeType": merge_type,
        }
    }


def col_width_request(sheet_id, col, px):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": col,
                "endIndex": col + 1,
            },
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def row_height_request(sheet_id, row, px):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": row,
                "endIndex": row + 1,
            },
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def border_request(sheet_id, r1, r2, c1, c2, color=None, style="SOLID"):
    if color is None:
        color = hex_to_rgb("#DEE2E6")
    border = {"style": style, "color": color}
    return {
        "updateBorders": {
            "range": grid_range(sheet_id, r1, r2, c1, c2),
            "top": border,
            "bottom": border,
            "left": border,
            "right": border,
        }
    }


def conditional_format_request(sheet_id, r1, r2, c1, c2, positive_color, negative_color):
    requests = []
    rng = grid_range(sheet_id, r1, r2, c1, c2)
    requests.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [rng],
                "booleanRule": {
                    "condition": {
                        "type": "NUMBER_GREATER",
                        "values": [{"userEnteredValue": "0"}],
                    },
                    "format": {"textFormat": {"foregroundColor": positive_color, "bold": True}},
                },
            },
            "index": 0,
        }
    })
    requests.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [rng],
                "booleanRule": {
                    "condition": {
                        "type": "NUMBER_LESS",
                        "values": [{"userEnteredValue": "0"}],
                    },
                    "format": {"textFormat": {"foregroundColor": negative_color, "bold": True}},
                },
            },
            "index": 1,
        }
    })
    return requests


def clear_sheet(dash_sh, ws, sid):
    """Clear all values, merges, formatting, charts, and conditional formats from a sheet."""
    max_rows = ws.row_count
    max_cols = ws.col_count
    ws.clear()
    dash_sh.batch_update({
        "requests": [
            {"unmergeCells": {"range": grid_range(sid, 0, max_rows, 0, max_cols)}},
            {
                "repeatCell": {
                    "range": grid_range(sid, 0, max_rows, 0, max_cols),
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 1, "green": 1, "blue": 1},
                            "textFormat": {
                                "foregroundColor": {"red": 0, "green": 0, "blue": 0},
                                "fontSize": 10,
                                "bold": False,
                            },
                            "horizontalAlignment": "LEFT",
                            "verticalAlignment": "BOTTOM",
                            "numberFormat": {"type": "TEXT"},
                            "wrapStrategy": "OVERFLOW_CELL",
                        }
                    },
                    "fields": "userEnteredFormat",
                }
            },
        ]
    })

    full_meta = dash_sh.fetch_sheet_metadata()
    sheet_meta = None
    for s in full_meta.get("sheets", []):
        if s["properties"]["sheetId"] == sid:
            sheet_meta = s
            break

    cleanup_reqs = []
    if sheet_meta and "charts" in sheet_meta:
        for c in sheet_meta["charts"]:
            cleanup_reqs.append({"deleteEmbeddedObject": {"objectId": c["chartId"]}})
    if sheet_meta and "conditionalFormats" in sheet_meta:
        for _ in sheet_meta["conditionalFormats"]:
            cleanup_reqs.append({"deleteConditionalFormatRule": {"sheetId": sid, "index": 0}})
    if cleanup_reqs:
        dash_sh.batch_update({"requests": cleanup_reqs})
