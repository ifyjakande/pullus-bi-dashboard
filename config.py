import json
import os

import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

# ---------- Config ----------
SERVICE_ACCOUNT_FILE = os.path.join(
    os.path.dirname(__file__), "pullus-pipeline-40a5302e034d.json"
)
PURCHASE_SHEET_ID = os.getenv("PURCHASE_SHEET_ID")
PURCHASE_TAB = os.getenv("PURCHASE_TRACKER_SHEET_NAME")
DASHBOARD_SHEET_ID = os.getenv("DASHBOARD_SHEET_ID")
DOC_PRICE_SHEET_ID = os.getenv("DOC_PRICE_SHEET_ID")
DOC_PRICE_SHEET_NAME = os.getenv("DOC_PRICE_SHEET_NAME")
COMPETITOR_SELLING_PRICE_SHEET_ID = os.getenv("COMPETITOR_SELLING_PRICE_SHEET_ID")
COMPETITOR_SELLING_PRICE_TABS = os.getenv("COMPETITOR_SELLING_PRICE_TABS", "").split(",")
COMPETITOR_BUYING_PRICE_SHEET_ID = os.getenv("COMPETITOR_BUYING_PRICE_SHEET_ID")
COMPETITOR_BUYING_PRICE_TAB = os.getenv("COMPETITOR_BUYING_PRICE_TAB")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ---------- Color Palette (0-1 floats) ----------
def hex_to_rgb(h):
    h = h.lstrip("#")
    return {
        "red": int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue": int(h[4:6], 16) / 255,
    }


NAVY = hex_to_rgb("#1B2A4A")
TEAL = hex_to_rgb("#2E86AB")
GREEN = hex_to_rgb("#27AE60")
RED = hex_to_rgb("#E74C3C")
LIGHT_BG = hex_to_rgb("#F8F9FA")
WHITE = hex_to_rgb("#FFFFFF")
DARK_TEXT = hex_to_rgb("#2C3E50")
CARD_BLUE = hex_to_rgb("#3498DB")
CARD_ORANGE = hex_to_rgb("#E67E22")
CARD_PURPLE = hex_to_rgb("#9B59B6")


# ---------- Authenticate ----------
def get_client():
    sa_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_json:
        info = json.loads(sa_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return gspread.authorize(creds)
