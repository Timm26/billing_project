"""
Freight Billing Report — a Streamlit dashboard.

Two self-contained reports behind one login:

  1. Billing report — joins a billing export (one row per charge line) to a
     shipment listing report (one row per shipment), renders on-screen
     summaries and produces a downloadable multi-sheet Excel report.

  2. Container breakdown — allocates every charge line across the containers on
     its shipment, groups the shipments into Domestic / Export / Import, and
     allocates each leg of each move to its Blind Purchase Order. The shipment
     listing on its own is enough to produce the structure: upload it without a
     billing export and the workbook comes out with every container against its
     group and BPO, ready for the figures.

Everything deployment-specific — app title, theme colours, currency codes, the
source-system column names, the charge-code-to-leg map and the BPO table —
lives in the CONFIG blocks below. Nothing in the logic assumes a particular
organisation, so this file can be published as is and adapted by changing the
constants.
"""

import hashlib
import hmac
import io
import re
from string import Template

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────

APP_TITLE = "Freight Billing Report"
APP_SUBTITLE = "Shipment & Charge Consolidation"
REPORT_FILENAME = "billing_report.xlsx"

# Currency the report is totalled in, and the secondary currency reported
# alongside it. Amounts in any currency still roll up into the base via the
# billing export's own local-total column.
BASE_CCY = "AUD"
FX_CCY = "USD"
LOCAL_LABEL = f"Local Total ({BASE_CCY})"

# Names and formatting that appear in the exported workbook. These are kept
# byte-identical to the established report format so downstream spreadsheets
# and any formulas pointing at them keep working.
SUPPLIER_TOTAL_COL = f"Local_Total_{BASE_CCY}"
REPORT_HEADER_FILL = "1A56A0"  # Analysis section header fill

# Column names as they appear in the two source exports.
BILLING_JOB_COL = "Job"          # job/shipment key in the billing export
SHIPMENT_KEY_COL = "Shipment"    # job/shipment key in the shipment listing
SHIPMENT_USECOLS = range(1, 22)  # columns to read from the shipment listing
JOB_ID_PATTERN = r"^[A-Za-z]{1,3}\d+"  # what a job reference looks like

# Source column -> internal name. Add or edit entries to match your export.
SHIPMENT_COLUMN_MAP = {
    SHIPMENT_KEY_COL: "Shipment Job",
    "Order Ref": "Order Reference",
    "INCO": "Incoterms",
    "Pack Mode": "Mode",
    "Consignor Name": "Supplier Name",
    "Origin": "Loading Port",
    "Dest.": "Destination Port",
    "Carrier Booking Reference": "Booking Ref",
    "Vessel": "Vessel / Voyage",
    "No. of Cont.": "Container Count",
    "Container #": "Containers",
}
DATE_COLS = ("ETD", "ETA", "ATD", "ATA")

# Theme — neutral defaults, no brand assets or brand colours.
PRIMARY = "#33475B"
PRIMARY_LIGHT = "#48627E"
ACCENT = "#0E8074"
ACCENT_DARK = "#0A6158"
BG = "#EDF1F5"
SURFACE = "#FFFFFF"
TEXT = "#1B2A38"
MUTED = "#7A8CA0"
BORDER = "#D3DCE6"
GRID = "#E6EDF3"
CONTROL_BG = "#3E5670"

PALETTE = [ACCENT, PRIMARY, "#5B8FB9", "#8FBF9F", "#B08EA2", "#C98B5E", "#6E7B8B"]

st.set_page_config(
    page_title=APP_TITLE,
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Styling ───────────────────────────────────────────────────────────────────

CSS = Template("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background-color: $bg; color: $text; }
.stMultiSelect [data-baseweb="tag"] { background-color: $accent !important; color: #ffffff !important; }
.app-header { background: linear-gradient(90deg,$primary 0%,$primary_light 100%); border-bottom:4px solid $accent; padding:20px 32px; margin:-1rem -1rem 2rem -1rem; display:flex; align-items:center; gap:20px; }
.app-header h1 { font-size:1.7rem; font-weight:700; color:#fff; margin:0; letter-spacing:.01em; }
.app-subtitle { font-size:.78rem; color:#c8d6e5; letter-spacing:.08em; text-transform:uppercase; margin-top:4px; }
.section-title { font-size:.95rem; font-weight:600; color:$primary; text-transform:uppercase; letter-spacing:.08em; border-left:4px solid $accent; padding-left:10px; margin:20px 0 10px 0; }
[data-testid="stDataFrame"] { border:1px solid $border; border-radius:6px; background-color:$surface; }
.stTabs [data-baseweb="tab-list"] { background-color:$surface; border-bottom:2px solid $border; gap:0; }
.stTabs [data-baseweb="tab"] { font-weight:600; font-size:.85rem; letter-spacing:.04em; text-transform:uppercase; color:$muted !important; background:transparent !important; border:none !important; padding:12px 24px; }
.stTabs [aria-selected="true"] { color:$primary !important; border-bottom:3px solid $accent !important; }
.stDownloadButton button, .stButton button { background-color:$accent !important; color:#fff !important; font-weight:600 !important; letter-spacing:.06em !important; text-transform:uppercase !important; border:none !important; border-radius:4px !important; padding:10px 28px !important; }
.stDownloadButton button:hover, .stButton button:hover { background-color:$accent_dark !important; }
[data-testid="stFileUploader"] { background-color:$surface; border:2px dashed $border; border-radius:6px; padding:10px; }
[data-testid="metric-container"] { background-color:$surface; border:1px solid $border; border-top:4px solid $accent; padding:16px; border-radius:6px; box-shadow:0 2px 8px rgba(27,42,56,.08); }
[data-testid="stMetricValue"] { font-size:1.8rem !important; font-weight:700 !important; color:$primary !important; }
[data-testid="stMetricLabel"] { color:$muted !important; font-size:.7rem !important; text-transform:uppercase; letter-spacing:.1em; }
div[data-testid="stVerticalBlock"] { gap:0.5rem; }
.stAlert { background-color:$grid !important; border-color:$primary !important; color:$text !important; }
[data-testid="stToolbar"] { display: none !important; }
header[data-testid="stHeader"] { background: transparent !important; }
/* No sidebar in this app: hide it and its expand control entirely */
[data-testid="stSidebar"],
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"] { display: none !important; }
</style>
""").safe_substitute(
    bg=BG, text=TEXT, primary=PRIMARY, primary_light=PRIMARY_LIGHT,
    accent=ACCENT, accent_dark=ACCENT_DARK, surface=SURFACE, muted=MUTED,
    border=BORDER, grid=GRID, control=CONTROL_BG,
)
st.markdown(CSS, unsafe_allow_html=True)

PLOTLY_LAYOUT = dict(
    paper_bgcolor=SURFACE, plot_bgcolor="#F7FAFC",
    font=dict(family="Inter, sans-serif", color=TEXT, size=11),
    xaxis=dict(gridcolor=GRID, linecolor=BORDER, tickcolor=TEXT, tickfont=dict(color=TEXT)),
    yaxis=dict(gridcolor=GRID, linecolor=BORDER, tickcolor=TEXT, tickfont=dict(color=TEXT)),
    margin=dict(l=20, r=20, t=36, b=20),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=TEXT)),
)


# ── Access control ────────────────────────────────────────────────────────────
#
# Credentials are NEVER stored in this file. They live in .streamlit/secrets.toml
# locally (git-ignored) or in the host's secrets manager when deployed, as
# salted PBKDF2-SHA256 hashes:
#
#   [users]
#   someuser = "pbkdf2_sha256$240000$<salt hex>$<hash hex>"
#
# Generate a hash with tools/hash_password.py. If no users are configured the
# app refuses all access rather than falling open.

MAX_ATTEMPTS = 6

# Tab ids used in the [access] section of the secrets file, in display order.
TAB_BILLING = "billing"
TAB_CONTAINER = "container"
TAB_LABELS = {
    TAB_BILLING: "Billing report",
    TAB_CONTAINER: "Container breakdown",
}


def load_users():
    """Username -> password hash, keyed on the lowercased username.

    Usernames are matched case-insensitively, so SeaImports, seaimports and
    SEAIMPORTS are the same account. Passwords remain case-sensitive.
    """
    try:
        return {str(name).strip().lower(): value
                for name, value in dict(st.secrets["users"]).items()}
    except Exception:
        return {}


def load_display_names():
    """Lowercased username -> the spelling used in the secrets file."""
    try:
        return {str(name).strip().lower(): str(name).strip()
                for name in dict(st.secrets["users"])}
    except Exception:
        return {}


def load_access():
    """Lowercased username -> list of tab ids that account may see.

    Configured in an [access] section, one comma-separated list per user:

        [access]
        someuser = "billing"
        another  = "billing, container"

    A user with no entry sees every tab.
    """
    try:
        raw = dict(st.secrets["access"])
    except Exception:
        return {}
    access = {}
    for name, value in raw.items():
        if isinstance(value, str):
            tabs = [part.strip().lower() for part in value.split(",") if part.strip()]
        else:
            tabs = [str(part).strip().lower() for part in value]
        access[str(name).strip().lower()] = tabs
    return access


def allowed_tabs(username):
    """Tab ids this account may see, in canonical order."""
    granted = load_access().get(str(username).strip().lower())
    if not granted:
        return list(TAB_LABELS)
    return [tab for tab in TAB_LABELS if tab in granted]


def password_matches(stored, supplied):
    """Constant-time check of a supplied password against a stored hash."""
    try:
        scheme, iterations, salt_hex, digest_hex = str(stored).split("$")
        if scheme != "pbkdf2_sha256":
            return False
        calculated = hashlib.pbkdf2_hmac(
            "sha256", supplied.encode("utf-8"), bytes.fromhex(salt_hex), int(iterations)
        )
    except (ValueError, TypeError):
        return False
    return hmac.compare_digest(calculated.hex(), digest_hex)


def login_screen():
    """Render the sign-in form. Never returns if the visitor is not signed in."""
    st.markdown(f"""
    <div class="app-header">
      <div>
        <h1>{APP_TITLE}</h1>
        <div class="app-subtitle">Sign in to continue</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    users = load_users()
    _, mid, _ = st.columns([1, 1.4, 1])
    with mid:
        if not users:
            st.error("No accounts are configured. Set the [users] section in the app's secrets.")
            st.stop()

        if st.session_state.get("login_attempts", 0) >= MAX_ATTEMPTS:
            st.error("Too many failed attempts. Reload the page to try again.")
            st.stop()

        with st.form("sign_in"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign in")

        if submitted:
            key = username.strip().lower()
            stored = users.get(key)
            if stored is not None and password_matches(stored, password):
                st.session_state["auth_user"] = key
                st.session_state["auth_display"] = load_display_names().get(key, key)
                st.session_state["login_attempts"] = 0
                st.rerun()
            st.session_state["login_attempts"] = st.session_state.get("login_attempts", 0) + 1
            # Deliberately vague: don't reveal which of the two was wrong.
            st.error("Incorrect username or password.")

        st.caption("Access is restricted. Contact the report owner if you need an account.")
    st.stop()


def require_login():
    if not st.session_state.get("auth_user"):
        login_screen()


require_login()


# ── Loading ───────────────────────────────────────────────────────────────────

def read_billing(file_obj):
    try:
        return pd.read_excel(file_obj)
    except Exception as e:
        st.error(f"Could not read the billing export: {e}")
        return None


def read_shipment(file_obj):
    """Read the shipment listing, locating the header row before parsing."""
    try:
        raw = pd.read_excel(file_obj, header=None)
        header_row = 0
        for i, row in raw.iterrows():
            vals = [str(v).strip() for v in row if str(v).strip() not in ("", "nan")]
            if SHIPMENT_KEY_COL in vals:
                header_row = i
                break
        file_obj.seek(0)
        df = pd.read_excel(file_obj, header=header_row, usecols=SHIPMENT_USECOLS)
        df.columns = [str(c).strip() for c in df.columns]
        if SHIPMENT_KEY_COL in df.columns:
            df = df[df[SHIPMENT_KEY_COL].astype(str).str.match(JOB_ID_PATTERN)]
            df = df.reset_index(drop=True)
            df = df.rename(columns={k: v for k, v in SHIPMENT_COLUMN_MAP.items() if k in df.columns})
            for col in DATE_COLS:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce")
            if "Container Count" in df.columns:
                df["Container Count"] = (
                    pd.to_numeric(df["Container Count"], errors="coerce").fillna(0).astype(int)
                )
        return df
    except Exception as e:
        st.error(f"Could not read the shipment listing: {e}")
        return None


def load_data(billing_files, shipment_file):
    sheets = [read_billing(f) for f in billing_files]
    sheets = [s for s in sheets if s is not None]
    if not sheets:
        return None
    billing = pd.concat(sheets, ignore_index=True)
    shipment = read_shipment(shipment_file)
    if shipment is None:
        return None
    merged = pd.merge(
        billing, shipment, how="left",
        left_on=BILLING_JOB_COL, right_on="Shipment Job",
    )
    return merged, shipment


# ── Table builders ────────────────────────────────────────────────────────────

def build_shipment_summary(shipment_df):
    cols = ["Shipment Job", "Supplier Name", "Loading Port", "Destination Port",
            "Order Reference", "Incoterms", "Containers", "Container Count",
            "Mode", "Vessel / Voyage", "Booking Ref", *DATE_COLS]
    return shipment_df[[c for c in cols if c in shipment_df.columns]].copy().reset_index(drop=True)


def build_billing_detail(data):
    cols = ["Shipment Job", "Order Reference", "Supplier Name", "Loading Port",
            "Destination Port", "Incoterms", "Containers", "Description", "Currency",
            "Amount", "Tax", "Total", "Local Total", "Exchange Rate"]
    return data[[c for c in cols if c in data.columns]].copy().reset_index(drop=True)


def build_billing_summary(data, shipment_df):
    base = (data[data["Currency"] == BASE_CCY].groupby("Shipment Job")["Total"].sum()
            .reset_index().rename(columns={"Total": BASE_CCY}))
    fx = (data[data["Currency"] == FX_CCY].groupby("Shipment Job")["Total"].sum()
          .reset_index().rename(columns={"Total": FX_CCY}))
    local = (data.groupby("Shipment Job")["Local Total"].sum()
             .reset_index().rename(columns={"Local Total": LOCAL_LABEL}))
    cur = pd.merge(base, fx, how="outer", on="Shipment Job").fillna(0)
    cur = pd.merge(cur, local, how="left", on="Shipment Job").fillna(0)
    ctx_cols = ["Shipment Job", "Supplier Name", "Loading Port", "Destination Port",
                "Order Reference", "Incoterms", "Containers"]
    ctx = shipment_df[[c for c in ctx_cols if c in shipment_df.columns]].drop_duplicates("Shipment Job")
    return pd.merge(ctx, cur, how="right", on="Shipment Job").reset_index(drop=True)


def build_supplier_summary(data, shipment_df):
    bd = build_billing_detail(data)
    sup = bd.groupby("Supplier Name").agg(
        Shipments=("Shipment Job", "nunique"),
        Charge_Lines=("Shipment Job", "count"),
        **{SUPPLIER_TOTAL_COL: ("Local Total", "sum")},
    ).reset_index()
    if "Container Count" in shipment_df.columns:
        cnt = shipment_df.groupby("Supplier Name")["Container Count"].sum().reset_index()
        sup = pd.merge(sup, cnt, how="left", on="Supplier Name")
    return sup.sort_values(SUPPLIER_TOTAL_COL, ascending=False).reset_index(drop=True)


def add_month_col(df, ship_sum):
    """Add a Month label and sort key to a billing detail frame, keyed off ETD."""
    if "ETD" not in ship_sum.columns:
        return df, []
    etd_map = ship_sum.set_index("Shipment Job")["ETD"]
    out = df.copy()
    out["_etd"] = pd.to_datetime(out["Shipment Job"].map(etd_map), errors="coerce")
    out["_sort"] = out["_etd"].dt.to_period("M").dt.to_timestamp()
    out["Month"] = out["_etd"].dt.strftime("%b %Y")
    order = out[["Month", "_sort"]].drop_duplicates().sort_values("_sort")["Month"].tolist()
    return out, order


def split_amount(total, parts):
    """Split an amount into `parts` shares that sum back exactly to the total.

    Works in whole cents and hands the leftover cents to the earliest shares
    (largest-remainder), so 218.00 over 11 containers becomes nine at 19.82
    and two at 19.81 rather than eleven at 19.82.
    """
    if parts <= 0:
        return []
    cents = int(round(float(total) * 100))
    sign = -1 if cents < 0 else 1
    base, remainder = divmod(abs(cents), parts)
    return [sign * (base + (1 if i < remainder else 0)) / 100 for i in range(parts)]


# ── Project 2 helpers: per-container charge breakdown ────────────────────────
#
# A different client with a different shipment listing layout: the header sits
# part-way down the sheet, the job reference column is named differently, and
# container numbers carry their size. Nothing here is shared with project 1, so
# changes on this side cannot affect the billing report.

CB_LOCAL_CCY = "NZD"            # currency the export's "Local" columns are in
CB_CHARGE_CODE_COL = "Charges"
CB_JOB_COL = "Job"
CB_REPORT_FILENAME = "container_charge_breakdown.xlsx"
CB_NO_CONTAINER = "(no container listed)"

# A job reference is a single letter followed by digits (S04944911). Carrier
# and B/L references carry two or three letters, so this discriminates.
CB_JOB_PATTERN = r"^[A-Za-z]\d{6,}$"

# File types the container tab accepts. Legacy .xls (what CargoWise exports as
# "XLS") needs the xlrd package installed alongside openpyxl.
CB_UPLOAD_TYPES = ["xlsx", "xls", "csv"]

# A complete ISO 6346 container number: four letters and seven digits. Anything
# shorter that starts the same way is a number the export cut off.
CB_ISO_CONTAINER = r"^[A-Z]{4}\d{7}$"
CB_TRUNCATED_CONTAINER = r"^[A-Z]{4}\d{0,6}$"

# Generic markers used to locate the header row without relying on any
# organisation-specific column name.
CB_HEADER_MARKERS = ("Container #", "No. of Cont.", "INCO", "Incoterm", "Origin",
                     "Dest.", "Pack Mode", "Consignor Name")

# Source column -> internal name for the shipment listing.
CB_SHIPMENT_MAP = {
    "Consignor Name": "Consignor",
    "Goods Description": "Goods",
    "Trans": "Transport",
    "Pack Mode": "Mode",
    "Shipping Line": "Shipping Line",
    "Carrier Master B/L": "Master B/L",
    "Origin": "Origin",
    "Dest.": "Destination",
    "Final Discharge Port": "Discharge Port",
    "Depart Vessel": "Vessel",
    "Load ETD": "ETD",
    "Final Discharge Port ETA": "ETA",
    "Final Discharge Port ATA": "ATA",
    "INCO": "Incoterms",
    "Incoterm": "Incoterms",
    "No. of Cont.": "Container Count",
    "TEU": "TEU",
    "Container #": "Container List",
    "Entry Ref": "Entry Ref",
    "Import Cartage Company": "Cartage Company",
    "Delivery Address": "Delivery Address",
    "Weight": "Weight",
    "Order Ref": "Order Reference",
}

CB_CONTEXT_COLS = ["Consignor", "Origin", "Destination", "Discharge Port",
                   "Vessel", "Shipping Line", "Mode", "Incoterms", "ETD", "ETA"]


def cb_find_header_row(raw, min_markers=3):
    """Row index whose cells contain at least `min_markers` known headers."""
    for i, row in raw.iterrows():
        cells = {str(v).strip() for v in row.tolist()}
        if sum(marker in cells for marker in CB_HEADER_MARKERS) >= min_markers:
            return i
    return 0


def cb_detect_job_column(df):
    """Column holding shipment job references, found by value shape not by name.

    Some listings carry the job reference in two columns (a shipment column and
    a house-bill column holding the same value). When scores tie, a column whose
    name says it is the shipment or job reference wins over one that happens to
    hold the same values.
    """
    scores = {}
    for col in df.columns:
        values = df[col].dropna().astype(str).str.strip()
        if values.empty:
            continue
        score = int(values.str.match(CB_JOB_PATTERN).sum())
        if score:
            scores[col] = score
    if not scores:
        return None
    top = max(scores.values())
    leaders = [c for c, v in scores.items() if v == top]
    for hint in ("shipment", "job", "ref"):
        for col in leaders:
            if hint in str(col).lower():
                return col
    return leaders[0]


def cb_read_table(file_obj, header=None):
    """Read an uploaded .xlsx, .xls or .csv from the first sheet."""
    name = str(getattr(file_obj, "name", "")).lower()
    file_obj.seek(0)
    if name.endswith(".csv"):
        return pd.read_csv(file_obj, header=header)
    return pd.read_excel(file_obj, header=header)


def cb_read_shipment(file_obj):
    """Read one shipment listing into normalised columns.

    Handles the different listing layouts the system produces: the consignee
    view (job reference in its own column, 'Load ETD', 'INCO') and the origin
    view (job reference under 'Shipment', 'Origin ETD', 'Incoterm'), in .xlsx,
    legacy .xls or .csv.
    """
    raw = cb_read_table(file_obj, header=None)
    header_row = cb_find_header_row(raw)
    df = cb_read_table(file_obj, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

    job_col = cb_detect_job_column(df)
    if job_col is None:
        return None, None
    df = df[df[job_col].astype(str).str.strip().str.match(CB_JOB_PATTERN)].copy()
    if job_col != "Job" and "Job" in df.columns:
        df = df.drop(columns=["Job"])
    df = df.rename(columns={job_col: "Job"})

    # Rename to the internal names. A listing can already hold a column under a
    # target name with different content (the origin view has 'Consignor' as a
    # code beside 'Consignor Name'); the mapped source column wins, so the
    # clashing one is dropped first rather than producing two 'Consignor's.
    renames = {k: v for k, v in CB_SHIPMENT_MAP.items() if k in df.columns}
    clashes = [v for k, v in renames.items()
               if v != k and v in df.columns and v not in renames]
    df = df.drop(columns=clashes).rename(columns=renames)
    df = df.loc[:, ~df.columns.duplicated()]
    df["Job"] = df["Job"].astype(str).str.strip()

    # Layouts that carry no load-port ETD or discharge-port ETA fall back to the
    # shipment-level dates, which are the same event on a direct sailing.
    if "ETD" not in df.columns and "Origin ETD" in df.columns:
        df["ETD"] = df["Origin ETD"]
    if "ETA" not in df.columns:
        for fallback in ("Disch. ETA", "Dest. ETA"):
            if fallback in df.columns:
                df["ETA"] = df[fallback]
                break

    for col in ("ETD", "ETA", "ATA"):
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    for col in ("Container Count", "TEU", "Weight"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    keep = ["Job"] + [c for c in CB_SHIPMENT_MAP.values() if c in df.columns]
    df = df[[c for c in dict.fromkeys(keep)]]
    return df.drop_duplicates("Job").reset_index(drop=True), job_col


def cb_parse_containers(value):
    """'CAAU6057065 (40HC), CMAU8473940 (40HC)' -> [(number, size), ...].

    Copes with what the listing export does to long lists: numbers broken by a
    line feed ('BEAU44331\\n98 (40HC)'), a list cut off part-way through a
    number ('TXGU84521'), and bookings whose numbers are not issued yet (' (20GP)').
    Returns (entries, sizes): the containers that could be read, and every size
    seen, so a short list can be padded out to the shipment's container count.
    """
    if pd.isna(value):
        return [], []
    entries, sizes = [], []
    for part in str(value).replace(";", ",").split(","):
        if not part.strip() or part.strip().lower() == "nan":
            continue
        number, _, rest = part.partition("(")
        size = rest.split(")")[0].strip() if rest else ""
        number = "".join(number.split()).upper()   # drop embedded line breaks
        if size:
            sizes.append(size)
        if not number:
            continue                                 # booked, no number yet
        if re.match(CB_ISO_CONTAINER, number):
            entries.append((number, size))
        elif re.match(CB_TRUNCATED_CONTAINER, number):
            entries.append((f"{number}… (cut off in listing)", size))
        else:
            entries.append((number, size))           # e.g. LOOSE on LCL / air
    return entries, sizes


def cb_container_table(shipment_df):
    """One row per container, with its shipment context.

    The listing's container column is capped in length, so on large shipments
    it lists fewer containers than the shipment carries. Where the container
    count says there are more, the missing ones are added as numbered
    placeholders so charges are split across the real number of containers.
    """
    rows = []
    for _, shipment in shipment_df.iterrows():
        job = shipment["Job"]
        containers, sizes = cb_parse_containers(shipment.get("Container List"))

        count = pd.to_numeric(shipment.get("Container Count"), errors="coerce")
        count = int(count) if pd.notna(count) and count > 0 else 0
        if count > len(containers):
            size = max(set(sizes), key=sizes.count) if sizes else ""
            for k in range(len(containers) + 1, count + 1):
                containers.append((f"{job} #{k:02d} (not listed)", size))

        if not containers:
            containers = [(CB_NO_CONTAINER, "")]
        for position, (number, size) in enumerate(containers, start=1):
            row = {"Job": job, "Container": number,
                   "Container Size": size, "Position": position}
            for col in CB_CONTEXT_COLS:
                if col in shipment_df.columns:
                    row[col] = shipment.get(col)
            rows.append(row)
    return pd.DataFrame(rows)


def cb_read_shipments(files):
    """Read several listings and stack them, first listing winning on a clash.

    The consignee view and the origin view each carry part of the picture (the
    domestic legs in one, the exports in the other), so both can be uploaded
    together.
    """
    frames, job_cols, errors = [], [], []
    for uploaded in files:
        try:
            frame, job_col = cb_read_shipment(uploaded)
        except Exception as e:
            errors.append(f"{uploaded.name}: {e}")
            continue
        if frame is None or frame.empty:
            errors.append(f"{uploaded.name}: no shipment rows found")
            continue
        frames.append(frame)
        job_cols.append(f"{uploaded.name} → '{job_col}'")
    if not frames:
        return None, job_cols, errors
    combined = pd.concat(frames, ignore_index=True)
    combined = combined.drop_duplicates("Job", keep="first").reset_index(drop=True)
    return combined, job_cols, errors


def cb_build_charge_detail(billing_df, shipment_df):
    """Explode every charge line across the containers on its job.

    Returns (detail, diagnostics). Each charge line's amounts are split evenly
    across the job's containers in whole cents, so the parts add back exactly.
    """
    diagnostics = {"missing_columns": [], "unmatched_jobs": [],
                   "jobs_without_containers": []}
    for col in (CB_CHARGE_CODE_COL, CB_JOB_COL, "Local Amount"):
        if col not in billing_df.columns:
            diagnostics["missing_columns"].append(col)
    if diagnostics["missing_columns"]:
        return pd.DataFrame(), diagnostics

    containers = cb_container_table(shipment_df)
    by_job = {job: group for job, group in containers.groupby("Job")}

    numeric_cols = [c for c in ("Amount", "Tax", "Total", "Local Amount", "Local Total")
                    if c in billing_df.columns]

    records = []
    for line in billing_df.to_dict("records"):
        job = str(line[CB_JOB_COL]).strip()
        group = by_job.get(job)
        if group is None:
            if job not in diagnostics["unmatched_jobs"]:
                diagnostics["unmatched_jobs"].append(job)
            continue
        if list(group["Container"]) == [CB_NO_CONTAINER]:
            if job not in diagnostics["jobs_without_containers"]:
                diagnostics["jobs_without_containers"].append(job)

        parts = len(group)
        shares = {col: split_amount(pd.to_numeric(line.get(col), errors="coerce") or 0.0, parts)
                  for col in numeric_cols}

        for i, (_, container) in enumerate(group.iterrows()):
            record = {
                "Job": job,
                "Container": container["Container"],
                "Container Size": container["Container Size"],
                "Charge Code": str(line[CB_CHARGE_CODE_COL]).strip().upper(),
                "Description": line.get("Description"),
                "Currency": line.get("Currency"),
                "Containers on Job": parts,
            }
            for col in CB_CONTEXT_COLS:
                if col in container.index:
                    record[col] = container[col]
            record["Line Total (ex tax)"] = pd.to_numeric(
                line.get("Local Amount"), errors="coerce")
            for col in numeric_cols:
                record[f"Allocated {col}"] = shares[col][i]
            if "Tax Date" in line:
                record["Charge Date"] = pd.to_datetime(line["Tax Date"], errors="coerce")
            records.append(record)

    if not records:
        return pd.DataFrame(), diagnostics

    detail = pd.DataFrame(records)
    ordered = ["Job", "Container", "Container Size", "Charge Code", "Description",
               "Currency", "Containers on Job", "Line Total (ex tax)"]
    allocated = [c for c in detail.columns if c.startswith("Allocated")]
    context = [c for c in CB_CONTEXT_COLS if c in detail.columns]
    tail = [c for c in ("Charge Date",) if c in detail.columns]
    detail = detail[ordered + allocated + context + tail]
    return (detail.sort_values(["Job", "Container", "Charge Code"])
            .reset_index(drop=True), diagnostics)


CB_ALLOC_EX = "Allocated Local Amount"
CB_ALLOC_TAX = "Allocated Local Tax"
CB_ALLOC_INC = "Allocated Local Total"


# ── Trade groups and BPO allocation ───────────────────────────────────────────
#
# Every shipment falls into one trade group, judged by direction against the
# home country:
#
#     NZ  ->  NZ          Domestic
#     NZ  ->  anywhere    Export
#     anywhere  ->  NZ    Import
#
# and every charge line falls into one leg of the move (Door to Wharf / Ocean
# freight / Wharf to Door). The group and the leg together decide which Blind
# Purchase Order the charge is billed against, and which entity and vendor code
# sit behind it. A BPO that has not been issued yet is carried as the reference
# "TBA" and behaves like any other BPO until the number arrives.
#
# Rows are split by charge currency and by BPO, so every row on the container
# matrix belongs to exactly one BPO and can be invoiced as it stands. Each row
# also carries its local equivalent, which is what the report reconciles on.
#
# Three tables normally need editing: BPO_HOME_COUNTRY / BPO_GROUP_RULES,
# BPO_LEG_BY_CODE and BPO_TABLE. Anything a rule does not recognise is never
# dropped — an unrecognised lane lands in BPO_GROUP_OTHER and an unknown charge
# code lands in BPO_LEG_UNMAPPED, both on a TBA reference and flagged on screen.

BPO_HOME_COUNTRY = "NZ"  # UN/LOCODE prefix the groups are judged against

BPO_GROUP_DOMESTIC = "Domestic"
BPO_GROUP_EXPORT = "Export"
BPO_GROUP_IMPORT = "Import"
BPO_GROUP_OTHER = "Other"

BPO_GROUP_ORDER = [BPO_GROUP_DOMESTIC, BPO_GROUP_EXPORT,
                   BPO_GROUP_IMPORT, BPO_GROUP_OTHER]

# Overrides for named lane pairs, if one ever has to sit outside the direction
# rule: {("AU", "NZ"): BPO_GROUP_IMPORT}. Normally left empty.
BPO_GROUP_RULES = {}

# Who receives the invoice for each group.
BPO_INVOICE_TO = {
    BPO_GROUP_DOMESTIC: "NZ (Eileen)",
    BPO_GROUP_EXPORT: "NZ (Eileen)",
    BPO_GROUP_IMPORT: "AU (Hayley)",
    BPO_GROUP_OTHER: "TBA",
}

BPO_LEG_ORIGIN = "Door to Wharf"
BPO_LEG_FREIGHT = "Ocean freight"
BPO_LEG_DEST = "Wharf to Door"
BPO_LEG_UNMAPPED = "Unmapped"

BPO_LEG_ORDER = [BPO_LEG_ORIGIN, BPO_LEG_FREIGHT, BPO_LEG_DEST, BPO_LEG_UNMAPPED]
BPO_BILLABLE_LEGS = [BPO_LEG_ORIGIN, BPO_LEG_FREIGHT, BPO_LEG_DEST]

# Charge code -> leg. Codes are matched upper-cased and stripped.
BPO_LEG_BY_CODE = {
    # Pick-up and origin side: everything up to the ship's rail.
    "OTHC": BPO_LEG_ORIGIN,       # Origin Terminal Handling
    "ODOC": BPO_LEG_ORIGIN,       # Origin Documentation Fee
    "OCART": BPO_LEG_ORIGIN,      # Origin cartage, if it ever appears
    "OCARTFSC": BPO_LEG_ORIGIN,
    "LOAD": BPO_LEG_ORIGIN,
    "EXPENT": BPO_LEG_ORIGIN,     # Export entry
    "OPUP": BPO_LEG_ORIGIN,       # Pick up Charges
    "OFSC": BPO_LEG_ORIGIN,       # Fuel Surcharge on Pick Up
    "OPRA": BPO_LEG_ORIGIN,       # Booking Fee - Origin
    "OVBF": BPO_LEG_ORIGIN,       # Vehicle Booking Fee - Origin
    "OPCH": BPO_LEG_ORIGIN,       # Origin Port Charges
    "OTSS": BPO_LEG_ORIGIN,       # Port Security Fee - Origin
    "OHDL": BPO_LEG_ORIGIN,       # Rail Handling Fee at the load port
    "OBLF": BPO_LEG_ORIGIN,       # Bill of Lading Fee Export
    "OCCL": BPO_LEG_ORIGIN,       # Export Customs Clearance
    "FMF": BPO_LEG_ORIGIN,        # Freight Management Fee (NZ-side fee; confirm)

    # The ocean move itself and the surcharges that ride on the freight rate.
    "FRT": BPO_LEG_FREIGHT,       # Ocean / International Freight
    "DFRT": BPO_LEG_FREIGHT,      # Domestic (coastal) Freight
    "BAF": BPO_LEG_FREIGHT,       # Bunker Adjustment Factor
    "CAF": BPO_LEG_FREIGHT,
    "PSS": BPO_LEG_FREIGHT,
    "EBS": BPO_LEG_FREIGHT,
    "GRI": BPO_LEG_FREIGHT,
    "ISPS": BPO_LEG_FREIGHT,

    # Discharge port through to the delivery door, including clearance.
    "DTHC": BPO_LEG_DEST,         # Destination Terminal Handling
    "DDOC": BPO_LEG_DEST,         # Destination Documentation Fee
    "DDOF": BPO_LEG_DEST,         # Destination Delivery Order Fee
    "DTIF": BPO_LEG_DEST,         # Terminal Infrastructure Fee
    "DCART": BPO_LEG_DEST,        # Cartage to the delivery address
    "DCARTFSC": BPO_LEG_DEST,     # Cartage Fuel Surcharge
    "DEMUR": BPO_LEG_DEST,        # Container Demurrage
    "DWASH": BPO_LEG_DEST,        # Container Wash
    "CUSENT": BPO_LEG_DEST,       # Customs Entry Lodgement
    "MPIENT": BPO_LEG_DEST,       # Biosecurity Entry Fee
    "MPIINS": BPO_LEG_DEST,       # Biosecurity Inspection
    "AGENCY": BPO_LEG_DEST,       # Agency Fee
    "DHDL": BPO_LEG_DEST,         # Handling Fee - Destination
    "DADM": BPO_LEG_DEST,         # Administration Fee - Destination
    "STORAGE": BPO_LEG_DEST,
    "UNPACK": BPO_LEG_DEST,
}

BPO_TBA = "TBA"


def _bpo_record(bpo, currency, entity, vendor):
    return {"BPO": bpo, "BPO Currency": currency,
            "BPO Entity": entity, "BPO Vendor": vendor}


BPO_UNALLOCATED = _bpo_record(BPO_TBA, CB_LOCAL_CCY, BPO_TBA, BPO_TBA)

# (group, leg) -> BPO record. A pair not listed here is carried on TBA.
BPO_TABLE = {
    # Coastal NZ (AKL to CHC, ANL & Maersk), invoiced to NZ. One BPO covers the
    # whole domestic move, delivery leg included.
    (BPO_GROUP_DOMESTIC, BPO_LEG_ORIGIN):
        _bpo_record("77263466-330920", "NZD", "Rohlig NZ", "30098"),
    (BPO_GROUP_DOMESTIC, BPO_LEG_FREIGHT):
        _bpo_record("77263466-330920", "NZD", "Rohlig NZ", "30098"),
    (BPO_GROUP_DOMESTIC, BPO_LEG_DEST):
        _bpo_record("77263466-330920", "NZD", "Rohlig NZ", "30098"),

    # NZ to overseas, invoiced to NZ.
    (BPO_GROUP_EXPORT, BPO_LEG_ORIGIN):
        _bpo_record("77263466-330920", "NZD", "Rohlig NZ", "30098"),
    (BPO_GROUP_EXPORT, BPO_LEG_FREIGHT):
        _bpo_record("77263749-330924", "USD", "Rohlig NZ", "62170"),
    (BPO_GROUP_EXPORT, BPO_LEG_DEST):
        _bpo_record("77260204-330932", "AUD", "Rohlig AU", "61935"),

    # Overseas to NZ, invoiced to AU. The Wharf to Door BPO has not been issued
    # yet, so it is carried as TBA and treated like any other BPO.
    (BPO_GROUP_IMPORT, BPO_LEG_ORIGIN):
        _bpo_record("77250782-714720", "AUD", "Rohlig AU", "29803"),
    (BPO_GROUP_IMPORT, BPO_LEG_FREIGHT):
        _bpo_record("77254657-709817", "USD", "Rohlig AU", "61934"),
    (BPO_GROUP_IMPORT, BPO_LEG_DEST):
        _bpo_record(BPO_TBA, "NZD", "Rohlig NZ", "30494"),
}

# Amounts on the charge detail. The first pair is in the charge's own currency,
# the second is the local equivalent everything reconciles on.
CB_AMOUNT_CCY = "Allocated Amount"
CB_INC_CCY = "Allocated Total"

CB_LOCAL_EX_LABEL = f"Local ex Tax ({CB_LOCAL_CCY})"
CB_LOCAL_INC_LABEL = f"Local inc Tax ({CB_LOCAL_CCY})"

# One matrix row per container, currency and BPO: every row is billable as it
# stands, with no mixing of BPOs inside a row.
CB_MATRIX_KEYS = ["Job", "Container", "Container Size", "Currency", "BPO"]


def bpo_country(port_code):
    """'NZAKL' -> 'NZ'. Anything unusable comes back empty."""
    text = str(port_code).strip().upper()
    return text[:2] if len(text) >= 2 and text[:2].isalpha() else ""


def bpo_classify_group(origin, destination):
    """Trade group for one shipment, from its origin and destination ports."""
    origin_country, dest_country = bpo_country(origin), bpo_country(destination)
    override = BPO_GROUP_RULES.get((origin_country, dest_country))
    if override:
        return override
    at_home = (origin_country == BPO_HOME_COUNTRY, dest_country == BPO_HOME_COUNTRY)
    if at_home == (True, True):
        return BPO_GROUP_DOMESTIC
    if at_home == (True, False):
        return BPO_GROUP_EXPORT
    if at_home == (False, True):
        return BPO_GROUP_IMPORT
    return BPO_GROUP_OTHER


def bpo_leg_for_code(charge_code):
    """Leg of the move a charge code belongs to."""
    return BPO_LEG_BY_CODE.get(str(charge_code).strip().upper(), BPO_LEG_UNMAPPED)


def bpo_for(group, leg):
    """BPO record for a (group, leg) pair, or the TBA placeholder."""
    return BPO_TABLE.get((group, leg), BPO_UNALLOCATED)


def bpo_label(group, leg):
    """Column header for a leg subtotal: 'Door to Wharf (BPO 772... / NZD)'."""
    record = bpo_for(group, leg)
    return f"{leg} (BPO {record['BPO']} / {record['BPO Currency']})"


def bpo_group_rank(group):
    return (BPO_GROUP_ORDER.index(group) if group in BPO_GROUP_ORDER
            else len(BPO_GROUP_ORDER))


def bpo_annotate(detail):
    """Add Group, Leg and BPO columns to a charge detail frame."""
    if detail is None or detail.empty:
        return detail

    out = detail.copy()
    origin = out["Origin"] if "Origin" in out.columns else pd.Series("", index=out.index)
    dest = (out["Destination"] if "Destination" in out.columns
            else pd.Series("", index=out.index))

    out["Group"] = [bpo_classify_group(o, d) for o, d in zip(origin, dest)]
    out["Leg"] = [bpo_leg_for_code(c) for c in out["Charge Code"]]

    records = [bpo_for(g, l) for g, l in zip(out["Group"], out["Leg"])]
    out["BPO"] = [r["BPO"] for r in records]
    out["BPO Vendor"] = [r["BPO Vendor"] for r in records]
    out["BPO Entity"] = [r["BPO Entity"] for r in records]
    out["BPO Currency"] = [r["BPO Currency"] for r in records]
    out["Invoice To"] = [BPO_INVOICE_TO.get(g, BPO_TBA) for g in out["Group"]]

    if "Currency" in out.columns:
        out["Currency"] = (out["Currency"].fillna(CB_LOCAL_CCY).astype(str)
                           .str.strip().str.upper())
    else:
        out["Currency"] = CB_LOCAL_CCY

    out["_group_rank"] = out["Group"].map(
        {g: i for i, g in enumerate(BPO_GROUP_ORDER)}).fillna(len(BPO_GROUP_ORDER))
    out["_leg_rank"] = out["Leg"].map(
        {l: i for i, l in enumerate(BPO_LEG_ORDER)}).fillna(len(BPO_LEG_ORDER))
    out = out.sort_values(["_group_rank", "Job", "Container", "Currency",
                           "_leg_rank", "Charge Code"])
    return out.drop(columns=["_group_rank", "_leg_rank"]).reset_index(drop=True)


def bpo_shipment_skeleton(shipment_df):
    """A charge detail built from the shipment listing alone, with no amounts.

    One row per container per billable leg, so every container arrives in the
    workbook against its group and its BPO. Amounts are zero and the charge
    code is blank, so the matrix comes out as the structure waiting for the
    figures rather than as a set of false charges.
    """
    containers = cb_container_table(shipment_df)
    if containers.empty:
        return pd.DataFrame()

    rows = []
    for record in containers.to_dict("records"):
        group = bpo_classify_group(record.get("Origin"), record.get("Destination"))
        for leg in BPO_BILLABLE_LEGS:
            bpo = bpo_for(group, leg)
            row = dict(record)
            row.update({
                "Charge Code": "",
                "Description": "",
                "Currency": bpo["BPO Currency"],
                "Containers on Job": None,
                "Group": group,
                "Leg": leg,
                "BPO": bpo["BPO"],
                "BPO Vendor": bpo["BPO Vendor"],
                "BPO Entity": bpo["BPO Entity"],
                "BPO Currency": bpo["BPO Currency"],
                "Invoice To": BPO_INVOICE_TO.get(group, BPO_TBA),
                CB_AMOUNT_CCY: 0.0,
                "Allocated Tax": 0.0,
                CB_INC_CCY: 0.0,
                CB_ALLOC_EX: 0.0,
                CB_ALLOC_INC: 0.0,
            })
            rows.append(row)

    detail = pd.DataFrame(rows).drop(columns=["Position"], errors="ignore")
    detail["_g"] = detail["Group"].map(bpo_group_rank)
    detail["_l"] = detail["Leg"].map({l: i for i, l in enumerate(BPO_LEG_ORDER)})
    return (detail.sort_values(["_g", "Job", "Container", "_l"])
            .drop(columns=["_g", "_l"]).reset_index(drop=True))


def bpo_has_charges(detail):
    """True when the detail carries real charge lines rather than a skeleton."""
    if detail is None or detail.empty:
        return False
    return bool(detail["Charge Code"].astype(str).str.strip().ne("").any())


def bpo_sort_by_group(df, then=()):
    """Sort any annotated frame into the canonical group order."""
    if df is None or df.empty or "Group" not in df.columns:
        return df
    out = df.copy()
    out["_rank"] = out["Group"].map(bpo_group_rank)
    keys = ["_rank"] + [c for c in then if c in out.columns]
    return out.sort_values(keys).drop(columns="_rank").reset_index(drop=True)


def bpo_currencies_in(frame):
    """Currencies present, local currency first then the rest alphabetically."""
    seen = sorted(set(frame["Currency"].dropna().astype(str)))
    return (([CB_LOCAL_CCY] if CB_LOCAL_CCY in seen else [])
            + [c for c in seen if c != CB_LOCAL_CCY])


def bpo_group_summary(detail):
    """One row per trade group, in local currency. On screen only."""
    summary = detail.groupby("Group", observed=True).agg(
        Shipments=("Job", "nunique"),
        Containers=("Container", "nunique"),
        **{CB_LOCAL_EX_LABEL: (CB_ALLOC_EX, "sum"),
           CB_LOCAL_INC_LABEL: (CB_ALLOC_INC, "sum")},
    ).reset_index()
    summary["Invoice To"] = summary["Group"].map(BPO_INVOICE_TO).fillna(BPO_TBA)
    summary["Avg_per_Container"] = (
        summary[CB_LOCAL_EX_LABEL] / summary["Containers"].replace(0, pd.NA)).round(2)
    for col in (CB_LOCAL_EX_LABEL, CB_LOCAL_INC_LABEL):
        summary[col] = summary[col].round(2)
    summary["_rank"] = summary["Group"].map(bpo_group_rank)
    ordered = ["Group", "Invoice To", "Shipments", "Containers",
               CB_LOCAL_EX_LABEL, CB_LOCAL_INC_LABEL, "Avg_per_Container"]
    return (summary.sort_values("_rank").drop(columns="_rank")
            .reset_index(drop=True)[ordered])


def bpo_summary(detail):
    """Sheet 1, upper table: one row per BPO and currency."""
    summary = detail.groupby(["Group", "Leg", "BPO", "Currency"], observed=True).agg(
        Shipments=("Job", "nunique"),
        Containers=("Container", "nunique"),
        Total_ex_Tax=(CB_AMOUNT_CCY, "sum"),
        **{CB_LOCAL_EX_LABEL: (CB_ALLOC_EX, "sum")},
    ).reset_index()

    records = [bpo_for(g, l) for g, l in zip(summary["Group"], summary["Leg"])]
    summary["BPO Vendor"] = [r["BPO Vendor"] for r in records]
    summary["BPO Entity"] = [r["BPO Entity"] for r in records]
    summary["BPO Currency"] = [r["BPO Currency"] for r in records]
    summary["Invoice To"] = summary["Group"].map(BPO_INVOICE_TO).fillna(BPO_TBA)
    for col in ("Total_ex_Tax", CB_LOCAL_EX_LABEL):
        summary[col] = summary[col].round(2)

    summary["_g"] = summary["Group"].map(bpo_group_rank)
    summary["_l"] = summary["Leg"].map(
        {l: i for i, l in enumerate(BPO_LEG_ORDER)}).fillna(len(BPO_LEG_ORDER))
    ordered = ["Group", "Invoice To", "Leg", "BPO", "BPO Vendor", "BPO Entity",
               "BPO Currency", "Currency", "Shipments", "Containers",
               "Total_ex_Tax", CB_LOCAL_EX_LABEL]
    return (summary.sort_values(["_g", "_l", "Currency"]).drop(columns=["_g", "_l"])
            .reset_index(drop=True)[ordered])


def bpo_code_summary(detail):
    """Sheet 1, lower table: the local ex-tax total split by charge code."""
    work = detail[detail["Charge Code"].astype(str).str.strip() != ""]
    if work.empty:
        return pd.DataFrame()
    summary = work.groupby(["Group", "Leg", "BPO", "Charge Code"],
                           observed=True).agg(
        Description=("Description", "first"),
        Containers=("Container", "nunique"),
        Charge_Lines=("Charge Code", "count"),
        **{CB_LOCAL_EX_LABEL: (CB_ALLOC_EX, "sum")},
    ).reset_index()
    summary[CB_LOCAL_EX_LABEL] = summary[CB_LOCAL_EX_LABEL].round(2)
    summary["Share_of_BPO_%"] = (
        100 * summary[CB_LOCAL_EX_LABEL]
        / summary.groupby(["Group", "BPO"])[CB_LOCAL_EX_LABEL].transform("sum")
    ).round(2)

    summary["_g"] = summary["Group"].map(bpo_group_rank)
    summary["_l"] = summary["Leg"].map(
        {l: i for i, l in enumerate(BPO_LEG_ORDER)}).fillna(len(BPO_LEG_ORDER))
    ordered = ["Group", "Leg", "BPO", "Charge Code", "Description", "Containers",
               "Charge_Lines", CB_LOCAL_EX_LABEL, "Share_of_BPO_%"]
    return (summary.sort_values(["_g", "_l", "BPO", CB_LOCAL_EX_LABEL],
                                ascending=[True, True, True, False])
            .drop(columns=["_g", "_l"]).reset_index(drop=True)[ordered])


def bpo_shipment_groups(detail):
    """Sheet 3: one row per shipment, currency and BPO."""
    rows = []
    for (group, job, currency, bpo), part in detail.groupby(
            ["Group", "Job", "Currency", "BPO"], observed=True):
        legs = sorted(set(part["Leg"]), key=lambda l: BPO_LEG_ORDER.index(l))
        row = {
            "Group": group,
            "Job": job,
            "BPO": bpo,
            "Leg": ", ".join(legs),
            "Currency": currency,
            "Containers": part["Container"].nunique(),
            "Total_ex_Tax": round(part[CB_AMOUNT_CCY].sum(), 2),
            CB_LOCAL_EX_LABEL: round(part[CB_ALLOC_EX].sum(), 2),
        }
        for col in ("Consignor", "Origin", "Destination", "ETD"):
            if col in part.columns:
                row[col] = part[col].iloc[0]
        rows.append(row)

    frame = pd.DataFrame(rows)
    frame["_g"] = frame["Group"].map(bpo_group_rank)
    return (frame.sort_values(["_g", "Job", "Currency", "BPO"])
            .drop(columns="_g").reset_index(drop=True))


def bpo_charge_detail(detail):
    """Sheet 4: the charge lines, trimmed to the columns that get read."""
    cols = ["Group", "Job", "Container", "Container Size", "Charge Code",
            "Description", "Leg", "BPO", "Currency", CB_AMOUNT_CCY,
            CB_ALLOC_EX, "Consignor", "Origin", "Destination", "ETD"]
    out = detail[[c for c in cols if c in detail.columns]].copy()
    return out.rename(columns={CB_AMOUNT_CCY: "Amount (charge ccy)",
                               CB_ALLOC_EX: CB_LOCAL_EX_LABEL})


def bpo_unmapped_codes(detail):
    """Charge codes seen in the data that BPO_LEG_BY_CODE does not cover."""
    unmapped = detail[detail["Leg"] == BPO_LEG_UNMAPPED]
    if unmapped.empty:
        return []
    return sorted(c for c in unmapped["Charge Code"].unique() if str(c).strip())


def bpo_group_matrix(detail, group):
    """Sheet 2, one block: a container matrix for a single group.

    One row per container, currency and BPO, so a row is never a mix of BPOs.
    Domestic shipments carry a single BPO across all three legs, so they come
    out as one row per container and currency; exports and imports split into
    the Door to Wharf, Ocean freight and Wharf to Door BPOs.
    """
    part = detail[detail["Group"] == group].copy()
    if part.empty:
        return pd.DataFrame(), {}

    # A blank container size would drop the whole row out of the pivot, which
    # is how shipments billed without a container number used to disappear.
    part["Container Size"] = (part["Container Size"].fillna("").astype(str)
                              if "Container Size" in part.columns else "")

    matrix = part.pivot_table(
        index=CB_MATRIX_KEYS, columns="Charge Code", values=CB_AMOUNT_CCY,
        aggfunc="sum", observed=True).reset_index()
    matrix.columns.name = None

    # The skeleton has a blank charge code; it is a placeholder, not a column.
    codes = [c for c in matrix.columns
             if c not in CB_MATRIX_KEYS and str(c).strip() != ""]
    matrix = matrix.drop(columns=[c for c in matrix.columns
                                  if c not in CB_MATRIX_KEYS and c not in codes])

    legs = {}
    for leg in BPO_LEG_ORDER:
        leg_codes = sorted(c for c in codes if bpo_leg_for_code(c) == leg)
        if leg_codes:
            legs[leg] = leg_codes

    ordered_codes = [c for leg in legs for c in legs[leg]]
    matrix = matrix[CB_MATRIX_KEYS + ordered_codes]

    subtotal_cols = {}
    for leg, leg_codes in legs.items():
        label = bpo_label(group, leg)
        totals = matrix[leg_codes].sum(axis=1).round(2)
        # A leg with nothing in this row stays blank rather than showing a
        # 0.00 that reads like a real charge.
        matrix[label] = totals.mask(totals == 0)
        subtotal_cols[leg] = label

    matrix["TOTAL"] = (matrix[ordered_codes].sum(axis=1).round(2)
                       if ordered_codes else 0.0)

    local = (part.groupby(CB_MATRIX_KEYS, observed=True)[CB_ALLOC_EX].sum()
             .round(2).reset_index().rename(columns={CB_ALLOC_EX: CB_LOCAL_EX_LABEL}))
    matrix = pd.merge(matrix, local, how="left", on=CB_MATRIX_KEYS)

    matrix = (matrix.sort_values(["Job", "Container", "Currency", "BPO"])
              .reset_index(drop=True))
    return matrix, subtotal_cols


def bpo_matrix_flat(detail):
    """All groups stacked into one frame, with Group as the first column.

    Leg subtotals get generic headers here, with the BPO already in its own
    key column, so the columns line up across groups. The Excel sheet uses
    per-group blocks instead. On screen only.
    """
    frames = []
    for group in BPO_GROUP_ORDER:
        matrix, subtotal_cols = bpo_group_matrix(detail, group)
        if matrix.empty:
            continue
        for leg, label in subtotal_cols.items():
            matrix = matrix.rename(columns={label: f"{leg} ex Tax"})
        matrix.insert(0, "Group", group)
        frames.append(matrix)
    if not frames:
        return pd.DataFrame()

    flat = pd.concat(frames, ignore_index=True)
    keys = ["Group"] + CB_MATRIX_KEYS
    legs = [f"{leg} ex Tax" for leg in BPO_LEG_ORDER
            if f"{leg} ex Tax" in flat.columns]
    tail = ["TOTAL", CB_LOCAL_EX_LABEL]
    codes = [c for c in flat.columns if c not in keys + legs + tail]
    codes = sorted(codes, key=lambda c: (BPO_LEG_ORDER.index(bpo_leg_for_code(c)), c))
    return flat[keys + codes + legs + tail]


# ── Excel writing (project 2) ────────────────────────────────────────────────

CB_BAND_FILL = "F1F5F9"          # very light tint used to separate groups of rows
BPO_HEADER_FILL = PRIMARY.lstrip("#")   # column headers
BPO_GROUP_FILL = ACCENT.lstrip("#")     # group title row
BPO_LEGEND_FILL = "E4EDF2"              # BPO legend row under the title
BPO_SUBTOTAL_FILL = "DCE6EE"            # currency and group total rows

_BPO_THIN = Side(style="thin", color="B7C4D1")


def _bpo_write_row(ws, row, values, fill=None, bold=False, colour=None,
                   number_format=None, border_top=False):
    for offset, value in enumerate(values, start=1):
        cell = ws.cell(row=row, column=offset, value=value)
        if fill:
            cell.fill = PatternFill("solid", start_color=fill)
        if bold or colour:
            cell.font = Font(bold=bold, color=colour or "000000", size=11)
        if number_format and isinstance(value, (int, float)):
            cell.number_format = number_format
        if border_top:
            cell.border = Border(top=_BPO_THIN)
    return row + 1


def _bpo_fit_columns(ws, skip_rows=(), limit=(11, 34)):
    """Size every column to its content, ignoring full-width layout rows."""
    for column in ws.columns:
        letter = column[0].column_letter
        longest = max((len(str(cell.value)) for cell in column
                       if cell.value is not None and cell.row not in skip_rows),
                      default=10)
        ws.column_dimensions[letter].width = min(max(longest + 2, limit[0]), limit[1])


def cb_write_sheet(writer, name, df, number_format="#,##0.00", band_by=None):
    """Write a sheet with a styled header row and sensible column widths.

    band_by: optional list of columns whose combined value defines a group. The
    background alternates between plain and a very light tint each time that
    value changes, so consecutive rows belonging to one job read as a block.
    """
    if df is None or df.empty:
        return
    sheet = name[:31]
    df.to_excel(writer, sheet_name=sheet, index=False)
    ws = writer.sheets[sheet]
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", start_color=BPO_HEADER_FILL)
    ws.freeze_panes = "A2"
    for column in ws.columns:
        letter = column[0].column_letter
        longest = max((len(str(c.value)) for c in column[:200] if c.value is not None),
                      default=10)
        ws.column_dimensions[letter].width = min(max(longest + 2, 10), 46)
        for cell in column[1:]:
            if isinstance(cell.value, (int, float)):
                cell.number_format = number_format

    if band_by:
        keys = [c for c in band_by if c in df.columns]
        if not keys:
            return
        band = PatternFill("solid", start_color=CB_BAND_FILL)
        shaded = False
        previous = None
        for offset, group in enumerate(df[keys].astype(str).agg("|".join, axis=1)):
            if previous is not None and group != previous:
                shaded = not shaded
            previous = group
            if shaded:
                for cell in ws[offset + 2]:
                    cell.fill = band
    return


def bpo_write_summary_sheet(writer, detail, sheet_name="BPO Summary",
                            second_table_row=12, number_format="#,##0.00"):
    """Sheet 1: the BPO table, then the same spend split by charge code.

    The second table starts at `second_table_row` (or below the first table if
    that has grown past it), so its position is predictable for anything
    pointing at the sheet.
    """
    ws = writer.book.create_sheet(sheet_name[:31])
    layout_rows = set()

    def write_table(row, title, frame):
        if frame is None or frame.empty:
            return row
        layout_rows.add(row)
        row = _bpo_write_row(ws, row, [title], fill=BPO_GROUP_FILL, bold=True,
                             colour="FFFFFF")
        row = _bpo_write_row(ws, row, list(frame.columns), fill=BPO_HEADER_FILL,
                             bold=True, colour="FFFFFF")
        shaded = False
        previous = None
        for record in frame.to_dict("records"):
            key = str(record.get("BPO", "")) + str(record.get("Group", ""))
            if previous is not None and key != previous:
                shaded = not shaded
            previous = key
            row = _bpo_write_row(ws, row, [record[c] for c in frame.columns],
                                 fill=CB_BAND_FILL if shaded else None,
                                 number_format=number_format)
        return row

    row = write_table(1, "BPO ALLOCATION", bpo_summary(detail))
    codes = bpo_code_summary(detail)
    if not codes.empty:
        row = write_table(max(row + 2, second_table_row),
                          f"{CB_LOCAL_EX_LABEL.upper()} BY CHARGE CODE", codes)

    _bpo_fit_columns(ws, skip_rows=layout_rows)
    return ws


def bpo_write_grouped_matrix(writer, detail, sheet_name="Container Matrix",
                             number_format="#,##0.00"):
    """Sheet 2: the container matrix as one block per trade group.

    Each block opens with a title row naming the group and who it is invoiced
    to, a legend row per leg spelling out the BPO, then the container rows —
    one per container, currency and BPO — a total per currency, and the group
    total in local currency.
    """
    ws = writer.book.create_sheet(sheet_name[:31])
    row = 1
    layout_rows = set()  # title and legend rows, excluded from column sizing

    for group in BPO_GROUP_ORDER:
        matrix, subtotal_cols = bpo_group_matrix(detail, group)
        if matrix.empty:
            continue

        part = detail[detail["Group"] == group]
        title = (f"{group.upper()}  ·  Invoice to {BPO_INVOICE_TO.get(group, BPO_TBA)}"
                 f"  ·  {part['Job'].nunique()} shipments"
                 f"  ·  {part['Container'].nunique()} containers")
        layout_rows.add(row)
        row = _bpo_write_row(ws, row, [title], fill=BPO_GROUP_FILL, bold=True,
                             colour="FFFFFF")

        for leg in BPO_LEG_ORDER:
            record = bpo_for(group, leg)
            if leg not in subtotal_cols and leg not in set(part["Leg"]):
                continue
            layout_rows.add(row)
            row = _bpo_write_row(ws, row, [
                f"{leg}:  BPO {record['BPO']}   ·   {record['BPO Currency']}   ·   "
                f"{record['BPO Entity']}   ·   Vendor {record['BPO Vendor']}"
            ], fill=BPO_LEGEND_FILL)

        header_row = row
        row = _bpo_write_row(ws, row, list(matrix.columns), fill=BPO_HEADER_FILL,
                             bold=True, colour="FFFFFF")
        for cell in ws[header_row]:
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=True)

        shaded = False
        previous_job = None
        for record in matrix.to_dict("records"):
            if previous_job is not None and record["Job"] != previous_job:
                shaded = not shaded
            previous_job = record["Job"]
            row = _bpo_write_row(ws, row, [record[c] for c in matrix.columns],
                                 fill=CB_BAND_FILL if shaded else None,
                                 number_format=number_format)

        value_cols = list(matrix.columns[len(CB_MATRIX_KEYS):])
        first = True
        for currency in bpo_currencies_in(matrix):
            rows_in_ccy = matrix[matrix["Currency"] == currency]
            totals = [f"TOTAL {currency}", "", "", currency, ""]
            for column in value_cols:
                total = round(float(rows_in_ccy[column].sum()), 2)
                totals.append(total if total else "")
            row = _bpo_write_row(ws, row, totals, fill=BPO_SUBTOTAL_FILL, bold=True,
                                 number_format=number_format, border_top=first)
            first = False

        # Only the local column can be added across currencies, so the group
        # total carries that one alone.
        local_total = ([f"GROUP TOTAL ({CB_LOCAL_CCY})"]
                       + [""] * (len(matrix.columns) - 2)
                       + [round(float(matrix[CB_LOCAL_EX_LABEL].sum()), 2)])
        row = _bpo_write_row(ws, row, local_total, fill=BPO_SUBTOTAL_FILL, bold=True,
                             number_format=number_format)

        row += 2  # blank line between groups

    _bpo_fit_columns(ws, skip_rows=layout_rows, limit=(11, 30))
    ws.freeze_panes = "F1"
    return ws


def cb_create_report(detail):
    """The four-sheet workbook: BPO summary, matrix, shipment groups, detail."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        bpo_write_summary_sheet(writer, detail)               # 1
        bpo_write_grouped_matrix(writer, detail)              # 2
        cb_write_sheet(writer, "Shipment Groups",             # 3
                       bpo_shipment_groups(detail), band_by=["Job"])
        cb_write_sheet(writer, "Charge Detail",               # 4
                       bpo_charge_detail(detail), band_by=["Job"])
    buf.seek(0)
    return buf


# ── Project 2 on-screen summaries ────────────────────────────────────────────

def cb_container_summary(detail):
    keys = [c for c in ("Group", "Job", "Container", "Container Size")
            if c in detail.columns]
    summary = detail.groupby(keys, observed=True).agg(
        Charge_Lines=("Charge Code", "count"),
        Total_ex_Tax=(CB_ALLOC_EX, "sum"),
        Total_inc_Tax=(CB_ALLOC_INC, "sum"),
    ).reset_index()
    context = detail.drop_duplicates(["Job", "Container"])[
        ["Job", "Container"] + [c for c in CB_CONTEXT_COLS if c in detail.columns]]
    summary = pd.merge(summary, context, how="left", on=["Job", "Container"])
    for col in ("Total_ex_Tax", "Total_inc_Tax"):
        summary[col] = summary[col].round(2)
    return summary.sort_values("Total_ex_Tax", ascending=False).reset_index(drop=True)


def cb_shipment_summary(detail, shipment_df):
    keys = [c for c in ("Group", "Job") if c in detail.columns]
    summary = detail.groupby(keys, observed=True).agg(
        Containers=("Container", "nunique"),
        Charge_Lines=("Charge Code", "count"),
        Total_ex_Tax=(CB_ALLOC_EX, "sum"),
        Total_inc_Tax=(CB_ALLOC_INC, "sum"),
    ).reset_index()
    summary["Cost_per_Container"] = (
        summary["Total_ex_Tax"] / summary["Containers"].replace(0, pd.NA)).round(2)
    context_cols = ["Job"] + [c for c in
                              ["Consignor", "Origin", "Destination", "Discharge Port",
                               "Vessel", "Shipping Line", "Mode", "Incoterms",
                               "Container Count", "TEU", "ETD", "ETA", "ATA",
                               "Entry Ref", "Order Reference"]
                              if c in shipment_df.columns]
    summary = pd.merge(summary, shipment_df[context_cols], how="left", on="Job")
    for col in ("Total_ex_Tax", "Total_inc_Tax"):
        summary[col] = summary[col].round(2)
    return summary.sort_values("Total_ex_Tax", ascending=False).reset_index(drop=True)


def cb_charge_code_summary(detail):
    work = detail[detail["Charge Code"].astype(str).str.strip() != ""]
    if work.empty:
        return pd.DataFrame()
    summary = work.groupby("Charge Code", observed=True).agg(
        Shipments=("Job", "nunique"),
        Containers=("Container", "nunique"),
        Charge_Lines=("Charge Code", "count"),
        Total_ex_Tax=(CB_ALLOC_EX, "sum"),
    ).reset_index()
    summary["Avg_per_Container"] = (
        summary["Total_ex_Tax"] / summary["Containers"].replace(0, pd.NA)).round(2)
    summary["Share_of_Spend_%"] = (
        100 * summary["Total_ex_Tax"] / summary["Total_ex_Tax"].sum()).round(2)
    summary["Total_ex_Tax"] = summary["Total_ex_Tax"].round(2)
    return summary.sort_values("Total_ex_Tax", ascending=False).reset_index(drop=True)


# ── Excel report (project 1) ──────────────────────────────────────────────────

def build_analysis_summary(billing_det, shipment_sum):
    """Build the flat analysis sections written to the Excel report."""
    sections = {}

    containers = ""
    if "Container Count" in shipment_sum.columns:
        containers = int(shipment_sum["Container Count"].fillna(0).sum())

    sections["kpis"] = pd.DataFrame([
        {"Metric": "Total Shipments", "Value": billing_det["Shipment Job"].nunique()},
        {"Metric": "Total Containers", "Value": containers},
        {"Metric": f"{BASE_CCY} Charges",
         "Value": billing_det[billing_det["Currency"] == BASE_CCY]["Total"].sum()},
        {"Metric": f"{FX_CCY} Charges",
         "Value": billing_det[billing_det["Currency"] == FX_CCY]["Total"].sum()},
        {"Metric": f"Total Billed ({BASE_CCY})", "Value": billing_det["Local Total"].sum()},
    ])

    sup = (billing_det.groupby("Supplier Name")["Local Total"].sum()
           .sort_values(ascending=False).reset_index())
    sup.columns = ["Supplier Name", LOCAL_LABEL]
    sections["supplier"] = sup

    bd_m, _ = add_month_col(billing_det, shipment_sum)
    if bd_m is not None and "Month" in bd_m.columns:
        monthly = (bd_m.groupby(["Month", "Currency"])["Local Total"].sum()
                   .unstack(fill_value=0).reset_index())
        monthly.columns.name = None
        sections["monthly"] = monthly

    top10 = (billing_det.groupby("Description")["Local Total"].sum()
             .sort_values(ascending=False).head(10).reset_index())
    top10.columns = ["Charge Description", LOCAL_LABEL]
    sections["top10"] = top10

    if "Incoterms" in billing_det.columns:
        inco = (billing_det.groupby("Incoterms")["Local Total"].sum()
                .sort_values(ascending=False).reset_index())
        inco.columns = ["Incoterms", LOCAL_LABEL]
        sections["incoterm"] = inco

    return sections


def create_report(shipment_sum, billing_sum, billing_det, supplier_sum):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        billing_det.to_excel(writer, sheet_name="Billing Detail", index=False)
        shipment_sum.to_excel(writer, sheet_name="Shipment Summary", index=False)
        billing_sum.to_excel(writer, sheet_name="Billing Summary", index=False)
        supplier_sum.to_excel(writer, sheet_name="Supplier Summary", index=False)

        analysis = build_analysis_summary(billing_det, shipment_sum)
        ws = writer.book.create_sheet("Analysis")
        row = 1

        section_labels = {
            "kpis": "KEY METRICS",
            "supplier": f"SPEND BY SUPPLIER (Local {BASE_CCY})",
            "monthly": f"MONTHLY CHARGES — {BASE_CCY} vs {FX_CCY} (Local {BASE_CCY})",
            "top10": f"TOP 10 CHARGE TYPES (Local {BASE_CCY})",
            "incoterm": f"SPEND BY INCOTERM (Local {BASE_CCY})",
        }

        for key, label in section_labels.items():
            if key not in analysis:
                continue
            df = analysis[key]
            header_cell = ws.cell(row=row, column=1, value=label)
            header_cell.font = Font(bold=True, color="FFFFFF", size=11)
            header_cell.fill = PatternFill("solid", start_color=REPORT_HEADER_FILL)
            row += 1
            for col_idx, col_name in enumerate(df.columns, start=1):
                ws.cell(row=row, column=col_idx, value=col_name).font = Font(bold=True)
            row += 1
            for _, data_row in df.iterrows():
                for col_idx, val in enumerate(data_row, start=1):
                    ws.cell(row=row, column=col_idx, value=val)
                row += 1
            row += 1  # blank line between sections

        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 60)

    buf.seek(0)
    return buf


# ── Header ────────────────────────────────────────────────────────────────────

st.markdown(f"""
<div class="app-header">
  <div>
    <h1>{APP_TITLE}</h1>
    <div class="app-subtitle">{APP_SUBTITLE}</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ── Project 1: consolidated billing report ───────────────────────────────────
# The original dashboard, wrapped in a function so it can live inside a tab
# (st.stop() would halt the whole script, including tab 2, so the guards
# return instead). Uploads and the download button sit at the top of the tab.

def render_billing_report():
    # ── Uploads ───────────────────────────────────────────────────────────────

    st.markdown('<div class="section-title">Upload files</div>', unsafe_allow_html=True)
    up_left, up_right = st.columns(2)
    with up_left:
        billing_uploads = st.file_uploader(
            "Billing export(s)", type=["xlsx", "csv"],
            accept_multiple_files=True, key="br_billing")
    with up_right:
        shipment_upload = st.file_uploader(
            "Shipment listing report", type=["xlsx", "csv"], key="br_shipment")

    if not billing_uploads or not shipment_upload:
        st.info("Upload the billing export(s) and the shipment listing report to begin.")
        return

    with st.spinner("Processing files…"):
        result = load_data(billing_uploads, shipment_upload)
    if result is None:
        st.error("The files could not be loaded. Check that both exports match the expected columns.")
        return

    data, shipment_df = result
    ship_sum = build_shipment_summary(shipment_df)
    bill_det = build_billing_detail(data)
    bill_sum = build_billing_summary(data, shipment_df)
    supp_sum = build_supplier_summary(data, shipment_df)
    billed_jobs = set(data["Shipment Job"].dropna())

    bd_all, all_months = add_month_col(bill_det, ship_sum)

    # The download button sits directly under the upload fields. It is rendered
    # into this slot at the end, once the filters below have been applied.
    download_slot = st.container()

    # ── Filters ───────────────────────────────────────────────────────────────

    with st.expander("Filters", expanded=False):
        billed_ship = ship_sum[ship_sum["Shipment Job"].isin(billed_jobs)]

        def options(col):
            if col not in billed_ship.columns:
                return []
            return sorted(billed_ship[col].dropna().unique())

        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            f_month = st.multiselect("Month (ETD)", all_months, placeholder="All months")
            f_supplier = st.multiselect("Supplier", options("Supplier Name"), placeholder="All suppliers")
        with fc2:
            f_origin = st.multiselect("Loading port", options("Loading Port"), placeholder="All loading ports")
            f_dest = st.multiselect("Destination port", options("Destination Port"), placeholder="All destinations")
        with fc3:
            f_inco = st.multiselect("Incoterms", options("Incoterms"), placeholder="All Incoterms")
            f_mode = st.multiselect("Mode", options("Mode"), placeholder="All modes")

    def filter_df(df):
        m = pd.Series(True, index=df.index)
        for values, col in (
            (f_supplier, "Supplier Name"),
            (f_origin, "Loading Port"),
            (f_dest, "Destination Port"),
            (f_inco, "Incoterms"),
            (f_mode, "Mode"),
        ):
            if values and col in df.columns:
                m &= df[col].isin(values)
        return df[m]

    bill_sum_f = filter_df(bill_sum)

    if f_month:
        jobs_in_month = bd_all[bd_all["Month"].isin(f_month)]["Shipment Job"].unique()
        bill_sum_f = bill_sum_f[bill_sum_f["Shipment Job"].isin(jobs_in_month)]

    kept = bill_sum_f["Shipment Job"]
    ship_sum_f = ship_sum[ship_sum["Shipment Job"].isin(kept)]
    bill_det_f = bill_det[bill_det["Shipment Job"].isin(kept)]
    fdata = data[data["Shipment Job"].isin(kept)]
    fship = shipment_df[shipment_df["Shipment Job"].isin(kept)]
    supp_f = build_supplier_summary(fdata, fship) if not fdata.empty else supp_sum

    # ── Download (rendered into the slot under the uploads) ───────────────────

    with download_slot:
        st.download_button(
            label="Download report",
            data=create_report(ship_sum_f, bill_sum_f, bill_det_f, supp_f),
            file_name=REPORT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="br_download_top",
        )
        st.caption(f"{len(bill_sum_f)} shipments in the current export")

    # ── KPIs ──────────────────────────────────────────────────────────────────

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Shipments", f"{bill_sum_f['Shipment Job'].nunique():,}")
    k2.metric("Containers", f"{int(ship_sum_f['Container Count'].fillna(0).sum()):,}"
              if "Container Count" in ship_sum_f.columns else "—")
    k3.metric(f"{BASE_CCY} charges",
              f"{bill_det_f[bill_det_f['Currency'] == BASE_CCY]['Total'].sum():,.2f}")
    k4.metric(f"{FX_CCY} charges",
              f"{bill_det_f[bill_det_f['Currency'] == FX_CCY]['Total'].sum():,.2f}")
    k5.metric(f"Total ({BASE_CCY})", f"{bill_det_f['Local Total'].sum():,.2f}")

    # ── Tabs ──────────────────────────────────────────────────────────────────

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Shipment summary", "Billing summary", "Billing detail", "Analysis", "Download",
    ])

    with tab1:
        st.markdown('<div class="section-title">Shipment summary</div>', unsafe_allow_html=True)
        st.dataframe(ship_sum_f, use_container_width=True, hide_index=True)
        containers = (int(ship_sum_f["Container Count"].fillna(0).sum())
                      if "Container Count" in ship_sum_f.columns else 0)
        st.caption(f"{len(ship_sum_f)} shipments · {containers} containers")

    with tab2:
        st.markdown('<div class="section-title">Billing summary — per job</div>', unsafe_allow_html=True)
        st.dataframe(bill_sum_f, use_container_width=True, hide_index=True)
        base_total = bill_sum_f[BASE_CCY].sum() if BASE_CCY in bill_sum_f.columns else 0
        fx_total = bill_sum_f[FX_CCY].sum() if FX_CCY in bill_sum_f.columns else 0
        local_total = bill_sum_f[LOCAL_LABEL].sum() if LOCAL_LABEL in bill_sum_f.columns else 0
        st.caption(f"{BASE_CCY} {base_total:,.2f}  ·  {FX_CCY} {fx_total:,.2f}  "
                   f"·  {LOCAL_LABEL} {local_total:,.2f}")

    with tab3:
        st.markdown('<div class="section-title">Billing detail — all charge lines</div>', unsafe_allow_html=True)
        st.dataframe(bill_det_f, use_container_width=True, hide_index=True)
        st.caption(f"{len(bill_det_f):,} charge lines")

    with tab4:
        if bill_sum_f.empty:
            st.warning("No shipments match the current filters. Clear a filter to see data.")
        else:
            st.markdown(f'<div class="section-title">Total billed by supplier ({BASE_CCY})</div>',
                        unsafe_allow_html=True)
            fig1 = px.bar(
                supp_f.sort_values(SUPPLIER_TOTAL_COL),
                x=SUPPLIER_TOTAL_COL, y="Supplier Name", orientation="h",
                color_discrete_sequence=[PRIMARY],
                labels={SUPPLIER_TOTAL_COL: LOCAL_LABEL, "Supplier Name": ""},
                text=SUPPLIER_TOTAL_COL,
            )
            fig1.update_traces(texttemplate="%{text:,.0f}", textposition="outside",
                               marker_line_width=0, textfont_color=TEXT, marker_color=PRIMARY)
            fig1.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title="")
            st.plotly_chart(fig1, use_container_width=True)

            r2l, r2r = st.columns(2)

            with r2l:
                st.markdown(f'<div class="section-title">Top 10 charge types ({BASE_CCY})</div>',
                            unsafe_allow_html=True)
                top10 = (bill_det_f.groupby("Description")["Local Total"].sum()
                         .sort_values(ascending=False).head(10).index.tolist())
                charge_stack = (bill_det_f[bill_det_f["Description"].isin(top10)]
                                .groupby(["Description", "Currency"])["Local Total"].sum().reset_index())
                charge_stack["rank"] = charge_stack["Description"].map(
                    {d: i for i, d in enumerate(reversed(top10))})
                charge_stack = charge_stack.sort_values("rank")
                fig3 = px.bar(
                    charge_stack, x="Local Total", y="Description", color="Currency",
                    orientation="h", color_discrete_map={BASE_CCY: ACCENT, FX_CCY: PRIMARY},
                    barmode="stack", labels={"Local Total": LOCAL_LABEL, "Description": ""},
                )
                fig3.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title="")
                fig3.update_traces(marker_line_width=0)
                st.plotly_chart(fig3, use_container_width=True)
                st.caption(f"Non-{BASE_CCY} charges are shown converted to {BASE_CCY}.")

            with r2r:
                st.markdown('<div class="section-title">Shipments by Incoterms</div>', unsafe_allow_html=True)
                if "Incoterms" in ship_sum_f.columns:
                    inco_data = ship_sum_f.groupby("Incoterms")["Shipment Job"].nunique().reset_index()
                    inco_data.columns = ["Incoterms", "Shipments"]
                    fig4 = px.pie(inco_data, values="Shipments", names="Incoterms",
                                  color_discrete_sequence=PALETTE, hole=0.45)
                    fig4.update_layout(**PLOTLY_LAYOUT)
                    fig4.update_traces(textfont_color="#ffffff", textfont_size=13)
                    st.plotly_chart(fig4, use_container_width=True)

            st.markdown(f'<div class="section-title">Monthly spend by Incoterm ({BASE_CCY})</div>',
                        unsafe_allow_html=True)
            bd_m2, month_order2 = add_month_col(bill_det_f, ship_sum_f)
            if month_order2 and "Incoterms" in bd_m2.columns:
                inco_monthly = bd_m2.groupby(["Month", "Incoterms"])["Local Total"].sum().reset_index()
                inco_monthly = inco_monthly[inco_monthly["Local Total"] > 0]
                inco_monthly["Month"] = pd.Categorical(
                    inco_monthly["Month"], categories=month_order2, ordered=True)
                inco_monthly = inco_monthly.sort_values("Month")
                fig5 = px.line(
                    inco_monthly, x="Month", y="Local Total", color="Incoterms",
                    color_discrete_sequence=PALETTE, markers=True,
                    labels={"Local Total": LOCAL_LABEL, "Month": ""},
                )
                fig5.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title=LOCAL_LABEL)
                fig5.update_traces(line_width=2.5, marker_size=8)
                st.plotly_chart(fig5, use_container_width=True)
                st.caption("Each line tracks total spend per Incoterm by ETD month.")
            else:
                st.info("Add shipments with ETD dates and Incoterms to see the monthly trend.")

            st.markdown('<div class="section-title">Supplier summary</div>', unsafe_allow_html=True)
            st.dataframe(supp_f, use_container_width=True, hide_index=True)

    with tab5:
        st.markdown('<div class="section-title">Download consolidated report</div>', unsafe_allow_html=True)
        active = any([f_supplier, f_origin, f_dest, f_inco, f_mode, f_month])
        scope = (f"**{len(bill_sum_f)} of {len(bill_sum)} shipments** match the current filters."
                 if active else f"**All {len(bill_sum)} billed shipments** are included.")
        st.info(scope)
        st.download_button(
            label="Download Excel report",
            data=create_report(ship_sum_f, bill_sum_f, bill_det_f, supp_f),
            file_name=REPORT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.markdown(f"""
**Sheets included:**
- **Shipment Summary** — one row per shipment
- **Billing Summary** — per-job {BASE_CCY} / {FX_CCY} / {LOCAL_LABEL}
- **Billing Detail** — every charge line
- **Supplier Summary** — shipment count and spend per supplier
- **Analysis** — key metrics and spend breakdowns
        """)


# ── Project 2: per-container charge breakdown ─────────────────────────────────
# Self-contained: its own uploaders, so it never interferes with project 1.
# The shipment listing alone is enough — the billing export is optional.

def render_container_breakdown():
    st.markdown('<div class="section-title">Charge breakdown by container</div>',
                unsafe_allow_html=True)
    st.caption(
        "Shipments are grouped into Domestic, Export and Import, and every "
        "charge line is allocated to the BPO for its leg of the move. Amounts "
        "are split evenly across the job's containers in whole cents, so the "
        "parts always add back to the invoiced figure. The shipment listing on "
        "its own produces the structure — upload the billing export when the "
        "figures are ready."
    )

    st.markdown('<div class="section-title">Upload files</div>', unsafe_allow_html=True)
    up_left, up_right = st.columns(2)
    with up_left:
        shipment_files = st.file_uploader(
            "Shipment listing report(s) (required)", type=CB_UPLOAD_TYPES,
            accept_multiple_files=True, key="cb_shipment",
            help="One or more listings, .xlsx or .xls. Upload the consignee view and "
                 "the origin view together to cover domestic, export and import jobs.")
    with up_right:
        billing_files = st.file_uploader(
            "Billing export(s) (optional)", type=CB_UPLOAD_TYPES,
            accept_multiple_files=True, key="cb_billing")

    if not shipment_files:
        st.info("Upload the shipment listing report to begin. The billing export is "
                "optional: without it the workbook comes out as the grouping and "
                "BPO structure, with no amounts.")
        return

    download_slot = st.container()

    with st.spinner("Reading the shipment listing…"):
        shipment_df, job_cols, read_errors = cb_read_shipments(shipment_files)
    for message in read_errors:
        st.error(f"Could not read the shipment listing {message}")
    if shipment_df is None or shipment_df.empty:
        st.error("No shipment rows were found. Check that the listing contains a "
                 "job reference column and container numbers. Legacy .xls files "
                 "need the xlrd package installed.")
        return

    billing = None
    diag = {"missing_columns": [], "unmatched_jobs": [], "jobs_without_containers": []}

    if billing_files:
        with st.spinner("Allocating charges across containers…"):
            sheets = []
            for uploaded in billing_files:
                try:
                    sheets.append(cb_read_table(uploaded, header=0))
                except Exception as e:
                    st.error(f"Could not read {uploaded.name}: {e}")
            if not sheets:
                return
            billing = pd.concat(sheets, ignore_index=True)
            billing.columns = [str(c).strip() for c in billing.columns]
            detail, diag = cb_build_charge_detail(billing, shipment_df)
            detail = bpo_annotate(detail)

        if diag["missing_columns"]:
            st.error("The billing export is missing: " + ", ".join(diag["missing_columns"]))
            return
        if detail is None or detail.empty:
            st.warning("No charge lines could be matched to a shipment in the listing.")
            return
    else:
        detail = bpo_shipment_skeleton(shipment_df)
        if detail.empty:
            st.error("No containers could be read from the shipment listing.")
            return
        st.info("No billing export uploaded, so the workbook carries the grouping and "
                "BPO structure with no amounts.")

    with_charges = bpo_has_charges(detail)

    st.caption(f"Shipment listing read from {len(shipment_df)} shipments "
               f"(job reference column: {'; '.join(job_cols)}).")

    total_ex = detail[CB_ALLOC_EX].sum()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Shipments", f"{detail['Job'].nunique():,}")
    k2.metric("Containers", f"{detail['Container'].nunique():,}")
    k3.metric("BPOs", f"{detail['BPO'].nunique():,}")
    k4.metric(f"Total ex tax ({CB_LOCAL_CCY})", f"{total_ex:,.2f}")

    if diag["unmatched_jobs"] and billing is not None:
        missing = billing[billing[CB_JOB_COL].astype(str).str.strip()
                          .isin(diag["unmatched_jobs"])]
        missing_total = pd.to_numeric(missing["Local Amount"], errors="coerce").fillna(0).sum()
        st.warning(
            f"{len(diag['unmatched_jobs'])} job(s) on the billing export are not in any "
            f"uploaded shipment listing, so {CB_LOCAL_CCY} {missing_total:,.2f} ex tax is "
            "excluded from the breakdown: " + ", ".join(sorted(diag["unmatched_jobs"])) +
            ". Upload the listing that contains them to include them."
        )

    placeholders = detail["Container"].astype(str).str.contains(
        r"\(not listed\)|\(cut off in listing\)", regex=True)
    if placeholders.any():
        affected = detail.loc[placeholders, "Job"].nunique()
        st.info(
            f"On {affected} shipment(s) the listing names fewer containers than the "
            "shipment carries (the export cuts long container lists short, and some "
            "bookings have no numbers issued yet). The missing containers are added "
            "as numbered placeholders, so charges are still split across the full "
            "container count."
        )
    if diag["jobs_without_containers"]:
        st.info(
            f"{len(diag['jobs_without_containers'])} shipment(s) have no container number, "
            f"so their charges sit against a '{CB_NO_CONTAINER}' row rather than being dropped."
        )

    unmapped = bpo_unmapped_codes(detail)
    if unmapped:
        st.warning(
            "These charge codes are not mapped to a leg of the move, so they are "
            "carried on a TBA reference: " + ", ".join(unmapped) +
            ". Add them to BPO_LEG_BY_CODE to allocate them."
        )

    # ── Group filter ──────────────────────────────────────────────────────────
    # The filter drives the on-screen views only. The download always covers
    # everything that was uploaded.

    groups_present = [g for g in BPO_GROUP_ORDER if g in set(detail["Group"])]
    chosen = st.multiselect("Trade group", groups_present,
                            placeholder="All groups", key="cb_groups")
    view_detail = detail[detail["Group"].isin(chosen)] if chosen else detail

    if view_detail.empty:
        st.warning("No rows match the selected group.")
        return

    views = ["Container matrix", "Per BPO", "Per group", "Per shipment"]
    if with_charges:
        views = ["Charge detail"] + views + ["Per container", "Per charge code"]
    view = st.radio("View", views, horizontal=True, key="cb_view",
                    label_visibility="collapsed")

    if view == "Charge detail":
        st.dataframe(bpo_charge_detail(view_detail), use_container_width=True,
                     hide_index=True)
        st.caption(f"{len(view_detail):,} allocated charge lines")
    elif view == "Container matrix":
        matrix = bpo_matrix_flat(view_detail)
        st.dataframe(matrix, use_container_width=True, hide_index=True)
        st.caption(f"{len(matrix):,} rows — one per container, currency and BPO. "
                   "The Excel download splits these into one block per group.")
    elif view == "Per BPO":
        st.dataframe(bpo_summary(view_detail), use_container_width=True, hide_index=True)
        st.caption("Compare the Currency and BPO Currency columns: a difference means "
                   "the charge was raised in a currency other than the BPO's.")
    elif view == "Per group":
        st.dataframe(bpo_group_summary(view_detail), use_container_width=True,
                     hide_index=True)
    elif view == "Per shipment":
        st.dataframe(bpo_shipment_groups(view_detail), use_container_width=True,
                     hide_index=True)
    elif view == "Per container":
        st.dataframe(cb_container_summary(view_detail), use_container_width=True,
                     hide_index=True)
    else:
        code_summary = cb_charge_code_summary(view_detail)
        st.dataframe(code_summary, use_container_width=True, hide_index=True)
        if not code_summary.empty:
            fig = px.bar(code_summary.sort_values("Total_ex_Tax").tail(15),
                         x="Total_ex_Tax", y="Charge Code", orientation="h",
                         color_discrete_sequence=[PRIMARY],
                         labels={"Total_ex_Tax": f"Total ex tax ({CB_LOCAL_CCY})",
                                 "Charge Code": ""})
            fig.update_layout(**PLOTLY_LAYOUT, xaxis_title="", yaxis_title="")
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True)

    with download_slot:
        st.download_button(
            label="Download Excel breakdown",
            data=cb_create_report(detail),
            file_name=CB_REPORT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="cb_download",
        )
        st.caption(f"{detail['Container'].nunique():,} containers · "
                   f"4 sheets: BPO Summary, Container Matrix, Shipment Groups, "
                   f"Charge Detail")

    if with_charges and billing is not None:
        source_total = pd.to_numeric(billing["Local Amount"], errors="coerce").fillna(0).sum()
        unmatched_total = 0.0
        if diag["unmatched_jobs"]:
            unmatched = billing[billing[CB_JOB_COL].astype(str).str.strip()
                                .isin(diag["unmatched_jobs"])]
            unmatched_total = pd.to_numeric(unmatched["Local Amount"],
                                            errors="coerce").fillna(0).sum()
        difference = round(total_ex + unmatched_total - source_total, 2)
        if abs(difference) < 0.01:
            st.success(f"Reconciled: allocated {CB_LOCAL_CCY} {total_ex:,.2f} plus "
                       f"{CB_LOCAL_CCY} {unmatched_total:,.2f} unmatched equals the "
                       f"export's {CB_LOCAL_CCY} {source_total:,.2f}.")
        else:
            st.error(f"Out by {CB_LOCAL_CCY} {difference:,.2f} against the export's "
                     f"{CB_LOCAL_CCY} {source_total:,.2f}.")


# ── Projects ──────────────────────────────────────────────────────────────────

session_left, session_right = st.columns([5, 1])
with session_right:
    if st.button("Sign out", key="sign_out"):
        st.session_state.clear()
        st.rerun()
with session_left:
    st.caption(f"Signed in as {st.session_state.get('auth_display', st.session_state['auth_user'])}")

# Only the tabs this account is granted. Content for a tab the user cannot see
# is never rendered, so it is not sent to the browser at all.
granted = allowed_tabs(st.session_state["auth_user"])

if not granted:
    st.error("This account has no reports assigned. Contact the report owner.")
    st.stop()

renderers = {
    TAB_BILLING: render_billing_report,
    TAB_CONTAINER: render_container_breakdown,
}

for tab, container in zip(granted, st.tabs([TAB_LABELS[t] for t in granted])):
    with container:
        renderers[tab]()
