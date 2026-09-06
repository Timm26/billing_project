"""
Trade-lane grouping and BPO allocation for the per-container charge breakdown.

Every shipment falls into one trade group, judged against New Zealand as the
home country:

    NZ  ->  NZ          Domestic
    NZ  ->  anywhere    Export
    anywhere  ->  NZ    Import

and every charge line falls into one leg of the move (Door to Wharf / Ocean
freight / Wharf to Door). The group and the leg together decide which Blind
Purchase Order the charge is billed against, and which Rohlig entity and vendor
code sit behind it.

Rows are split by charge currency: a container billed partly in NZD and partly
in USD appears twice, once per currency, so each row can be invoiced against
the BPO in the currency that BPO is raised in. Every row also carries its local
NZD equivalent, which is what the whole report still reconciles on.

Three tables normally need editing:

    HOME_COUNTRY / GROUP_RULES   how a lane maps to a group
    LEG_BY_CODE                  which charge code belongs to which leg
    BPO_TABLE                    the BPO / vendor / entity per (group, leg)

Anything a rule does not recognise is never dropped: a lane with neither end in
the home country lands in GROUP_OTHER, an unknown charge code lands in
LEG_UNMAPPED, and both are billed to TBA and listed on the exceptions sheet.
"""

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Groups ────────────────────────────────────────────────────────────────────

HOME_COUNTRY = "NZ"  # two-letter UN/LOCODE prefix the groups are judged against

GROUP_DOMESTIC = "Domestic (within New Zealand)"
GROUP_EXPORT = "Export (New Zealand to overseas)"
GROUP_IMPORT = "Import (overseas to New Zealand)"
GROUP_OTHER = "Other (neither end in New Zealand)"

GROUP_ORDER = [GROUP_DOMESTIC, GROUP_EXPORT, GROUP_IMPORT, GROUP_OTHER]

# Overrides for named lane pairs, if one ever has to sit outside the direction
# rule: {("AU", "NZ"): GROUP_IMPORT}. Normally left empty.
GROUP_RULES = {}

# Who receives the invoice for each group.
INVOICE_TO = {
    GROUP_DOMESTIC: "NZ (Eileen)",
    GROUP_EXPORT: "NZ (Eileen)",
    GROUP_IMPORT: "AU (Hayley)",
    GROUP_OTHER: "TBA",
}

# ── Legs ──────────────────────────────────────────────────────────────────────

LEG_ORIGIN = "Door to Wharf"
LEG_FREIGHT = "Ocean freight"
LEG_DEST = "Wharf to Door"
LEG_UNMAPPED = "Unmapped"

LEG_ORDER = [LEG_ORIGIN, LEG_FREIGHT, LEG_DEST, LEG_UNMAPPED]

# Charge code -> leg. Codes are matched upper-cased and stripped.
LEG_BY_CODE = {
    # Pick-up and origin side: everything up to the ship's rail.
    "OTHC": LEG_ORIGIN,       # Origin Terminal Handling
    "ODOC": LEG_ORIGIN,       # Origin Documentation Fee
    "OCART": LEG_ORIGIN,      # Origin cartage, if it ever appears
    "OCARTFSC": LEG_ORIGIN,
    "LOAD": LEG_ORIGIN,
    "EXPENT": LEG_ORIGIN,     # Export entry

    # The ocean move itself and the surcharges that ride on the freight rate.
    "FRT": LEG_FREIGHT,       # Ocean Freight
    "BAF": LEG_FREIGHT,       # Bunker Adjustment Factor
    "CAF": LEG_FREIGHT,
    "PSS": LEG_FREIGHT,
    "EBS": LEG_FREIGHT,
    "GRI": LEG_FREIGHT,
    "ISPS": LEG_FREIGHT,

    # Discharge port through to the delivery door, including clearance.
    "DTHC": LEG_DEST,         # Destination Terminal Handling
    "DDOC": LEG_DEST,         # Destination Documentation Fee
    "DDOF": LEG_DEST,         # Destination Delivery Order Fee
    "DTIF": LEG_DEST,         # Terminal Infrastructure Fee
    "DCART": LEG_DEST,        # Cartage to the delivery address
    "DCARTFSC": LEG_DEST,     # Cartage Fuel Surcharge
    "DEMUR": LEG_DEST,        # Container Demurrage
    "DWASH": LEG_DEST,        # Container Wash
    "CUSENT": LEG_DEST,       # Customs Entry Lodgement
    "MPIENT": LEG_DEST,       # Biosecurity Entry Fee
    "MPIINS": LEG_DEST,       # Biosecurity Inspection
    "AGENCY": LEG_DEST,       # Agency Fee
    "STORAGE": LEG_DEST,
    "UNPACK": LEG_DEST,
}

# ── BPOs ──────────────────────────────────────────────────────────────────────

TBA = "TBA"


def _bpo(bpo, currency, entity, vendor):
    return {"BPO": bpo, "BPO Currency": currency,
            "BPO Entity": entity, "BPO Vendor": vendor}


UNALLOCATED = _bpo(TBA, TBA, TBA, TBA)

# (group, leg) -> BPO record. A pair not listed here is billed to TBA.
BPO_TABLE = {
    # Coastal NZ (AKL to CHC, ANL & Maersk), invoiced to NZ. One BPO covers the
    # whole domestic move, delivery leg included.
    (GROUP_DOMESTIC, LEG_ORIGIN): _bpo("77263466-330920", "NZD", "Rohlig NZ", "30098"),
    (GROUP_DOMESTIC, LEG_FREIGHT): _bpo("77263466-330920", "NZD", "Rohlig NZ", "30098"),
    (GROUP_DOMESTIC, LEG_DEST): _bpo("77263466-330920", "NZD", "Rohlig NZ", "30098"),

    # NZ to overseas, invoiced to NZ.
    (GROUP_EXPORT, LEG_ORIGIN): _bpo("77263466-330920", "NZD", "Rohlig NZ", "30098"),
    (GROUP_EXPORT, LEG_FREIGHT): _bpo("77263749-330924", "USD", "Rohlig NZ", "62170"),
    (GROUP_EXPORT, LEG_DEST): _bpo("77260204-330932", "AUD", "Rohlig AU", "61935"),

    # Overseas to NZ, invoiced to AU.
    (GROUP_IMPORT, LEG_ORIGIN): _bpo("77250782-714720", "AUD", "Rohlig AU", "29803"),
    (GROUP_IMPORT, LEG_FREIGHT): _bpo("77254657-709817", "USD", "Rohlig AU", "61934"),
    (GROUP_IMPORT, LEG_DEST): _bpo(f"{TBA}-{TBA}", "NZD", "Rohlig NZ", "30494"),
}

# Columns this module adds to the charge detail.
ANNOTATION_COLS = ["Group", "Leg", "BPO", "BPO Vendor", "BPO Entity",
                   "BPO Currency", "Invoice To"]

# Amounts on the charge detail. The first pair is in the charge's own currency,
# the second is the local equivalent the rest of the report totals on.
AMOUNT_CCY = "Allocated Amount"
INC_CCY = "Allocated Total"
ALLOC_EX = "Allocated Local Amount"
ALLOC_INC = "Allocated Local Total"
LOCAL_CCY = "NZD"

LOCAL_EX_LABEL = f"Local ex Tax ({LOCAL_CCY})"
LOCAL_INC_LABEL = f"Local inc Tax ({LOCAL_CCY})"


# ── Classification ────────────────────────────────────────────────────────────

def country(port_code):
    """'NZAKL' -> 'NZ'. Anything unusable comes back empty."""
    text = str(port_code).strip().upper()
    return text[:2] if len(text) >= 2 and text[:2].isalpha() else ""


def classify_group(origin, destination):
    """Trade group for one shipment, from its origin and destination ports.

    Judged by direction against HOME_COUNTRY, so every overseas origin is an
    import and every overseas destination is an export, whichever country it
    is. GROUP_RULES overrides the direction rule for named lane pairs.
    """
    origin_country, dest_country = country(origin), country(destination)
    override = GROUP_RULES.get((origin_country, dest_country))
    if override:
        return override
    at_home = (origin_country == HOME_COUNTRY, dest_country == HOME_COUNTRY)
    if at_home == (True, True):
        return GROUP_DOMESTIC
    if at_home == (True, False):
        return GROUP_EXPORT
    if at_home == (False, True):
        return GROUP_IMPORT
    return GROUP_OTHER


def leg_for_code(charge_code):
    """Leg of the move a charge code belongs to."""
    return LEG_BY_CODE.get(str(charge_code).strip().upper(), LEG_UNMAPPED)


def bpo_for(group, leg):
    """BPO record for a (group, leg) pair, or the TBA placeholder."""
    return BPO_TABLE.get((group, leg), UNALLOCATED)


def bpo_label(group, leg):
    """Column header for a leg subtotal: 'Door to Wharf (BPO 772... / NZD)'."""
    record = bpo_for(group, leg)
    return f"{leg} (BPO {record['BPO']} / {record['BPO Currency']})"


def annotate(detail):
    """Add Group, Leg and BPO columns to a charge detail frame.

    Expects the frame produced by cb_build_charge_detail: one row per charge
    line per container, with Origin, Destination and Charge Code columns.
    """
    if detail is None or detail.empty:
        return detail

    out = detail.copy()
    origin = out["Origin"] if "Origin" in out.columns else pd.Series("", index=out.index)
    dest = out["Destination"] if "Destination" in out.columns else pd.Series("", index=out.index)

    out["Group"] = [classify_group(o, d) for o, d in zip(origin, dest)]
    out["Leg"] = [leg_for_code(c) for c in out["Charge Code"]]

    records = [bpo_for(g, l) for g, l in zip(out["Group"], out["Leg"])]
    out["BPO"] = [r["BPO"] for r in records]
    out["BPO Vendor"] = [r["BPO Vendor"] for r in records]
    out["BPO Entity"] = [r["BPO Entity"] for r in records]
    out["BPO Currency"] = [r["BPO Currency"] for r in records]
    out["Invoice To"] = [INVOICE_TO.get(g, TBA) for g in out["Group"]]

    if "Currency" in out.columns:
        out["Currency"] = (out["Currency"].fillna(LOCAL_CCY).astype(str)
                           .str.strip().str.upper())
    else:
        out["Currency"] = LOCAL_CCY

    out["_group_rank"] = out["Group"].map(
        {g: i for i, g in enumerate(GROUP_ORDER)}).fillna(len(GROUP_ORDER))
    out["_leg_rank"] = out["Leg"].map(
        {l: i for i, l in enumerate(LEG_ORDER)}).fillna(len(LEG_ORDER))
    out = out.sort_values(["_group_rank", "Job", "Container", "Currency",
                           "_leg_rank", "Charge Code"])
    return out.drop(columns=["_group_rank", "_leg_rank"]).reset_index(drop=True)


def group_rank(group):
    return GROUP_ORDER.index(group) if group in GROUP_ORDER else len(GROUP_ORDER)


def sort_by_group(df, then=()):
    """Sort any annotated frame into the canonical group order."""
    if df is None or df.empty or "Group" not in df.columns:
        return df
    out = df.copy()
    out["_rank"] = out["Group"].map(group_rank)
    keys = ["_rank"] + [c for c in then if c in out.columns]
    return out.sort_values(keys).drop(columns="_rank").reset_index(drop=True)


def currencies_in(frame):
    """Currencies present, local currency first then the rest alphabetically."""
    seen = sorted(set(frame["Currency"].dropna().astype(str)))
    return ([LOCAL_CCY] if LOCAL_CCY in seen else []) + [c for c in seen if c != LOCAL_CCY]


# ── Summaries ─────────────────────────────────────────────────────────────────

def group_summary(detail):
    """One row per trade group, in local currency."""
    summary = detail.groupby("Group", observed=True).agg(
        Shipments=("Job", "nunique"),
        Containers=("Container", "nunique"),
        Charge_Lines=("Charge Code", "count"),
        **{LOCAL_EX_LABEL: (ALLOC_EX, "sum"), LOCAL_INC_LABEL: (ALLOC_INC, "sum")},
    ).reset_index()
    summary["Invoice To"] = summary["Group"].map(INVOICE_TO).fillna(TBA)
    summary["Currencies"] = summary["Group"].map(
        detail.groupby("Group", observed=True)["Currency"]
        .agg(lambda s: ", ".join(sorted(set(s)))))
    summary["Avg_per_Container"] = (
        summary[LOCAL_EX_LABEL] / summary["Containers"].replace(0, pd.NA)).round(2)
    summary["Share_of_Spend_%"] = (
        100 * summary[LOCAL_EX_LABEL] / summary[LOCAL_EX_LABEL].sum()).round(2)
    for col in (LOCAL_EX_LABEL, LOCAL_INC_LABEL):
        summary[col] = summary[col].round(2)
    summary["_rank"] = summary["Group"].map(group_rank)
    ordered = ["Group", "Invoice To", "Currencies", "Shipments", "Containers",
               "Charge_Lines", LOCAL_EX_LABEL, LOCAL_INC_LABEL,
               "Avg_per_Container", "Share_of_Spend_%"]
    return (summary.sort_values("_rank").drop(columns="_rank")
            .reset_index(drop=True)[ordered])


def bpo_summary(detail):
    """One row per BPO and currency: what to raise, and in what.

    A BPO carrying charges in more than one currency gets a row each, so an
    ocean-freight BPO raised in USD is never blended with NZD charges.
    """
    summary = detail.groupby(["Group", "Leg", "BPO", "Currency"], observed=True).agg(
        Shipments=("Job", "nunique"),
        Containers=("Container", "nunique"),
        Charge_Lines=("Charge Code", "count"),
        Charge_Codes=("Charge Code", lambda s: ", ".join(sorted(set(s)))),
        Total_ex_Tax=(AMOUNT_CCY, "sum"),
        Total_inc_Tax=(INC_CCY, "sum"),
        **{LOCAL_EX_LABEL: (ALLOC_EX, "sum")},
    ).reset_index()

    records = [bpo_for(g, l) for g, l in zip(summary["Group"], summary["Leg"])]
    summary["BPO Vendor"] = [r["BPO Vendor"] for r in records]
    summary["BPO Entity"] = [r["BPO Entity"] for r in records]
    summary["BPO Currency"] = [r["BPO Currency"] for r in records]
    summary["Invoice To"] = summary["Group"].map(INVOICE_TO).fillna(TBA)
    summary["Currency Matches BPO"] = [
        "yes" if str(a) == str(b) else "check"
        for a, b in zip(summary["Currency"], summary["BPO Currency"])]
    for col in ("Total_ex_Tax", "Total_inc_Tax", LOCAL_EX_LABEL):
        summary[col] = summary[col].round(2)

    summary["_g"] = summary["Group"].map(group_rank)
    summary["_l"] = summary["Leg"].map(
        {l: i for i, l in enumerate(LEG_ORDER)}).fillna(len(LEG_ORDER))
    ordered = ["Group", "Invoice To", "Leg", "BPO", "BPO Vendor", "BPO Entity",
               "BPO Currency", "Currency", "Currency Matches BPO", "Shipments",
               "Containers", "Charge_Lines", "Total_ex_Tax", "Total_inc_Tax",
               LOCAL_EX_LABEL, "Charge_Codes"]
    return (summary.sort_values(["_g", "_l", "Currency"]).drop(columns=["_g", "_l"])
            .reset_index(drop=True)[ordered])


def shipment_groups(detail):
    """One row per shipment and currency, with the BPO behind each leg."""
    rows = []
    for (group, job, currency), part in detail.groupby(
            ["Group", "Job", "Currency"], observed=True):
        row = {
            "Group": group,
            "Invoice To": INVOICE_TO.get(group, TBA),
            "Job": job,
            "Currency": currency,
            "Containers": part["Container"].nunique(),
            "Total_ex_Tax": round(part[AMOUNT_CCY].sum(), 2),
            "Total_inc_Tax": round(part[INC_CCY].sum(), 2),
            LOCAL_EX_LABEL: round(part[ALLOC_EX].sum(), 2),
        }
        for col in ("Consignor", "Origin", "Destination", "Vessel", "Shipping Line",
                    "Incoterms", "ETD", "ETA"):
            if col in part.columns:
                row[col] = part[col].iloc[0]
        for leg in LEG_ORDER:
            leg_part = part[part["Leg"] == leg]
            if leg_part.empty:
                continue
            row[f"{leg} BPO"] = bpo_for(group, leg)["BPO"]
            row[f"{leg} ex Tax"] = round(leg_part[AMOUNT_CCY].sum(), 2)
        rows.append(row)
    return sort_by_group(pd.DataFrame(rows), then=["Job", "Currency"])


def exceptions(detail):
    """Charges that could not be placed in a group or against a leg.

    A BPO that is simply not issued yet is not an exception — it is a known
    gap, and tba_summary reports it. This sheet is for lanes with neither end
    in the home country and charge codes the leg map does not know, which is
    what actually needs a decision.
    """
    mask = (detail["Group"] == GROUP_OTHER) | (detail["Leg"] == LEG_UNMAPPED)
    if not mask.any():
        return pd.DataFrame()
    cols = ["Group", "Job", "Container", "Charge Code", "Description", "Leg",
            "BPO", "Origin", "Destination", "Currency", AMOUNT_CCY, ALLOC_EX]
    out = detail.loc[mask, [c for c in cols if c in detail.columns]].copy()
    out["Reason"] = [
        "Neither origin nor destination is in the home country" if g == GROUP_OTHER
        else "Charge code is not mapped to a leg"
        for g in out["Group"]]
    return out.reset_index(drop=True)


def tba_summary(detail):
    """Spend sitting on a BPO that has not been issued yet."""
    mask = detail["BPO"].astype(str).str.contains(TBA)
    if not mask.any():
        return pd.DataFrame()
    summary = detail[mask].groupby(["Group", "Leg", "BPO", "Currency"],
                                   observed=True).agg(
        Shipments=("Job", "nunique"),
        Containers=("Container", "nunique"),
        Total_ex_Tax=(AMOUNT_CCY, "sum"),
        **{LOCAL_EX_LABEL: (ALLOC_EX, "sum")},
    ).reset_index()
    for col in ("Total_ex_Tax", LOCAL_EX_LABEL):
        summary[col] = summary[col].round(2)
    return sort_by_group(summary, then=["Leg", "Currency"])


def unmapped_codes(detail):
    """Charge codes seen in the data that LEG_BY_CODE does not cover."""
    unmapped = detail[detail["Leg"] == LEG_UNMAPPED]
    if unmapped.empty:
        return []
    return sorted(unmapped["Charge Code"].unique())


# ── Grouped container matrix ──────────────────────────────────────────────────

MATRIX_KEYS = ["Job", "Container", "Container Size", "Currency"]


def group_matrix(detail, group):
    """Container matrix for one group, one row per container and currency.

    Charge-code values are in the row's own currency; the last column carries
    the local equivalent so the sheet still reconciles to the export. Codes run
    in leg order, each leg followed by a subtotal column headed with its BPO.
    """
    part = detail[detail["Group"] == group].copy()
    if part.empty:
        return pd.DataFrame(), {}

    # A blank container size would drop the whole row out of the pivot, which
    # is how shipments billed without a container number used to disappear.
    part["Container Size"] = (part["Container Size"].fillna("").astype(str)
                              if "Container Size" in part.columns else "")

    matrix = part.pivot_table(
        index=MATRIX_KEYS, columns="Charge Code", values=AMOUNT_CCY,
        aggfunc="sum", observed=True).reset_index()
    matrix.columns.name = None

    codes = [c for c in matrix.columns if c not in MATRIX_KEYS]
    legs = {}
    for leg in LEG_ORDER:
        leg_codes = sorted(c for c in codes if leg_for_code(c) == leg)
        if leg_codes:
            legs[leg] = leg_codes

    ordered_codes = [c for leg in legs for c in legs[leg]]
    matrix = matrix[MATRIX_KEYS + ordered_codes]

    subtotal_cols = {}
    for leg, leg_codes in legs.items():
        label = bpo_label(group, leg)
        totals = matrix[leg_codes].sum(axis=1).round(2)
        # A leg with nothing in this row's currency stays blank rather than
        # showing a 0.00 that reads like a real charge.
        matrix[label] = totals.mask(totals == 0)
        subtotal_cols[leg] = label

    matrix["TOTAL"] = matrix[ordered_codes].sum(axis=1).round(2)

    local = (part.groupby(MATRIX_KEYS, observed=True)[ALLOC_EX].sum()
             .round(2).reset_index().rename(columns={ALLOC_EX: LOCAL_EX_LABEL}))
    matrix = pd.merge(matrix, local, how="left", on=MATRIX_KEYS)

    matrix = matrix.sort_values(["Job", "Container", "Currency"]).reset_index(drop=True)
    return matrix, subtotal_cols


def grouped_matrix_flat(detail):
    """All groups stacked into one frame, with Group as the first column.

    Leg subtotals get generic headers here, with the BPO in its own column, so
    the columns line up across groups. The Excel sheet uses per-group blocks.
    """
    frames = []
    for group in GROUP_ORDER:
        matrix, subtotal_cols = group_matrix(detail, group)
        if matrix.empty:
            continue
        for leg, label in subtotal_cols.items():
            matrix = matrix.rename(columns={label: f"{leg} ex Tax"})
            matrix[f"{leg} BPO"] = bpo_for(group, leg)["BPO"]
        matrix.insert(0, "Group", group)
        matrix.insert(1, "Invoice To", INVOICE_TO.get(group, TBA))
        frames.append(matrix)
    if not frames:
        return pd.DataFrame()

    flat = pd.concat(frames, ignore_index=True)
    keys = ["Group", "Invoice To"] + MATRIX_KEYS
    legs = [f"{leg} {suffix}" for leg in LEG_ORDER
            for suffix in ("BPO", "ex Tax") if f"{leg} {suffix}" in flat.columns]
    tail = ["TOTAL", LOCAL_EX_LABEL]
    codes = [c for c in flat.columns if c not in keys + legs + tail]
    codes = sorted(codes, key=lambda c: (LEG_ORDER.index(leg_for_code(c)), c))
    return flat[keys + codes + legs + tail]


# ── Excel writing ─────────────────────────────────────────────────────────────

HEADER_FILL = "33475B"    # column headers
GROUP_FILL = "0E8074"     # group title row
LEGEND_FILL = "E4EDF2"    # BPO legend row under the title
BAND_FILL = "F1F5F9"      # alternating job band
SUBTOTAL_FILL = "DCE6EE"  # currency and group total rows

_thin = Side(style="thin", color="B7C4D1")


def _write_row(ws, row, values, fill=None, bold=False, colour=None,
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
            cell.border = Border(top=_thin)
    return row + 1


def write_grouped_matrix(writer, detail, sheet_name="Container Matrix",
                         number_format="#,##0.00"):
    """Write the container matrix as one block per trade group.

    Each block opens with a title row naming the group and who it is invoiced
    to, a legend row per leg spelling out the BPO, then the container rows —
    one per container per currency — a total per currency, and the group total
    in local currency.
    """
    ws = writer.book.create_sheet(sheet_name[:31])
    row = 1
    widest = 0
    layout_rows = set()  # title and legend rows, excluded from column sizing

    for group in GROUP_ORDER:
        matrix, subtotal_cols = group_matrix(detail, group)
        if matrix.empty:
            continue

        part = detail[detail["Group"] == group]
        title = (f"{group.upper()}  ·  Invoice to {INVOICE_TO.get(group, TBA)}  ·  "
                 f"{part['Job'].nunique()} shipments  ·  "
                 f"{part['Container'].nunique()} containers  ·  "
                 f"{', '.join(currencies_in(part))}")
        layout_rows.add(row)
        row = _write_row(ws, row, [title], fill=GROUP_FILL, bold=True, colour="FFFFFF")

        for leg in LEG_ORDER:
            if leg not in subtotal_cols:
                continue
            record = bpo_for(group, leg)
            layout_rows.add(row)
            row = _write_row(ws, row, [
                f"{leg}:  BPO {record['BPO']}   ·   {record['BPO Currency']}   ·   "
                f"{record['BPO Entity']}   ·   Vendor {record['BPO Vendor']}"
            ], fill=LEGEND_FILL)

        header_row = row
        row = _write_row(ws, row, list(matrix.columns), fill=HEADER_FILL,
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
            row = _write_row(ws, row, [record[c] for c in matrix.columns],
                             fill=BAND_FILL if shaded else None,
                             number_format=number_format)

        value_cols = list(matrix.columns[len(MATRIX_KEYS):])
        first = True
        for currency in currencies_in(matrix):
            rows_in_ccy = matrix[matrix["Currency"] == currency]
            totals = [f"TOTAL {currency}", "", "", currency]
            for column in value_cols:
                total = round(float(rows_in_ccy[column].sum()), 2)
                totals.append(total if total else "")
            row = _write_row(ws, row, totals, fill=SUBTOTAL_FILL, bold=True,
                             number_format=number_format, border_top=first)
            first = False

        # Only the local column can be added across currencies, so the group
        # total carries that one alone.
        local_total = ([f"GROUP TOTAL ({LOCAL_CCY})"] + [""] * (len(matrix.columns) - 2)
                       + [round(float(matrix[LOCAL_EX_LABEL].sum()), 2)])
        row = _write_row(ws, row, local_total, fill=SUBTOTAL_FILL, bold=True,
                         number_format=number_format)

        widest = max(widest, len(matrix.columns))
        row += 2  # blank line between groups

    for index in range(1, widest + 1):
        letter = get_column_letter(index)
        longest = max((len(str(cell.value)) for cell in ws[letter]
                       if cell.value is not None and cell.row not in layout_rows),
                      default=10)
        ws.column_dimensions[letter].width = min(max(longest + 2, 11), 30)
    ws.freeze_panes = "E1"
    return ws
