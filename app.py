import streamlit as st
import pandas as pd
import io
import os
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Röhlig | TGG Billing Report",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');
html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
.stApp { background-color: #D6E4F0; color: #0d2d5e; }
[data-testid="stSidebar"] { background-color: #1A56A0 !important; border-right: none; }
[data-testid="stSidebar"] * { color: #ffffff !important; }
[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] { background-color: #F26522 !important; color: #ffffff !important; }
[data-testid="stSidebar"] [data-baseweb="select"] { background-color: #1e63b8 !important; }
[data-testid="stSidebar"] [data-baseweb="select"] * { background-color: #1e63b8 !important; color: #ffffff !important; border-color: #2d7dd6 !important; }
.rohlig-header { background: linear-gradient(90deg,#1A56A0 0%,#1e6ec8 100%); border-bottom:4px solid #F26522; padding:20px 32px; margin:-1rem -1rem 2rem -1rem; display:flex; align-items:center; gap:20px; }
.rohlig-header h1 { font-family:'Barlow Condensed',sans-serif; font-size:2rem; font-weight:700; color:#fff; margin:0; letter-spacing:.04em; text-transform:uppercase; }
.rohlig-header span { color:#F26522; }
.rohlig-subtitle { font-size:.78rem; color:#a8cff5; letter-spacing:.08em; text-transform:uppercase; margin-top:3px; }
.section-title { font-family:'Barlow Condensed',sans-serif; font-size:1rem; font-weight:700; color:#1A56A0; text-transform:uppercase; letter-spacing:.1em; border-left:4px solid #F26522; padding-left:10px; margin:20px 0 10px 0; }
[data-testid="stDataFrame"] { border:1px solid #c5d9f0; border-radius:6px; background-color:#ffffff; }
.stTabs [data-baseweb="tab-list"] { background-color:#ffffff; border-bottom:2px solid #c5d9f0; gap:0; }
.stTabs [data-baseweb="tab"] { font-family:'Barlow Condensed',sans-serif; font-weight:600; font-size:.85rem; letter-spacing:.06em; text-transform:uppercase; color:#6b8cba !important; background:transparent !important; border:none !important; padding:12px 24px; }
.stTabs [aria-selected="true"] { color:#1A56A0 !important; border-bottom:3px solid #F26522 !important; }
.stDownloadButton button, .stButton button { background-color:#F26522 !important; color:#fff !important; font-family:'Barlow Condensed',sans-serif !important; font-weight:700 !important; letter-spacing:.08em !important; text-transform:uppercase !important; border:none !important; border-radius:4px !important; padding:10px 28px !important; }
.stDownloadButton button:hover, .stButton button:hover { background-color:#d4541a !important; }
[data-testid="stFileUploader"] { background-color:#1e63b8; border:2px dashed #5a9fd4; border-radius:6px; padding:8px; }
[data-testid="metric-container"] { background-color:#ffffff; border:1px solid #c5d9f0; border-top:4px solid #F26522; padding:16px; border-radius:6px; box-shadow:0 2px 8px rgba(26,86,160,.08); }
[data-testid="stMetricValue"] { font-family:'Barlow Condensed',sans-serif; font-size:2rem !important; color:#1A56A0 !important; }
[data-testid="stMetricLabel"] { color:#6b8cba !important; font-size:.7rem !important; text-transform:uppercase; letter-spacing:.1em; }
div[data-testid="stVerticalBlock"] { gap:0.5rem; }
.stAlert { background-color:#ddeeff !important; border-color:#1A56A0 !important; color:#0d2d5e !important; }
</style>
""", unsafe_allow_html=True)

PLOTLY_LAYOUT = dict(
    paper_bgcolor="#ffffff", plot_bgcolor="#f5f9ff",
    font=dict(family="Barlow, sans-serif", color="#000000", size=11),
    xaxis=dict(gridcolor="#ddeeff", linecolor="#c5d9f0", tickcolor="#000000", tickfont=dict(color="#000000")),
    yaxis=dict(gridcolor="#ddeeff", linecolor="#c5d9f0", tickcolor="#000000", tickfont=dict(color="#000000")),
    margin=dict(l=20, r=20, t=36, b=20),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#000000")),
)
ORANGE  = "#F26522"
BLUE    = "#1A56A0"
LBLUE   = "#5A9FD4"
PALETTE = [ORANGE, BLUE, LBLUE, "#2ecc71", "#9b59b6", "#e74c3c", "#f1c40f"]


# ── Functions ─────────────────────────────────────────────────────────────────

def read_billing(file_obj):
    try:
        return pd.read_excel(file_obj)
    except Exception as e:
        st.error(f"Error reading billing file: {e}")
        return None

def read_shipment(file_obj):
    try:
        raw = pd.read_excel(file_obj, header=None)
        header_row = 0
        for i, row in raw.iterrows():
            vals = [str(v).strip() for v in row if str(v).strip() not in ('', 'nan')]
            if 'Shipment' in vals:
                header_row = i
                break
        file_obj.seek(0)
        df = pd.read_excel(file_obj, header=header_row, usecols=range(1, 22))
        df.columns = [str(c).strip() for c in df.columns]
        if 'Shipment' in df.columns:
            df = df[df['Shipment'].astype(str).str.match(r'^[SB]\d+')].reset_index(drop=True)
            rename = {
                'Shipment': 'Shipment Job', 'Order Ref': 'Order Reference',
                'INCO': 'Incoterms', 'Pack Mode': 'Mode',
                'Consignor Name': 'Supplier Name', 'Origin': 'Loading Port',
                'Dest.': 'Destination Port', 'Carrier Booking Reference': 'Booking Ref',
                'Vessel': 'Vessel / Voyage', 'No. of Cont.': 'Container Count',
                'Container #': 'Containers',
            }
            df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
            for col in ('ETD', 'ETA', 'ATD', 'ATA'):
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            if 'Container Count' in df.columns:
                df['Container Count'] = pd.to_numeric(df['Container Count'], errors='coerce').fillna(0).astype(int)
        return df
    except Exception as e:
        st.error(f"Error reading shipment file: {e}")
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
    return pd.merge(billing, shipment, how='left', left_on='Job', right_on='Shipment Job'), shipment

def build_shipment_summary(shipment_df):
    cols = ['Shipment Job','Supplier Name','Loading Port','Destination Port',
            'Order Reference','Incoterms','Containers','Container Count',
            'Mode','Vessel / Voyage','Booking Ref','ETD','ETA','ATD','ATA']
    return shipment_df[[c for c in cols if c in shipment_df.columns]].copy().reset_index(drop=True)

def build_billing_detail(data):
    cols = ["Shipment Job","Order Reference","Supplier Name","Loading Port",
            "Destination Port","Incoterms","Description","Currency",
            "Amount","Tax","Total","Local Total","Exchange Rate","Containers"]
    return data[[c for c in cols if c in data.columns]].copy().reset_index(drop=True)

def build_billing_summary(data, shipment_df):
    aud   = data[data['Currency']=='AUD'].groupby('Shipment Job')['Total'].sum().reset_index().rename(columns={'Total':'AUD'})
    usd   = data[data['Currency']=='USD'].groupby('Shipment Job')['Total'].sum().reset_index().rename(columns={'Total':'USD'})
    local = data.groupby('Shipment Job')['Local Total'].sum().reset_index().rename(columns={'Local Total':'Local Total (AUD)'})
    cur   = pd.merge(aud, usd, how='outer', on='Shipment Job').fillna(0)
    cur   = pd.merge(cur, local, how='left', on='Shipment Job').fillna(0)
    ctx_cols = ['Shipment Job','Supplier Name','Loading Port','Destination Port','Order Reference','Incoterms','Containers']
    ctx = shipment_df[[c for c in ctx_cols if c in shipment_df.columns]].drop_duplicates('Shipment Job')
    return pd.merge(ctx, cur, how='right', on='Shipment Job').reset_index(drop=True)

def build_supplier_summary(data, shipment_df):
    bd  = build_billing_detail(data)
    sup = bd.groupby('Supplier Name').agg(
        Shipments=('Shipment Job','nunique'),
        Charge_Lines=('Shipment Job','count'),
        Local_Total_AUD=('Local Total','sum')
    ).reset_index()
    if 'Container Count' in shipment_df.columns:
        cnt = shipment_df.groupby('Supplier Name')['Container Count'].sum().reset_index()
        sup = pd.merge(sup, cnt, how='left', on='Supplier Name')
    return sup.sort_values('Local_Total_AUD', ascending=False).reset_index(drop=True)

def create_report(shipment_sum, billing_sum, billing_det, supplier_sum):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        shipment_sum.to_excel(writer, sheet_name="Shipment Summary", index=False)
        billing_sum.to_excel(writer, sheet_name="Billing Summary", index=False)
        billing_det.to_excel(writer, sheet_name="Billing Detail", index=False)
        supplier_sum.to_excel(writer, sheet_name="Supplier Summary", index=False)
    buf.seek(0)
    return buf

def add_month_col(df, ship_sum):
    """Add Month string + sort key to a billing detail df using ETD from ship_sum."""
    if 'ETD' not in ship_sum.columns:
        return df, []
    etd_map = ship_sum.set_index('Shipment Job')['ETD']
    out = df.copy()
    out['_etd']   = pd.to_datetime(out['Shipment Job'].map(etd_map), errors='coerce')
    out['_sort']  = out['_etd'].dt.to_period('M').dt.to_timestamp()
    out['Month']  = out['_etd'].dt.strftime('%b %Y')
    order = out[['Month','_sort']].drop_duplicates().sort_values('_sort')['Month'].tolist()
    return out, order


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="rohlig-header">
  <div>
    <h1>Röhlig <span>·</span> TGG Billing Report</h1>
    <div class="rohlig-subtitle">Shipment &amp; Charge Consolidation — The Green Group (Aust) Pty Ltd</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Sidebar uploads ───────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="section-title" style="color:#fff;border-color:#F26522">Upload Files</div>', unsafe_allow_html=True)
    billing_uploads = st.file_uploader("CargoWise Billing Export(s)", type=["xlsx","csv"], accept_multiple_files=True)
    shipment_upload = st.file_uploader("Shipment Listing Report", type=["xlsx","csv"])

if not billing_uploads or not shipment_upload:
    st.info("⬅  Upload the CargoWise billing export(s) and the Shipment Listing Report to get started.")
    st.stop()

with st.spinner("Processing files…"):
    result = load_data(billing_uploads, shipment_upload)
if result is None:
    st.error("Could not load files.")
    st.stop()

data, shipment_df = result
ship_sum = build_shipment_summary(shipment_df)
bill_det = build_billing_detail(data)
bill_sum = build_billing_summary(data, shipment_df)
supp_sum = build_supplier_summary(data, shipment_df)
billed_jobs = set(data['Shipment Job'].dropna())

# Build month list for filter
bd_all, all_months = add_month_col(bill_det, ship_sum)

# ── Sidebar filters ───────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown('<div class="section-title" style="color:#fff;border-color:#F26522">Filters</div>', unsafe_allow_html=True)
    billed_ship = ship_sum[ship_sum['Shipment Job'].isin(billed_jobs)]

    f_month = st.multiselect("Month (ETD)",
        all_months, placeholder="All months")
    f_supplier = st.multiselect("Supplier",
        sorted(billed_ship['Supplier Name'].dropna().unique()), placeholder="All suppliers")
    f_origin = st.multiselect("Loading Port",
        sorted(billed_ship['Loading Port'].dropna().unique()), placeholder="All loading ports")
    f_dest = st.multiselect("Destination Port",
        sorted(billed_ship['Destination Port'].dropna().unique()), placeholder="All destinations")
    f_inco = st.multiselect("Incoterms",
        sorted(billed_ship['Incoterms'].dropna().unique()) if 'Incoterms' in billed_ship.columns else [],
        placeholder="All Incoterms")
    f_mode = st.multiselect("Mode",
        sorted(billed_ship['Mode'].dropna().unique()) if 'Mode' in billed_ship.columns else [],
        placeholder="All modes")
    st.markdown("---")
    st.markdown('<div class="section-title" style="color:#fff;border-color:#F26522">Download</div>', unsafe_allow_html=True)


def filter_df(df):
    m = pd.Series(True, index=df.index)
    if f_supplier and 'Supplier Name' in df.columns:
        m &= df['Supplier Name'].isin(f_supplier)
    if f_origin and 'Loading Port' in df.columns:
        m &= df['Loading Port'].isin(f_origin)
    if f_dest and 'Destination Port' in df.columns:
        m &= df['Destination Port'].isin(f_dest)
    if f_inco and 'Incoterms' in df.columns:
        m &= df['Incoterms'].isin(f_inco)
    if f_mode and 'Mode' in df.columns:
        m &= df['Mode'].isin(f_mode)
    return df[m]

# Apply supplier/port/incoterm/mode filters first
bill_sum_f = filter_df(bill_sum)

# Apply month filter on top — find jobs in those months
if f_month:
    jobs_in_month = bd_all[bd_all['Month'].isin(f_month)]['Shipment Job'].unique()
    bill_sum_f = bill_sum_f[bill_sum_f['Shipment Job'].isin(jobs_in_month)]

ship_sum_f = ship_sum[ship_sum['Shipment Job'].isin(bill_sum_f['Shipment Job'])]
bill_det_f = bill_det[bill_det['Shipment Job'].isin(bill_sum_f['Shipment Job'])]
fdata      = data[data['Shipment Job'].isin(bill_sum_f['Shipment Job'])]
fship      = shipment_df[shipment_df['Shipment Job'].isin(bill_sum_f['Shipment Job'])]
supp_f     = build_supplier_summary(fdata, fship) if not fdata.empty else supp_sum

# ── KPIs ──────────────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Shipments",   f"{bill_sum_f['Shipment Job'].nunique():,}")
k2.metric("Containers",  f"{int(ship_sum_f['Container Count'].fillna(0).sum()):,}")
k3.metric("AUD Charges", f"${bill_det_f[bill_det_f['Currency']=='AUD']['Total'].sum():,.2f}")
k4.metric("USD Charges", f"${bill_det_f[bill_det_f['Currency']=='USD']['Total'].sum():,.2f}")
k5.metric("Total (AUD)", f"${bill_det_f['Local Total'].sum():,.2f}")

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋  Shipment Summary","💲  Billing Summary","📄  Billing Detail","📊  Analysis","⬇   Download",
])

with tab1:
    st.markdown('<div class="section-title">Shipment Summary</div>', unsafe_allow_html=True)
    st.dataframe(ship_sum_f, use_container_width=True, hide_index=True)
    st.caption(f"{len(ship_sum_f)} shipments · {int(ship_sum_f['Container Count'].fillna(0).sum())} containers")

with tab2:
    st.markdown('<div class="section-title">Billing Summary — Per Job</div>', unsafe_allow_html=True)
    st.dataframe(bill_sum_f, use_container_width=True, hide_index=True)
    a = bill_sum_f['AUD'].sum() if 'AUD' in bill_sum_f.columns else 0
    u = bill_sum_f['USD'].sum() if 'USD' in bill_sum_f.columns else 0
    l = bill_sum_f['Local Total (AUD)'].sum() if 'Local Total (AUD)' in bill_sum_f.columns else 0
    st.caption(f"AUD ${a:,.2f}  ·  USD ${u:,.2f}  ·  Local Total AUD ${l:,.2f}")

with tab3:
    st.markdown('<div class="section-title">Billing Detail — All Charge Lines</div>', unsafe_allow_html=True)
    st.dataframe(bill_det_f, use_container_width=True, hide_index=True)
    st.caption(f"{len(bill_det_f):,} charge lines")

with tab4:
    if bill_sum_f.empty:
        st.warning("No data matches the current filters.")
    else:
        r1l, r1r = st.columns(2)

        # Chart 1 — Spend by Supplier
        with r1l:
            st.markdown('<div class="section-title">Total Billed by Supplier (AUD)</div>', unsafe_allow_html=True)
            fig1 = px.bar(
                supp_f.sort_values('Local_Total_AUD'),
                x='Local_Total_AUD', y='Supplier Name', orientation='h',
                color_discrete_sequence=[BLUE],
                labels={'Local_Total_AUD':'Local Total (AUD)','Supplier Name':''},
                text='Local_Total_AUD',
            )
            fig1.update_traces(texttemplate='$%{text:,.0f}', textposition='outside',
                               marker_line_width=0, textfont_color='#000000', marker_color=BLUE)
            fig1.update_layout(**PLOTLY_LAYOUT, xaxis_title='', yaxis_title='')
            st.plotly_chart(fig1, use_container_width=True)

        # Chart 2 — Monthly stacked bars AUD vs USD (go.Bar for reliable stacking)
        with r1r:
            st.markdown('<div class="section-title">Monthly Charges — AUD vs USD Converted (Local AUD)</div>', unsafe_allow_html=True)
            bd_m, month_order = add_month_col(bill_det_f, ship_sum_f)
            if month_order:
                monthly = bd_m.groupby(['Month','Currency'])['Local Total'].sum().reset_index()
                aud_vals = [monthly[(monthly['Month']==m)&(monthly['Currency']=='AUD')]['Local Total'].sum() for m in month_order]
                usd_vals = [monthly[(monthly['Month']==m)&(monthly['Currency']=='USD')]['Local Total'].sum() for m in month_order]
                fig2 = go.Figure()
                fig2.add_trace(go.Bar(name='AUD', x=month_order, y=aud_vals, marker_color=ORANGE, marker_line_width=0))
                fig2.add_trace(go.Bar(name='USD (converted)', x=month_order, y=usd_vals, marker_color=BLUE, marker_line_width=0))
                fig2.update_layout(**PLOTLY_LAYOUT, barmode='stack', xaxis_title='', yaxis_title='Amount (AUD)')
                st.plotly_chart(fig2, use_container_width=True)
                st.caption("USD charges converted to AUD using the exchange rate on each invoice line.")

        r2l, r2r = st.columns(2)

        # Chart 3 — Top 10 charges stacked by currency
        with r2l:
            st.markdown('<div class="section-title">Top 10 Charge Types (Local AUD)</div>', unsafe_allow_html=True)
            top10 = bill_det_f.groupby('Description')['Local Total'].sum().sort_values(ascending=False).head(10).index.tolist()
            charge_stack = (bill_det_f[bill_det_f['Description'].isin(top10)]
                           .groupby(['Description','Currency'])['Local Total'].sum().reset_index())
            charge_stack['rank'] = charge_stack['Description'].map({d:i for i,d in enumerate(reversed(top10))})
            charge_stack = charge_stack.sort_values('rank')
            fig3 = px.bar(
                charge_stack, x='Local Total', y='Description', color='Currency',
                orientation='h', color_discrete_map={'AUD':ORANGE,'USD':BLUE},
                barmode='stack', labels={'Local Total':'Amount (AUD)','Description':''},
            )
            fig3.update_layout(**PLOTLY_LAYOUT, xaxis_title='', yaxis_title='')
            fig3.update_traces(marker_line_width=0)
            st.plotly_chart(fig3, use_container_width=True)
            st.caption("Blue = USD charges converted to AUD.")

        # Chart 4 — Incoterms donut
        with r2r:
            st.markdown('<div class="section-title">Shipments by Incoterms</div>', unsafe_allow_html=True)
            if 'Incoterms' in ship_sum_f.columns:
                inco_data = ship_sum_f.groupby('Incoterms')['Shipment Job'].nunique().reset_index()
                inco_data.columns = ['Incoterms','Shipments']
                fig4 = px.pie(inco_data, values='Shipments', names='Incoterms',
                              color_discrete_sequence=PALETTE, hole=0.45)
                fig4.update_layout(**PLOTLY_LAYOUT)
                fig4.update_traces(textfont_color='#ffffff', textfont_size=13)
                st.plotly_chart(fig4, use_container_width=True)

        # Chart 5 — Spend by Incoterm over time (line chart)
        st.markdown('<div class="section-title">Monthly Spend by Incoterm (Local AUD)</div>', unsafe_allow_html=True)
        bd_m2, month_order2 = add_month_col(bill_det_f, ship_sum_f)
        if month_order2 and 'Incoterms' in bd_m2.columns:
            inco_monthly = bd_m2.groupby(['Month','Incoterms'])['Local Total'].sum().reset_index()
            inco_monthly = inco_monthly[inco_monthly['Local Total'] > 0]
            inco_monthly['Month'] = pd.Categorical(inco_monthly['Month'], categories=month_order2, ordered=True)
            inco_monthly = inco_monthly.sort_values('Month')
            fig5 = px.line(
                inco_monthly, x='Month', y='Local Total', color='Incoterms',
                color_discrete_sequence=PALETTE, markers=True,
                labels={'Local Total':'Amount (AUD)','Month':''},
            )
            fig5.update_layout(**PLOTLY_LAYOUT, xaxis_title='', yaxis_title='Amount (AUD)')
            fig5.update_traces(line_width=2.5, marker_size=8)
            st.plotly_chart(fig5, use_container_width=True)
            st.caption("Each line tracks total AUD-equivalent spend per Incoterm by ETD month.")
        else:
            st.info("Not enough data for Incoterm trend.")

        st.markdown('<div class="section-title">Supplier Summary</div>', unsafe_allow_html=True)
        st.dataframe(supp_f, use_container_width=True, hide_index=True)

with tab5:
    st.markdown('<div class="section-title">Download Consolidated Report</div>', unsafe_allow_html=True)
    active = any([f_supplier, f_origin, f_dest, f_inco, f_mode, f_month])
    scope  = (f"**{len(bill_sum_f)} of {len(bill_sum)} shipments** match current filters."
              if active else f"**All {len(bill_sum)} billed shipments** included.")
    st.info(scope)
    buf = create_report(ship_sum_f, bill_sum_f, bill_det_f, supp_f)
    st.download_button(label="⬇  Download Excel Report", data=buf,
        file_name="TGG_Billing_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown("""
**Sheets included:**
- **Shipment Summary** — one row per shipment
- **Billing Summary** — per-job AUD / USD / Local Total (AUD)
- **Billing Detail** — every charge line
- **Supplier Summary** — shipment count and spend per supplier
    """)

with st.sidebar:
    buf2 = create_report(ship_sum_f, bill_sum_f, bill_det_f, supp_f)
    st.download_button(label="⬇  Download Report", data=buf2,
        file_name="TGG_Billing_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption(f"{len(bill_sum_f)} shipments in current export")
