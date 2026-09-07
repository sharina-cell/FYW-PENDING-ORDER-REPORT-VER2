"""
FYW Pending Order Report — Streamlit App
Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from report_engine import (
    load_fyw_csv,
    load_shopee_files,
    build_report_rows,
    generate_excel,
    merge_manual_orders,
)
from tc_reconciliation import (
    run_full_reconciliation,
    apply_manual_overrides,
    manual_confirmed_to_df,
    extract_order_detail,
    build_mp_detail_map,
)

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FYW Pending Order Report",
    page_icon="📦",
    layout="wide",
)

st.title("📦 FYW Pending Order Report")
st.markdown("Upload your daily FYW dashboard export and marketplace files to generate reports.")

# Persist manually-confirmed "pushed to TC" orders across reruns/reuploads.
# Structure: { marketplace_label: [ {eOrder Number, Invoice Number, ...}, ... ] }
if 'manual_confirmed' not in st.session_state:
    st.session_state.manual_confirmed = {}

# Holds the most recent auto-extraction preview, keyed by "label||order_id",
# so the user can review before confirming.
if 'extracted_preview' not in st.session_state:
    st.session_state.extracted_preview = {}

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR — Shared file uploads
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("📁 Upload Files")

    fyw_csv = st.file_uploader(
        "FYW Dashboard CSV (ALL-DD-Mon-YYYY.csv) *",
        type=["csv"],
        help="Required — the main FYW order export (this is 'TC')"
    )

    st.markdown("---")
    st.markdown("**Shopee Seller Centre Files** _(per brand)_")
    melissa_file = st.file_uploader("Melissa Shopee Order (.xlsx)", type=["xlsx"], key="melissa_shopee")
    ipanema_file = st.file_uploader("Ipanema Shopee Order (.xlsx)", type=["xlsx"], key="ipanema_shopee")
    cspace_file  = st.file_uploader("CSpace Shopee Order (.xlsx)",  type=["xlsx"], key="cspace_shopee")

    st.markdown("---")
    st.markdown("**Other Marketplace Files** _(for TC reconciliation)_")
    st.markdown("**TikTok** _(per brand)_")
    tiktok_melissa_file = st.file_uploader("TikTok - Melissa Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="tiktok_melissa")
    tiktok_ipanema_file = st.file_uploader("TikTok - Ipanema Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="tiktok_ipanema")
    tiktok_cspace_file  = st.file_uploader("TikTok - CSpace Order Export (.xlsx/.csv)",  type=["xlsx", "csv"], key="tiktok_cspace")

    st.markdown("**Lazada** _(per brand)_")
    lazada_melissa_file = st.file_uploader("Lazada - Melissa Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="lazada_melissa")
    lazada_ipanema_file = st.file_uploader("Lazada - Ipanema Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="lazada_ipanema")
    lazada_cspace_file  = st.file_uploader("Lazada - CSpace Order Export (.xlsx/.csv)",  type=["xlsx", "csv"], key="lazada_cspace")

    zalora_file = st.file_uploader("Zalora Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="zalora_file")

    st.markdown("---")
    report_date = st.date_input("Report Date", value=datetime.today())

    with st.expander("🔄 Reupload TC order file (refresh reconciliation)"):
        st.caption(
            "If you've since re-exported the FYW dashboard CSV (e.g. after pushing more "
            "orders), upload it here to re-run the check. Orders already confirmed below "
            "stay confirmed either way."
        )
        tc_reupload = st.file_uploader("Updated FYW Dashboard CSV", type=["csv"], key="tc_reupload_csv")

if not fyw_csv:
    st.info("👈 Upload the FYW CSV file to get started.")
    st.stop()

# ════════════════════════════════════════════════════════════════════════════
# BUILD BASE REPORT
# ════════════════════════════════════════════════════════════════════════════
mp_files = {
    'Shopee - Melissa': melissa_file,
    'Shopee - Ipanema': ipanema_file,
    'Shopee - CSpace':  cspace_file,
    'TikTok - Melissa': tiktok_melissa_file,
    'TikTok - Ipanema': tiktok_ipanema_file,
    'TikTok - CSpace':  tiktok_cspace_file,
    'Lazada - Melissa': lazada_melissa_file,
    'Lazada - Ipanema': lazada_ipanema_file,
    'Lazada - CSpace':  lazada_cspace_file,
    'Zalora':           zalora_file,
}
uploaded_mp_files = {k: v for k, v in mp_files.items() if v is not None}

with st.spinner("Loading files..."):
    try:
        pending = load_fyw_csv(fyw_csv)
    except Exception as e:
        st.error(f"❌ Failed to load FYW CSV: {e}")
        st.stop()

    shopee_map = {'MELISSA': melissa_file, 'IPANEMA': ipanema_file, 'CSPACE': cspace_file}
    mp_sla_map, shopee_tracking = load_shopee_files(shopee_map)

    # Merge all uploaded TikTok brand files into one order_id -> detail map, so
    # Melissa/Ipanema/CSpace all get identical, real-data-driven MP SLA treatment.
    tiktok_detail_map = {}
    for tiktok_label in ('TikTok - Melissa', 'TikTok - Ipanema', 'TikTok - CSpace'):
        tiktok_detail_map.update(build_mp_detail_map(tiktok_label, mp_files.get(tiktok_label)))

    base_df_report = build_report_rows(pending, mp_sla_map, shopee_tracking, mp_detail_map=tiktok_detail_map)

# ════════════════════════════════════════════════════════════════════════════
# TC RECONCILIATION — runs before the report is finalized so NOT PUSHED
# orders are already reflected in the same preview/report below.
# ════════════════════════════════════════════════════════════════════════════
results = None
if uploaded_mp_files:
    with st.spinner("Cross-checking marketplace orders against TC..."):
        try:
            active_tc_file = tc_reupload if tc_reupload is not None else fyw_csv
            active_tc_file.seek(0)
            fyw_raw_df = pd.read_csv(active_tc_file)
            base_results = run_full_reconciliation(fyw_raw_df, uploaded_mp_files)
            results = apply_manual_overrides(base_results, st.session_state.manual_confirmed)
        except Exception as e:
            st.error(f"❌ TC reconciliation failed: {e}")
            results = None

    if results is not None and tc_reupload is not None:
        st.success("✅ Reconciliation refreshed using the reuploaded TC file.")

# Fold in any orders manually confirmed as pushed so FILTERED DATA / PIVOT /
# the download all reflect them immediately.
df_report = merge_manual_orders(base_df_report, st.session_state.manual_confirmed)
manual_row_count = len(df_report) - len(base_df_report)
total_not_pushed = sum(r.get('missing_count', 0) for r in results.values()) if results else 0

# ════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ════════════════════════════════════════════════════════════════════════════
st.subheader("📊 Summary")
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Pending Items", len(df_report))
col2.metric("Unique Orders",       df_report['Order ID'].nunique())
col3.metric("Total Value (MYR)",   f"{df_report['Sold Price (MYR)'].sum():,.2f}")
col4.metric("Active Channels",     df_report['Channel'].nunique())

if results is not None:
    n1, n2, n3 = st.columns(3)
    n1.metric("Manually Confirmed (added)", manual_row_count)
    n2.metric(
        "Not Pushed to TC",
        total_not_pushed,
        delta="✅ All synced" if total_not_pushed == 0 else f"{total_not_pushed} pending",
        delta_color="normal" if total_not_pushed == 0 else "inverse",
    )
    n3.metric("Marketplaces Checked", len(results))
else:
    st.caption(
        "👉 Upload marketplace files (Shopee/TikTok/Lazada/Zalora) in the sidebar to also check "
        "which orders haven't been pushed to TC yet."
    )

# ════════════════════════════════════════════════════════════════════════════
# PIVOT
# ════════════════════════════════════════════════════════════════════════════
st.subheader("🗂️ Pivot: Orders by Nickname × FYW SLA")
pivot_raw = df_report.pivot_table(
    index='Nickname', columns='FYW SLA', values='Order ID',
    aggfunc='count', fill_value=0
)
pivot_raw['Grand Total'] = pivot_raw.sum(axis=1)
total_row = pivot_raw.sum(axis=0)
total_row.name = 'Grand Total'
pivot_display = pd.concat([pivot_raw, total_row.to_frame().T]).astype(int)
pivot_display = pivot_display.replace(0, '-')


def style_pivot(df):
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    today = datetime.today().date()
    yest  = (datetime.today() - pd.Timedelta(days=1)).date()
    for col in df.columns:
        if col == 'Grand Total':
            styles[col] = 'background-color: #DDEEFF; font-weight: bold'
        else:
            try:
                col_date = datetime.strptime(col, '%d/%m/%Y').date()
                if col_date < yest:
                    styles[col] = 'background-color: #FFCCCC; color: #9C0006; font-weight: bold'
                elif col_date == yest:
                    styles[col] = 'background-color: #FCE4D6; color: #833C00; font-weight: bold'
                elif col_date == today:
                    styles[col] = 'background-color: #C6EFCE; color: #006100; font-weight: bold'
            except Exception:
                pass
    return styles


st.dataframe(pivot_display.style.apply(style_pivot, axis=None), use_container_width=True)

col_a, col_b, col_c, col_d = st.columns(4)
col_a.markdown("🔴 **Past due**")
col_b.markdown("🟠 **Yesterday**")
col_c.markdown("🟢 **Today**")
col_d.markdown("🔵 **Grand Total**")

# ════════════════════════════════════════════════════════════════════════════
# BRAND SUMMARY
# ════════════════════════════════════════════════════════════════════════════
st.subheader("🏷️ Brand Summary")
brand_summary = df_report.groupby('Brand').agg(
    Orders=('Order ID', 'nunique'),
    Items=('Qty', 'sum'),
    Value=('Sold Price (MYR)', 'sum'),
    Channels=('Channel', lambda x: ', '.join(sorted(set(x))))
).reset_index()
brand_summary['Value'] = brand_summary['Value'].apply(lambda x: f"MYR {x:,.2f}")
st.dataframe(brand_summary, use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# ALL PENDING ORDERS
# ════════════════════════════════════════════════════════════════════════════
st.subheader("📋 All Pending Orders")
with st.expander("🔍 Filter options"):
    c1, c2, c3 = st.columns(3)
    brand_filter   = c1.multiselect("Brand",   options=sorted(df_report['Brand'].unique()),   default=list(df_report['Brand'].unique()))
    channel_filter = c2.multiselect("Channel", options=sorted(df_report['Channel'].unique()), default=list(df_report['Channel'].unique()))
    status_filter  = c3.multiselect("Status",  options=sorted(df_report['FYW Status'].unique()), default=list(df_report['FYW Status'].unique()))

filtered = df_report[
    df_report['Brand'].isin(brand_filter) &
    df_report['Channel'].isin(channel_filter) &
    df_report['FYW Status'].isin(status_filter)
]
display_df = filtered[['Order ID', 'FYW Status', 'Payment Status', 'Order Date', 'Nickname', 'MP SLA']].rename(columns={
    'Order ID':   'Order Number',
    'FYW Status': 'Order Status',
    'Order Date': 'Ordered Date',
})
st.dataframe(display_df, use_container_width=True, hide_index=True)
st.caption(f"Showing {len(filtered)} of {len(df_report)} orders")

# ════════════════════════════════════════════════════════════════════════════
# NOT PUSHED TO TC — click an order, auto-extract details, confirm
# ════════════════════════════════════════════════════════════════════════════
if results is not None:
    st.subheader("🔴 Not Pushed to TC")

    if total_not_pushed == 0:
        st.success("🎉 All marketplace orders have been successfully pushed to TC!")
    else:
        not_pushed_rows = []
        for label, res in results.items():
            for oid in res.get('missing_ids', []):
                not_pushed_rows.append({'Marketplace': label, 'Order ID / Order Number': oid})
        not_pushed_df = pd.DataFrame(not_pushed_rows)
        st.dataframe(not_pushed_df, use_container_width=True, hide_index=True)

        st.markdown("**Already pushed to TC?** Select an order below to pull its details straight from the marketplace file.")

        order_options = ["— Select an order —"] + [f"{r['Marketplace']} | {r['Order ID / Order Number']}" for r in not_pushed_rows]
        selected = st.selectbox("Order to update", options=order_options, key="not_pushed_selector")

        if selected != order_options[0]:
            sel_label, sel_oid = [s.strip() for s in selected.split("|", 1)]
            preview_key = f"{sel_label}||{sel_oid}"

            with st.expander(f"🔄 Update — {sel_label} | {sel_oid}", expanded=True):
                # Auto-extract the first time this order is opened.
                if preview_key not in st.session_state.extracted_preview:
                    detail = extract_order_detail(sel_label, mp_files.get(sel_label), sel_oid,
                                                   id_col_name=results.get(sel_label, {}).get('column_used'))
                    st.session_state.extracted_preview[preview_key] = detail or {}
                    # Pre-seed the widget state directly so the fields actually show the
                    # extracted values (passing value= alone is ignored once a widget key
                    # already exists in session_state on a rerun).
                    seed = detail or {}
                    st.session_state[f"inv_{preview_key}"] = seed.get('Invoice Number', '')
                    st.session_state[f"ois_{preview_key}"] = seed.get('Order Item Status', '')
                    st.session_state[f"pay_{preview_key}"] = seed.get('Payment Status', '')
                    st.session_state[f"odt_{preview_key}"] = seed.get('Ordered Date', '')
                    st.session_state[f"sla_{preview_key}"] = seed.get('MP SLA', '')

                if not st.session_state.extracted_preview.get(preview_key):
                    st.warning(
                        "Couldn't auto-extract this order (unsupported marketplace or order not "
                        "found in the file). Fill it in manually below."
                    )
                else:
                    st.caption("✅ Auto-filled from the marketplace file — adjust anything below if needed.")

                invoice_number = st.text_input("Invoice Number (not in marketplace file — enter manually)",
                                                key=f"inv_{preview_key}")
                p1, p2 = st.columns(2)
                order_item_status = p1.text_input("Order Item Status", key=f"ois_{preview_key}")
                payment_status    = p2.text_input("Payment Status",    key=f"pay_{preview_key}")
                p3, p4 = st.columns(2)
                ordered_date = p3.text_input("Ordered Date", key=f"odt_{preview_key}")
                mp_sla       = p4.text_input("MP SLA",       key=f"sla_{preview_key}")

                if st.button(f"✅ Confirm & Update Report", key=f"confirm_{preview_key}", type="primary"):
                    entry = {
                        'eOrder Number':     sel_oid,
                        'Invoice Number':    invoice_number,
                        'Payment Status':    payment_status,
                        'Order Item Status': order_item_status,
                        'Ordered Date':      ordered_date,
                        'MP SLA':            mp_sla,
                    }
                    existing = st.session_state.manual_confirmed.setdefault(sel_label, [])
                    existing_ids = {r['eOrder Number'] for r in existing}
                    if sel_oid in existing_ids:
                        st.warning("This order has already been confirmed.")
                    else:
                        existing.append(entry)
                        st.session_state.extracted_preview.pop(preview_key, None)
                        st.success(f"Updated report — {sel_oid} moved out of Not Pushed and into Filtered Data / Pivot.")
                        st.rerun()

        if st.session_state.manual_confirmed and any(st.session_state.manual_confirmed.values()):
            with st.expander("✅ Manually Confirmed Orders (applied)", expanded=False):
                st.dataframe(
                    manual_confirmed_to_df(st.session_state.manual_confirmed),
                    use_container_width=True,
                    hide_index=True,
                )
                if st.button("🗑️ Clear all manual confirmations"):
                    st.session_state.manual_confirmed = {}
                    st.session_state.extracted_preview = {}
                    st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# DOWNLOAD
# ════════════════════════════════════════════════════════════════════════════
st.subheader("⬇️ Download Excel Report")
with st.spinner("Generating Excel..."):
    excel_buffer = generate_excel(
        df_report,
        report_date=datetime.combine(report_date, datetime.min.time()),
        not_pushed_results=results,
    )

filename = f"FYW_Pending_Orders_{report_date.strftime('%d%b%Y')}.xlsx"
st.download_button(
    label="📥 Download Excel Report",
    data=excel_buffer,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    key="download_pending_report",
)
st.success(f"✅ **{filename}** — {len(df_report)} items | MYR {df_report['Sold Price (MYR)'].sum():,.2f}")
