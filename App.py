import os
import io
import datetime
import pandas as pd
import numpy as np
import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(
    page_title="Quick-Commerce Master Dashboard & Summary",
    page_icon="🔵",
    layout="wide"
)

st.markdown("""
<style>
    :root {
        --dash-blue: #4A90C2;
        --dash-blue-dark: #285B7A;
        --dash-border: #C8DCEB;
        --dash-text: #29465B;
        --lavender-blue-mix: linear-gradient(135deg, #F3EBFB 0%, #EEF6FC 100%);
        --lavender-border: #D5C6E8;
    }
    .stApp {
        background: linear-gradient(180deg, #F6FAFD 0%, #FFFFFF 38%);
        color: var(--dash-text);
    }
    [data-testid="stHeader"] {
        background: rgba(255,255,255,0.92);
    }
    h1, h2, h3 {
        color: var(--dash-blue-dark) !important;
        font-weight: 700 !important;
        text-align: center !important;
    }
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #FFFFFF 0%, #EEF6FF 100%);
        border: 1px solid var(--dash-border);
        border-left: 5px solid #6AA6CF;
        border-radius: 14px;
        padding: 14px 18px;
        box-shadow: 0 3px 12px rgba(74, 144, 194, 0.10);
        text-align: center !important;
    }
    [data-testid="stMetricLabel"] {
        color: #52708D !important;
        font-weight: 600 !important;
        justify-content: center;
        display: flex;
    }
    [data-testid="stMetricValue"] {
        color: var(--dash-blue-dark) !important;
        font-weight: 750 !important;
        justify-content: center;
        display: flex;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background: var(--lavender-blue-mix);
        border: 1px solid var(--lavender-border);
        border-radius: 16px;
        padding: 6px;
        justify-content: center;
    }
    .stTabs [data-baseweb="tab"] {
        height: 42px;
        border-radius: 12px;
        color: #4A3B6C;
        font-weight: 600;
        background: #FFFFFF;
        border: 1px solid #E2D6F3;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #6AA6CF, #4A90C2) !important;
        color: #FFFFFF !important;
        box-shadow: 0 3px 9px rgba(74, 144, 194, 0.25);
        border: 0 !important;
    }
    [data-testid="stFileUploader"] {
        background: #FFFFFF;
        border: 1px solid var(--dash-border);
        border-radius: 14px;
        padding: 8px;
    }
    [data-testid="stDataFrame"] {
        border: 1px solid var(--dash-border);
        border-radius: 12px;
        overflow: hidden;
    }
    [data-testid="stDataFrame"] th {
        background: #E5F1F9 !important;
        color: #285B7A !important;
        font-weight: 700 !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>Quick-Commerce Master Dashboard & Summary</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload Ad Campaign and Sales reports to generate consolidated analytics and master dumps.</p>", unsafe_allow_html=True)

CITY_REGION_MASTER = {
    "Ahmedabad": "T-10", "Bengaluru": "T-10", "Chennai": "T-10", "Delhi": "T-10",
    "Gurgaon": "T-10", "Hyderabad": "T-10", "Kolkata": "T-10", "Lucknow": "T-10",
    "Mumbai": "T-10", "Pune": "T-10"
}

def get_region_for_city(city_val):
    if pd.isna(city_val): return "ROI"
    return CITY_REGION_MASTER.get(str(city_val).strip().title(), "ROI")

def format_duration_month(dt):
    if pd.isna(dt): return np.nan
    month_num = dt.month
    short_mon = dt.strftime('%b')
    yr = dt.strftime('%y')
    return f"{month_num}. {short_mon}'{yr}"

def format_duration_week(dt):
    if pd.isna(dt): return np.nan
    iso_yr, iso_wk, iso_day = dt.isocalendar()
    start_dt = dt - datetime.timedelta(days=iso_day - 1)
    end_dt = start_dt + datetime.timedelta(days=6)
    return f"Week {iso_wk:02d} ({start_dt.strftime('%b %d')} to {end_dt.strftime('%b %d')})"

def style_and_export_pivot(pivot_df, sheet_name="Sheet1"):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    
    top_header_fill = PatternFill(start_color="4A90C2", end_color="4A90C2", fill_type="solid")
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10, color="000000")

    for col_idx, col_name in enumerate(pivot_df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=str(col_name))
        cell.fill = top_header_fill
        cell.font = top_header_font

    for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=2):
        for c_idx, (col_name, val) in enumerate(zip(pivot_df.columns, row_data), start=1):
            val_cell = ws.cell(row=r_idx, column=c_idx, value=val)
            val_cell.font = data_font

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max(max_len + 3, 10)
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

st.markdown("### 📤 Upload Ad Campaign Reports (Up to 5)")
col1, col2, col3, col4, col5 = st.columns(5)
uploaded_files = []
file_types = ["csv", "xlsx", "xls", "xlsb", "xlsm"]
with col1:
    f1 = st.file_uploader("Upload Ad File 1", type=file_types, key="file1")
    if f1: uploaded_files.append(f1)
with col2:
    f2 = st.file_uploader("Upload Ad File 2", type=file_types, key="file2")
    if f2: uploaded_files.append(f2)
with col3:
    f3 = st.file_uploader("Upload Ad File 3", type=file_types, key="file3")
    if f3: uploaded_files.append(f3)
with col4:
    f4 = st.file_uploader("Upload Ad File 4", type=file_types, key="file4")
    if f4: uploaded_files.append(f4)
with col5:
    f5 = st.file_uploader("Upload Ad File 5", type=file_types, key="file5")
    if f5: uploaded_files.append(f5)

st.markdown("### 📥 Upload Sales Reports (Up to 5)")
scol1, scol2, scol3, scol4, scol5 = st.columns(5)
uploaded_sales_files = []
with scol1:
    s1 = st.file_uploader("Upload Sales File 1", type=file_types, key="sales1")
    if s1: uploaded_sales_files.append(s1)
with scol2:
    s2 = st.file_uploader("Upload Sales File 2", type=file_types, key="sales2")
    if s2: uploaded_sales_files.append(s2)
with scol3:
    s3 = st.file_uploader("Upload Sales File 3", type=file_types, key="sales3")
    if s3: uploaded_sales_files.append(s3)
with scol4:
    s4 = st.file_uploader("Upload Sales File 4", type=file_types, key="sales4")
    if s4: uploaded_sales_files.append(s4)
with scol5:
    s5 = st.file_uploader("Upload Sales File 5", type=file_types, key="sales5")
    if s5: uploaded_sales_files.append(s5)

master_sales_df = pd.DataFrame()
if uploaded_sales_files:
    sales_dfs_list = []
    for s_file in uploaded_sales_files:
        try:
            if s_file.name.lower().endswith('.csv'):
                sales_dfs_list.append(pd.read_csv(s_file, low_memory=False))
            else:
                s_xls = pd.ExcelFile(s_file)
                for s_sheet in s_xls.sheet_names:
                    sales_dfs_list.append(pd.read_excel(s_xls, sheet_name=s_sheet))
        except Exception as e:
            st.error(f"Error reading sales file {s_file.name}: {str(e)}")

    if sales_dfs_list:
        raw_sales_concat = pd.concat(sales_dfs_list, ignore_index=True)
        prod_col = next((c for c in ['PRODUCT_NAME', 'Product Name', 'Item Name', 'product_name'] if c in raw_sales_concat.columns), None)
        date_col = next((c for c in ['ORDER_DATE', 'Order Date', 'date', 'Date'] if c in raw_sales_concat.columns), None)
        qty_col = next((c for c in ['UNITS_SOLD', 'Quantity', 'Qty', 'qty_sold'] if c in raw_sales_concat.columns), None)
        mrp_col = next((c for c in ['GMV', 'BASE_MRP', 'MRP (Rs)', 'Selling Price (Rs)', 'mrp'] if c in raw_sales_concat.columns), None)
        city_col = next((c for c in ['CITY', 'Supply City', 'city_name', 'City'] if c in raw_sales_concat.columns), None)
        brand_col = next((c for c in ['Brand Name', 'BRAND', 'Franchise', 'Brand'] if c in raw_sales_concat.columns), None)
        cat_col = next((c for c in ['Business Category', 'L1_CATEGORY', 'category', 'Category'] if c in raw_sales_concat.columns), None)
        
        master_sales_df = raw_sales_concat.copy()
        master_sales_df['_product_name'] = master_sales_df[prod_col].astype(str) if prod_col else ""
        master_sales_df['_brand'] = master_sales_df[brand_col].astype(str) if brand_col else ""
        master_sales_df['_category'] = master_sales_df[cat_col].astype(str) if cat_col else ""
        
        if date_col:
            master_sales_df['_sales_dt'] = pd.to_datetime(master_sales_df[date_col], errors='coerce')
            master_sales_df['Duration (Month)'] = master_sales_df['_sales_dt'].apply(format_duration_month)
            master_sales_df['Duration (Week)'] = master_sales_df['_sales_dt'].apply(format_duration_week)
            master_sales_df['date'] = master_sales_df['_sales_dt'].dt.strftime('%Y/%m/%d')
        else:
            master_sales_df['Duration (Month)'] = np.nan
            master_sales_df['Duration (Week)'] = np.nan
            master_sales_df['date'] = np.nan
            
        master_sales_df['_qty'] = pd.to_numeric(master_sales_df[qty_col].astype(str).str.replace(',', ''), errors='coerce').fillna(1.0) if qty_col else 1.0
        master_sales_df['_gross_amt'] = pd.to_numeric(master_sales_df[mrp_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0) if mrp_col else 0.0
        master_sales_df['city_name'] = master_sales_df[city_col] if city_col else "Unknown"
        master_sales_df['Region'] = master_sales_df['city_name'].apply(get_region_for_city)

if uploaded_files:
    consolidated_dfs, consolidated_raw_tabs = [], {}
    for uploaded_file in uploaded_files:
        try:
            if uploaded_file.name.lower().endswith('.csv'):
                df = pd.read_csv(uploaded_file, low_memory=False)
                consolidated_raw_tabs.setdefault("Sheet1", []).append(df.copy())
                consolidated_dfs.append(df)
            else:
                xls = pd.ExcelFile(uploaded_file)
                campaign_type_map = {}
                if 'MTD Claimables' in xls.sheet_names:
                    df_claim = pd.read_excel(xls, sheet_name='MTD Claimables')
                    if 'Campaign ID' in df_claim.columns and 'Campaign Type' in df_claim.columns:
                        campaign_type_map = df_claim.set_index('Campaign ID')['Campaign Type'].to_dict()

                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    if sheet_name.lower() not in ['summary', 'consolidated_master', 'mtd claimables']:
                        if 'Campaign Type' not in df.columns and 'Campaign ID' in df.columns and campaign_type_map:
                            df['Campaign Type'] = df['Campaign ID'].map(campaign_type_map)
                        consolidated_raw_tabs.setdefault(sheet_name, []).append(df.copy())
                        consolidated_dfs.append(df)
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")

    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True, sort=False)

        def get_numeric(cols):
            for c in cols:
                if c in final_df.columns: return pd.to_numeric(final_df[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
            return pd.Series(0.0, index=final_df.index)

        date_col_ad = next((c for c in ['date_ist', 'Date', 'date'] if c in final_df.columns), None)
        final_df['_date_dt'] = pd.to_datetime(final_df[date_col_ad], errors='coerce') if date_col_ad else pd.NaT
        final_df['date_ist'] = final_df['_date_dt'].dt.strftime('%Y/%m/%d')
        final_df['Duration (Month)'] = final_df['_date_dt'].apply(format_duration_month)
        final_df['Duration (Week)'] = final_df['_date_dt'].apply(format_duration_week)

        final_df['_impressions'] = get_numeric(['Impressions'])
        final_df['_atc'] = get_numeric(['Direct ATC']) + get_numeric(['Indirect ATC'])
        final_df['_orders'] = get_numeric(['Direct Quantities Sold', 'Direct Orders']) + get_numeric(['Indirect Quantities Sold', 'Indirect Orders'])
        final_df['_sales'] = get_numeric(['Direct Sales', 'Claimables'])
        final_df['_budget_consumed'] = get_numeric(['served_budget_consumed', 'Estimated Budget Consumed', 'Budget Consumed', 'Spend'])

        st.markdown("### 🔍 Global Month Filter")
        avail_months = sorted([m for m in final_df['Duration (Month)'].dropna().unique() if str(m) != 'nan'])
        sel_month = st.selectbox("Select Month:", ["All Months"] + avail_months)
        
        filt_df = final_df if sel_month == "All Months" else final_df[final_df['Duration (Month)'] == sel_month]

        tot_ad_sales = round(filt_df['_sales'].sum(), 2)
        tot_spends = round(filt_df['_budget_consumed'].sum(), 2)
        ad_roas = round(tot_ad_sales / tot_spends, 2) if tot_spends > 0 else 0.0

        if not master_sales_df.empty:
            filt_sales_df = master_sales_df if sel_month == "All Months" else master_sales_df[master_sales_df['Duration (Month)'] == sel_month]
            tot_overall_sales = round(filt_sales_df['_gross_amt'].sum(), 2)
            troas = round(tot_overall_sales / tot_spends, 2) if tot_spends > 0 else 0.0
            tacos = round((tot_spends / tot_overall_sales) * 100, 2) if tot_overall_sales > 0 else 0.0
        else:
            tot_overall_sales = 0.0
            troas = 0.0
            tacos = 0.0

        st.markdown("### 📈 Overall Campaign Performance Metrics")
        row1_c1, row1_c2, row1_c3 = st.columns(3)
        with row1_c1: st.metric("Ad Sales", f"₹{int(round(tot_ad_sales)):,}")
        with row1_c2: st.metric("Total Sales", f"₹{int(round(tot_overall_sales)):,}")
        with row1_c3: st.metric("Ad Spends", f"₹{int(round(tot_spends)):,}")

        row2_c1, row2_c2, row2_c3 = st.columns(3)
        with row2_c1: st.metric("Ad ROAS", f"{ad_roas:.2f}x")
        with row2_c2: st.metric("TROAS", f"{troas:.2f}x")
        with row2_c3: st.metric("TACOS", f"{tacos:.2f}%")

        st.divider()
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📅 Daily Overview", "🏙️ City Overview", "🗓️ Week Overview", "📑 ADs Dump Preview", "🛒 Overall Dump Preview"])
        
        with tab1:
            st.markdown("#### Day-Wise Performance Breakdown")
            daily_pivot = filt_df.groupby('date_ist').agg(
                Impressions=('_impressions', 'sum'),
                Spends=('_budget_consumed', 'sum'),
                Ad_Sales=('_sales', 'sum')
            ).reset_index()
            daily_pivot['Ad ROAS'] = (daily_pivot['Ad_Sales'] / daily_pivot['Spends']).round(2).fillna(0.0)
            st.dataframe(daily_pivot, use_container_width=True)

        with tab2:
            st.markdown("#### City & Region Performance Breakdown")
            city_col = next((c for c in ['City', 'Supply City', 'city_name'] if c in filt_df.columns), None)
            if city_col:
                city_pivot = filt_df.groupby(city_col).agg(
                    Impressions=('_impressions', 'sum'),
                    Spends=('_budget_consumed', 'sum'),
                    Ad_Sales=('_sales', 'sum')
                ).reset_index()
                city_pivot['Region'] = city_pivot[city_col].apply(get_region_for_city)
                city_pivot['Ad ROAS'] = (city_pivot['Ad_Sales'] / city_pivot['Spends']).round(2).fillna(0.0)
                st.dataframe(city_pivot, use_container_width=True)
            else:
                st.info("City column not found in uploaded ad report.")

        with tab3:
            st.markdown("#### Weekly Performance Breakdown")
            week_pivot = filt_df.groupby('Duration (Week)').agg(
                Impressions=('_impressions', 'sum'),
                Spends=('_budget_consumed', 'sum'),
                Ad_Sales=('_sales', 'sum')
            ).reset_index()
            week_pivot['Ad ROAS'] = (week_pivot['Ad_Sales'] / week_pivot['Spends']).round(2).fillna(0.0)
            st.dataframe(week_pivot, use_container_width=True)

        def safe_get(df_src, possible):
            for p in possible:
                if p in df_src.columns: return df_src[p]
            return np.nan

        # Prepare ADs Dump DataFrame
        final_ads_export = pd.DataFrame({
            'Duration (Month)': filt_df['Duration (Month)'],
            'Duration (Week)': filt_df['Duration (Week)'],
            'Manufacturer': safe_get(filt_df, ['Manufacturer', 'Manufacturer Name']),
            'date_ist': filt_df['date_ist'],
            'Campaign ID': safe_get(filt_df, ['Campaign ID', 'CampaignId']),
            'Campaign Name': safe_get(filt_df, ['Campaign Name', 'Campaign']),
            'City': safe_get(filt_df, ['City', 'Supply City', 'city_name']),
            'Ad type': safe_get(filt_df, ['Ad type', 'Ad Type', 'Campaign Type']),
            'Targeting Type': safe_get(filt_df, ['Targeting Type', 'Targeting']),
            'Franchise': safe_get(filt_df, ['Franchise', 'Brand', 'BRAND']),
            'Category': safe_get(filt_df, ['Category', 'category', 'L1_CATEGORY']),
            'Sub Category': safe_get(filt_df, ['Sub Category', 'Subcategory', 'L2_CATEGORY']),
            'Keyword': safe_get(filt_df, ['Keyword', 'Search Term']),
            'Most Viewed Position': safe_get(filt_df, ['Most Viewed Position', 'Average Position']),
            'Match Type': safe_get(filt_df, ['Match Type']),
            'Pacing Type': safe_get(filt_df, ['Pacing Type']),
            'CPM': safe_get(filt_df, ['CPM']),
            'Total Budget': safe_get(filt_df, ['Total Budget', 'Budget']),
            'Impressions': filt_df['_impressions'],
            'Direct ATC': safe_get(filt_df, ['Direct ATC']),
            'Indirect ATC': safe_get(filt_df, ['Indirect ATC']),
            'New Users': safe_get(filt_df, ['New Users']),
            'Direct Sales': safe_get(filt_df, ['Direct Sales']),
            'Indirect Sales': safe_get(filt_df, ['Indirect Sales']),
            'Direct Quantities Sold': safe_get(filt_df, ['Direct Quantities Sold', 'Direct Orders']),
            'Indirect Quantities Sold': safe_get(filt_df, ['Indirect Quantities Sold', 'Indirect Orders']),
            'Estimated Budget Consumed': filt_df['_budget_consumed'],
            'Direct RoAS': safe_get(filt_df, ['Direct RoAS', 'Direct ROAS']),
            'Total RoAS': safe_get(filt_df, ['Total RoAS', 'Total ROAS', 'ROAS']),
            'latest on hold timestamp': safe_get(filt_df, ['latest on hold timestamp'])
        })

        with tab4:
            st.markdown("#### Preview: ADs Dump Output Table")
            st.dataframe(final_ads_export.head(100), use_container_width=True)

        # Prepare Overall Dump DataFrame
        if not master_sales_df.empty:
            filt_s = master_sales_df if sel_month == "All Months" else master_sales_df[master_sales_df['Duration (Month)'] == sel_month]
            final_overall_export = pd.DataFrame({
                'item_id': safe_get(filt_s, ['ITEM_CODE', 'Item Id', 'item_id', 'SKU ID']),
                'item_name': filt_s['_product_name'],
                'Franchise': filt_s['_brand'],
                'manufacturer_id': safe_get(filt_s, ['manufacturer_id', 'Manufacturer ID']),
                'manufacturer_name': safe_get(filt_s, ['manufacturer_name', 'Manufacturer']),
                'city_id': safe_get(filt_s, ['STORE_ID', 'city_id', 'City ID', 'Order Id']),
                'city_name': filt_s['city_name'],
                'category': filt_s['_category'],
                'date': filt_s['date'],
                'Duration (Month)': filt_s['Duration (Month)'],
                'Duration (Week)': filt_s['Duration (Week)'],
                'qty_sold': filt_s['_qty'],
                'mrp': filt_s['_gross_amt'],
                'Region': filt_s['Region']
            })
        else:
            final_overall_export = pd.DataFrame(columns=['item_id', 'item_name', 'Franchise', 'manufacturer_id', 'manufacturer_name', 'city_id', 'city_name', 'category', 'date', 'Duration (Month)', 'Duration (Week)', 'qty_sold', 'mrp', 'Region'])

        with tab5:
            st.markdown("#### Preview: Overall Dump Output Table")
            st.dataframe(final_overall_export.head(100), use_container_width=True)

        st.divider()
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            final_ads_export.to_excel(writer, sheet_name='ADs Dump', index=False)
            final_overall_export.to_excel(writer, sheet_name='Overall Dump', index=False)
            for r_name, d_list in consolidated_raw_tabs.items():
                pd.concat(d_list, ignore_index=True, sort=False).to_excel(writer, sheet_name=r_name[:31], index=False)
        buf.seek(0)

        st.download_button(
            label="📥 Download Complete Excel Workbook (ADs Dump + Overall Dump + Raw Tabs)",
            data=buf,
            file_name="Master_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.info("Please upload valid ad campaign reports.")
else:
    st.info("Awaiting file uploads.")
