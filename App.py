import datetime
import io
import re
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Quick-Commerce Master Dashboard & Summary",
    page_icon="🔵",
    layout="wide",
)

st.markdown(
    """<style>
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
        text-align: center !important;
    }
    [data-testid="stDataFrame"] td {
        text-align: center !important;
        justify-content: center !important;
    }
    table {
        margin: auto;
        text-align: center !important;
    }
    th, td {
        text-align: center !important;
    }
</style>""",
    unsafe_allow_html=True,
)

st.markdown(
    "<h1 style='text-align: center;'>Quick-Commerce Master Dashboard &"
    " Summary</h1>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align: center;'>Upload Ad Campaign and Sales reports to"
    " generate consolidated analytics and master dumps.</p>",
    unsafe_allow_html=True,
)

# ==========================================
# TIER & CITY MASTER MAPPING (FROM PDF)
# ==========================================
TIER_1_CITIES = {
    "Ahmedabad",
    "Bengaluru",
    "Chennai",
    "Gurgaon",
    "Hyderabad",
    "Kolkata",
    "Mumbai",
    "New Delhi",
    "Pune",
}

TIER_2_CITIES = {
    "Agra",
    "Ambala",
    "Amravati",
    "Amritsar",
    "Ananthapur",
    "Jabalpur",
    "Jaipur",
    "Jalandhar",
    "Jalgaon",
    "Jamnagar",
    "Nizamabad",
    "Noida",
    "Panchkula",
    "Patna",
    "Prayagraj",
    "Asansol",
    "Jamshedpur",
    "Puducherry",
    "Belagavi",
    "Bhavnagar",
    "Bhiwandi",
    "Bhopal",
    "Bhubaneswar",
    "Calicut",
    "Jodhpur",
    "Raipur",
    "Kadapa",
    "Rajahmundry",
    "Kakinada",
    "Rajkot",
    "Kalyan Dombivli",
    "Ranchi",
    "Kanpur",
    "Rohtak",
    "Karnal",
    "Sagar",
    "Ch.Sambhajinagar (Aurangabad)",
    "Kochi",
    "Salem",
    "Chandigarh",
    "Coimbatore",
    "Cuttack",
    "Kolhapur",
    "Sangli",
    "Kota",
    "Satara",
    "Kurnool",
    "Solapur",
    "Davangere",
    "Latur",
    "Lucknow",
    "Sonipat",
    "Dhanbad",
    "Ludhiana",
    "Surat",
    "Durg-Bhilai",
    "Madurai",
    "Thane",
    "Gandhingar",
    "Gandhingar",
    "Malegaon",
    "Thrissur",
    "Mangaluru",
    "Tirupathi",
    "Goa",
    "Mira Bhyander",
    "Ghaziabad",
    "Greater Noida",
    "Guntur",
    "Guwahati",
    "Gwalior",
    "Hisar",
    "Howrah-Hooghly",
    "Huballi & Dharwad",
    "Indore",
    "Mohali",
    "Moradabad",
    "Mysuru",
    "Nagpur",
    "Nanded",
    "Nashik",
    "Navi Mumbai",
    "Nellore",
    "Trivandrum",
    "Udaipur",
    "Vadodara",
    "Varanasi",
    "Vijayawada",
    "Vishakhapatnam",
    "Warangal",
    "Yamuna",
}

TIER_3_CITIES = {
    "Ahmednagar",
    "Akola",
    "Aligarh",
    "Alwar",
    "Amalner",
    "Ambajogai",
    "Amreli",
    "Anand",
    "Baramati",
    "Bardoli",
    "Barshi",
    "Bathinda",
    "Becharaji",
    "Beed",
    "Begusarai",
    "Berhampur",
    "Bhadradri Kothagudam",
    "Bhandara",
    "Bharatpur",
    "Bharuch",
    "Indapur",
    "Islampur",
    "Jagityal",
    "Jalna",
    "Jaysingpur",
    "Jhansi",
    "Jharsuguda",
    "Junagadh",
    "Kachchh",
    "Kadi",
    "Kagal",
    "Kalol",
    "Kamareddy",
    "Karad",
    "Karim Nagar",
    "Karwar",
    "Khamgaon",
    "Khammam",
    "Kharar",
    "Kopargaon-Shirdi",
    "Palanpur",
    "Palghar Boisar",
    "Pandharpur",
    "Parbhani",
    "Patan",
    "Phaltan",
    "Pimpalgaon Baswant",
    "Puri",
    "Raigad",
    "Raigarh",
    "Ratnagiri",
    "Ropar",
    "Sabarkantha",
    "Sakri",
    "Sambalpur",
    "Sangrur",
    "Sawantwadi",
    "Shahada-Nandurbar",
    "Shahapur Murbad",
    "Shirpur",
    "Bhimavaram",
    "Machilipatnam",
    "Siddipet",
    "Bhusawal",
    "Mahabubnagar",
    "Silvassa",
    "Bidar",
    "Mahad",
    "Sindhudurg",
    "Bilaspur",
    "Malvan",
    "Sinnar",
    "Buldhana",
    "Mancherial",
    "Srikakulam",
    "Chalisgaon",
    "Mathura",
    "Tirunelveli",
    "Chandrapur",
    "Mehsana",
    "Tuni",
    "Chiplun",
    "Modasa",
    "Udgir",
    "Dahod",
    "Mundra",
    "Udupi",
    "Daund",
    "Muzaffarpur",
    "Ujjain",
    "Dharmapuri",
    "Nadiad",
    "Unjha",
    "Dhule",
    "Nalgonda",
    "Uran Dronagiri",
    "Eluru",
    "Nandyala",
    "Valsad",
    "Erode",
    "Narsipatnam",
    "Vapi",
    "Gadhinglaj",
    "Navsari",
    "Vijapur",
    "Gadwal",
    "Nilanga",
    "Vijayapura",
    "Gandhidham",
    "Nirmal",
    "Visnagar",
    "Gaya",
    "North Bengal",
    "Vizianagaram",
    "Godhra",
    "Omerga",
    "Vyara",
    "Gondia",
    "Ongole",
    "Wanaparthy",
    "Gurdaspur",
    "Osmanabad / Dharashiv",
    "Wani",
    "Hingoli",
    "Ozar",
    "Wardha",
    "Hoshangabad",
    "Pachora",
    "Washim",
    "Hosur",
    "Palakkad",
    "Yavatmal",
    "Ichalkaranji",
    "Palakollu",
    "Zirakpur",
}


def get_city_tier(city_val):
  if pd.isna(city_val):
    return "Tier 3"
  clean_city = str(city_val).strip()
  if clean_city in TIER_1_CITIES:
    return "Tier 1"
  elif clean_city in TIER_2_CITIES:
    return "Tier 2"
  elif clean_city in TIER_3_CITIES:
    return "Tier 3"
  return "Tier 3"


T10_STATES = [
    "Maharashtra",
    "Delhi",
    "Karnataka",
    "Tamil Nadu",
    "Telangana",
    "West Bengal",
    "Uttar Pradesh",
    "Gujarat",
]


def get_region_for_state(state_val):
  if pd.isna(state_val):
    return "ROI"
  st_clean = str(state_val).strip().title()
  return "T-10" if st_clean in T10_STATES else "ROI"


MONTH_ORDER_MAP = {
    "JANUARY": 1,
    "FEBRUARY": 2,
    "MARCH": 3,
    "APRIL": 4,
    "MAY": 5,
    "JUNE": 6,
    "JULY": 7,
    "AUGUST": 8,
    "SEPTEMBER": 9,
    "OCTOBER": 10,
    "NOVEMBER": 11,
    "DECEMBER": 12,
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}


def get_chronological_month_mapping(month_names_series):
  unique_months = [
      str(m).strip()
      for m in month_names_series.dropna().unique()
      if str(m) != "nan"
  ]
  sorted_months = sorted(
      unique_months, key=lambda x: MONTH_ORDER_MAP.get(x[:3].upper(), 99)
  )
  mapping = {}
  for idx, m in enumerate(sorted_months, start=1):
    mapping[m] = f"{idx}.{m}"
  return mapping


def format_duration_week(dt):
  if pd.isna(dt):
    return np.nan
  iso_yr, iso_wk, iso_day = dt.isocalendar()
  start_dt = dt - datetime.timedelta(days=iso_day - 1)
  end_dt = start_dt + datetime.timedelta(days=6)
  return (
      f"Week {iso_wk:02d} ({start_dt.strftime('%b %d')} to"
      f" {end_dt.strftime('%b %d')})"
  )


def extract_week_num(week_str):
  if pd.isna(week_str):
    return 0
  match = re.search(r"Week (\d+)", str(week_str))
  return int(match.group(1)) if match else 0


st.markdown("### 📤 Upload Ad Campaign Reports (Up to 5)")
col1, col2, col3, col4, col5 = st.columns(5)
uploaded_files = []
file_types = ["csv", "xlsx", "xls", "xlsb", "xlsm"]
with col1:
  f1 = st.file_uploader("Upload Ad File 1", type=file_types, key="file1")
  if f1:
    uploaded_files.append(f1)
with col2:
  f2 = st.file_uploader("Upload Ad File 2", type=file_types, key="file2")
  if f2:
    uploaded_files.append(f2)
with col3:
  f3 = st.file_uploader("Upload Ad File 3", type=file_types, key="file3")
  if f3:
    uploaded_files.append(f3)
with col4:
  f4 = st.file_uploader("Upload Ad File 4", type=file_types, key="file4")
  if f4:
    uploaded_files.append(f4)
with col5:
  f5 = st.file_uploader("Upload Ad File 5", type=file_types, key="file5")
  if f5:
    uploaded_files.append(f5)

st.markdown("### 📥 Upload Sales Reports (Up to 5)")
scol1, scol2, scol3, scol4, scol5 = st.columns(5)
uploaded_sales_files = []
with scol1:
  s1 = st.file_uploader("Upload Sales File 1", type=file_types, key="sales1")
  if s1:
    uploaded_sales_files.append(s1)
with scol2:
  s2 = st.file_uploader("Upload Sales File 2", type=file_types, key="sales2")
  if s2:
    uploaded_sales_files.append(s2)
with scol3:
  s3 = st.file_uploader("Upload Sales File 3", type=file_types, key="sales3")
  if s3:
    uploaded_sales_files.append(s3)
with scol4:
  s4 = st.file_uploader("Upload Sales File 4", type=file_types, key="sales4")
  if s4:
    uploaded_sales_files.append(s4)
with scol5:
  s5 = st.file_uploader("Upload Sales File 5", type=file_types, key="sales5")
  if s5:
    uploaded_sales_files.append(s5)

master_sales_df = pd.DataFrame()
consolidated_sales_raw_tabs = {}

if uploaded_sales_files:
  sales_dfs_list = []
  for s_file in uploaded_sales_files:
    try:
      if s_file.name.lower().endswith(".csv"):
        sales_dfs_list.append(pd.read_csv(s_file, low_memory=False))
      else:
        s_xls = pd.ExcelFile(s_file)
        for s_sheet in s_xls.sheet_names:
          df_s_sheet = pd.read_excel(s_xls, sheet_name=s_sheet)
          sales_dfs_list.append(df_s_sheet)
          consolidated_sales_raw_tabs.setdefault(s_sheet, []).append(
              df_s_sheet.copy()
          )
    except Exception as e:
      st.error(f"Error reading sales file {s_file.name}: {str(e)}")

  if sales_dfs_list:
    raw_sales_concat = pd.concat(sales_dfs_list, ignore_index=True)
    prod_col = next(
        (
            c
            for c in [
                "PRODUCT_NAME",
                "Product Name",
                "Item Name",
                "product_name",
            ]
            if c in raw_sales_concat.columns
        ),
        None,
    )
    date_col = next(
        (
            c
            for c in ["ORDER_DATE", "Order Date", "date", "Date"]
            if c in raw_sales_concat.columns
        ),
        None,
    )
    qty_col = next(
        (
            c
            for c in ["UNITS_SOLD", "Quantity", "Qty", "qty_sold"]
            if c in raw_sales_concat.columns
        ),
        None,
    )
    mrp_col = next(
        (
            c
            for c in [
                "Selling Price (Rs)",
                "MRP (Rs)",
                "GMV",
                "BASE_MRP",
                "mrp",
            ]
            if c in raw_sales_concat.columns
        ),
        None,
    )
    cust_city_col = next(
        (
            c
            for c in [
                "Customer City",
                "CITY",
                "Supply City",
                "city_name",
                "City",
            ]
            if c in raw_sales_concat.columns
        ),
        None,
    )
    cust_state_col = next(
        (
            c
            for c in ["Customer State", "Supply State", "State", "Region"]
            if c in raw_sales_concat.columns
        ),
        None,
    )
    brand_col = next(
        (
            c
            for c in ["Brand Name", "BRAND", "Franchise", "Brand"]
            if c in raw_sales_concat.columns
        ),
        None,
    )
    cat_col = next(
        (
            c
            for c in [
                "Business Category",
                "L1_CATEGORY",
                "category",
                "Category",
            ]
            if c in raw_sales_concat.columns
        ),
        None,
    )
    order_id_col = next(
        (
            c
            for c in ["Order Id", "Order ID", "order_id"]
            if c in raw_sales_concat.columns
        ),
        None,
    )

    master_sales_df = raw_sales_concat.copy()
    master_sales_df["_product_name"] = (
        master_sales_df[prod_col].astype(str) if prod_col else ""
    )
    master_sales_df["_brand"] = (
        master_sales_df[brand_col].astype(str) if brand_col else ""
    )
    master_sales_df["_category"] = (
        master_sales_df[cat_col].astype(str) if cat_col else ""
    )
    master_sales_df["_order_id"] = (
        master_sales_df[order_id_col] if order_id_col else np.nan
    )

    if date_col:
      master_sales_df["_sales_dt"] = pd.to_datetime(
          master_sales_df[date_col], errors="coerce"
      )
      master_sales_df["_raw_month"] = master_sales_df["_sales_dt"].dt.strftime(
          "%B"
      )
      master_sales_df["Duration (Week)"] = master_sales_df[
          "_sales_dt"
      ].apply(format_duration_week)
      master_sales_df["Date"] = master_sales_df["_sales_dt"].dt.strftime(
          "%Y/%m/%d"
      )
    else:
      master_sales_df["_raw_month"] = "Unknown"
      master_sales_df["Duration (Week)"] = np.nan
      master_sales_df["Date"] = np.nan

    master_sales_df["_qty"] = (
        pd.to_numeric(
            master_sales_df[qty_col].astype(str).str.replace(",", ""),
            errors="coerce",
        ).fillna(1.0)
        if qty_col
        else 1.0
    )
    master_sales_df["_gross_amt"] = (
        pd.to_numeric(
            master_sales_df[mrp_col].astype(str).str.replace(",", ""),
            errors="coerce",
        ).fillna(0.0)
        if mrp_col
        else 0.0
    )
    master_sales_df["Customer City"] = (
        master_sales_df[cust_city_col] if cust_city_col else "Unknown"
    )
    master_sales_df["Customer State"] = (
        master_sales_df[cust_state_col] if cust_state_col else "Unknown"
    )

    # Derive Customer Tier and Region strictly from Customer State & City
    master_sales_df["Customer Tier"] = master_sales_df["Customer City"].apply(
        get_city_tier
    )
    master_sales_df["Region"] = master_sales_df["Customer State"].apply(
        get_region_for_state
    )

if uploaded_files:
  consolidated_dfs, consolidated_raw_tabs = [], {}
  for uploaded_file in uploaded_files:
    try:
      if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file, low_memory=False)
        consolidated_raw_tabs.setdefault("Sheet1", []).append(df.copy())
        consolidated_dfs.append(df)
      else:
        xls = pd.ExcelFile(uploaded_file)
        campaign_type_map = {}
        if "MTD Claimables" in xls.sheet_names:
          df_claim = pd.read_excel(xls, sheet_name="MTD Claimables")
          if (
              "Campaign ID" in df_claim.columns
              and "Campaign Type" in df_claim.columns
          ):
            campaign_type_map = df_claim.set_index("Campaign ID")[
                "Campaign Type"
            ].to_dict()
        for sheet_name in xls.sheet_names:
          df = pd.read_excel(xls, sheet_name=sheet_name)
          if sheet_name.lower() not in [
              "summary",
              "consolidated_master",
              "mtd claimables",
          ]:
            if (
                "Campaign Type" not in df.columns
                and "Campaign ID" in df.columns
                and campaign_type_map
            ):
              df["Campaign Type"] = df["Campaign ID"].map(campaign_type_map)
            consolidated_raw_tabs.setdefault(sheet_name, []).append(df.copy())
            consolidated_dfs.append(df)
    except Exception as e:
      st.error(f"Error processing {uploaded_file.name}: {str(e)}")

  if consolidated_dfs:
    final_df = pd.concat(consolidated_dfs, ignore_index=True, sort=False)

    def get_numeric(cols):
      for c in cols:
        if c in final_df.columns:
          return (
              pd.to_numeric(
                  final_df[c].astype(str).str.replace(",", ""), errors="coerce"
              )
              .fillna(0.0)
          )
      return pd.Series(0.0, index=final_df.index)

    date_col_ad = next(
        (c for c in ["date_ist", "Date", "date"] if c in final_df.columns), None
    )
    final_df["_date_dt"] = (
        pd.to_datetime(final_df[date_col_ad], errors="coerce")
        if date_col_ad
        else pd.NaT
    )
    final_df["Date"] = final_df["_date_dt"].dt.strftime("%Y/%m/%d")
    final_df["_raw_month"] = final_df["_date_dt"].dt.strftime("%B")
    final_df["Duration (Week)"] = final_df["_date_dt"].apply(
        format_duration_week
    )
    final_df["_impressions"] = get_numeric(["Impressions"])
    final_df["_atc"] = get_numeric(["Direct ATC"]) + get_numeric(
        ["Indirect ATC"]
    )
    final_df["_direct_qty"] = get_numeric(
        ["Direct Quantities Sold", "Direct Orders"]
    )
    final_df["_indirect_qty"] = get_numeric(
        ["Indirect Quantities Sold", "Indirect Orders"]
    )
    final_df["_ad_qty"] = final_df["_direct_qty"] + final_df["_indirect_qty"]
    final_df["_direct_sales"] = get_numeric(["Direct Sales", "Claimables"])
    final_df["_indirect_sales"] = get_numeric(["Indirect Sales"])
    final_df["_sales"] = final_df["_direct_sales"] + final_df["_indirect_sales"]
    final_df["_budget_consumed"] = get_numeric([
        "served_budget_consumed",
        "Estimated Budget Consumed",
        "Budget Consumed",
        "Spend",
    ])

    all_raw_months = pd.concat([
        final_df["_raw_month"],
        master_sales_df.get("_raw_month", pd.Series()),
    ])
    chrono_map = get_chronological_month_mapping(all_raw_months)
    final_df["Duration (Month)"] = final_df["_raw_month"].map(chrono_map)
    if not master_sales_df.empty:
      master_sales_df["Duration (Month)"] = master_sales_df["_raw_month"].map(
          chrono_map
      )

    st.markdown("### 🔍 Global Month Filter")
    avail_months = sorted([
        m
        for m in final_df["Duration (Month)"].dropna().unique()
        if str(m) != "nan"
    ])
    sel_month = st.selectbox("Select Month:", ["All Months"] + avail_months)

    filt_df = (
        final_df
        if sel_month == "All Months"
        else final_df[final_df["Duration (Month)"] == sel_month]
    )
    tot_ad_sales = round(filt_df["_sales"].sum(), 2)
    tot_spends = round(filt_df["_budget_consumed"].sum(), 2)
    ad_roas = (
        round(tot_ad_sales / tot_spends, 2) if tot_spends > 0 else 0.0
    )

    if not master_sales_df.empty:
      filt_sales_df = (
          master_sales_df
          if sel_month == "All Months"
          else master_sales_df[master_sales_df["Duration (Month)"] == sel_month]
      )
      tot_overall_sales = round(filt_sales_df["_gross_amt"].sum(), 2)
      troas = (
          round(tot_overall_sales / tot_spends, 2) if tot_spends > 0 else 0.0
      )
      tacos = (
          round((tot_spends / tot_overall_sales) * 100, 2)
          if tot_overall_sales > 0
          else 0.0
      )
    else:
      tot_overall_sales = 0.0
      troas = 0.0
      tacos = 0.0

    st.markdown("### 📈 Overall Campaign Performance Metrics")
    row1_c1, row1_c2, row1_c3 = st.columns(3)
    with row1_c1:
      st.metric("Ad Sales", f"₹{int(round(tot_ad_sales)):,}")
    with row1_c2:
      st.metric("Total Sales", f"₹{int(round(tot_overall_sales)):,}")
    with row1_c3:
      st.metric("Ad Spends", f"₹{int(round(tot_spends)):,}")

    row2_c1, row2_c2, row2_c3 = st.columns(3)
    with row2_c1:
      st.metric("Ad ROAS", f"{ad_roas:.2f}x")
    with row2_c2:
      st.metric("TROAS", f"{troas:.2f}x")
    with row2_c3:
      st.metric("TACOS", f"{tacos:.2f}%")

    st.divider()
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📅 Daily Overview",
        "🏙️ City Overview",
        "🗓️ Week Overview",
        "📑 ADs Dump Preview",
        "🛒 Overall Dump Preview",
    ])

    with tab1:
      st.markdown("#### Day-Wise Performance Breakdown")
      daily_avail_months = sorted([
          m
          for m in final_df["Duration (Month)"].dropna().unique()
          if str(m) != "nan"
      ])
      daily_sel_month = st.selectbox(
          "Filter Month (Daily Overview):",
          ["All Months"] + daily_avail_months,
          key="daily_month_filter",
      )
      daily_filt_ad = (
          final_df
          if daily_sel_month == "All Months"
          else final_df[final_df["Duration (Month)"] == daily_sel_month]
      )

      ads_daily_agg = (
          daily_filt_ad.groupby("Date")
          .agg(
              Imp=("_impressions", "sum"),
              Spends=("_budget_consumed", "sum"),
              Ad_sales=("_sales", "sum"),
              AD_QTY=("_ad_qty", "sum"),
          )
          .reset_index()
      )

      if not master_sales_df.empty:
        daily_filt_sales = (
            master_sales_df
            if daily_sel_month == "All Months"
            else master_sales_df[
                master_sales_df["Duration (Month)"] == daily_sel_month
            ]
        )
        sales_daily_agg = (
            daily_filt_sales.groupby("Date")
            .agg(
                Overall_sales=("_gross_amt", "sum"),
                OV_QTY=("_qty", "sum"),
                Orders=("_order_id", "nunique"),
            )
            .reset_index()
        )
        daily_overview = pd.merge(
            ads_daily_agg, sales_daily_agg, on="Date", how="outer"
        ).fillna(0)
      else:
        daily_overview = ads_daily_agg.copy()
        daily_overview["Overall_sales"] = 0.0
        daily_overview["OV_QTY"] = 0.0
        daily_overview["Orders"] = 0

      daily_overview["CPM"] = (
          (daily_overview["Spends"] / daily_overview["Imp"] * 1000)
          .round(2)
          .fillna(0.0)
      )
      daily_overview["Ad ROAS"] = (
          (daily_overview["Ad_sales"] / daily_overview["Spends"])
          .round(2)
          .fillna(0.0)
      )
      daily_overview["TROAS"] = (
          (daily_overview["Overall_sales"] / daily_overview["Spends"])
          .round(2)
          .fillna(0.0)
      )
      daily_overview["TACOS"] = (
          (
              (daily_overview["Spends"] / daily_overview["Overall_sales"])
              * 100
          )
          .round(2)
          .fillna(0.0)
          .astype(str)
          + "%"
      )
      daily_overview["AOV"] = (
          (daily_overview["Overall_sales"] / daily_overview["Orders"])
          .round(2)
          .fillna(0.0)
      )

      daily_cols = [
          "Date",
          "Imp",
          "CPM",
          "Spends",
          "Ad_sales",
          "Overall_sales",
          "TACOS",
          "Ad ROAS",
          "TROAS",
          "OV_QTY",
          "AD_QTY",
          "AOV",
      ]
      daily_overview = daily_overview[
          [c for c in daily_cols if c in daily_overview.columns]
      ].sort_values(by="Date")
      st.dataframe(daily_overview, use_container_width=True)

    with tab2:
      st.markdown("#### Customer City Performance Breakdown (From Overall Raw)")
      if not master_sales_df.empty:
        filt_s_city = (
            master_sales_df
            if sel_month == "All Months"
            else master_sales_df[master_sales_df["Duration (Month)"] == sel_month]
        )
        sales_city_agg = (
            filt_s_city.groupby(["Customer City", "Customer Tier", "Region"])
            .agg(
                Overall_sales=("_gross_amt", "sum"),
                OV_QTY=("_qty", "sum"),
                Orders=("_order_id", "nunique"),
            )
            .reset_index()
        )

        city_col_ad = next(
            (
                c
                for c in [
                    "Customer City",
                    "City",
                    "Supply City",
                    "city_name",
                ]
                if c in filt_df.columns
            ),
            None,
        )
        if city_col_ad:
          ads_city_agg = (
              filt_df.groupby(city_col_ad)
              .agg(
                  Imp=("_impressions", "sum"),
                  Spends=("_budget_consumed", "sum"),
                  Ad_sales=("_sales", "sum"),
                  AD_QTY=("_ad_qty", "sum"),
              )
              .reset_index()
              .rename(columns={city_col_ad: "Customer City"})
          )
          city_overview = pd.merge(
              sales_city_agg, ads_city_agg, on="Customer City", how="left"
          ).fillna(0)
        else:
          city_overview = sales_city_agg.copy()
          city_overview["Imp"] = 0.0
          city_overview["Spends"] = 0.0
          city_overview["Ad_sales"] = 0.0
          city_overview["AD_QTY"] = 0.0

        city_overview["CPM"] = (
            (city_overview["Spends"] / city_overview["Imp"] * 1000)
            .round(2)
            .fillna(0.0)
        )
        city_overview["Ad ROAS"] = (
            (city_overview["Ad_sales"] / city_overview["Spends"])
            .round(2)
            .fillna(0.0)
        )
        city_overview["TROAS"] = (
            (city_overview["Overall_sales"] / city_overview["Spends"])
            .round(2)
            .fillna(0.0)
        )
        city_overview["TACOS"] = (
            (
                (city_overview["Spends"] / city_overview["Overall_sales"])
                * 100
            )
            .round(2)
            .fillna(0.0)
            .astype(str)
            + "%"
        )
        city_overview["AOV"] = (
            (city_overview["Overall_sales"] / city_overview["Orders"])
            .round(2)
            .fillna(0.0)
        )

        city_cols = [
            "Customer City",
            "Customer Tier",
            "Region",
            "Imp",
            "CPM",
            "Spends",
            "Ad_sales",
            "Overall_sales",
            "TACOS",
            "Ad ROAS",
            "TROAS",
            "OV_QTY",
            "AD_QTY",
            "AOV",
        ]
        city_overview = city_overview[
            [c for c in city_cols if c in city_overview.columns]
        ].sort_values(by="Overall_sales", ascending=False)
        st.dataframe(city_overview, use_container_width=True)

        st.markdown("### 📊 Region-Wise & Performance Distribution Charts")
        chart_col1, chart_col2 = st.columns(2)

        # Professional Blue/Lavender Color Palette
        blue_palette = [
            "#4A90C2",
            "#6AA6CF",
            "#285B7A",
            "#93C5FD",
            "#A78BFA",
            "#3B82F6",
        ]

        with chart_col1:
          st.markdown(
              "<h4 style='text-align: center;'>Region-Wise Overall Sales"
              " Distribution</h4>",
              unsafe_allow_html=True,
          )
          region_sales = (
              city_overview.groupby("Region")["Overall_sales"]
              .sum()
              .reset_index()
          )
          if not region_sales.empty and region_sales["Overall_sales"].sum() > 0:
            fig, ax = plt.subplots(figsize=(6, 4))
            ax.pie(
                region_sales["Overall_sales"],
                labels=region_sales["Region"],
                autopct="%1.1f%%",
                startangle=140,
                colors=blue_palette[: len(region_sales)],
                wedgeprops={
                    "edgecolor": "white",
                    "linewidth": 2,
                    "antialiased": True,
                },
                textprops={"color": "#29465B", "weight": "bold"},
            )
            ax.axis("equal")
            st.pyplot(fig)
          else:
            st.info("No sales data available for region distribution chart.")

        with chart_col2:
          st.markdown(
              "<h4 style='text-align: center;'>Region-Wise Ad Sales vs Spends"
              " Distribution</h4>",
              unsafe_allow_html=True,
          )
          region_ads = (
              city_overview.groupby("Region")[["Ad_sales", "Spends"]]
              .sum()
              .reset_index()
          )
          if not region_ads.empty and region_ads["Ad_sales"].sum() > 0:
            fig, ax = plt.subplots(figsize=(6, 4))
            ax.pie(
                region_ads["Ad_sales"],
                labels=region_ads["Region"],
                autopct="%1.1f%%",
                startangle=140,
                colors=blue_palette[2:],
                wedgeprops={
                    "edgecolor": "white",
                    "linewidth": 2,
                    "antialiased": True,
                },
                textprops={"color": "#29465B", "weight": "bold"},
            )
            ax.axis("equal")
            st.pyplot(fig)
          else:
            st.info("No ad sales data available for distribution chart.")

      else:
        st.info("Sales report required to render Customer City overview.")

    with tab3:
      st.markdown("#### Weekly Performance Breakdown")
      week_avail_months = sorted([
          m
          for m in final_df["Duration (Month)"].dropna().unique()
          if str(m) != "nan"
      ])
      wk_sel_month = st.selectbox(
          "Filter Month (Weekly Overview):",
          ["All Months"] + week_avail_months,
          key="wk_month_filter",
      )
      wk_filt_df = (
          final_df
          if wk_sel_month == "All Months"
          else final_df[final_df["Duration (Month)"] == wk_sel_month]
      )

      week_pivot = (
          wk_filt_df.groupby("Duration (Week)")
          .agg(
              Impressions=("_impressions", "sum"),
              Spends=("_budget_consumed", "sum"),
              Ad_Sales=("_sales", "sum"),
              AD_QTY=("_ad_qty", "sum"),
          )
          .reset_index()
      )
      week_pivot["_wk_num"] = week_pivot["Duration (Week)"].apply(
          extract_week_num
      )
      week_pivot = week_pivot.sort_values(by="_wk_num").drop(
          columns=["_wk_num"]
      )
      week_pivot["Ad ROAS"] = (
          (week_pivot["Ad_Sales"] / week_pivot["Spends"])
          .round(2)
          .fillna(0.0)
      )
      st.dataframe(week_pivot, use_container_width=True)

    def safe_get(df_src, possible):
      for p in possible:
        if p in df_src.columns:
          return df_src[p]
      return np.nan

    def derive_ad_type_and_targeting(df_in):
      ad_type_series = safe_get(df_in, ["Ad type", "Ad Type", "Campaign Type"])
      if (
          type(ad_type_series) == type(np.nan)
          or ad_type_series is np.nan
          or isinstance(ad_type_series, float)
      ):
        ad_type_series = pd.Series("", index=df_in.index)
      targeting_list = []
      for val in ad_type_series:
        v_str = str(val).lower()
        if "product booster" in v_str or "pb" in v_str:
          targeting_list.append("Keyword")
        elif "recommendation" in v_str or "rec" in v_str:
          targeting_list.append("Recommendation")
        else:
          targeting_list.append("Keyword")
      return ad_type_series, pd.Series(targeting_list, index=df_in.index)

    ad_types_col, targeting_types_col = derive_ad_type_and_targeting(filt_df)

    ads_export_df = pd.DataFrame({
        "Duration (Month)": filt_df["Duration (Month)"],
        "Duration (Week)": filt_df["Duration (Week)"],
        "Manufacturer": safe_get(filt_df, ["Manufacturer", "Manufacturer Name"]),
        "Date": filt_df["Date"],
        "Campaign ID": safe_get(filt_df, ["Campaign ID", "CampaignId"]),
        "Campaign Name": safe_get(filt_df, ["Campaign Name", "Campaign"]),
        "City": safe_get(
            filt_df, ["City", "Customer City", "Supply City", "city_name"]
        ),
        "Ad type": ad_types_col,
        "Targeting Type": targeting_types_col,
        "Franchise": safe_get(filt_df, ["Franchise", "Brand", "BRAND"]),
        "Category": safe_get(filt_df, ["Category", "category", "L1_CATEGORY"]),
        "Sub Category": safe_get(
            filt_df, ["Sub Category", "Subcategory", "L2_CATEGORY"]
        ),
        "Keyword": safe_get(filt_df, ["Keyword", "Search Term"]),
        "Most Viewed Position": safe_get(
            filt_df, ["Most Viewed Position", "Average Position"]
        ),
        "Match Type": safe_get(filt_df, ["Match Type"]),
        "Pacing Type": safe_get(filt_df, ["Pacing Type"]),
        "CPM": safe_get(filt_df, ["CPM"]),
        "Total Budget": safe_get(filt_df, ["Total Budget", "Budget"]),
        "Impressions": filt_df["_impressions"],
        "Direct ATC": safe_get(filt_df, ["Direct ATC"]),
        "Indirect ATC": safe_get(filt_df, ["Indirect ATC"]),
        "New Users": safe_get(filt_df, ["New Users"]),
        "Direct Sales": safe_get(filt_df, ["Direct Sales"]),
        "Indirect Sales": safe_get(filt_df, ["Indirect Sales"]),
        "Direct Quantities Sold": safe_get(
            filt_df, ["Direct Quantities Sold", "Direct Orders"]
        ),
        "Indirect Quantities Sold": safe_get(
            filt_df, ["Indirect Quantities Sold", "Indirect Orders"]
        ),
        "Estimated Budget Consumed": filt_df["_budget_consumed"],
        "Direct RoAS": safe_get(filt_df, ["Direct RoAS", "Direct ROAS"]),
        "Total RoAS": safe_get(filt_df, ["Total RoAS", "Total ROAS", "ROAS"]),
        "latest on hold timestamp": safe_get(
            filt_df, ["latest on hold timestamp"]
        ),
    })

    with tab4:
      st.markdown("#### Preview: ADs Dump Output Table")
      st.dataframe(ads_export_df.head(100), use_container_width=True)

    if not master_sales_df.empty:
      filt_s = (
          master_sales_df
          if sel_month == "All Months"
          else master_sales_df[master_sales_df["Duration (Month)"] == sel_month]
      )
      final_overall_export = pd.DataFrame({
          "item_id": safe_get(
              filt_s, ["ITEM_CODE", "Item Id", "item_id", "SKU ID"]
          ),
          "item_name": filt_s["_product_name"],
          "Franchise": filt_s["_brand"],
          "manufacturer_id": safe_get(
              filt_s, ["manufacturer_id", "Manufacturer ID"]
          ),
          "manufacturer_name": safe_get(
              filt_s, ["manufacturer_name", "Manufacturer"]
          ),
          "city_id": safe_get(
              filt_s, ["STORE_ID", "city_id", "City ID", "Order Id"]
          ),
          "city_name": filt_s["Customer City"],
          "Customer Tier": filt_s["Customer Tier"],
          "Region": filt_s["Region"],
          "category": filt_s["_category"],
          "Date": filt_s["Date"],
          "Duration (Month)": filt_s["Duration (Month)"],
          "Duration (Week)": filt_s["Duration (Week)"],
          "qty_sold": filt_s["_qty"],
          "mrp": filt_s["_gross_amt"],
      })
    else:
      final_overall_export = pd.DataFrame(columns=[
          "item_id",
          "item_name",
          "Franchise",
          "manufacturer_id",
          "manufacturer_name",
          "city_id",
          "city_name",
          "Customer Tier",
          "Region",
          "category",
          "Date",
          "Duration (Month)",
          "Duration (Week)",
          "qty_sold",
          "mrp",
      ])

    with tab5:
      st.markdown("#### Preview: Overall Dump Output Table")
      st.dataframe(final_overall_export.head(100), use_container_width=True)

    st.divider()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
      ads_export_df.to_excel(writer, sheet_name="ADs Dump", index=False)
      final_overall_export.to_excel(writer, sheet_name="Overall Dump", index=False)
      for r_name, d_list in consolidated_raw_tabs.items():
        pd.concat(d_list, ignore_index=True, sort=False).to_excel(
            writer, sheet_name=r_name[:31], index=False
        )
      for r_name, d_list in consolidated_sales_raw_tabs.items():
        if r_name[:31] not in writer.sheets:
          pd.concat(d_list, ignore_index=True, sort=False).to_excel(
              writer, sheet_name=r_name[:31], index=False
          )

    buf.seek(0)
    st.download_button(
        label=(
            "📥 Download Complete Excel Workbook (ADs Dump + Overall Dump +"
            " Raw Tabs)"
        ),
        data=buf,
        file_name="Master_Report.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        type="primary",
    )
  else:
    st.info("Please upload valid ad campaign reports.")
else:
    st.info("Awaiting file uploads.")
