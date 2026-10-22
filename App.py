import os
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(
    page_title="STAZE Blinkit Summary",
    page_icon="data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA1MTIgNTEyIj48cmVjdCB3aWR0aD0iNTEyIiBoZWlnaHQ9IjUxMiIgcng9IjExMiIgZmlsbD0iI2ZkZDA0MyIvPjx0ZXh0IHg9IjI1NiIgeT0iMzIwIiBmb250LWZhbWlseT0iQXJpYWwsIHNhbnMtc2VyaWYiIGZvbnQtd2VpZ2h0PSJib2x0IiBmb250LXNpemU9IjIxMCIgdGV4dC1hbmNob3I9Im1pZGRsZSI+PGZzcGFuIGZpbGw9IiMwMDAwMDAiPmJsaW5rPC9mc3Bhbj48ZnNwYW4gZmlsbD0iIzAwODEwNyI+aXQ8L2ZzcBhbj48L3RleHQ+PC9zdmc+",
    layout="wide"
)

st.markdown("""
<style>
    :root {
        --dash-blue: #4A90C2;
        --dash-blue-dark: #285B7A;
        --dash-blue-light: #EEF6FC;
        --dash-blue-soft: #DCECF8;
        --dash-border: #C8DCEB;
        --dash-text: #29465B;
        --lavender-light: #F3EBFB;
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
        align-items: center !important;
        justify-content: center !important;
    }

    [data-testid="stMetricLabel"] {
        color: #52708D !important;
        font-weight: 600 !important;
        text-align: center !important;
        display: flex;
        justify-content: center;
    }

    [data-testid="stMetricValue"] {
        color: var(--dash-blue-dark) !important;
        font-weight: 750 !important;
        text-align: center !important;
        display: flex;
        justify-content: center;
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
        text-align: center;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #6AA6CF, #4A90C2) !important;
        color: #FFFFFF !important;
        box-shadow: 0 3px 9px rgba(74, 144, 194, 0.25);
        border: 0 !important;
    }

    [data-baseweb="select"] > div {
        background: var(--lavender-blue-mix) !important;
        border-color: var(--lavender-border) !important;
        border-radius: 12px !important;
    }

    [data-testid="stFileUploader"] {
        background: #FFFFFF;
        border: 1px solid var(--dash-border);
        border-radius: 14px;
        padding: 8px;
        box-shadow: 0 2px 8px rgba(74, 144, 194, 0.07);
    }

    [data-testid="stFileUploaderDropzone"] {
        background: #F7FBFE;
        border: 1px dashed #9FC4DE;
        border-radius: 10px;
    }

    [data-testid="stFileUploader"] button,
    .stDownloadButton button,
    .stButton button {
        background: linear-gradient(135deg, #6AA6CF, #4A90C2) !important;
        color: #FFFFFF !important;
        border: 0 !important;
        border-radius: 12px !important;
        font-weight: 650 !important;
        box-shadow: 0 2px 7px rgba(11, 92, 173, 0.20);
    }

    [data-testid="stFileUploader"] button:hover,
    .stDownloadButton button:hover,
    .stButton button:hover {
        background: #397BAA !important;
        color: #FFFFFF !important;
    }

    [data-testid="stDataFrame"] {
        border: 1px solid var(--dash-border);
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 2px 9px rgba(11, 92, 173, 0.06);
    }

    [data-testid="stDataFrame"] div[data-testid="stTable"] table,
    [data-testid="stDataFrame"] table,
    [data-testid="stDataFrame"] th,
    [data-testid="stDataFrame"] td,
    .stDataFrame td, .stDataFrame th {
        text-align: center !important;
        justify-content: center !important;
        alignment: center !important;
    }

    [data-testid="stDataFrame"] th {
        background: #E5F1F9 !important;
        color: #285B7A !important;
        font-weight: 700 !important;
        text-align: center !important;
    }

    [data-testid="stRadio"] label,
    [data-testid="stSelectbox"] label,
    [data-testid="stMultiSelect"] label {
        color: #315A7D !important;
        font-weight: 600 !important;
    }

    hr {
        border-color: #D7E4F2 !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="display: flex; align-items: center; justify-content: center; gap: 14px; margin-bottom: 5px;">
    <svg width="48" height="48" viewBox="0 0 512 512" style="border-radius: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.15);">
        <rect width="512" height="512" rx="112" fill="#fdd043"/>
        <text x="256" y="320" font-family="Arial, sans-serif" font-weight="bold" font-size="210" text-anchor="middle">
            <tspan fill="#000000">blink</tspan><tspan fill="#008107">it</tspan>
        </text>
    </svg>
    <h1 style="margin: 0; padding: 0; line-height: 1.1;">STAZE Blinkit Summary</h1>
</div>
""", unsafe_allow_html=True)

st.markdown("<p style='text-align: center;'>Upload Ad Campaign reports and Overall Sales reports. System automatically filters campaigns and products containing <b>'staze'</b>, computing total sales by summing MRP values directly across all levels.</p>", unsafe_allow_html=True)

MONTH_ORDER_ASC = {
    "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4,
    "MAY": 5, "JUNE": 6, "JULY": 7, "AUGUST": 8,
    "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12
}

def get_calendar_months_asc(df_input):
    if df_input is None or 'Month' not in df_input.columns:
        return []
    months = [str(x).strip().upper() for x in df_input['Month'].dropna().unique()]
    return sorted(months, key=lambda x: MONTH_ORDER_ASC.get(x, 99))

def style_and_export_pivot(pivot_df, sheet_name="Comparison"):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    
    top_header_fill = PatternFill(start_color="4A90C2", end_color="4A90C2", fill_type="solid")
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    sec_header_fill = PatternFill(start_color="E5F1F9", end_color="E5F1F9", fill_type="solid")
    sec_header_font = Font(name="Calibri", size=10, bold=True, color="285B7A")
    index_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    index_font = Font(name="Calibri", size=10, bold=True, color="000000")
    total_fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")
    data_font = Font(name="Calibri", size=10, color="000000")
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )

    if isinstance(pivot_df.columns, pd.MultiIndex):
        metrics = pivot_df.columns.get_level_values(1)
        entity_title = pivot_df.index.name if pivot_df.index.name else ""

        ws.cell(row=1, column=1, value="").fill = top_header_fill
        ws.cell(row=1, column=1).border = thin_border
        
        current_col = 2
        for month in pivot_df.columns.get_level_values(0).unique():
            sub_cols = [c for c in pivot_df.columns if c[0] == month]
            num_sub = len(sub_cols)
            if num_sub > 1:
                ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + num_sub - 1)
            cell = ws.cell(row=1, column=current_col, value=str(month))
            cell.fill = top_header_fill
            cell.font = top_header_font
            cell.alignment = align_center
            
            for c in range(current_col, current_col + num_sub):
                ws.cell(row=1, column=c).border = thin_border
                ws.cell(row=1, column=c).fill = top_header_fill
            current_col += num_sub

        cell_a2 = ws.cell(row=2, column=1, value=str(entity_title))
        cell_a2.fill = index_fill
        cell_a2.font = index_font
        cell_a2.alignment = align_center
        cell_a2.border = thin_border

        for col_idx, metric in enumerate(metrics, start=2):
            cell = ws.cell(row=2, column=col_idx, value=str(metric))
            cell.fill = sec_header_fill
            cell.font = sec_header_font
            cell.alignment = align_center
            cell.border = thin_border

        start_data_row = 3
        for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=start_data_row):
            is_grand_total = (str(idx_val).strip().lower() in ['grand total', 'aggregate / average'])
            
            idx_cell = ws.cell(row=r_idx, column=1, value=str(idx_val))
            idx_cell.fill = total_fill if is_grand_total else index_fill
            idx_cell.font = total_font if is_grand_total else index_font
            idx_cell.alignment = align_center
            idx_cell.border = thin_border

            for c_idx, (m_col, val) in enumerate(zip(metrics, row_data), start=2):
                val_cell = ws.cell(row=r_idx, column=c_idx)
                val_cell.font = total_font if is_grand_total else data_font
                if is_grand_total:
                    val_cell.fill = total_fill
                val_cell.alignment = align_center
                val_cell.border = thin_border

                if isinstance(val, (int, float, np.number)):
                    if str(m_col).upper() in ['ACOS', 'TACOS']:
                        val_cell.value = val / 100.0 if val > 1 else val
                        val_cell.number_format = '0.00%'
                    elif str(m_col).upper() in ['SPENDS', 'AD SALES', 'SALES', 'TOTAL SALES', 'OVERALL SALES']:
                        val_cell.value = int(round(val))
                        val_cell.number_format = '#,##0'
                    else:
                        val_cell.value = round(val, 2)
                        val_cell.number_format = '#,##0.00' if not float(val).is_integer() else '#,##0'
                else:
                    val_cell.value = val
    else:
        for col_idx, col_name in enumerate(pivot_df.columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=str(col_name))
            cell.fill = top_header_fill
            cell.font = top_header_font
            cell.alignment = align_center
            cell.border = thin_border

        for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=2):
            is_grand_total = (str(idx_val).strip().lower() in ['grand total', 'aggregate / average'])
            for c_idx, (col_name, val) in enumerate(zip(pivot_df.columns, row_data), start=1):
                val_cell = ws.cell(row=r_idx, column=c_idx)
                val_cell.font = total_font if is_grand_total else data_font
                if is_grand_total:
                    val_cell.fill = total_fill
                val_cell.alignment = align_center
                val_cell.border = thin_border
                
                if str(col_name).upper() in ['ACOS', 'TACOS'] and isinstance(val, (int, float, np.number)):
                    val_cell.value = val / 100.0 if val > 1 else val
                    val_cell.number_format = '0.00%'
                elif str(col_name).upper() in ['SPENDS', 'AD SALES', 'SALES', 'TOTAL SALES', 'OVERALL SALES'] and isinstance(val, (int, float, np.number)):
                    val_cell.value = int(round(val))
                    val_cell.number_format = '#,##0'
                elif isinstance(val, (int, float, np.number)):
                    val_cell.value = round(val, 2)
                    val_cell.number_format = '#,##0.00' if not float(val).is_integer() else '#,##0'
                else:
                    val_cell.value = val

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def convert_all_pivots_to_excel(pivot_dict):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    top_header_fill = PatternFill(start_color="4A90C2", end_color="4A90C2", fill_type="solid")
    top_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    sec_header_fill = PatternFill(start_color="E5F1F9", end_color="E5F1F9", fill_type="solid")
    sec_header_font = Font(name="Calibri", size=10, bold=True, color="285B7A")
    index_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    index_font = Font(name="Calibri", size=10, bold=True, color="000000")
    total_fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
    total_font = Font(name="Calibri", size=10, bold=True, color="000000")
    data_font = Font(name="Calibri", size=10, color="000000")
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )

    for sheet_name, pivot_df in pivot_dict.items():
        if pivot_df is not None and not pivot_df.empty:
            ws = wb.create_sheet(title=sheet_name[:31])

            if isinstance(pivot_df.columns, pd.MultiIndex):
                metrics = pivot_df.columns.get_level_values(1)
                entity_title = pivot_df.index.name if pivot_df.index.name else ""

                ws.cell(row=1, column=1, value="").fill = top_header_fill
                ws.cell(row=1, column=1).border = thin_border
                
                current_col = 2
                for month in pivot_df.columns.get_level_values(0).unique():
                    sub_cols = [c for c in pivot_df.columns if c[0] == month]
                    num_sub = len(sub_cols)
                    if num_sub > 1:
                        ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + num_sub - 1)
                    cell = ws.cell(row=1, column=current_col, value=str(month))
                    cell.fill = top_header_fill
                    cell.font = top_header_font
                    cell.alignment = align_center
                    
                    for c in range(current_col, current_col + num_sub):
                        ws.cell(row=1, column=c).border = thin_border
                        ws.cell(row=1, column=c).fill = top_header_fill
                    current_col += num_sub

                cell_a2 = ws.cell(row=2, column=1, value=str(entity_title))
                cell_a2.fill = index_fill
                cell_a2.font = index_font
                cell_a2.alignment = align_center
                cell_a2.border = thin_border

                for col_idx, metric in enumerate(metrics, start=2):
                    cell = ws.cell(row=2, column=col_idx, value=str(metric))
                    cell.fill = sec_header_fill
                    cell.font = sec_header_font
                    cell.alignment = align_center
                    cell.border = thin_border

                for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=3):
                    is_grand_total = (str(idx_val).strip().lower() in ['grand total', 'aggregate / average'])
                    
                    idx_cell = ws.cell(row=r_idx, column=1, value=str(idx_val))
                    idx_cell.fill = total_fill if is_grand_total else index_fill
                    idx_cell.font = total_font if is_grand_total else index_font
                    idx_cell.alignment = align_center
                    idx_cell.border = thin_border

                    for c_idx, (m_col, val) in enumerate(zip(metrics, row_data), start=2):
                        val_cell = ws.cell(row=r_idx, column=c_idx)
                        val_cell.font = total_font if is_grand_total else data_font
                        if is_grand_total:
                            val_cell.fill = total_fill
                        val_cell.alignment = align_center
                        val_cell.border = thin_border

                        if isinstance(val, (int, float, np.number)):
                            if str(m_col).upper() in ['ACOS', 'TACOS']:
                                val_cell.value = val / 100.0 if val > 1 else val
                                val_cell.number_format = '0.00%'
                            elif str(m_col).upper() in ['SPENDS', 'AD SALES', 'SALES', 'TOTAL SALES', 'OVERALL SALES']:
                                val_cell.value = int(round(val))
                                val_cell.number_format = '#,##0'
                            else:
                                val_cell.value = round(val, 2)
                                val_cell.number_format = '#,##0.00' if not float(val).is_integer() else '#,##0'
                        else:
                            val_cell.value = val
            else:
                for col_idx, col_name in enumerate(pivot_df.columns, start=1):
                    cell = ws.cell(row=1, column=col_idx, value=str(col_name))
                    cell.fill = top_header_fill
                    cell.font = top_header_font
                    cell.alignment = align_center
                    cell.border = thin_border

                for r_idx, (idx_val, row_data) in enumerate(pivot_df.iterrows(), start=2):
                    is_grand_total = (str(idx_val).strip().lower() in ['grand total', 'aggregate / average'])
                    for c_idx, (col_name, val) in enumerate(zip(pivot_df.columns, row_data), start=1):
                        val_cell = ws.cell(row=r_idx, column=c_idx)
                        val_cell.font = total_font if is_grand_total else data_font
                        if is_grand_total:
                            val_cell.fill = total_fill
                        val_cell.alignment = align_center
                        val_cell.border = thin_border
                        
                        if str(col_name).upper() in ['ACOS', 'TACOS'] and isinstance(val, (int, float, np.number)):
                            val_cell.value = val / 100.0 if val > 1 else val
                            val_cell.number_format = '0.00%'
                        elif str(col_name).upper() in ['SPENDS', 'AD SALES', 'SALES', 'TOTAL SALES', 'OVERALL SALES'] and isinstance(val, (int, float, np.number)):
                            val_cell.value = int(round(val))
                            val_cell.number_format = '#,##0'
                        elif isinstance(val, (int, float, np.number)):
                            val_cell.value = round(val, 2)
                            val_cell.number_format = '#,##0.00' if not float(val).is_integer() else '#,##0'
                        else:
                            val_cell.value = val

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def convert_dataframe_to_excel_download(df_to_export, sheet_name="Sheet1"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_to_export.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    buffer.seek(0)
    return buffer.getvalue()

st.markdown("### 📤 Upload Ad Campaign Reports (Up to 5)")
col1, col2, col3, col4, col5 = st.columns(5)
uploaded_files = []

ad_file_types = ["csv", "xlsx", "xls", "xlsb", "xlsm"]

with col1:
    f1 = st.file_uploader("Upload Ad File 1", type=ad_file_types, key="file1")
    if f1: uploaded_files.append(f1)
with col2:
    f2 = st.file_uploader("Upload Ad File 2", type=ad_file_types, key="file2")
    if f2: uploaded_files.append(f2)
with col3:
    f3 = st.file_uploader("Upload Ad File 3", type=ad_file_types, key="file3")
    if f3: uploaded_files.append(f3)
with col4:
    f4 = st.file_uploader("Upload Ad File 4", type=ad_file_types, key="file4")
    if f4: uploaded_files.append(f4)
with col5:
    f5 = st.file_uploader("Upload Ad File 5", type=ad_file_types, key="file5")
    if f5: uploaded_files.append(f5)

st.markdown("### 📥 Upload Sales Reports (Supports multiple schemas - Up to 5)")
scol1, scol2, scol3, scol4, scol5 = st.columns(5)
uploaded_sales_files = []

sales_file_types = ["csv", "xlsx", "xls", "xlsb", "xlsm"]

with scol1:
    s1 = st.file_uploader("Upload Sales File 1", type=sales_file_types, key="sales1")
    if s1: uploaded_sales_files.append(s1)
with scol2:
    s2 = st.file_uploader("Upload Sales File 2", type=sales_file_types, key="sales2")
    if s2: uploaded_sales_files.append(s2)
with scol3:
    s3 = st.file_uploader("Upload Sales File 3", type=sales_file_types, key="sales3")
    if s3: uploaded_sales_files.append(s3)
with scol4:
    s4 = st.file_uploader("Upload Sales File 4", type=sales_file_types, key="sales4")
    if s4: uploaded_sales_files.append(s4)
with scol5:
    s5 = st.file_uploader("Upload Sales File 5", type=sales_file_types, key="sales5")
    if s5: uploaded_sales_files.append(s5)

sales_dfs_list = []
if uploaded_sales_files:
    for s_file in uploaded_sales_files:
        try:
            fname_lower = s_file.name.lower()
            if fname_lower.endswith('.csv'):
                sdf = pd.read_csv(s_file, low_memory=False)
                sales_dfs_list.append(sdf)
            else:
                s_xls = pd.ExcelFile(s_file)
                for s_sheet in s_xls.sheet_names:
                    sdf = pd.read_excel(s_xls, sheet_name=s_sheet)
                    sales_dfs_list.append(sdf)
        except Exception as e:
            st.error(f"Error reading sales file {s_file.name}: {str(e)}")

master_sales_df = pd.DataFrame()
if sales_dfs_list:
    raw_sales_concat = pd.concat(sales_dfs_list, ignore_index=True)
    
    prod_col = next((c for c in ['Product Name', 'Item Name', 'product_name', 'item_name', 'Product'] if c in raw_sales_concat.columns), None)
    date_col = next((c for c in ['Order Date', 'date', 'Date', 'ORDER_DATE', 'timestamp'] if c in raw_sales_concat.columns), None)
    qty_col = next((c for c in ['Quantity', 'Qty', 'Units', 'qty_sold'] if c in raw_sales_concat.columns), None)
    mrp_col = next((c for c in ['MRP (Rs)', 'mrp', 'MRP', 'Selling Price (Rs)'] if c in raw_sales_concat.columns), None)
    city_col = next((c for c in ['Supply City', 'city_name', 'City'] if c in raw_sales_concat.columns), None)
    state_col = next((c for c in ['Supply State', 'Region', 'State'] if c in raw_sales_concat.columns), None)

    master_sales_df = raw_sales_concat.copy()
    if prod_col:
        master_sales_df['_product_name'] = master_sales_df[prod_col].astype(str)
    else:
        master_sales_df['_product_name'] = ""

    # Filter overall sales strictly for products with "staze" mentioned in the product name
    master_sales_df = master_sales_df[master_sales_df['_product_name'].str.lower().str.contains('staze', na=False)]

    if date_col:
        master_sales_df['_sales_dt'] = pd.to_datetime(master_sales_df[date_col], errors='coerce')
        master_sales_df['Month'] = master_sales_df['_sales_dt'].dt.strftime('%B').str.upper()
        master_sales_df['Date_str'] = master_sales_df['_sales_dt'].dt.strftime('%Y-%m-%d')
        master_sales_df['Day'] = master_sales_df['_sales_dt'].dt.day
    else:
        master_sales_df['Month'] = "UNKNOWN"
        master_sales_df['Date_str'] = np.nan
        master_sales_df['Day'] = 1

    if qty_col:
        master_sales_df['_qty'] = pd.to_numeric(master_sales_df[qty_col].astype(str).str.replace(',', ''), errors='coerce').fillna(1.0)
    else:
        master_sales_df['_qty'] = 1.0

    # Total sales calculation: Just add all MRP (no multiplication with quantity)
    if mrp_col:
        master_sales_df['_gross_amt'] = pd.to_numeric(master_sales_df[mrp_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
    else:
        gross_col = next((c for c in ['Total Gross Bill Amount', 'Gross Amount', 'Total Amount', 'Bill Amount', 'revenue', 'sales_amount'] if c in master_sales_df.columns), None)
        if gross_col:
            master_sales_df['_gross_amt'] = pd.to_numeric(master_sales_df[gross_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
        else:
            master_sales_df['_gross_amt'] = 0.0

    if city_col:
        master_sales_df['Supply City'] = master_sales_df[city_col]
    else:
        master_sales_df['Supply City'] = "Unknown"

    if state_col:
        master_sales_df['Supply State'] = master_sales_df[state_col]
    else:
        master_sales_df['Supply State'] = "Unknown"

if uploaded_files:
    consolidated_dfs = []
    raw_files_dict = {}
    consolidated_raw_tabs = {}

    for uploaded_file in uploaded_files:
        file_name_lower = uploaded_file.name.lower()
        fallback_month_name = os.path.splitext(uploaded_file.name)[0].upper()
        raw_files_dict[uploaded_file.name] = {}
        
        try:
            if file_name_lower.endswith('.csv'):
                df = pd.read_csv(uploaded_file, low_memory=False)
                sheet_name = "Sheet1"
                raw_files_dict[uploaded_file.name][sheet_name] = df.copy()
                
                date_col = next((c for c in ['Date', 'date', 'Day', 'DATE', 'date_ist', 'timestamp'] if c in df.columns), None)

                if date_col:
                    dt_series = pd.to_datetime(df[date_col], format='mixed', dayfirst=False, errors='coerce')
                    month_series = dt_series.dt.strftime('%B').str.upper()
                    df['Month'] = month_series.fillna(fallback_month_name)
                    df['_date_dt'] = dt_series
                    df['Date'] = dt_series.dt.strftime('%Y-%m-%d')
                else:
                    df['Month'] = fallback_month_name
                    df['_date_dt'] = pd.NaT
                    df['Date'] = np.nan

                df_raw = df.copy()
                if sheet_name not in consolidated_raw_tabs:
                    consolidated_raw_tabs[sheet_name] = []
                consolidated_raw_tabs[sheet_name].append(df_raw)
                
                df_consolidated = df.copy()
                if 'Month' in df_consolidated.columns:
                    col_month = df_consolidated.pop('Month')
                    df_consolidated.insert(0, 'Month', col_month)
                else:
                    df_consolidated.insert(0, 'Month', fallback_month_name)
                    
                df_consolidated.insert(1, 'Tab Name', sheet_name)
                
                if 'Date' in df_consolidated.columns:
                    col_date = df_consolidated.pop('Date')
                    df_consolidated.insert(2, 'Date', col_date)
                else:
                    df_consolidated.insert(2, 'Date', np.nan)

                consolidated_dfs.append(df_consolidated)
            else:
                xls = pd.ExcelFile(uploaded_file)
                sheets_to_process = xls.sheet_names[1:] if len(xls.sheet_names) > 1 else xls.sheet_names
                valid_sheet_names = [s for s in sheets_to_process if s.strip().lower() not in ['consolidated_master', 'summary', 'overall', 'ads dump', 'overall dump']]

                for sheet_name in valid_sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    raw_files_dict[uploaded_file.name][sheet_name] = df.copy()
                    
                    date_col = next((c for c in ['Date', 'date', 'Day', 'DATE', 'date_ist', 'timestamp'] if c in df.columns), None)

                    if date_col:
                        dt_series = pd.to_datetime(df[date_col], format='mixed', dayfirst=False, errors='coerce')
                        month_series = dt_series.dt.strftime('%B').str.upper()
                        df['Month'] = month_series.fillna(fallback_month_name)
                        df['_date_dt'] = dt_series
                        df['Date'] = dt_series.dt.strftime('%Y-%m-%d')
                    else:
                        df['Month'] = fallback_month_name
                        df['_date_dt'] = pd.NaT
                        df['Date'] = np.nan

                    df_raw = df.copy()
                    if sheet_name not in consolidated_raw_tabs:
                        consolidated_raw_tabs[sheet_name] = []
                    consolidated_raw_tabs[sheet_name].append(df_raw)
                    
                    df_consolidated = df.copy()
                    if 'Month' in df_consolidated.columns:
                        col_month = df_consolidated.pop('Month')
                        df_consolidated.insert(0, 'Month', col_month)
                    else:
                        df_consolidated.insert(0, 'Month', fallback_month_name)
                        
                    df_consolidated.insert(1, 'Tab Name', sheet_name)
                    
                    if 'Date' in df_consolidated.columns:
                        col_date = df_consolidated.pop('Date')
                        df_consolidated.insert(2, 'Date', col_date)
                    else:
                        df_consolidated.insert(2, 'Date', np.nan)

                    consolidated_dfs.append(df_consolidated)
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")

    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True)
        
        camp_col_candidates = [c for c in ['Campaign Name', 'Campaign', 'CampaignId', 'Campaign ID'] if c in final_df.columns]
        if camp_col_candidates:
            c_col = camp_col_candidates[0]
            final_df = final_df[final_df[c_col].astype(str).str.lower().str.contains('staze', na=False)]

        base_cols = ['Month', 'Tab Name', 'Date']
        product_listing_cols = []
        for df in consolidated_dfs:
            if 'PRODUCT_LISTING' in str(df['Tab Name'].values).upper():
                pl_df = df[df['Tab Name'].str.upper() == 'PRODUCT_LISTING']
                product_listing_cols = [c for c in pl_df.columns if c not in base_cols]
                break
        
        remaining_cols = [c for c in final_df.columns if c not in base_cols and c not in product_listing_cols]
        final_df = final_df.reindex(columns=base_cols + product_listing_cols + remaining_cols)

        match_col = 'Match Type' if 'Match Type' in final_df.columns else None
        target_col = 'Targeting Type' if 'Targeting Type' in final_df.columns else None
        tab_col = 'Tab Name' if 'Tab Name' in final_df.columns else None

        def is_na_series(series):
            if series is None or series.empty:
                return pd.Series(True, index=final_df.index)
            cleaned = series.astype(str).str.strip().str.upper()
            return series.isna() | cleaned.isin(['NA', 'N/A', 'NAN', 'NONE', ''])

        if match_col:
            na_match = is_na_series(final_df[match_col])
            if target_col:
                valid_target = ~is_na_series(final_df[target_col])
                fill_from_target = na_match & valid_target
                final_df.loc[fill_from_target, match_col] = final_df.loc[fill_from_target, target_col]
                na_match = is_na_series(final_df[match_col])

            if tab_col:
                final_df.loc[na_match, match_col] = final_df.loc[na_match, tab_col]
        elif target_col:
            final_df['Match Type'] = final_df[target_col]
            na_match = is_na_series(final_df['Match Type'])
            if tab_col:
                final_df.loc[na_match, 'Match Type'] = final_df.loc[na_match, tab_col]
            match_col = 'Match Type'

        def parse_robust_numeric(series):
            if series is None:
                return pd.Series(0.0, index=[])
            s_str = series.astype(str).str.replace(',', '', regex=False)
            return pd.to_numeric(s_str, errors='coerce').fillna(0.0)

        def get_numeric_col(df, possible_cols):
            for col in possible_cols:
                if col in df.columns:
                    return parse_robust_numeric(df[col])
            return pd.Series(0.0, index=df.index)

        final_df['_impressions'] = get_numeric_col(final_df, ['Impressions'])
        final_df['_direct_atc'] = get_numeric_col(final_df, ['Direct ATC'])
        final_df['_indirect_atc'] = get_numeric_col(final_df, ['Indirect ATC'])
        final_df['_atc'] = final_df['_direct_atc'] + final_df['_indirect_atc']

        final_df['_direct_orders'] = get_numeric_col(final_df, ['Direct Quantities Sold', 'Direct Orders'])
        final_df['_indirect_orders'] = get_numeric_col(final_df, ['Indirect Quantities Sold', 'Indirect Orders'])
        final_df['_orders'] = final_df['_direct_orders'] + final_df['_indirect_orders']

        final_df['_direct_sales'] = get_numeric_col(final_df, ['Direct Sales'])
        final_df['_indirect_sales'] = get_numeric_col(final_df, ['Indirect Sales'])
        final_df['_sales'] = final_df['_direct_sales'] + final_df['_indirect_sales']

        final_df['_budget_consumed'] = get_numeric_col(final_df, ['Estimated Budget Consumed', 'Budget Consumed', 'Spend'])
        final_df['_most_viewed_position'] = get_numeric_col(final_df, ['Most Viewed Position', 'Viewed Position', 'Average Position', 'Position'])

        if match_col and match_col in final_df.columns:
            final_df['Ad Type Combined'] = final_df[match_col].fillna("Other").astype(str).str.title()
        else:
            final_df['Ad Type Combined'] = "Other"

        def create_monthly_summary_table(df_input, selected_months=None):
            working_df = df_input.dropna(subset=['Month']).copy()
            if selected_months:
                working_df = working_df[working_df['Month'].isin(selected_months)]

            if working_df.empty:
                return pd.DataFrame()

            all_months = get_calendar_months_asc(working_df)
            has_sales = not master_sales_df.empty and '_sales_dt' in master_sales_df.columns

            rows_list = []
            for m in all_months:
                m_df = working_df[working_df['Month'] == m]
                imp = m_df['_impressions'].sum()
                atc = m_df['_atc'].sum()
                ords = m_df['_orders'].sum()
                spends = round(m_df['_budget_consumed'].sum(), 2)
                sales = round(m_df['_sales'].sum(), 2)
                cpm = round((spends / imp) * 1000, 2) if imp > 0 else 0.0
                roas = round(sales / spends, 2) if spends > 0 else 0.0
                acos = round((spends / sales) * 100, 2) if sales > 0 else 0.0
                
                row_dict = {
                    'Month': m,
                    'Impressions': imp,
                    'CPM': cpm,
                    'ATC': atc,
                    'Spends': spends,
                    'Ad Sales': sales,
                }
                
                if has_sales:
                    m_sales_df = master_sales_df[master_sales_df['Month'] == m]
                    ov_sales = round(m_sales_df['_gross_amt'].sum(), 2)
                    ov_orders = len(m_sales_df)
                    tacos = round((spends / ov_sales) * 100, 2) if ov_sales > 0 else 0.0
                    troas = round(ov_sales / spends, 2) if spends > 0 else 0.0
                    
                    row_dict['Overall Sales'] = ov_sales
                    row_dict['Ad ROAS'] = roas
                    row_dict['TROAS'] = troas
                    row_dict['Ad Orders'] = ords
                    row_dict['Overall Orders'] = ov_orders
                    row_dict['TACOS'] = f"{tacos:.2f}%"
                else:
                    row_dict['Ad ROAS'] = roas
                    row_dict['Ad Orders'] = ords

                row_dict.update({
                    'ACOS': f"{acos:.2f}%"
                })
                rows_list.append(row_dict)

            return pd.DataFrame(rows_list)

        def create_monthly_summary_daywise_table(selected_month):
            if master_sales_df.empty or '_sales_dt' not in master_sales_df.columns:
                return pd.DataFrame()

            m_df = master_sales_df[master_sales_df['Month'] == selected_month].copy()
            if m_df.empty:
                return pd.DataFrame()

            ad_df = final_df[final_df['Month'] == selected_month].copy() if 'Month' in final_df.columns else pd.DataFrame()

            rows = []
            for day_i in range(1, 32):
                day_sales = m_df[m_df['Day'] == day_i]
                ov_sales = round(day_sales['_gross_amt'].sum(), 2)
                ov_orders = len(day_sales)
                
                day_spends = 0.0
                day_imp = 0.0
                day_atc = 0.0
                day_ad_sales = 0.0
                day_ad_orders = 0.0

                if not ad_df.empty and 'Date' in ad_df.columns:
                    matched_ad = ad_df[ad_df['_date_dt'].dt.day == day_i]
                    day_spends = round(matched_ad['_budget_consumed'].sum(), 2)
                    day_imp = matched_ad['_impressions'].sum()
                    day_atc = matched_ad['_atc'].sum()
                    day_ad_sales = round(matched_ad['_sales'].sum(), 2)
                    day_ad_orders = int(matched_ad['_orders'].sum())

                cpm = round((day_spends / day_imp) * 1000, 2) if day_imp > 0 else 0.0
                roas = round(day_ad_sales / day_spends, 2) if day_spends > 0 else 0.0
                acos = round((day_spends / day_ad_sales) * 100, 2) if day_ad_sales > 0 else 0.0
                tacos = round((day_spends / ov_sales) * 100, 2) if ov_sales > 0 else 0.0
                troas = round(ov_sales / day_spends, 2) if day_spends > 0 else 0.0

                rows.append({
                    'Month': f"Day {day_i}",
                    'Impressions': int(day_imp),
                    'CPM': cpm,
                    'ATC': int(day_atc),
                    'Spends': day_spends,
                    'Ad Sales': day_ad_sales,
                    'Overall Sales': ov_sales,
                    'Ad ROAS': roas,
                    'TROAS': troas,
                    'Ad Orders': day_ad_orders,
                    'Overall Orders': ov_orders,
                    'TACOS': f"{tacos:.2f}%",
                    'ACOS': f"{acos:.2f}%"
                })

            res_df = pd.DataFrame(rows)
            total_sales_sum = res_df['Overall Sales'].sum()
            total_orders_sum = res_df['Overall Orders'].sum()
            total_spends_sum = res_df['Spends'].sum()
            total_ad_sales_sum = res_df['Ad Sales'].sum()
            total_ad_orders_sum = res_df['Ad Orders'].sum()
            total_atc_sum = res_df['ATC'].sum()
            total_imp_sum = res_df['Impressions'].sum()
            avg_cpm = round(res_df['CPM'].mean(), 2)
            avg_tacos = round((total_spends_sum / total_sales_sum) * 100, 2) if total_sales_sum > 0 else 0.0
            avg_acos = round((total_spends_sum / total_ad_sales_sum) * 100, 2) if total_ad_sales_sum > 0 else 0.0

            summary_row = pd.DataFrame([{
                'Month': 'Aggregate / Average',
                'Impressions': total_imp_sum,
                'CPM': avg_cpm,
                'ATC': int(total_atc_sum),
                'Spends': total_spends_sum,
                'Ad Sales': total_ad_sales_sum,
                'Overall Sales': total_sales_sum,
                'Ad Orders': total_ad_orders_sum,
                'Overall Orders': total_orders_sum,
                'TACOS': f"{avg_tacos:.2f}%",
                'ACOS': f"{avg_acos:.2f}%"
            }])
            res_df = pd.concat([res_df, summary_row], ignore_index=True)
            return res_df

        def create_mom_comparison_table(df_input, entity_col, selected_months=None):
            if entity_col not in df_input.columns:
                return pd.DataFrame()
            
            working_df = df_input.dropna(subset=[entity_col, 'Month']).copy()
            if selected_months:
                working_df = working_df[working_df['Month'].isin(selected_months)]

            if working_df.empty:
                return pd.DataFrame()
            
            grouped = working_df.groupby([entity_col, 'Month']).agg(
                Impressions=('_impressions', 'sum'),
                ATC=('_atc', 'sum'),
                Orders=('_orders', 'sum'),
                Spends=('_budget_consumed', 'sum'),
                Sales=('_sales', 'sum'),
                Most_Viewed_Position=('_most_viewed_position', 'mean')
            ).reset_index()

            grouped['CPM'] = grouped.apply(lambda r: round((r['Spends'] / r['Impressions']) * 1000, 2) if r['Impressions'] > 0 else 0.0, axis=1)
            grouped['ROAS'] = grouped.apply(lambda r: round(r['Sales'] / r['Spends'], 2) if r['Spends'] > 0 else 0.0, axis=1)
            grouped['ACOS'] = grouped.apply(lambda r: round((r['Spends'] / r['Sales']) * 100, 2) if r['Sales'] > 0 else 0.0, axis=1)
            grouped['Most Viewed Position'] = grouped['Most_Viewed_Position'].fillna(0).round().astype(int)

            pivot_df = grouped.pivot(
                index=entity_col,
                columns='Month',
                values=['Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS', 'Most Viewed Position']
            )

            metrics_order = ['Impressions', 'CPM', 'ATC', 'Orders', 'Spends', 'Sales', 'ROAS', 'ACOS', 'Most Viewed Position']
            all_months_asc = get_calendar_months_asc(working_df)
            all_months_sorted = sorted(all_months_asc, key=lambda x: MONTH_ORDER_ASC.get(str(x).strip().upper(), 99), reverse=False)
            
            pivot_df = pivot_df.reorder_levels([1, 0], axis=1)
            sorted_cols = pd.MultiIndex.from_product([all_months_sorted, metrics_order], names=['Month', 'Metric'])
            pivot_df = pivot_df.reindex(columns=sorted_cols).fillna(0)

            total_spends_per_entity = grouped.groupby(entity_col)['Spends'].sum()
            sorted_entities = total_spends_per_entity.sort_values(ascending=False).index
            pivot_df = pivot_df.reindex(sorted_entities)

            grand_total_series = {}
            for month in all_months_sorted:
                month_df = working_df[working_df['Month'] == month]
                total_imp = month_df['_impressions'].sum()
                total_atc = month_df['_atc'].sum()
                total_orders = month_df['_orders'].sum()
                total_spends = month_df['_budget_consumed'].sum()
                total_sales = month_df['_sales'].sum()
                mean_pos = month_df['_most_viewed_position'].mean()

                grand_total_series[(month, 'Impressions')] = total_imp
                grand_total_series[(month, 'CPM')] = round((total_spends / total_imp) * 1000, 2) if total_imp > 0 else 0.0
                grand_total_series[(month, 'ATC')] = total_atc
                grand_total_series[(month, 'Orders')] = total_orders
                grand_total_series[(month, 'Spends')] = round(total_spends, 2)
                grand_total_series[(month, 'Sales')] = round(total_sales, 2)
                grand_total_series[(month, 'ROAS')] = round(total_sales / total_spends, 2) if total_spends > 0 else 0.0
                grand_total_series[(month, 'ACOS')] = round((total_spends / total_sales) * 100, 2) if total_sales > 0 else 0.0
                grand_total_series[(month, 'Most Viewed Position')] = int(round(mean_pos)) if pd.notna(mean_pos) else 0

            pivot_df.loc['Grand Total'] = grand_total_series
            pivot_df.index.name = entity_col.upper()
            return pivot_df

        def create_weekly_comparison_table_with_sales(df_input, selected_months=None):
            def assign_week_formatted_dashboard(dt):
                if pd.isna(dt): return np.nan
                iso_year, iso_week, iso_day = dt.isocalendar()
                start_of_week = dt - pd.Timedelta(days=iso_day - 1)
                end_of_week = start_of_week + pd.Timedelta(days=6)
                return f"Week {iso_week} ({start_of_week.strftime('%b %d')} to {end_of_week.strftime('%b %d')})"

            working_df = df_input.dropna(subset=['_date_dt', 'Month']).copy()
            working_df['Dashboard_Week'] = working_df['_date_dt'].apply(assign_week_formatted_dashboard)
            
            if selected_months:
                working_df = working_df[working_df['Month'].isin(selected_months)]
            if working_df.empty:
                return pd.DataFrame()

            weeks_list = sorted(working_df['Dashboard_Week'].dropna().unique(), key=lambda w: int(''.join(filter(str.isdigit, str(w))) or 99))
            has_sales = not master_sales_df.empty and '_sales_dt' in master_sales_df.columns

            rows = []
            for w in weeks_list:
                w_df = working_df[working_df['Dashboard_Week'] == w]
                imp = w_df['_impressions'].sum()
                spends = round(w_df['_budget_consumed'].sum(), 2)
                sales = round(w_df['_sales'].sum(), 2)
                orders = int(w_df['_orders'].sum())
                cpm = round((spends / imp) * 1000, 2) if imp > 0 else 0.0
                ad_roas = round(sales / spends, 2) if spends > 0 else 0.0
                
                row_data = {
                    'Week': w,
                    'Impressions': imp,
                    'CPM': cpm,
                    'Spends': spends,
                    'Ad Sales': sales,
                }
                
                if has_sales:
                    w_dates = w_df['Date'].dropna().unique()
                    w_sales_df = master_sales_df[master_sales_df['Date_str'].isin(w_dates)]
                    ov_sales = round(w_sales_df['_gross_amt'].sum(), 2)
                    ov_orders = len(w_sales_df)
                    tacos = round((spends / ov_sales) * 100, 2) if ov_sales > 0 else 0.0
                    troas = round(ov_sales / spends, 2) if spends > 0 else 0.0
                    
                    row_data['Overall Sales'] = ov_sales
                    row_data['Ad ROAS'] = ad_roas
                    row_data['TROAS'] = troas
                    row_data['Ad Orders'] = orders
                    row_data['Overall Orders'] = ov_orders
                    row_data['TACOS'] = f"{tacos:.2f}%"
                else:
                    row_data['Ad ROAS'] = ad_roas
                    row_data['Ad Orders'] = orders

                rows.append(row_data)

            res_df = pd.DataFrame(rows)
            if not res_df.empty:
                n_weeks = len(weeks_list)
                tot_imp = res_df['Impressions'].sum()
                avg_cpm = round(res_df['CPM'].mean(), 2)
                tot_spends = round(res_df['Spends'].sum(), 2)
                tot_ad_sales = round(res_df['Ad Sales'].sum(), 2)
                avg_ad_roas = round(tot_ad_sales / tot_spends, 2) if tot_spends > 0 else 0.0
                tot_ad_orders = int(res_df['Ad Orders'].sum()) if 'Ad Orders' in res_df.columns else 0

                summary_dict = {
                    'Week': f'Aggregate / Average ({n_weeks} Wks)',
                    'Impressions': tot_imp,
                    'CPM': avg_cpm,
                    'Spends': tot_spends,
                    'Ad Sales': tot_ad_sales,
                }
                if has_sales:
                    tot_ov_sales = round(res_df['Overall Sales'].sum(), 2) if 'Overall Sales' in res_df.columns else 0.0
                    tot_ov_orders = int(res_df['Overall Orders'].sum()) if 'Overall Orders' in res_df.columns else 0
                    avg_troas = round(tot_ov_sales / tot_spends, 2) if tot_spends > 0 else 0.0
                    avg_tacos = round((tot_spends / tot_ov_sales) * 100, 2) if tot_ov_sales > 0 else 0.0
                    summary_dict['Overall Sales'] = tot_ov_sales
                    summary_dict['Ad ROAS'] = avg_ad_roas
                    summary_dict['TROAS'] = avg_troas
                    summary_dict['Ad Orders'] = tot_ad_orders
                    summary_dict['Overall Orders'] = tot_ov_orders
                    summary_dict['TACOS'] = f"{avg_tacos:.2f}%"
                else:
                    summary_dict['Ad ROAS'] = avg_ad_roas
                    summary_dict['Ad Orders'] = tot_ad_orders

                summary_row = pd.DataFrame([summary_dict])
                res_df = pd.concat([res_df, summary_row], ignore_index=True)

            return res_df

        def create_daily_tracker_table(selected_month):
            if master_sales_df.empty or '_sales_dt' not in master_sales_df.columns:
                return pd.DataFrame()

            m_df = master_sales_df[master_sales_df['Month'] == selected_month].copy()
            if m_df.empty:
                return pd.DataFrame()

            ad_df = final_df[final_df['Month'] == selected_month].copy() if 'Month' in final_df.columns else pd.DataFrame()

            rows = []
            for day_i in range(1, 32):
                day_sales = m_df[m_df['Day'] == day_i]
                ov_sales = round(day_sales['_gross_amt'].sum(), 2)
                ov_orders = len(day_sales)
                
                day_spends = 0.0
                day_imp = 0.0
                day_atc = 0.0
                day_ad_sales = 0.0
                day_ad_orders = 0.0

                if not ad_df.empty and 'Date' in ad_df.columns:
                    matched_ad = ad_df[ad_df['_date_dt'].dt.day == day_i]
                    day_spends = round(matched_ad['_budget_consumed'].sum(), 2)
                    day_imp = matched_ad['_impressions'].sum()
                    day_atc = matched_ad['_atc'].sum()
                    day_ad_sales = round(matched_ad['_sales'].sum(), 2)
                    day_ad_orders = int(matched_ad['_orders'].sum())

                cpm = round((day_spends / day_imp) * 1000, 2) if day_imp > 0 else 0.0
                roas = round(day_ad_sales / day_spends, 2) if day_spends > 0 else 0.0
                tacos = round((day_spends / ov_sales) * 100, 2) if ov_sales > 0 else 0.0
                troas = round(ov_sales / day_spends, 2) if day_spends > 0 else 0.0

                rows.append({
                    'Date (Day)': f"Day {day_i}",
                    'CPM': cpm,
                    'Impressions': int(day_imp),
                    'Spends': day_spends,
                    'Ad Sales': day_ad_sales,
                    'Total Sales': ov_sales,
                    'Ad Orders': day_ad_orders,
                    'Total Orders': ov_orders,
                    'ATC': int(day_atc),
                    'TACOS': f"{tacos:.2f}%",
                    'Ad ROAS': roas,
                    'TROAS': troas
                })

            res_df = pd.DataFrame(rows)
            total_sales_sum = res_df['Total Sales'].sum()
            total_orders_sum = res_df['Total Orders'].sum()
            total_spends_sum = res_df['Spends'].sum()
            total_ad_sales_sum = res_df['Ad Sales'].sum()
            total_ad_orders_sum = res_df['Ad Orders'].sum()
            total_atc_sum = res_df['ATC'].sum()
            total_imp_sum = res_df['Impressions'].sum()
            avg_cpm = round(res_df['CPM'].mean(), 2)
            avg_roas = round(total_ad_sales_sum / total_spends_sum, 2) if total_spends_sum > 0 else 0.0
            avg_tacos = round((total_spends_sum / total_sales_sum) * 100, 2) if total_sales_sum > 0 else 0.0
            avg_troas = round(total_sales_sum / total_spends_sum, 2) if total_spends_sum > 0 else 0.0

            summary_row = pd.DataFrame([{
                'Date (Day)': 'Aggregate / Average',
                'CPM': avg_cpm,
                'Impressions': total_imp_sum,
                'Spends': total_spends_sum,
                'Ad Sales': total_ad_sales_sum,
                'Total Sales': total_sales_sum,
                'Ad Orders': total_ad_orders_sum,
                'Total Orders': total_orders_sum,
                'ATC': int(total_atc_sum),
                'TACOS': f"{avg_tacos:.2f}%",
                'Ad ROAS': avg_roas,
                'TROAS': avg_troas
            }])
            res_df = pd.concat([res_df, summary_row], ignore_index=True)
            return res_df

        def _format_dashboard_value(value, column_name=""):
            if pd.isna(value):
                return value
            col = str(column_name).strip().upper()
            is_percentage = ('%' in col) or ('ACOS' in col) or ('TACOS' in col)

            if isinstance(value, (int, np.integer)) and not isinstance(value, bool):
                return f"{int(value):,}" if not is_percentage else f"{int(value)}%"

            if isinstance(value, (float, np.floating)) and np.isfinite(value):
                if is_percentage:
                    return f"{value:.2f}%"
                if col in ['SPENDS', 'AD SALES', 'SALES', 'TOTAL SALES', 'OVERALL SALES']:
                    return f"{int(round(value)):,}"
                if float(value).is_integer():
                    return f"{int(value):,}"
                return f"{value:,.2f}"
            return value

        def format_dashboard_dataframe(df):
            out = df.copy()
            for col in out.columns:
                out[col] = out[col].map(lambda v: _format_dashboard_value(v, col))
            return out

        def render_unified_single_table(df_to_show, key_prefix="mom"):
            if df_to_show is None or df_to_show.empty:
                return
            display_df = format_dashboard_dataframe(df_to_show)
            styled_df = display_df.style.set_properties(**{
                'text-align': 'center',
                'vertical-align': 'middle'
            }).set_table_styles([
                {'selector': 'th', 'props': [('text-align', 'center'), ('vertical-align', 'middle'), ('background-color', '#E5F1F9'), ('color', '#285B7A')]},
                {'selector': 'td', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]}
            ])
            st.dataframe(
                styled_df,
                use_container_width=True,
                hide_index=False if (isinstance(df_to_show.index, pd.Index) and df_to_show.index.name) else True,
                key=f"{key_prefix}_grid"
            )

        st.markdown("### 🔍 Global Dashboard Filters")
        available_months_asc = get_calendar_months_asc(final_df)
        selected_month = st.selectbox("Select Month Across Dashboard (Excluding Comparison Tables)", ["All Months"] + available_months_asc)

        filtered_df = final_df.copy()
        if selected_month != "All Months":
            filtered_df = filtered_df[filtered_df['Month'] == selected_month]

        total_ad_sales = round(filtered_df['_sales'].sum(), 2)
        total_ad_spends = round(filtered_df['_budget_consumed'].sum(), 2)
        overall_ad_roas = round(total_ad_sales / total_ad_spends, 2) if total_ad_spends > 0 else 0.0

        has_sales_data = not master_sales_df.empty and '_sales_dt' in master_sales_df.columns

        st.markdown("### 📈 Overall Campaign Performance Dashboard (Filtered for 'staze')")
        
        if has_sales_data:
            f_sales_df = master_sales_df if selected_month == "All Months" else master_sales_df[master_sales_df['Month'] == selected_month]
            total_overall_sales = round(f_sales_df['_gross_amt'].sum(), 2)
            total_overall_orders = len(f_sales_df)
            overall_tacos = round((total_ad_spends / total_overall_sales) * 100, 2) if total_overall_sales > 0 else 0.0
            overall_troas = round(total_overall_sales / total_ad_spends, 2) if total_ad_spends > 0 else 0.0

            dcol1, dcol2, dcol3 = st.columns(3)
            with dcol1: st.metric("Ad Sales", f"₹{int(round(total_ad_sales)):,}")
            with dcol2: st.metric("Total Sales", f"₹{int(round(total_overall_sales)):,}")
            with dcol3: st.metric("Ad ROAS", f"{overall_ad_roas:.2f}x")

            dcol4, dcol5, dcol6 = st.columns(3)
            with dcol4: st.metric("TACOS", f"{overall_tacos:.2f}%")
            with dcol5: st.metric("TROAS", f"{overall_troas:.2f}")
            with dcol6: st.metric("Spends", f"₹{int(round(total_ad_spends)):,}")
        else:
            dcol1, dcol2, dcol3 = st.columns(3)
            with dcol1: st.metric("Ad Sales", f"₹{int(round(total_ad_sales)):,}")
            with dcol2: st.metric("Spends", f"₹{int(round(total_ad_spends)):,}")
            with dcol3: st.metric("Ad ROAS", f"{overall_ad_roas:.2f}x")

        st.divider()

        st.markdown("### 📑 Navigation & Performance Breakdown")
        
        main_tab1, main_tab2, main_tab3, main_tab4, main_tab5 = st.tabs([
            "📄 Raw Files Preview",
            "📌 Consolidated Master",
            "📊 Comparison Tables (MoM)",
            "📅 Daily Tracker",
            "📍 City Level"
        ])

        with main_tab1:
            st.caption("Inspect individual files or use flexible substring search (e.g. typing 'Sta' reveals all matching campaigns and products).")
            
            raw_sub_tab1, raw_sub_tab2, raw_sub_tab3 = st.tabs([
                "📂 Standard File & Sheet Inspector",
                "🎯 Ads Raw Filter & Overview",
                "📦 Overall Sales Raw Filter & Overview"
            ])

            with raw_sub_tab1:
                selected_file_name = st.selectbox("Select Uploaded File to Preview:", list(raw_files_dict.keys()), key="raw_file_sel")
                if selected_file_name:
                    sheets = raw_files_dict[selected_file_name]
                    selected_sheet = st.selectbox("Select Sheet Tab:", list(sheets.keys()), key="raw_sheet_sel")
                    if selected_sheet:
                        raw_styled = format_dashboard_dataframe(sheets[selected_sheet].head(100)).style.set_properties(**{
                            'text-align': 'center', 'vertical-align': 'middle'
                        }).set_table_styles([
                            {'selector': 'th', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]},
                            {'selector': 'td', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]}
                        ])
                        st.dataframe(raw_styled, use_container_width=True, hide_index=True)

            with raw_sub_tab2:
                st.markdown("#### 🎯 Ads Campaign Raw Data Filter & Overview")
                campaign_col_candidates = [c for c in ['Campaign Name', 'Campaign', 'CampaignId', 'Campaign ID'] if c in final_df.columns]
                if campaign_col_candidates:
                    camp_col_name = campaign_col_candidates[0]
                    
                    campaign_search_input = st.text_input("Type Campaign Name Syllable / Keyword (e.g., 'Sta', 'staze'):", "Sta", key="ads_raw_search_input")
                    
                    all_campaigns_list = sorted([str(x) for x in final_df[camp_col_name].dropna().unique()])
                    filtered_campaigns_list = [c for c in all_campaigns_list if campaign_search_input.strip().lower() in c.lower()]

                    selected_campaign_filter = st.selectbox("Select Matching Campaign:", ["All Matching Campaigns"] + filtered_campaigns_list, key="ads_raw_campaign_filter")

                    ads_filtered_raw = final_df.copy()
                    if campaign_search_input.strip():
                        ads_filtered_raw = ads_filtered_raw[ads_filtered_raw[camp_col_name].astype(str).str.lower().str.contains(campaign_search_input.strip().lower(), na=False)]

                    if selected_campaign_filter != "All Matching Campaigns":
                        ads_filtered_raw = ads_filtered_raw[ads_filtered_raw[camp_col_name].astype(str) == selected_campaign_filter]

                    f_ad_sales = round(ads_filtered_raw['_sales'].sum(), 2)
                    f_ad_spends = round(ads_filtered_raw['_budget_consumed'].sum(), 2)
                    f_ad_imp = int(ads_filtered_raw['_impressions'].sum())
                    f_ad_orders = int(ads_filtered_raw['_orders'].sum())
                    f_ad_roas = round(f_ad_sales / f_ad_spends, 2) if f_ad_spends > 0 else 0.0

                    acol1, acol2, acol3, acol4 = st.columns(4)
                    with acol1: st.metric("Filtered Ad Sales", f"₹{int(round(f_ad_sales)):,}")
                    with acol2: st.metric("Filtered Spends", f"₹{int(round(f_ad_spends)):,}")
                    with acol3: st.metric("Filtered Ad ROAS", f"{f_ad_roas:.2f}x")
                    with acol4: st.metric("Filtered Impressions", f"{f_ad_imp:,}")

                    st.markdown("##### Filtered Ads Raw Data Preview:")
                    preview_ads_raw = ads_filtered_raw.drop(columns=[c for c in ads_filtered_raw.columns if c.startswith('_')], errors='ignore')
                    
                    ads_raw_styled = format_dashboard_dataframe(preview_ads_raw.head(200)).style.set_properties(**{
                        'text-align': 'center', 'vertical-align': 'middle'
                    }).set_table_styles([
                        {'selector': 'th', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]},
                        {'selector': 'td', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]}
                    ])
                    st.dataframe(ads_raw_styled, use_container_width=True, hide_index=True)

                    ads_raw_excel_bytes = convert_dataframe_to_excel_download(preview_ads_raw, sheet_name="Ads_Raw_Filtered")
                    st.download_button(
                        label="📥 Download Filtered Ads Raw Data (.xlsx)",
                        data=ads_raw_excel_bytes,
                        file_name=f"Ads_Raw_Filtered_{campaign_search_input.strip().replace(' ', '_')[:15]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_ads_raw_filtered"
                    )
                else:
                    st.info("Campaign Name column not found in ads data.")

            with raw_sub_tab3:
                st.markdown("#### 📦 Overall Sales Raw Data Product Syllable Filter & Overview")
                if not master_sales_df.empty:
                    product_col_candidates = [c for c in ['Product Name', 'Item Name', 'product_name', 'item_name', 'Product'] if c in master_sales_df.columns]
                    if product_col_candidates:
                        prod_col_name = product_col_candidates[0]
                        product_syllable_input = st.text_input("Type Product Name Syllable / Keyword (e.g., 'Sta', 'staze'):", "Sta", key="sales_raw_syllable_input")

                        sales_filtered_raw = master_sales_df.copy()
                        if product_syllable_input.strip():
                            query_str = product_syllable_input.strip().lower()
                            sales_filtered_raw = sales_filtered_raw[sales_filtered_raw[prod_col_name].astype(str).str.lower().str.contains(query_str, na=False)]

                        f_sales_amt = round(sales_filtered_raw['_gross_amt'].sum(), 2) if '_gross_amt' in sales_filtered_raw.columns else 0.0
                        f_sales_qty = int(sales_filtered_raw['_qty'].sum()) if '_qty' in sales_filtered_raw.columns else len(sales_filtered_raw)
                        f_sales_orders = len(sales_filtered_raw)
                        f_aov = round(f_sales_amt / f_sales_orders, 2) if f_sales_orders > 0 else 0.0

                        scol1, scol2, scol3, scol4 = st.columns(4)
                        with scol1: st.metric("Filtered Sales Amount", f"₹{int(round(f_sales_amt)):,}")
                        with scol2: st.metric("Filtered Units Sold", f"{f_sales_qty:,}")
                        with scol3: st.metric("Filtered Total Orders", f"{f_sales_orders:,}")
                        with scol4: st.metric("Filtered AOV", f"₹{f_aov:,.2f}")

                        st.markdown("##### Filtered Sales Raw Data Preview:")
                        preview_sales_raw = sales_filtered_raw.drop(columns=[c for c in sales_filtered_raw.columns if c.startswith('_')], errors='ignore')

                        sales_raw_styled = format_dashboard_dataframe(preview_sales_raw.head(200)).style.set_properties(**{
                            'text-align': 'center', 'vertical-align': 'middle'
                        }).set_table_styles([
                            {'selector': 'th', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]},
                            {'selector': 'td', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]}
                        ])
                        st.dataframe(sales_raw_styled, use_container_width=True, hide_index=True)

                        sales_raw_excel_bytes = convert_dataframe_to_excel_download(preview_sales_raw, sheet_name="Overall_Sales_Raw_Filtered")
                        st.download_button(
                            label="📥 Download Filtered Sales Raw Data (.xlsx)",
                            data=sales_raw_excel_bytes,
                            file_name=f"Overall_Sales_Filtered_{product_syllable_input.strip().replace(' ', '_')[:15]}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_dl_sales_raw_filtered"
                        )
                    else:
                        st.info("Product/Item Name column not found in sales reports.")
                else:
                    st.info("Please upload Sales Reports to use this feature.")

        with main_tab2:
            st.caption("Preview the combined dataset across all uploaded files.")
            preview_clean_df = final_df.drop(columns=['_impressions', '_direct_atc', '_indirect_atc', '_atc', '_direct_orders', '_indirect_orders', '_orders', '_direct_sales', '_indirect_sales', '_sales', '_budget_consumed', '_date_dt', '_most_viewed_position', 'Ad Type Combined'], errors='ignore')
            master_styled = format_dashboard_dataframe(preview_clean_df.head(100)).style.set_properties(**{
                'text-align': 'center', 'vertical-align': 'middle'
            }).set_table_styles([
                {'selector': 'th', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]},
                {'selector': 'td', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]}
            ])
            st.dataframe(master_styled, use_container_width=True, hide_index=True)

        with main_tab3:
            st.subheader("📊 Month-on-Month Comparison Tables")
            
            mom_dict = {
                "Monthly_Summary": create_monthly_summary_table(final_df),
                "Campaign_MoM": create_mom_comparison_table(final_df, 'Campaign Name') if 'Campaign Name' in final_df.columns else None,
                "AdType_MoM": create_mom_comparison_table(final_df, 'Ad Type Combined') if 'Ad Type Combined' in final_df.columns else None,
                "Keyword_MoM": create_mom_comparison_table(final_df, next((c for c in ['Search Term', 'Keyword', 'Targeting Value'] if c in final_df.columns), None)) if next((c for c in ['Search Term', 'Keyword', 'Targeting Value'] if c in final_df.columns), None) else None,
                "Weekly_MoM": create_weekly_comparison_table_with_sales(final_df)
            }
            all_pivots_bytes = convert_all_pivots_to_excel(mom_dict)
            st.download_button(
                label="📥 Download All MoM Comparison Tables (.xlsx - One Master File with All Sub-Tabs)",
                data=all_pivots_bytes,
                file_name="All_MoM_Comparison_Tables_Combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_all_mom_pivots",
                type="primary"
            )
            st.markdown("---")

            comp_sub_tab0, comp_sub_tab1, comp_sub_tab2, comp_sub_tab3, comp_sub_tab4 = st.tabs([
                "📅 Monthly Summary",
                "🎯 Campaign Comparison",
                "📢 Ad Type Comparison",
                "🔎 Keyword Comparison",
                "📅 Weekly Comparison"
            ])

            with comp_sub_tab0:
                st.markdown("#### Monthly Summary Table")
                ms_filter_options = ["All Months"] + available_months_asc
                selected_ms_month = st.selectbox("Select Month for Monthly Summary (Expands Day-Wise 1-31):", ms_filter_options, key="ms_month_filter_sel")

                if selected_ms_month == "All Months":
                    monthly_summary_df = create_monthly_summary_table(final_df)
                    sheet_label = "Monthly_Summary"
                else:
                    monthly_summary_df = create_monthly_summary_daywise_table(selected_ms_month)
                    sheet_label = f"Monthly_Summary_{selected_ms_month}"

                if not monthly_summary_df.empty:
                    render_unified_single_table(monthly_summary_df, key_prefix="monthly_summary_tab")
                    excel_month_summary = style_and_export_pivot(monthly_summary_df, sheet_name=sheet_label)
                    st.download_button(
                        label="📥 Download Monthly Summary Table (.xlsx)",
                        data=excel_month_summary,
                        file_name=f"{sheet_label}_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_month_summary"
                    )

            with comp_sub_tab1:
                st.markdown("#### Campaign Comparison Table")
                camp_pivot = create_mom_comparison_table(final_df, 'Campaign Name') if 'Campaign Name' in final_df.columns else None
                if camp_pivot is not None and not camp_pivot.empty:
                    render_unified_single_table(camp_pivot, key_prefix="camp_pivot_tab")
                    excel_camp = style_and_export_pivot(camp_pivot, sheet_name="Campaign_MoM")
                    st.download_button(
                        label="📥 Download Campaign Comparison Table (.xlsx)",
                        data=excel_camp,
                        file_name="Campaign_MoM_Comparison_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_camp_pivot"
                    )

            with comp_sub_tab2:
                st.markdown("#### Ad Type Comparison Table")
                ad_pivot = create_mom_comparison_table(final_df, 'Ad Type Combined') if 'Ad Type Combined' in final_df.columns else None
                if ad_pivot is not None and not ad_pivot.empty:
                    render_unified_single_table(ad_pivot, key_prefix="ad_pivot_tab")
                    excel_ad = style_and_export_pivot(ad_pivot, sheet_name="AdType_MoM")
                    st.download_button(
                        label="📥 Download Ad Type Comparison Table (.xlsx)",
                        data=excel_ad,
                        file_name="Ad_Type_MoM_Comparison_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_ad_pivot"
                    )

            with comp_sub_tab3:
                st.markdown("#### Keyword Performance & Comparison")
                
                kw_col = next((c for c in ['Search Term', 'Keyword', 'Targeting Value'] if c in final_df.columns), None)
                if kw_col:
                    with st.container():
                        st.markdown("""
                        <div style="background: linear-gradient(135deg, #F3EBFB 0%, #EEF6FC 100%); border: 1px solid #D5C6E8; border-radius: 14px; padding: 16px; margin-bottom: 20px;">
                            <h4 style="margin-top: 0; color: #285B7A; font-size: 16px; text-align: center;">Advanced Filter Parameters</h4>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        kcol1, kcol2 = st.columns(2)
                        
                        with kcol1:
                            selected_kw_months = st.multiselect(
                                "Filter by Month(s):",
                                options=available_months_asc,
                                default=available_months_asc,
                                key="kw_months_multi"
                            )
                        
                        with kcol2:
                            valid_dates_df = final_df.dropna(subset=['_date_dt'])
                            if not valid_dates_df.empty:
                                min_dt = valid_dates_df['_date_dt'].min().date()
                                max_dt = valid_dates_df['_date_dt'].max().date()
                                
                                kw_date_selection = st.date_input(
                                    "Filter by Date (Single Day or Date Range):",
                                    value=[min_dt, max_dt],
                                    min_value=min_dt,
                                    max_value=max_dt,
                                    key="kw_date_flexible"
                                )
                            else:
                                kw_date_selection = None

                    kw_filtered_df = final_df.copy()
                    if selected_kw_months:
                        kw_filtered_df = kw_filtered_df[kw_filtered_df['Month'].isin(selected_kw_months)]
                    
                    if kw_date_selection:
                        if isinstance(kw_date_selection, (tuple, list)):
                            if len(kw_date_selection) == 2:
                                start_d, end_d = kw_date_selection
                                kw_filtered_df = kw_filtered_df[
                                    kw_filtered_df['_date_dt'].notna() &
                                    (kw_filtered_df['_date_dt'].dt.date >= start_d) &
                                    (kw_filtered_df['_date_dt'].dt.date <= end_d)
                                ]
                            elif len(kw_date_selection) == 1:
                                single_d = kw_date_selection[0]
                                kw_filtered_df = kw_filtered_df[
                                    kw_filtered_df['_date_dt'].notna() &
                                    (kw_filtered_df['_date_dt'].dt.date == single_d)
                                ]
                        else:
                            single_d = kw_date_selection
                            kw_filtered_df = kw_filtered_df[
                                kw_filtered_df['_date_dt'].notna() &
                                (kw_filtered_df['_date_dt'].dt.date == single_d)
                            ]

                    kw_pivot = create_mom_comparison_table(kw_filtered_df, kw_col) if not kw_filtered_df.empty else None
                    if kw_pivot is not None and not kw_pivot.empty:
                        render_unified_single_table(kw_pivot, key_prefix="kw_pivot_tab")
                        excel_kw = style_and_export_pivot(kw_pivot, sheet_name="Keyword_MoM")
                        st.download_button(
                            label="📥 Download Keyword Comparison Table (.xlsx)",
                            data=excel_kw,
                            file_name="Keyword_MoM_Comparison_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_dl_kw_pivot"
                        )
                    else:
                        st.info("No keyword data available for the selected filters.")
                else:
                    st.info("Keyword column not found in uploaded ad files.")

            with comp_sub_tab4:
                st.markdown("#### Weekly Comparison Table")
                weekly_comp_df = create_weekly_comparison_table_with_sales(final_df)
                if not weekly_comp_df.empty:
                    render_unified_single_table(weekly_comp_df, key_prefix="weekly_comp_tab")
                    excel_week = style_and_export_pivot(weekly_comp_df, sheet_name="Weekly_MoM")
                    st.download_button(
                        label="📥 Download Weekly Comparison Table (.xlsx)",
                        data=excel_week,
                        file_name="Weekly_MoM_Comparison_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_week_pivot"
                    )

        with main_tab4:
            st.subheader("📅 Daily Tracker")
            if has_sales_data:
                tracker_month = st.selectbox("Select Month for Daily Tracker:", available_months_asc, key="tracker_month_sel")
                if tracker_month:
                    daily_df = create_daily_tracker_table(tracker_month)
                    if not daily_df.empty:
                        render_unified_single_table(daily_df, key_prefix="daily_tracker_tab")
                        excel_daily = style_and_export_pivot(daily_df, sheet_name="Daily_Tracker")
                        st.download_button(
                            label="📥 Download Daily Tracker Table (.xlsx)",
                            data=excel_daily,
                            file_name=f"Daily_Tracker_{tracker_month}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_dl_daily_tracker"
                        )
                    else:
                        st.info("No sales data available for the selected month.")
            else:
                st.info("Please upload a Sales Report to view the Daily Tracker.")

        with main_tab5:
            st.subheader("📍 City Level Overview & Product Performance")
            if has_sales_data:
                sales_months_avail = ["All Months"] + sorted(master_sales_df['Month'].dropna().unique().tolist(), key=lambda x: MONTH_ORDER_ASC.get(str(x).strip().upper(), 99))
                sel_city_month = st.selectbox("Select Month for City View:", sales_months_avail, key="city_month_filter")

                city_filtered_sales = master_sales_df.copy()
                if sel_city_month != "All Months":
                    city_filtered_sales = city_filtered_sales[city_filtered_sales['Month'] == sel_city_month]

                state_agg = city_filtered_sales.groupby('Supply State').agg(
                    Total_Sales=('_gross_amt', 'sum'),
                    Total_Orders=('Supply State', 'count')
                ).reset_index()

                all_states_total_sales = city_filtered_sales['_gross_amt'].sum()
                all_states_total_orders = len(city_filtered_sales)

                state_options = [f"🌐 All States  •  Sales: ₹{int(round(all_states_total_sales)):,}  •  Orders: {all_states_total_orders:,}"]
                state_mapping = {state_options[0]: "All States"}

                for _, row in state_agg.sort_values(by='Total_Sales', ascending=False).iterrows():
                    st_name = str(row['Supply State']).title()
                    st_sales = row['Total_Sales']
                    st_orders = int(row['Total_Orders'])
                    display_label = f"📍 {st_name}  •  Sales: ₹{int(round(st_sales)):,}  •  Orders: {st_orders:,}"
                    state_options.append(display_label)
                    state_mapping[display_label] = row['Supply State']

                sel_state_label = st.selectbox("Select Supply State:", state_options, key="city_state_sel")
                sel_state = state_mapping.get(sel_state_label, "All States")

                if sel_state != "All States" and 'Supply State' in city_filtered_sales.columns:
                    city_filtered_sales = city_filtered_sales[city_filtered_sales['Supply State'] == sel_state]

                city_agg = city_filtered_sales.groupby('Supply City').agg(
                    Total_Sales=('_gross_amt', 'sum'),
                    Total_Orders=('Supply City', 'count')
                ).reset_index()

                all_cities_total_sales = city_filtered_sales['_gross_amt'].sum()
                all_cities_total_orders = len(city_filtered_sales)

                city_options = [f"🌐 All Cities  •  Sales: ₹{int(round(all_cities_total_sales)):,}  •  Orders: {all_cities_total_orders:,}"]
                city_mapping = {city_options[0]: "All Cities"}

                for _, row in city_agg.sort_values(by='Total_Sales', ascending=False).iterrows():
                    c_name = str(row['Supply City']).title()
                    c_sales = row['Total_Sales']
                    c_orders = int(row['Total_Orders'])
                    display_label = f"🏙️ {c_name}  •  Sales: ₹{int(round(c_sales)):,}  •  Orders: {c_orders:,}"
                    city_options.append(display_label)
                    city_mapping[display_label] = row['Supply City']

                sel_city_label = st.selectbox("Select Supply City:", city_options, key="city_city_sel")
                sel_city = city_mapping.get(sel_city_label, "All Cities")

                if sel_city != "All Cities" and 'Supply City' in city_filtered_sales.columns:
                    city_filtered_sales = city_filtered_sales[city_filtered_sales['Supply City'] == sel_city]

                prod_col_name = next((c for c in ['Product Name', 'Item Name', 'product_name', 'item_name', 'Product'] if c in city_filtered_sales.columns), None)
                qty_col_name = next((c for c in ['Quantity', 'Qty', 'Units', '_qty'] if c in city_filtered_sales.columns), None)
                amt_col_name = next((c for c in ['_gross_amt', 'MRP (Rs)', 'mrp', 'MRP'] if c in city_filtered_sales.columns), None)

                if prod_col_name and qty_col_name and amt_col_name:
                    prod_agg = city_filtered_sales.groupby(prod_col_name).agg(
                        Units_Sold=(qty_col_name, 'sum'),
                        Total_Sales=(amt_col_name, 'sum')
                    ).reset_index().sort_values(by='Total_Sales', ascending=False)

                    prod_agg = prod_agg.rename(columns={
                        prod_col_name: 'PRODUCT NAME',
                        'Units_Sold': 'UNITS SOLD',
                        'Total_Sales': 'TOTAL SALES AMOUNT (₹)'
                    })
                    prod_styled = format_dashboard_dataframe(prod_agg).style.set_properties(**{
                        'text-align': 'center', 'vertical-align': 'middle'
                    }).set_table_styles([
                        {'selector': 'th', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]},
                        {'selector': 'td', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]}
                    ])
                    st.dataframe(prod_styled, use_container_width=True, hide_index=True)
                    
                    st.markdown("---")
                    st.markdown("#### 🥧 Top 5 States & Top 5 Cities Contribution to Overall Sales & Orders")
                    
                    state_contrib = state_agg[state_agg['Total_Sales'] > 0].sort_values(by='Total_Sales', ascending=False).head(5).copy()
                    state_contrib['Supply State'] = state_contrib['Supply State'].astype(str).str.title()
                    state_contrib['AOV'] = state_contrib.apply(lambda r: round(r['Total_Sales'] / r['Total_Orders'], 2) if r['Total_Orders'] > 0 else 0.0, axis=1)

                    full_city_agg = master_sales_df.groupby('Supply City').agg(
                        Total_Sales=('_gross_amt', 'sum'),
                        Total_Orders=('Supply City', 'count')
                    ).reset_index()
                    city_contrib = full_city_agg[full_city_agg['Total_Sales'] > 0].sort_values(by='Total_Sales', ascending=False).head(5).copy()
                    city_contrib['Supply City'] = city_contrib['Supply City'].astype(str).str.title()
                    city_contrib['AOV'] = city_contrib.apply(lambda r: round(r['Total_Sales'] / r['Total_Orders'], 2) if r['Total_Orders'] > 0 else 0.0, axis=1)

                    col_p1, col_p2 = st.columns(2)
                    
                    with col_p1:
                        if not state_contrib.empty:
                            fig_state_pie = px.pie(
                                state_contrib,
                                names='Supply State',
                                values='Total_Sales',
                                title=f'Top 5 States Contribution (₹)',
                                color_discrete_sequence=['#285B7A', '#4A90C2', '#6AA6CF', '#8ABCE3', '#B4D6F0']
                            )
                            fig_state_pie.update_traces(textposition='inside', textinfo='percent+label')
                            fig_state_pie.update_layout(template='plotly_white', height=400)
                            st.plotly_chart(fig_state_pie, use_container_width=True)

                    with col_p2:
                        if not city_contrib.empty:
                            fig_city_pie = px.pie(
                                city_contrib,
                                names='Supply City',
                                values='Total_Sales',
                                title=f'Top 5 Cities Contribution (₹)',
                                color_discrete_sequence=['#0F4C81', '#1D63B8', '#3882E6', '#60A5FA', '#93C5FD']
                            )
                            fig_city_pie.update_traces(textposition='inside', textinfo='percent+label')
                            fig_city_pie.update_layout(template='plotly_white', height=400)
                            st.plotly_chart(fig_city_pie, use_container_width=True)
                else:
                    st.info("Required product/sales columns not found in sales report.")
            else:
                st.info("Please upload a Sales Report to view City Level Overview.")

        st.markdown("---")
        st.subheader("💾 Download Consolidated Master Workbook")
        
        ads_dump_df = final_df.copy()
        
        def build_relative_month_map(dt_series_list):
            all_valid_dates = pd.concat(dt_series_list).dropna().unique()
            if len(all_valid_dates) == 0:
                return {}
            
            dt_df = pd.DataFrame({'dt': pd.to_datetime(all_valid_dates)})
            dt_df['year_month'] = dt_df['dt'].dt.to_period('M')
            sorted_unique_ym = sorted(dt_df['year_month'].unique())
            
            month_map = {}
            for idx, ym in enumerate(sorted_unique_ym, start=1):
                m_short = ym.strftime('%b')
                y_two = ym.strftime('%y')
                month_map[ym] = f"{idx}. {m_short}'{y_two}"
            return month_map

        all_dt_series = []
        if '_date_dt' in ads_dump_df.columns:
            all_dt_series.append(ads_dump_df['_date_dt'])
        if not master_sales_df.empty and '_sales_dt' in master_sales_df.columns:
            all_dt_series.append(master_sales_df['_sales_dt'])

        relative_month_mapping = build_relative_month_map(all_dt_series) if all_dt_series else {}

        def get_duration_month_relative(dt):
            if pd.isna(dt): return np.nan
            ym = dt.to_period('M')
            return relative_month_mapping.get(ym, np.nan)

        def get_duration_week(dt):
            if pd.isna(dt): return np.nan
            iso_year, iso_week, iso_day = dt.isocalendar()
            start_of_week = dt - pd.Timedelta(days=iso_day - 1)
            end_of_week = start_of_week + pd.Timedelta(days=6)
            return f"Week {iso_week} ({start_of_week.strftime('%b %d')} to {end_of_week.strftime('%b %d')})"

        if '_date_dt' in ads_dump_df.columns:
            ads_dump_df['Duration (Month)'] = ads_dump_df['_date_dt'].apply(get_duration_month_relative)
            ads_dump_df['Duration (Week)'] = ads_dump_df['_date_dt'].apply(get_duration_week)
            ads_dump_df['date_ist'] = ads_dump_df['_date_dt'].dt.strftime('%Y/%m/%d')
        else:
            ads_dump_df['Duration (Month)'] = np.nan
            ads_dump_df['Duration (Week)'] = np.nan
            ads_dump_df['date_ist'] = np.nan

        def safe_get_col(df_source, possible_names):
            for name in possible_names:
                if name in df_source.columns:
                    return df_source[name]
            return np.nan

        ads_export_mapping = {
            'Duration (Month)': ads_dump_df['Duration (Month)'],
            'Duration (Week)': ads_dump_df['Duration (Week)'],
            'Manufacturer': safe_get_col(ads_dump_df, ['Manufacturer', 'Manufacturer Name', 'manufacturer_name']),
            'date_ist': ads_dump_df['date_ist'],
            'Campaign ID': safe_get_col(ads_dump_df, ['Campaign ID', 'CampaignId']),
            'Campaign Name': safe_get_col(ads_dump_df, ['Campaign Name', 'Campaign']),
            'City': safe_get_col(ads_dump_df, ['City', 'Supply City', 'city_name']),
            'Ad type': safe_get_col(ads_dump_df, ['Ad type', 'Ad Type', 'Match Type', 'Ad Type Combined']),
            'Targeting Type': safe_get_col(ads_dump_df, ['Targeting Type', 'Targeting']),
            'Franchise': safe_get_col(ads_dump_df, ['Franchise', 'Brand']),
            'Category': safe_get_col(ads_dump_df, ['Category', 'category']),
            'Sub Category': safe_get_col(ads_dump_df, ['Sub Category', 'Subcategory']),
            'Keyword': safe_get_col(ads_dump_df, ['Keyword', 'Search Term', 'Targeting Value']),
            'Most Viewed Position': ads_dump_df['_most_viewed_position'],
            'Match Type': safe_get_col(ads_dump_df, ['Match Type', 'Targeting Type']),
            'Pacing Type': safe_get_col(ads_dump_df, ['Pacing Type']),
            'CPM': safe_get_col(ads_dump_df, ['CPM']),
            'Total Budget': safe_get_col(ads_dump_df, ['Total Budget', 'Budget']),
            'Impressions': ads_dump_df['_impressions'],
            'Direct ATC': ads_dump_df['_direct_atc'],
            'Indirect ATC': ads_dump_df['_indirect_atc'],
            'New Users': safe_get_col(ads_dump_df, ['New Users']),
            'Direct Sales': ads_dump_df['_direct_sales'],
            'Indirect Sales': ads_dump_df['_indirect_sales'],
            'Direct Quantities Sold': ads_dump_df['_direct_orders'],
            'Indirect Quantities Sold': ads_dump_df['_indirect_orders'],
            'Estimated Budget Consumed': ads_dump_df['_budget_consumed'],
            'Direct RoAS': safe_get_col(ads_dump_df, ['Direct RoAS', 'Direct ROAS']),
            'Total RoAS': safe_get_col(ads_dump_df, ['Total RoAS', 'Total ROAS', 'ROAS']),
            'latest on hold timestamp': safe_get_col(ads_dump_df, ['latest on hold timestamp'])
        }

        final_ads_dump_export = pd.DataFrame(ads_export_mapping)

        if not master_sales_df.empty:
            s_dump_df = master_sales_df.copy()
            if '_sales_dt' in s_dump_df.columns:
                s_dump_df['Duration (Month)'] = s_dump_df['_sales_dt'].apply(get_duration_month_relative)
                s_dump_df['Duration (Week)'] = s_dump_df['_sales_dt'].apply(get_duration_week)
                s_dump_df['date'] = s_dump_df['_sales_dt'].dt.strftime('%Y/%m/%d')
            else:
                s_dump_df['Duration (Month)'] = np.nan
                s_dump_df['Duration (Week)'] = np.nan
                s_dump_df['date'] = np.nan

            overall_export_mapping = {
                'item_id': safe_get_col(s_dump_df, ['Item Id', 'item_id', 'SKU ID', 'Product ID']),
                'item_name': s_dump_df['_product_name'],
                'Franchise': safe_get_col(s_dump_df, ['Brand Name', 'Franchise', 'Brand']),
                'manufacturer_id': safe_get_col(s_dump_df, ['manufacturer_id', 'Manufacturer ID']),
                'manufacturer_name': safe_get_col(s_dump_df, ['manufacturer_name', 'Manufacturer', 'Manufacturer Name']),
                'city_id': safe_get_col(s_dump_df, ['city_id', 'City ID']),
                'city_name': safe_get_col(s_dump_df, ['Supply City', 'city_name', 'City']),
                'category': safe_get_col(s_dump_df, ['Business Category', 'category', 'Category']),
                'date': s_dump_df['date'],
                'Duration (Month)': s_dump_df['Duration (Month)'],
                'Duration (Week)': s_dump_df['Duration (Week)'],
                'qty_sold': s_dump_df['_qty'],
                'mrp': safe_get_col(s_dump_df, ['MRP (Rs)', 'mrp', 'MRP', 'Selling Price (Rs)']),
                'Region': safe_get_col(s_dump_df, ['Supply State', 'Region', 'State'])
            }
            final_overall_dump_export = pd.DataFrame(overall_export_mapping)
        else:
            final_overall_dump_export = pd.DataFrame(columns=[
                'item_id', 'item_name', 'Franchise', 'manufacturer_id', 'manufacturer_name',
                'city_id', 'city_name', 'category', 'date', 'Duration (Month)',
                'Duration (Week)', 'qty_sold', 'mrp', 'Region'
            ])

        buffer_multi = io.BytesIO()
        with pd.ExcelWriter(buffer_multi, engine='openpyxl') as writer:
            final_ads_dump_export.to_excel(writer, sheet_name='ADs Dump', index=False)
            final_overall_dump_export.to_excel(writer, sheet_name='Overall Dump', index=False)

            for raw_tab_name, df_list in consolidated_raw_tabs.items():
                combined_raw_tab_df = pd.concat(df_list, ignore_index=True)
                clean_sheet_name = raw_tab_name[:31]
                combined_raw_tab_df.to_excel(writer, sheet_name=clean_sheet_name, index=False)

        buffer_multi.seek(0)
        st.download_button(
            label="📥 Download Complete Excel Workbook (Initial Dumps + Raw Tabs)",
            data=buffer_multi,
            file_name="STAZE_Blinkit_Consolidated_Master_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    else:
        st.info("Please upload valid ad campaign reports to start analysis.")
else:
    st.info("Awaiting file uploads. Please upload your ad campaign spreadsheets and sales reports.")
