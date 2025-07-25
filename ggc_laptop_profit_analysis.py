import streamlit as st
import pandas as pd
import plotly.express as px

# File path
EXCEL_PATH = r"GGC_Laptop_Profit_Sheet.xlsx"

# Sheet and column mapping (aligned with user-provided headers)
# NOTE: If your actual Excel file uses slightly different header spellings, feel free to
# add the alternative spellings here as additional list entries so they can be picked up too.
data_config = {
    "Purchase": [
        "Date of Purchase", "Month", "Branch", "Vendor Name", "Model No.", "Serial No", "Part Code", "Bill No.", "Price (Excuding GST)", "Extra Support Committed", "HP ST Support"
    ],
    "Orginal Sale": [
        "Date of Sales", "Month", "Branch", "Party Name", "Model No.", "Bill No.", "Part Code", "Price (Excuding GST)", "HP SO Support"
    ],
    "Tally-II Sale": [
        "Month", "Branch", "Party Name", "Model No.", "Serial No.", "Part Code", "HP SO Support"
    ],
    "Return": [
        "Date of Purchase", "Month", "Branch", "Party Name", "Model No.", "Serial No.", "Return date", "SRN No", "Return Value", "Actual Profit/Los"
    ]
}

def load_data():
    dfs = {}
    xls = pd.ExcelFile(EXCEL_PATH)
    for sheet, cols in data_config.items():
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, usecols=lambda x: x in cols)
            dfs[sheet] = df
    return dfs

def main():
    st.title("GGC Laptop Profit Sheet Data Analysis")
    st.write("---")
    dfs = load_data()
    sheet = st.sidebar.selectbox("Select Data Sheet", list(dfs.keys()))
    df = dfs[sheet]

    # Filter options
    filter_cols = [col for col in df.columns if df[col].dtype == 'object' or 'Date' in col or 'Month' in col]
    with st.sidebar:
        st.header("Filter Options")
        filters = {}
        # Standard filters
        for col in filter_cols:
            unique_vals = df[col].dropna().unique()
            if len(unique_vals) > 1 and len(unique_vals) < 100:
                selected = st.multiselect(f"Filter by {col}", options=unique_vals, default=unique_vals)
                filters[col] = selected
        # Vendor Name filter (case/space insensitive)
        vendor_col = next((c for c in df.columns if c.strip().lower() == 'vendor name'), None)
        if vendor_col:
            vendor_vals = df[vendor_col].dropna().unique()
            if len(vendor_vals) > 1 and len(vendor_vals) < 100:
                selected_vendor = st.multiselect("Filter by Vendor Name", options=vendor_vals, default=vendor_vals)
                df = df[df[vendor_col].isin(selected_vendor)]
        # Party Name filter (case/space insensitive)
        party_col = next((c for c in df.columns if c.strip().lower() == 'party name'), None)
        if party_col:
            party_vals = df[party_col].dropna().unique()
            if len(party_vals) > 1 and len(party_vals) < 100:
                selected_party = st.multiselect("Filter by Party Name", options=party_vals, default=party_vals)
                df = df[df[party_col].isin(selected_party)]
        # Apply standard filters
        for col, selected in filters.items():
            if selected:
                df = df[df[col].isin(selected)]

    st.subheader(f"Data from {sheet} Sheet")
    st.dataframe(df)

    # Pie Chart Options
    st.write("---")
    st.header("Pie Chart Visualization")
    pie_col = st.selectbox("Select column for Pie Chart", [col for col in df.columns if df[col].dtype == 'object' or 'Month' in col or 'Branch' in col])
    if pie_col:
        pie_data = df[pie_col].value_counts().reset_index()
        pie_data.columns = [pie_col, 'Count']
        fig = px.pie(pie_data, names=pie_col, values='Count', title=f"Distribution by {pie_col}")
        st.plotly_chart(fig)

    # Download filtered data
    st.write("---")
    st.download_button(
        label="Download Filtered Data as CSV",
        data=df.to_csv(index=False).encode('utf-8'),
        file_name=f"{sheet}_filtered.csv",
        mime='text/csv'
    )

if __name__ == "__main__":
    main()
