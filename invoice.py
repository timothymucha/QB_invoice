import streamlit as st
import pandas as pd
from io import StringIO

def remove_void_pairs(df):
    """
    Cancels out sale/void pairs considering quantities.
    Keeps remaining sales if only part of the quantity is voided.
    """
    df['key'] = df['Bill#'].astype(str) + "_" + df['Code'].astype(str)
    cleaned_rows = []

    for key, group in df.groupby('key'):
        sales = group[group['Type'].str.lower() == 'sale'].copy()
        voids = group[group['Type'].str.lower() == 'void'].copy()

        total_sales_qty = sales['Qty'].sum()
        total_void_qty = voids['Qty'].sum()

        net_qty = total_sales_qty - total_void_qty

        if net_qty > 0:
            # Build a single representative row from first sale entry
            row = sales.iloc[0].copy()
            row['Qty'] = net_qty

            # Adjust the total proportionally if 'Total' is available
            if 'Total' in row:
                price_per_unit = sales['Total'].sum() / total_sales_qty if total_sales_qty else 0
                row['Total'] = price_per_unit * net_qty

            cleaned_rows.append(row)

    return pd.DataFrame(cleaned_rows) if cleaned_rows else pd.DataFrame()

def generate_iif(df):
    output = StringIO()

    # IIF Headers
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
    output.write("!ENDTRNS\n")

    df.columns = df.columns.str.strip()
    df['Type'] = df['Type'].astype(str).str.strip().str.lower()

    df = remove_void_pairs(df)

    for bill_no, bill_df in df.groupby('Bill#'):
        raw_date = str(bill_df['Trans Date'].iloc[0])
        cleaned_date = raw_date.replace('.', ':', 2)
        trans_date = pd.to_datetime(cleaned_date, errors='coerce')
        if pd.isna(trans_date):
            continue

        date_str = trans_date.strftime('%m/%d/%Y')
        day = trans_date.day
        till = bill_df['Till#'].iloc[0]
        customer = "Walk In"
        memo = f"Till {till} Bill {bill_no}"
        docnum = f"INV{day:02d}/{int(bill_no):04d}"

        trnstype = "INVOICE"
        total_amount = bill_df['Total'].sum()

        output.write(f"TRNS\t{trnstype}\t{date_str}\tAccounts Receivable\t{customer}\t{memo}\t{total_amount:.2f}\t{docnum}\n")

        for _, row in bill_df.iterrows():
            desc = str(row.get('Description', '')).strip()
            item_code = str(row.get('Code', '')).strip()
            qty = row.get('Qty', 1)
            amount = -float(row['Total'])  # QuickBooks needs SPL amount negative
            invitem = desc[:31] if len(desc) > 31 else desc
            memo_line = f"{desc} {item_code}".strip()

            output.write(
                f"SPL\t{trnstype}\t{date_str}\tSales Revenue\t{customer}\t{memo_line}\t{amount:.2f}\t{qty}\t{invitem}\n"
            )

        output.write("ENDTRNS\n")

    return output.getvalue()


# Streamlit UI
st.set_page_config(page_title="QuickBooks IIF Generator", layout="wide")
st.title("🧾 QuickBooks IIF Generator from Excel")

uploaded_file = st.file_uploader("📤 Upload your POS Excel file (.xlsx only)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ File uploaded and read successfully.")
        df.columns = df.columns.str.strip()

        st.subheader("📋 Data Preview")
        st.dataframe(df)

        if st.button("🚀 Generate QuickBooks IIF"):
            iif_output = generate_iif(df)

            st.download_button(
                label="📥 Download .IIF File",
                data=iif_output,
                file_name="qb_sales_import.iif",
                mime="text/plain"
            )
    except Exception as e:
        st.error(f"❌ Failed to process file: {e}")
