import streamlit as st
import pandas as pd
from io import StringIO

def generate_iif(df):
    output = StringIO()

    # IIF Headers
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\n")
    output.write("!ENDTRNS\n")  # Required third line

    # Remove voided transactions
    df = df[~df['Type'].str.lower().str.contains("void")]

    # Group by Bill#
    for bill_no, bill_df in df.groupby('Bill#'):
        raw_date = bill_df['Trans Date'].iloc[0]
        cleaned_date = raw_date.replace('.', ':', 2)  # Fix malformed datetime
        trans_date = pd.to_datetime(cleaned_date, errors='coerce')

        if pd.isna(trans_date):
            continue  # skip malformed dates

        date_str = trans_date.strftime('%m/%d/%Y')
        day = trans_date.day
        till = bill_df['Till#'].iloc[0]
        customer = "Walk In"
        memo = f"Till {till} Bill {bill_no}"
        docnum = f"INV{day:02d}/{int(bill_no):02d}"

        # Determine transaction type
        is_return = all(bill_df['Type'].str.lower() == 'return')
        trnstype = "CREDIT MEMO" if is_return else "INVOICE"

        # Total amount reversed
        total_amount = -bill_df['Total'].sum()

        # TRNS Header
        output.write(f"TRNS\t{trnstype}\t{date_str}\tAccounts Receivable\t{customer}\t{memo}\t{total_amount:.2f}\t{docnum}\n")

        # SPL Lines
        for _, row in bill_df.iterrows():
            desc = row['Description']
            item_code = row['Code']
            amount = -row['Total']
            qty = row.get('Qty', 1)
            output.write(f"SPL\t{trnstype}\t{date_str}\tSales Revenue\t{customer}\t{desc} ({item_code})\t{amount:.2f}\t{qty}\n")

        # Transaction end
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
        df.columns = df.columns.str.strip()  # remove spaces around header names

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
