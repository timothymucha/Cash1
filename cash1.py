import streamlit as st
import pandas as pd
from io import StringIO

st.set_page_config(page_title="Cash Sales to IIF Converter", layout="wide")
st.title("📄 Convert Excel Cash Sales to QuickBooks IIF")

uploaded_file = st.file_uploader("📤 Upload Cash Sales Excel File", type=["xlsx"])

def cut_after_cash_total(df_raw):
    """
    Find the first row that contains 'Total Amount for' anywhere in the row,
    then keep everything BEFORE that row (exclude the total and anything after).
    """
    # True where any cell in the row contains the phrase
    mask_total = df_raw.apply(
        lambda r: r.astype(str).str.contains(r"\bTotal\s+Amount\s+for\b", case=False, regex=True, na=False).any(),
        axis=1
    )
    if mask_total.any():
        stop_idx = mask_total[mask_total].index[0]  # first occurrence
        return df_raw.iloc[:stop_idx]
    return df_raw

def generate_iif(df):
    output = StringIO()

    # IIF Headers
    output.write("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tDOCNUM\n")
    output.write("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tAMOUNT\tQNTY\tINVITEM\n")
    output.write("!ENDTRNS\n")

    for _, row in df.iterrows():
        trns_date = row['Date'].strftime("%m/%d/%Y")
        amount = float(str(row['Amount']).replace(",", ""))  # handle 1,160.00, etc.
        docnum = str(row['Bill No.'])
        till = str(row['Till No'])
        memo = f"Till {till} | Invoice {docnum}"

        # Write the transaction
        output.write(f"TRNS\tPAYMENT\t{trns_date}\tCash in Drawer\tWalk In\t{memo}\t{amount}\t{docnum}\n")
        output.write(f"SPL\tPAYMENT\t{trns_date}\tAccounts Receivable\tWalk In\t{memo}\t{-amount}\t\t\n")
        output.write("ENDTRNS\n")

    return output.getvalue()

if uploaded_file:
    try:
        # Read raw data from row 17 onwards (skip frozen header)
        df_raw = pd.read_excel(uploaded_file, header=None, skiprows=16)

        # Cut everything after the 'Total Amount for ...' row
        df_raw = cut_after_cash_total(df_raw)

        # Rename relevant columns (keep using your fixed positions)
        df_raw.rename(columns={
            4: "Till No",
            9: "Date",
            15: "Bill No.",
            25: "Amount"
        }, inplace=True)

        # Keep only the relevant columns
        df = df_raw[["Till No", "Date", "Bill No.", "Amount"]].copy()

        # Drop rows where essential fields are missing
        df.dropna(subset=["Date", "Amount"], inplace=True)

        # Normalize whitespace before parsing dates (handles double spaces)
        df["Date"] = (
            df["Date"].astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        # Convert Date (example format: '31-Jul-2025  10.24.49 AM')
        df["Date"] = pd.to_datetime(df["Date"], format="%d-%b-%Y %I.%M.%S %p", errors="coerce")
        df.dropna(subset=["Date"], inplace=True)

        st.subheader("🧾 Preview: First 10 Cleaned Cash Sales")
        st.dataframe(df.head(10))

        # Generate IIF
        iif_text = generate_iif(df)

        st.subheader("📥 Download .IIF File")
        st.download_button("Download IIF", iif_text, file_name="cash_sales.iif", mime="text/plain")

    except Exception as e:
        st.error(f"❌ Failed to process file: {e}")
else:
    st.info("👆 Please upload a cash sales report to begin.")
