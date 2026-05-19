import streamlit as st
import xlrd

from parser import detect_statement_type, extract_last4, parse_cc_statement, parse_bank_statement
from converter import (
    to_zoho_cc_xlsx,
    to_zoho_bank_xlsx,
    transactions_to_dataframe_cc,
    transactions_to_dataframe_bank,
)

st.set_page_config(
    page_title="Statement Converter for Zoho Books",
    page_icon="🏦",
    layout="centered",
)

st.title("Statement Converter for Zoho Books")
st.markdown(
    "Upload one or more HDFC Bank statements (credit card or account) and download "
    "Zoho Books-compatible Excel files ready for import."
)

uploaded_files = st.file_uploader(
    "Upload bank statements (.xls)",
    type=["xls"],
    accept_multiple_files=True,
    help="Supports HDFC credit card and account statements in XLS format.",
)

if not uploaded_files:
    st.stop()

for uploaded_file in uploaded_files:
    st.divider()
    st.subheader(uploaded_file.name)

    file_bytes = uploaded_file.read()

    try:
        workbook = xlrd.open_workbook(file_contents=file_bytes)
    except Exception as e:
        st.error(f"Could not open file: {e}")
        continue

    statement_type = detect_statement_type(workbook)

    if statement_type == "cc":
        st.success("Detected: **Credit Card Statement**")
        type_label = "credit card"
        type_prefix = "cc"
    elif statement_type == "bank":
        st.success("Detected: **Bank Account Statement**")
        type_label = "bank account"
        type_prefix = "bank"
    else:
        st.error(
            "Could not identify as an HDFC credit card or bank account statement. "
            "Please check that this is a supported HDFC statement."
        )
        continue

    try:
        last4 = extract_last4(workbook, statement_type)

        if statement_type == "cc":
            transactions = parse_cc_statement(workbook)
            df = transactions_to_dataframe_cc(transactions)
            xlsx_bytes = to_zoho_cc_xlsx(transactions)
        else:
            transactions = parse_bank_statement(workbook)
            df = transactions_to_dataframe_bank(transactions)
            xlsx_bytes = to_zoho_bank_xlsx(transactions)
    except Exception as e:
        st.error(f"Error parsing statement: {e}")
        continue

    if not transactions:
        st.warning("No transactions found in this file.")
        continue

    dates = [tx["date"] for tx in transactions]
    start_date = min(dates).strftime("%d%m%y")
    end_date = max(dates).strftime("%d%m%y")
    output_filename = f"{type_prefix}_{last4}_{start_date}_{end_date}.xlsx"

    st.markdown(
        f"**{len(transactions)} transactions** &nbsp;|&nbsp; "
        f"Account/Card: `...{last4}` &nbsp;|&nbsp; "
        f"{min(dates).strftime('%d %b %Y')} → {max(dates).strftime('%d %b %Y')}",
        unsafe_allow_html=True,
    )

    st.caption(f"Output filename: `{output_filename}`")

    st.dataframe(df.head(10), use_container_width=True)

    st.download_button(
        label=f"Download {output_filename}",
        data=xlsx_bytes,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=output_filename,
    )

    with st.expander("Show all transactions"):
        st.dataframe(df, use_container_width=True)
