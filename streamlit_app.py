import streamlit as st
import sys
import io
import contextlib
import tempfile
from pathlib import Path
import pandas as pd

st.set_page_config(page_title="Bank Statement to Excel", page_icon="🏦", layout="centered")

st.title("🏦 Bank Statement PDF to Excel")
st.write(
    "Upload a bank statement PDF (digital or scanned). "
    "The tool auto-detects the bank and format, extracts transactions, "
    "and flags any row where the balance math doesn't reconcile."
)

# Import the parser as a module (not via subprocess) so it always runs in
# the exact same environment as this app -- avoids environment-mismatch
# issues that can occur when spawning a subprocess in some hosting setups.
try:
    import bank_statement_parser_v6 as parser
    PARSER_IMPORT_ERROR = None
except Exception as e:
    parser = None
    PARSER_IMPORT_ERROR = e

if PARSER_IMPORT_ERROR:
    st.error(f"Could not import bank_statement_parser_v6.py: {PARSER_IMPORT_ERROR}")
    st.stop()

uploaded_file = st.file_uploader("Choose a PDF", type=["pdf"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name

    log_buffer = io.StringIO()
    transactions = None
    run_error = None

    with st.spinner("Processing... this can take a while for scanned PDFs (OCR)."):
        try:
            with contextlib.redirect_stdout(log_buffer):
                transactions = parser.extract(tmp_path)
        except Exception as e:
            run_error = e

    with st.expander("Log output"):
        st.code(log_buffer.getvalue() or "(no output)", language="text")
        if run_error:
            st.exception(run_error)

    if run_error:
        st.error("The parser raised an error. See the log above for details.")
    elif not transactions:
        st.warning("0 transactions extracted — this bank's format may not be configured yet in BANK_CONFIGS.")
    else:
        df = pd.DataFrame(transactions)
        df.insert(0, "row", range(1, len(df) + 1))
        df_out = df[["row", "date", "narration", "debit", "credit", "balance", "note"]]
        flagged_df = df_out[df_out["note"] != ""]

        st.success(f"Extracted {len(df_out)} transaction(s).")
        if len(flagged_df):
            st.warning(f"{len(flagged_df)} row(s) flagged for manual review "
                       f"({100*len(flagged_df)/len(df_out):.1f}%).")
        else:
            st.info("No rows flagged — all balances reconciled.")

        st.subheader("All transactions")
        st.dataframe(df_out, use_container_width=True)

        if len(flagged_df):
            st.subheader("Flagged rows")
            st.dataframe(flagged_df, use_container_width=True)

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df_out.to_excel(writer, sheet_name="Transactions", index=False)
            if len(flagged_df):
                flagged_df.to_excel(writer, sheet_name="Flagged for review", index=False)
        excel_buffer.seek(0)

        st.download_button(
            "⬇️ Download Excel file",
            data=excel_buffer,
            file_name=f"{Path(uploaded_file.name).stem}_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.divider()
st.caption(
    "Currently tuned for: HDFC, Axis (digital), Union Bank (scanned, via OCR). "
    "Other banks fall back to a generic pattern and may return 0 transactions."
)