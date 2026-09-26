import streamlit as st
import subprocess
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

PARSER_SCRIPT = "bank_statement_parser_v6.py"

uploaded_file = st.file_uploader("Choose a PDF", type=["pdf"])

if uploaded_file is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / "input.pdf"
        output_path = Path(tmpdir) / "output.xlsx"

        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        with st.spinner("Processing... this can take a while for scanned PDFs (OCR)."):
            result = subprocess.run(
                ["python3", PARSER_SCRIPT, str(input_path), str(output_path)],
                capture_output=True,
                text=True,
                timeout=600,
            )

        with st.expander("Log output"):
            st.code(result.stdout or "(no stdout)", language="text")
            if result.stderr:
                st.code(result.stderr, language="text")

        if output_path.exists():
            sheets = pd.read_excel(output_path, sheet_name=None)
            df = sheets.get("Transactions")

            if df is None or df.empty:
                st.warning("0 transactions extracted — this bank's format may not be configured yet in BANK_CONFIGS.")
            else:
                st.success(f"Extracted {len(df)} transaction(s).")

                flagged_df = sheets.get("Flagged for review")
                n_flagged = 0 if flagged_df is None else len(flagged_df)
                if n_flagged:
                    st.warning(f"{n_flagged} row(s) flagged for manual review ({100*n_flagged/len(df):.1f}%).")
                else:
                    st.info("No rows flagged — all balances reconciled.")

                st.subheader("All transactions")
                st.dataframe(df, use_container_width=True)

                if n_flagged:
                    st.subheader("Flagged rows")
                    st.dataframe(flagged_df, use_container_width=True)

            with open(output_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Excel file",
                    data=f.read(),
                    file_name=f"{Path(uploaded_file.name).stem}_extracted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.error("Parser did not produce an output file. Check the log above for errors.")

st.divider()
st.caption(
    "Currently tuned for: HDFC, Axis (digital), Union Bank (scanned, via OCR). "
    "Other banks fall back to a generic pattern and may return 0 transactions."
)
