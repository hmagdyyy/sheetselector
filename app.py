import io
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Select Sheets", page_icon="📄", layout="centered")
st.title("📄 Extract Selected Sheets from Excel")
st.caption("Uploads stay in memory. We preserve formulas, formatting, and macros (if .xlsm).")

xlsx_file = st.file_uploader("Upload Excel (.xlsx / .xlsm)", type=["xlsx", "xlsm"])
csv_file  = st.file_uploader("Upload CSV with a 'SheetName' column", type=["csv"])

if st.button("Build file", disabled=not (xlsx_file and csv_file)):
    try:
        # --- read CSV safely ---
        wanted_df = pd.read_csv(csv_file, dtype=str)
        if "SheetName" not in wanted_df.columns:
            st.error("CSV must contain a column named 'SheetName'.")
            st.stop()
        wanted_sheets = (
            wanted_df["SheetName"]
            .dropna()
            .map(lambda s: s.strip())
            .tolist()
        )
        if not wanted_sheets:
            st.error("No sheet names found in CSV.")
            st.stop()

        # --- detect macros & set output extension accordingly ---
        src_name = xlsx_file.name
        ext = Path(src_name).suffix.lower()
        keep_vba = (ext == ".xlsm")
        out_ext = ".xlsm" if keep_vba else ".xlsx"

        # Important: read file bytes once and wrap in BytesIO
        xlsx_bytes = xlsx_file.getvalue()
        xlsx_buf = io.BytesIO(xlsx_bytes)

        # --- load workbook with formulas & formatting preserved ---
        wb = load_workbook(filename=xlsx_buf, data_only=False, keep_vba=keep_vba)

        source_sheetnames = wb.sheetnames
        wanted_set = set(wanted_sheets)

        # warn about missing sheets
        missing = [s for s in wanted_sheets if s not in source_sheetnames]
        if missing:
            st.warning("Not found in source file: " + ", ".join(missing))

        # abort if nothing matches
        matched = [s for s in wanted_sheets if s in source_sheetnames]
        if not matched:
            st.error("No matching sheets found. Check your CSV and sheet names.")
            st.stop()

        # --- remove unwanted sheets (mutate in place) ---
        for name in list(source_sheetnames):  # copy list since we'll mutate
            if name not in wanted_set:
                ws = wb[name]
                wb.remove(ws)

        # --- save to bytes and serve ---
        out_buf = io.BytesIO()
        # For macros, ensure .xlsm is used so VBA is preserved
        wb.save(out_buf)
        out_buf.seek(0)

        out_name = f"SelectedSheets{out_ext}"
        st.success(f"✅ Done! Saved filtered workbook → {out_name}")
        st.download_button(
            "Download file",
            data=out_buf,
            file_name=out_name,
            mime="application/vnd.ms-excel" if keep_vba else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Optional: show what will be included
        st.write("**Included sheets:**", ", ".join(matched))
        if missing:
            st.write("**Missing sheets:**", ", ".join(missing))

    except Exception as e:
        st.error("Something went wrong.")
        st.exception(e)
