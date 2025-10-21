import io
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Select Sheets", page_icon="📄", layout="centered")
st.title("📄 Extract Selected Sheets from Excel")

# One uploader for both files → fewer state issues
files = st.file_uploader(
    "Upload your Excel (.xlsx/.xlsm) and CSV (with 'SheetName')",
    type=["xlsx", "xlsm", "csv"],
    accept_multiple_files=True,
    key="files_upl"
)

xlsx_file, csv_file = None, None
if files:
    for f in files:
        name = f.name.lower()
        if name.endswith((".xlsx", ".xlsm")) and xlsx_file is None:
            xlsx_file = f
        elif name.endswith(".csv") and csv_file is None:
            csv_file = f

    # Helpful UI feedback
    st.write("**Detected:**")
    st.write("• Excel:", xlsx_file.name if xlsx_file else "—")
    st.write("• CSV:", csv_file.name if csv_file else "—")

ready = (xlsx_file is not None and csv_file is not None)

if st.button("Build file", disabled=not ready, key="build_btn"):
    try:
        # read CSV
        wanted_df = pd.read_csv(csv_file, dtype=str)
        if "SheetName" not in wanted_df.columns:
            st.error("CSV must contain a 'SheetName' column.")
            st.stop()
        wanted = (wanted_df["SheetName"].dropna().map(str.strip).tolist())
        if not wanted:
            st.error("No sheet names found in CSV.")
            st.stop()

        # detect macros
        keep_vba = Path(xlsx_file.name).suffix.lower() == ".xlsm"
        out_ext = ".xlsm" if keep_vba else ".xlsx"

        # IMPORTANT: read bytes once; wrap in BytesIO
        xlsx_bytes = xlsx_file.getvalue()
        wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=False, keep_vba=keep_vba)

        # filter sheets
        source_sheetnames = wb.sheetnames
        wanted_set = set(wanted)
        matched = [s for s in wanted if s in source_sheetnames]
        missing = [s for s in wanted if s not in source_sheetnames]
        if not matched:
            st.error("No matching sheets found in the workbook."); st.stop()

        for name in list(source_sheetnames):
            if name not in wanted_set:
                wb.remove(wb[name])

        # save to bytes
        out_buf = io.BytesIO()
        wb.save(out_buf); out_buf.seek(0)
        out_name = f"SelectedSheets{out_ext}"

        st.success(f"✅ Done! {out_name}")
        st.download_button("Download file", data=out_buf, file_name=out_name)
        if missing:
            st.warning("Missing sheets: " + ", ".join(missing))

    except Exception as e:
        st.error(f"Error: {e}")
