import io
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Select Sheets", page_icon="📄", layout="centered")
st.title("📄 Extract Selected Sheets from Excel")

# Single uploader avoids state issues
files = st.file_uploader(
    "Upload your Excel (.xlsx/.xlsm) and CSV (with 'SheetName')",
    type=["xlsx", "xlsm", "csv"],
    accept_multiple_files=True,
    key="files_upl",
)

xlsx_file, csv_file = None, None
if files:
    for f in files:
        n = f.name.lower()
        if n.endswith((".xlsx", ".xlsm")) and xlsx_file is None:
            xlsx_file = f
        elif n.endswith(".csv") and csv_file is None:
            csv_file = f

st.write("**Detected**")
st.write("• Excel:", xlsx_file.name if xlsx_file else "—")
st.write("• CSV:", csv_file.name if csv_file else "—")

# NEW: output mode
mode = st.radio(
    "Choose output",
    ["Filtered workbook", "Summary sheet", "Both"],
    index=0,
)

ready = xlsx_file is not None and csv_file is not None

if st.button("Build file(s)", disabled=not ready, key="build_btn"):
    try:
        # --- read CSV for wanted tabs ---
        wanted_df = pd.read_csv(csv_file, dtype=str)
        if "SheetName" not in wanted_df.columns:
            st.error("CSV must contain a 'SheetName' column."); st.stop()
        wanted = wanted_df["SheetName"].dropna().map(str.strip).tolist()
        if not wanted:
            st.error("No sheet names found in CSV."); st.stop()

        # macro detection
        keep_vba = Path(xlsx_file.name).suffix.lower() == ".xlsm"
        out_ext = ".xlsm" if keep_vba else ".xlsx"

        # load workbook from bytes
        wb = load_workbook(io.BytesIO(xlsx_file.getvalue()), data_only=False, keep_vba=keep_vba)
        source_sheetnames = wb.sheetnames
        wanted_set = set(wanted)

        matched = [s for s in wanted if s in source_sheetnames]
        missing = [s for s in wanted if s not in source_sheetnames]
        if not matched:
            st.error("No matching sheets found in the workbook."); st.stop()

        # ---------- OPTION A: Filtered workbook ----------
        def build_filtered_workbook_bytes():
            # Work on a copy-like path: removing others in-place is fine since we already loaded from bytes
            for name in list(wb.sheetnames):
                if name not in wanted_set:
                    wb.remove(wb[name])
            buf = io.BytesIO()
            wb.save(buf); buf.seek(0)
            return buf, f"SelectedSheets{out_ext}"

        # ---------- OPTION B: Summary sheet ----------
        def build_summary_bytes():
            rows = []
            for name in matched:
                ws = wb[name]
                def v(cell):
                    try:
                        return ws[cell].value
                    except Exception:
                        return None
                row = {
                    "Sheet": name,
                    "Name (B4)": v("B4"),
                    "NAV (B114)": v("B114"),
                    "Cash (C27)": v("C27"),
                }
                rows.append(row)

            df = pd.DataFrame(rows)
            # Optional: try numeric conversion for NAV/Cash
            for col in ["NAV (B114)", "Cash (C27)"]:
                df[col] = pd.to_numeric(df[col], errors="coerce")

            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as xw:
                df.to_excel(xw, index=False, sheet_name="Summary")
            out.seek(0)
            return out, "Summary.xlsx"

        # Build according to mode
        if mode in ("Filtered workbook", "Both"):
            fbuf, fname = build_filtered_workbook_bytes()
            st.success(f"✅ Built {fname}")
            st.download_button("Download filtered workbook", data=fbuf, file_name=fname, key="dl_filtered")

        if mode in ("Summary sheet", "Both"):
            sbuf, sname = build_summary_bytes()
            st.success(f"✅ Built {sname}")
            st.download_button("Download summary sheet", data=sbuf, file_name=sname, key="dl_summary")

        if missing:
            st.warning("Missing sheets (not found in workbook): " + ", ".join(missing))

    except Exception as e:
        st.error(f"Error: {e}")

        if missing:
            st.warning("Missing sheets: " + ", ".join(missing))

    except Exception as e:
        st.error(f"Error: {e}")
