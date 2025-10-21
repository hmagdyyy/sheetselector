import io
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Select Sheets", page_icon="📄", layout="centered")
st.title("📄 Extract Selected Sheets from Excel")

# ---- Upload (single multi-file uploader to avoid state issues) ----
files = st.file_uploader(
    "Upload your Excel (.xlsx/.xlsm) and CSV (with a 'SheetName' column)",
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

# ---- Output mode ----
mode = st.radio(
    "Choose output",
    ["Filtered workbook", "Summary sheet", "Both"],
    index=0,
    key="mode_radio",
)

ready = (xlsx_file is not None and csv_file is not None)

# ---- Build button ----
if st.button("Build file(s)", disabled=not ready, key="build_btn"):
    try:
        # --- Read CSV for desired tabs ---
        wanted_df = pd.read_csv(csv_file, dtype=str)
        if "SheetName" not in wanted_df.columns:
            st.error("CSV must contain a column named 'SheetName'.")
            st.stop()
        wanted = wanted_df["SheetName"].dropna().map(str.strip).tolist()
        if not wanted:
            st.error("No sheet names found in CSV.")
            st.stop()

        # --- Macro detection & output extension ---
        keep_vba = Path(xlsx_file.name).suffix.lower() == ".xlsm"
        filtered_ext = ".xlsm" if keep_vba else ".xlsx"

        # --- Load workbook from bytes twice:
        # wb      -> preserves formulas/macros for filtered save
        # wb_vals -> data_only=True for reading values into the summary
        xlsx_bytes = xlsx_file.getvalue()
        wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=False, keep_vba=keep_vba)
        wb_vals = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)

        source_sheetnames = wb.sheetnames
        wanted_set = set(wanted)
        matched = [s for s in wanted if s in source_sheetnames]
        missing = [s for s in wanted if s not in source_sheetnames]
        if not matched:
            st.error("No matching sheets found in the workbook.")
            st.stop()

        # ---------- A) Filtered workbook (remove other sheets, preserve formatting/macros) ----------
        def build_filtered_workbook_bytes():
            # Mutate wb in-place; safe because we loaded from bytes
            for name in list(wb.sheetnames):
                if name not in wanted_set:
                    wb.remove(wb[name])
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf, f"SelectedSheets{filtered_ext}"

        # ---------- B) Summary sheet (Name=B4, Cash=C27, NAV by label search in col A → B) ----------
        def build_summary_bytes():
            TARGET = "total net asset value after ic fall"

            def norm(s):
                if s is None:
                    return ""
                # lowercase & collapse spaces
                return " ".join(str(s).strip().lower().split())

            rows = []
            for tab in matched:
                ws = wb_vals[tab]  # read values (data_only=True)

                # Name from B4
                try:
                    name_val = ws["B4"].value
                except Exception:
                    name_val = None

                # Cash from C27
                try:
                    cash_val = ws["C27"].value
                except Exception:
                    cash_val = None

                # NAV via search in col A for TARGET text, then take col B same row
                nav_val = None
                try:
                    for a_cell in ws["A"]:
                        label = norm(a_cell.value)
                        if label and (label == norm(TARGET) or norm(TARGET) in label):
                            nav_val = ws.cell(row=a_cell.row, column=2).value  # B
                            break
                except Exception:
                    nav_val = None

                rows.append({
                    "Name": name_val,
                    "NAV": nav_val,
                    "Cash": cash_val,
                })

            df = pd.DataFrame(rows)

            # Try to coerce numeric where possible
            for col in ["NAV", "Cash"]:
                df[col] = pd.to_numeric(df[col], errors="coerce")

            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as xw:
                df.to_excel(xw, index=False, sheet_name="Summary")
            out.seek(0)
            return out, "Summary.xlsx"

        # ---- Build per mode ----
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

