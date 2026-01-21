import streamlit as st
import pandas as pd
import os
import tempfile
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter
from openpyxl.pivot.table import PivotTable, PivotCache, PivotCacheDefinition

st.set_page_config(page_title="GSTR-1 Processor", layout="centered")
st.title("GSTR-1 Excel Processor")

# ---------------- Utilities ---------------- #

def normalize_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )
    return df

def find_column_by_keywords(df, keywords, label):
    for col in df.columns:
        if all(k.lower() in col.lower() for k in keywords):
            return col
    raise KeyError(f"{label} column not found")

def get_column_letter_by_header(ws, header):
    for c in range(1, ws.max_column + 1):
        if ws.cell(1, c).value == header:
            return ws.cell(1, c).column_letter
    raise KeyError(f"{header} not found in {ws.title}")

# ---------------- Session State ---------------- #

if "processed" not in st.session_state:
    st.session_state.processed = False
    st.session_state.outputs = {}

# ---------------- UI ---------------- #

company_code = st.text_input("Company Code")

sd_file = st.file_uploader("Upload SD File", type="xlsx")
sr_file = st.file_uploader("Upload SR File", type="xlsx")
tb_file = st.file_uploader("Upload TB File", type="xlsx")
gl_file = st.file_uploader("Upload GL Dump File", type="xlsx")

# ---------------- Processing ---------------- #

if st.button("Process Files"):

    if not all([company_code, sd_file, sr_file, tb_file, gl_file]):
        st.error("All inputs are mandatory")
        st.stop()

    try:
        outputs = {}

        with tempfile.TemporaryDirectory() as tmpdir:

            paths = {}
            for f, name in [
                (sd_file, "sd.xlsx"),
                (sr_file, "sr.xlsx"),
                (tb_file, "tb.xlsx"),
                (gl_file, "gl.xlsx"),
            ]:
                p = os.path.join(tmpdir, name)
                with open(p, "wb") as out:
                    out.write(f.getbuffer())
                paths[name] = p

            # ---------- SALES REGISTER ---------- #

            df_sales = pd.concat(
                [
                    normalize_columns(pd.read_excel(paths["sd.xlsx"])),
                    normalize_columns(pd.read_excel(paths["sr.xlsx"]))
                ],
                ignore_index=True
            )

            # ---------- WRITE BASE WORKBOOK ---------- #

            gstr_path = os.path.join(tmpdir, f"{company_code}_GSTR-1_Workbook.xlsx")
            with pd.ExcelWriter(gstr_path, engine="openpyxl") as writer:
                df_sales.to_excel(writer, "Sales register", index=False)

            # ---------- EXCEL POST PROCESS ---------- #

            wb = load_workbook(gstr_path)
            ws = wb["Sales register"]

            # Convert Sales register into Excel Table (required for pivot)
            end_col = get_column_letter(ws.max_column)
            end_row = ws.max_row
            table = Table(displayName="SalesRegisterTable", ref=f"A1:{end_col}{end_row}")
            ws.add_table(table)

            # ---------- CREATE PIVOT ---------- #

            pivot_cache = PivotCacheDefinition(
                cacheId=1,
                sourceRef=f"SalesRegisterTable"
            )

            pivot_table = PivotTable(
                name="SalesSummaryPivot",
                cacheId=1,
                ref="A3"
            )

            # Filters
            pivot_table.addReportFilter("GSTIN of Taxpayer")

            # Rows
            pivot_table.addRowField("Sales summary")

            # Values (order preserved)
            pivot_table.addDataField("Taxable value", "Sum of Taxable value")
            pivot_table.addDataField("IGST Amt", "Sum of IGST Amt")
            pivot_table.addDataField("CGST Amt", "Sum of CGST Amt")
            pivot_table.addDataField("SGST/UTGST Amt", "Sum of SGST/UTGST Amt")

            pivot_ws = wb.create_sheet("Sales summary")
            pivot_ws._pivots.append(pivot_table)
            wb._pivots.append(pivot_cache)

            wb.save(gstr_path)

            with open(gstr_path, "rb") as f:
                outputs["GSTR-1 Workbook.xlsx"] = f.read()

        st.session_state.outputs = outputs
        st.session_state.processed = True
        st.success("Processing completed successfully")

    except Exception as e:
        st.error(str(e))

# ---------------- Download ---------------- #

if st.session_state.processed:
    for k, v in st.session_state.outputs.items():
        st.download_button(f"Download {k}", v, file_name=k)
