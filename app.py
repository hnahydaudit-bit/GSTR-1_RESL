import streamlit as st
import pandas as pd
import os
import tempfile
import xlsxwriter
from openpyxl import load_workbook

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


def find_column(df, name):
    if name not in df.columns:
        raise KeyError(f"Column '{name}' not found")
    return name


def get_col_idx(headers, name):
    return headers.index(name)

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

            # ---------------- READ SALES REGISTER ---------------- #

            df_sales = pd.concat(
                [
                    normalize_columns(pd.read_excel(sd_file)),
                    normalize_columns(pd.read_excel(sr_file))
                ],
                ignore_index=True
            )

            # ---------------- WRITE USING XLSXWRITER ---------------- #

            gstr_path = os.path.join(tmpdir, f"{company_code}_GSTR-1_Workbook.xlsx")

            workbook = xlsxwriter.Workbook(gstr_path)
            ws_sales = workbook.add_worksheet("Sales register")
            ws_pivot = workbook.add_worksheet("Sales summary")

            # Write Sales register
            for c, col in enumerate(df_sales.columns):
                ws_sales.write(0, c, col)

            for r in range(len(df_sales)):
                for c, col in enumerate(df_sales.columns):
                    ws_sales.write(r + 1, c, df_sales.iloc[r, c])

            # Define table range
            last_row = len(df_sales)
            last_col = len(df_sales.columns) - 1

            ws_sales.add_table(
                0, 0, last_row, last_col,
                {"columns": [{"header": c} for c in df_sales.columns]}
            )

            # ---------------- CREATE PIVOT ---------------- #

            pivot = workbook.add_pivot_table({
                "name": "SalesSummaryPivot",
                "source": "Sales register!A1:{}{}".format(
                    xlsxwriter.utility.xl_col_to_name(last_col),
                    last_row + 1
                ),
                "destination": "Sales summary!A3",
                "filters": [
                    {"field": "GSTIN of Taxpayer"}
                ],
                "rows": [
                    {"field": "Sales summary"}
                ],
                "values": [
                    {"field": "Taxable value", "function": "sum"},
                    {"field": "IGST Amt", "function": "sum"},
                    {"field": "CGST Amt", "function": "sum"},
                    {"field": "SGST/UTGST Amt", "function": "sum"},
                ],
            })

            workbook.close()

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
        st.download_button(
            label=f"Download {k}",
            data=v,
            file_name=k
        )
