import streamlit as st
import pandas as pd
import tempfile
import os
import xlsxwriter

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

# ---------------- UI ---------------- #

company_code = st.text_input("Company Code")

sd_file = st.file_uploader("Upload SD File", type="xlsx")
sr_file = st.file_uploader("Upload SR File", type="xlsx")

# ---------------- Processing ---------------- #

if st.button("Process Files"):

    if not all([company_code, sd_file, sr_file]):
        st.error("Company Code, SD & SR files are mandatory")
        st.stop()

    try:
        with tempfile.TemporaryDirectory() as tmpdir:

            # ---- Read & consolidate Sales Register ---- #
            df_sales = pd.concat(
                [
                    normalize_columns(pd.read_excel(sd_file)),
                    normalize_columns(pd.read_excel(sr_file))
                ],
                ignore_index=True
            )

            output_path = os.path.join(
                tmpdir, f"{company_code}_GSTR-1_Workbook.xlsx"
            )

            # ---- Write using XLSXWRITER ---- #
            workbook = xlsxwriter.Workbook(output_path)
            ws_sales = workbook.add_worksheet("Sales register")
            ws_pivot = workbook.add_worksheet("Sales summary")

            # Write headers
            for col, header in enumerate(df_sales.columns):
                ws_sales.write(0, col, header)

            # Write data
            for row in range(len(df_sales)):
                for col, header in enumerate(df_sales.columns):
                    ws_sales.write(row + 1, col, df_sales.iloc[row, col])

            last_row = len(df_sales)
            last_col = len(df_sales.columns) - 1

            # Convert to Excel Table (mandatory for pivot)
            ws_sales.add_table(
                0, 0, last_row, last_col,
                {"columns": [{"header": c} for c in df_sales.columns]}
            )

            # ---- CREATE PIVOT TABLE ---- #
            workbook.add_pivot_table({
                "name": "SalesSummaryPivot",
                "source": f"'Sales register'!A1:{xlsxwriter.utility.xl_col_to_name(last_col)}{last_row+1}",
                "destination": "'Sales summary'!A3",
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

            with open(output_path, "rb") as f:
                st.download_button(
                    "Download GSTR-1 Workbook",
                    data=f.read(),
                    file_name=os.path.basename(output_path)
                )

    except Exception as e:
        st.error(str(e))
