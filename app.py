import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import os
import xlsxwriter

st.set_page_config(page_title="GSTR-1 Processor", layout="centered")
st.title("GSTR-1 Processor")

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
tb_file = st.file_uploader("Upload TB File", type="xlsx")
gl_file = st.file_uploader("Upload GL Dump File", type="xlsx")

# ---------------- Processing ---------------- #

if st.button("Process Files"):

    if not all([company_code, sd_file, sr_file, tb_file, gl_file]):
        st.error("All inputs are mandatory")
        st.stop()

    try:
        with tempfile.TemporaryDirectory() as tmpdir:

            # ---------- READ INPUTS ---------- #

            df_sales = pd.concat(
                [
                    normalize_columns(pd.read_excel(sd_file)),
                    normalize_columns(pd.read_excel(sr_file))
                ],
                ignore_index=True
            )

            df_gl = normalize_columns(pd.read_excel(gl_file))
            df_tb = normalize_columns(pd.read_excel(tb_file))

            # ---------- CREDIT NOTE NEGATIVES ---------- #

            amt_cols = ["Taxable value", "IGST Amt", "CGST Amt", "SGST/UTGST Amt"]
            mask_cn = df_sales["Document Type"] == "C"

            for c in amt_cols:
                df_sales.loc[mask_cn, c] = -df_sales.loc[mask_cn, c].abs()

            # ---------- SALES SUMMARY CLASSIFICATION ---------- #

            def classify(row):
                it, dt, tr = row["Invoice type"], row["Document Type"], float(row["Tax rate"] or 0)
                if it == "SEWOP":
                    return "SEZWOP"
                if it == "SEWP":
                    return "SEWP"
                if it not in ("SEWOP", "SEWP") and tr == 0:
                    return "Exempt supply"
                if it == "B2B" and dt == "C" and tr != 0:
                    return "B2B Credit Notes"
                if it == "B2B" and dt != "C" and tr != 0:
                    return "B2B Supplies"
                if it == "B2CS" and tr != 0:
                    return "B2C Supplies"
                return ""

            df_sales["Sales summary"] = df_sales.apply(classify, axis=1)

            # ---------- GST SUMMARY ---------- #

            gst_accounts = [
                "Central GST Payable",
                "Integrated GST Payable",
                "State GST Payable"
            ]

            df_gst = df_gl[df_gl["G/L Account: Long Text"].isin(gst_accounts)]
            df_tb_gst = df_tb[df_tb["G/L Acct Long Text"].isin(gst_accounts)].copy()
            df_tb_gst["Difference as per TB"] = df_tb_gst["Period 09 C"] - df_tb_gst["Period 09 D"]

            summary_df = (
                df_gst.groupby("G/L Account: Long Text", as_index=False)["Value"].sum()
                .rename(columns={
                    "G/L Account: Long Text": "GST Type",
                    "Value": "GST Payable as per GL"
                })
                .merge(
                    df_tb_gst.groupby("G/L Acct Long Text", as_index=False)["Difference as per TB"].sum()
                    .rename(columns={"G/L Acct Long Text": "GST Type"}),
                    on="GST Type",
                    how="left"
                )
                .fillna(0)
            )

            # ---------- WRITE EXCEL (XLSXWRITER) ---------- #

            output_path = os.path.join(tmpdir, f"{company_code}_GSTR-1_Workbook.xlsx")
            wb = xlsxwriter.Workbook(output_path)

            ws_sales = wb.add_worksheet("Sales register")
            ws_rev = wb.add_worksheet("Revenue")
            ws_gst = wb.add_worksheet("GST payable")
            ws_sum = wb.add_worksheet("GST Summary")
            ws_pivot = wb.add_worksheet("Sales summary")

            # ---------- WRITE SALES REGISTER ---------- #

            for c, col in enumerate(df_sales.columns):
                ws_sales.write(0, c, col)

            for r in range(len(df_sales)):
                for c, col in enumerate(df_sales.columns):
                    ws_sales.write(r + 1, c, df_sales.iloc[r, c])

            last_row = len(df_sales)
            last_col = len(df_sales.columns) - 1

            ws_sales.add_table(
                0, 0, last_row, last_col,
                {"columns": [{"header": c} for c in df_sales.columns]}
            )

            # ---------- GST SUMMARY ---------- #

            for c, col in enumerate(summary_df.columns):
                ws_sum.write(0, c, col)

            for r in range(len(summary_df)):
                for c, col in enumerate(summary_df.columns):
                    ws_sum.write(r + 1, c, summary_df.iloc[r, c])

            num_fmt = wb.add_format({"num_format": "0.00"})
            for r in range(1, len(summary_df) + 1):
                ws_sum.write_formula(r, 3, f"=B{r+1}+C{r+1}", num_fmt)

            ws_sum.write(0, 3, "Net Difference")

            # ---------- PIVOT TABLE ---------- #

            wb.add_pivot_table({
                "name": "SalesSummaryPivot",
                "source": f"'Sales register'!A1:{xlsxwriter.utility.xl_col_to_name(last_col)}{last_row+1}",
                "destination": "'Sales summary'!A3",
                "filters": [{"field": "GSTIN of Taxpayer"}],
                "rows": [{"field": "Sales summary"}],
                "values": [
                    {"field": "Taxable value", "function": "sum"},
                    {"field": "IGST Amt", "function": "sum"},
                    {"field": "CGST Amt", "function": "sum"},
                    {"field": "SGST/UTGST Amt", "function": "sum"},
                ],
            })

            wb.close()

            with open(output_path, "rb") as f:
                st.download_button(
                    "Download GSTR-1 Workbook",
                    data=f.read(),
                    file_name=os.path.basename(output_path)
                )

    except Exception as e:
        st.error(str(e))
