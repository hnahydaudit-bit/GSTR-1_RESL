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


def sanitize_dataframe(df):
    df.replace([np.inf, -np.inf], np.nan, inplace=True)
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d")
    df.fillna("", inplace=True)
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

    with tempfile.TemporaryDirectory() as tmpdir:

        # ---------- READ FILES ---------- #

        df_sales = pd.concat(
            [
                normalize_columns(pd.read_excel(sd_file)),
                normalize_columns(pd.read_excel(sr_file)),
            ],
            ignore_index=True,
        )

        df_gl = normalize_columns(pd.read_excel(gl_file))
        df_tb = normalize_columns(pd.read_excel(tb_file))

        df_sales = sanitize_dataframe(df_sales)
        df_gl = sanitize_dataframe(df_gl)
        df_tb = sanitize_dataframe(df_tb)

        # ---------- CREDIT NOTE NEGATIVES ---------- #

        amt_cols = ["Taxable value", "IGST Amt", "CGST Amt", "SGST/UTGST Amt"]
        df_sales["Document Type"] = df_sales["Document Type"].astype(str)

        for c in amt_cols:
            df_sales[c] = pd.to_numeric(df_sales[c], errors="coerce").fillna(0)
            df_sales.loc[df_sales["Document Type"] == "C", c] *= -1

        # ---------- SALES SUMMARY COLUMN ---------- #

        df_sales["Tax rate"] = pd.to_numeric(
            df_sales["Tax rate"], errors="coerce"
        ).fillna(0)

        def classify(r):
            if r["Invoice type"] == "SEWOP":
                return "SEZWOP"
            if r["Invoice type"] == "SEWP":
                return "SEWP"
            if r["Invoice type"] not in ("SEWOP", "SEWP") and r["Tax rate"] == 0:
                return "Exempt supply"
            if r["Invoice type"] == "B2B" and r["Document Type"] == "C":
                return "B2B Credit Notes"
            if r["Invoice type"] == "B2B":
                return "B2B Supplies"
            if r["Invoice type"] == "B2CS":
                return "B2C Supplies"
            return ""

        df_sales["Sales summary"] = df_sales.apply(classify, axis=1)

        # ---------- REVENUE & GST PAYABLE ---------- #

        df_revenue = df_gl[
            df_gl["G/L Account: Long Text"].str.contains("Revenue", case=False, na=False)
        ]

        df_gst_payable = df_gl[
            df_gl["G/L Account: Long Text"].str.contains("GST", case=False, na=False)
        ]

        # ---------- GST SUMMARY ---------- #

        df_gl["Company Code Currency Value"] = pd.to_numeric(
            df_gl["Company Code Currency Value"], errors="coerce"
        ).fillna(0)

        df_tb["Period 09 C"] = pd.to_numeric(df_tb["Period 09 C"], errors="coerce").fillna(0)
        df_tb["Period 09 D"] = pd.to_numeric(df_tb["Period 09 D"], errors="coerce").fillna(0)
        df_tb["Difference as per TB"] = df_tb["Period 09 C"] - df_tb["Period 09 D"]

        gst_summary = (
            df_gl
            .groupby("G/L Account: Long Text", as_index=False)["Company Code Currency Value"]
            .sum()
            .merge(
                df_tb
                .groupby("G/L Acct Long Text", as_index=False)["Difference as per TB"]
                .sum(),
                left_on="G/L Account: Long Text",
                right_on="G/L Acct Long Text",
                how="left",
            )
            .fillna(0)
        )

        # ---------- WRITE EXCEL ---------- #

        output_path = os.path.join(tmpdir, f"{company_code}_GSTR-1_Workbook.xlsx")
        wb = xlsxwriter.Workbook(output_path, {"nan_inf_to_errors": True})

        # ---- CHECK PIVOT SUPPORT (NO CRASH) ---- #
        if not hasattr(wb, "add_pivot_table"):
            st.error(
                "This environment does not support editable Excel Pivot Tables.\n\n"
                "Please upgrade xlsxwriter to version 3.1.0 or higher "
                "and redeploy the app."
            )
            st.stop()

        ws_sales = wb.add_worksheet("Sales register")
        ws_rev = wb.add_worksheet("Revenue")
        ws_gst = wb.add_worksheet("GST payable")
        ws_gst_sum = wb.add_worksheet("GST Summary")
        ws_pivot = wb.add_worksheet("Sales summary")

        def write_df(ws, df):
            for c, col in enumerate(df.columns):
                ws.write(0, c, col)
            for r in range(len(df)):
                for c, col in enumerate(df.columns):
                    ws.write(r + 1, c, df.iloc[r, c])

        write_df(ws_sales, df_sales)
        write_df(ws_rev, df_revenue)
        write_df(ws_gst, df_gst_payable)
        write_df(ws_gst_sum, gst_summary)

        last_row = len(df_sales)
        last_col = len(df_sales.columns) - 1

        # ---------- REAL EDITABLE EXCEL PIVOT ---------- #

        wb.add_pivot_table(
            {
                "name": "SalesSummaryPivot",
                "source": (
                    f"'Sales register'!A1:"
                    f"{xlsxwriter.utility.xl_col_to_name(last_col)}{last_row+1}"
                ),
                "destination": "'Sales summary'!A3",
                "filters": [{"field": "GSTIN of Taxpayer"}],
                "rows": [{"field": "Sales summary"}],
                "values": [
                    {"field": "Taxable value", "function": "sum"},
                    {"field": "IGST Amt", "function": "sum"},
                    {"field": "CGST Amt", "function": "sum"},
                    {"field": "SGST/UTGST Amt", "function": "sum"},
                ],
            }
        )

        wb.close()

        with open(output_path, "rb") as f:
            st.download_button(
                "Download GSTR-1 Workbook",
                f.read(),
                file_name=os.path.basename(output_path),
            )




