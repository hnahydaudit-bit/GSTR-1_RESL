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
    # Replace Inf / -Inf
    df.replace([np.inf, -np.inf], np.nan, inplace=True)

    # Convert datetime columns to string to avoid NaT errors
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

    try:
        with tempfile.TemporaryDirectory() as tmpdir:

            # ---------- READ INPUT FILES ---------- #

            df_sales = pd.concat(
                [
                    normalize_columns(pd.read_excel(sd_file)),
                    normalize_columns(pd.read_excel(sr_file)),
                ],
                ignore_index=True,
            )

            df_gl = normalize_columns(pd.read_excel(gl_file))
            df_tb = normalize_columns(pd.read_excel(tb_file))

            # ---------- SANITIZE DATA ---------- #

            df_sales = sanitize_dataframe(df_sales)
            df_gl = sanitize_dataframe(df_gl)
            df_tb = sanitize_dataframe(df_tb)

            # ---------- CREDIT NOTE NEGATIVES ---------- #

            amount_cols = ["Taxable value", "IGST Amt", "CGST Amt", "SGST/UTGST Amt"]

            df_sales["Document Type"] = df_sales["Document Type"].astype(str)

            for col in amount_cols:
                df_sales[col] = pd.to_numeric(df_sales[col], errors="coerce").fillna(0)
                df_sales.loc[df_sales["Document Type"] == "C", col] *= -1

            # ---------- SALES SUMMARY CLASSIFICATION ---------- #

            df_sales["Tax rate"] = pd.to_numeric(
                df_sales["Tax rate"], errors="coerce"
            ).fillna(0)

            def classify(row):
                it = row["Invoice type"]
                dt = row["Document Type"]
                tr = row["Tax rate"]

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

            # ---------- GST SUMMARY (GL vs TB) ---------- #

            gst_accounts = [
                "Central GST Payable",
                "Integrated GST Payable",
                "State GST Payable",
            ]

            df_gl["Company Code Currency Value"] = pd.to_numeric(
                df_gl["Company Code Currency Value"], errors="coerce"
            ).fillna(0)

            df_gst = df_gl[
                df_gl["G/L Account: Long Text"].isin(gst_accounts)
            ]

            df_tb["Period 09 C"] = pd.to_numeric(
                df_tb["Period 09 C"], errors="coerce"
            ).fillna(0)
            df_tb["Period 09 D"] = pd.to_numeric(
                df_tb["Period 09 D"], errors="coerce"
            ).fillna(0)

            df_tb_gst = df_tb[
                df_tb["G/L Acct Long Text"].isin(gst_accounts)
            ].copy()

            df_tb_gst["Difference as per TB"] = (
                df_tb_gst["Period 09 C"] - df_tb_gst["Period 09 D"]
            )

            summary_df = (
                df_gst
                .groupby("G/L Account: Long Text", as_index=False)[
                    "Company Code Currency Value"
                ]
                .sum()
                .rename(
                    columns={
                        "G/L Account: Long Text": "GST Type",
                        "Company Code Currency Value": "GST Payable as per GL",
                    }
                )
                .merge(
                    df_tb_gst
                    .groupby("G/L Acct Long Text", as_index=False)[
                        "Difference as per TB"
                    ]
                    .sum()
                    .rename(columns={"G/L Acct Long Text": "GST Type"}),
                    on="GST Type",
                    how="left",
                )
                .fillna(0)
            )

            # ---------- WRITE EXCEL (XLSXWRITER) ---------- #

            output_path = os.path.join(
                tmpdir, f"{company_code}_GSTR-1_Workbook.xlsx"
            )

            wb = xlsxwriter.Workbook(
                output_path,
                {"nan_inf_to_errors": True}
            )

            # ---- HARD CHECK FOR PIVOT SUPPORT ---- #
            if not hasattr(wb, "add_pivot_table"):
                raise RuntimeError(
                    "Installed xlsxwriter version does not support pivot tables. "
                    "Please upgrade xlsxwriter to >= 3.1.0."
                )

            ws_sales = wb.add_worksheet("Sales register")
            ws_sum = wb.add_worksheet("GST Summary")
            ws_pivot = wb.add_worksheet("Sales summary")

            # --- Sales register sheet --- #
            for c, col in enumerate(df_sales.columns):
                ws_sales.write(0, c, col)

            for r in range(len(df_sales)):
                for c, col in enumerate(df_sales.columns):
                    ws_sales.write(r + 1, c, df_sales.iloc[r, c])

            last_row = len(df_sales)
            last_col = len(df_sales.columns) - 1

            ws_sales.add_table(
                0,
                0,
                last_row,
                last_col,
                {"columns": [{"header": c} for c in df_sales.columns]},
            )

            # --- GST Summary sheet --- #
            for c, col in enumerate(summary_df.columns):
                ws_sum.write(0, c, col)

            for r in range(len(summary_df)):
                for c, col in enumerate(summary_df.columns):
                    ws_sum.write(r + 1, c, summary_df.iloc[r, c])

            num_fmt = wb.add_format({"num_format": "0.00"})
            ws_sum.write(0, 3, "Net Difference")

            for r in range(1, len(summary_df) + 1):
                ws_sum.write_formula(
                    r, 3, f"=B{r+1}+C{r+1}", num_fmt
                )

            # --- Sales summary Pivot --- #
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
                    data=f.read(),
                    file_name=os.path.basename(output_path),
                )

    except Exception as e:
        st.error(str(e))
