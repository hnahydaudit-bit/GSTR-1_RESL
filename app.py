import streamlit as st
import pandas as pd
import os
import tempfile
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


def find_column_by_keywords(df, keywords, label):
    for col in df.columns:
        col_l = col.lower()
        if all(k.lower() in col_l for k in keywords):
            return col
    raise KeyError(
        f"{label} column not found.\n"
        f"Expected keywords: {keywords}\n"
        f"Found columns: {list(df.columns)}"
    )

# ---------------- Session State ---------------- #

if "processed" not in st.session_state:
    st.session_state.processed = False
    st.session_state.outputs = {}

# ---------------- UI ---------------- #

company_code = st.text_input("Company Code (XXXX)")

sd_file = st.file_uploader("Upload SD File", type="xlsx")
sr_file = st.file_uploader("Upload SR File", type="xlsx")
tb_file = st.file_uploader("Upload TB File", type="xlsx")
gl_file = st.file_uploader("Upload GL Dump File", type="xlsx")

# ---------------- Processing ---------------- #

if st.button("Process Files"):

    if not company_code:
        st.error("Company Code is mandatory")
        st.stop()

    if not all([sd_file, sr_file, tb_file, gl_file]):
        st.error("All files must be uploaded")
        st.stop()

    try:
        outputs = {}

        with tempfile.TemporaryDirectory() as tmpdir:

            # Save uploaded files
            paths = {}
            for f, name in [
                (sd_file, "sd.xlsx"),
                (sr_file, "sr.xlsx"),
                (tb_file, "tb.xlsx"),
                (gl_file, "gl.xlsx"),
            ]:
                path = os.path.join(tmpdir, name)
                with open(path, "wb") as out:
                    out.write(f.getbuffer())
                paths[name] = path

            # ---------- SD + SR ---------- #

            df_sd = normalize_columns(pd.read_excel(paths["sd.xlsx"]))
            df_sr = normalize_columns(pd.read_excel(paths["sr.xlsx"]))

            df_sales = pd.concat([df_sd, df_sr], ignore_index=True)

            # ---------- GL ---------- #

            df_gl = normalize_columns(pd.read_excel(paths["gl.xlsx"]))

            gl_text_col = find_column_by_keywords(df_gl, ["g/l", "account", "long", "text"], "GL Text")
            gl_account_col = find_column_by_keywords(df_gl, ["g/l", "account"], "GL Account")
            value_col = find_column_by_keywords(df_gl, ["value"], "Amount")
            doc_col = find_column_by_keywords(df_gl, ["document"], "Document Number")

            gst_accounts = [
                "Central GST Payable",
                "Integrated GST Payable",
                "State GST Payable",
            ]

            df_gst = df_gl[df_gl[gl_text_col].isin(gst_accounts)]
            df_revenue = df_gl[df_gl[gl_account_col].astype(str).str.startswith("3")]

            # ---------- GSTR-1 Workbook ---------- #

            gstr_path = os.path.join(tmpdir, f"{company_code}_GSTR-1_Workbook.xlsx")

            with pd.ExcelWriter(gstr_path, engine="openpyxl") as writer:
                df_sales.to_excel(writer, sheet_name="Sales register", index=False)
                df_revenue.to_excel(writer, sheet_name="Revenue", index=False)
                df_gst.to_excel(writer, sheet_name="GST payable", index=False)

            # ---------- ADD VLOOKUP FORMULAS ---------- #

            wb = load_workbook(gstr_path)

            ws_sales = wb["Sales register"]
            ws_rev = wb["Revenue"]
            ws_gst = wb["GST payable"]

            sales_last_col = ws_sales.max_column
            ws_sales.cell(row=1, column=sales_last_col + 1, value="Revenue VLOOKUP")
            ws_sales.cell(row=1, column=sales_last_col + 2, value="GST Payable VLOOKUP")

            for r in range(2, ws_sales.max_row + 1):
                ws_sales.cell(
                    row=r, column=sales_last_col + 1,
                    value=f'=IFERROR(VLOOKUP(A{r},Revenue!A:A,1,FALSE),"Not Found")'
                )
                ws_sales.cell(
                    row=r, column=sales_last_col + 2,
                    value=f'=IFERROR(VLOOKUP(A{r},\'GST payable\'!A:A,1,FALSE),"Not Found")'
                )

            rev_last_col = ws_rev.max_column
            ws_rev.cell(row=1, column=rev_last_col + 1, value="Sales Register VLOOKUP")
            for r in range(2, ws_rev.max_row + 1):
                ws_rev.cell(
                    row=r, column=rev_last_col + 1,
                    value=f'=IFERROR(VLOOKUP(A{r},\'Sales register\'!A:A,1,FALSE),"Not Found")'
                )

            gst_last_col = ws_gst.max_column
            ws_gst.cell(row=1, column=gst_last_col + 1, value="Sales Register VLOOKUP")
            for r in range(2, ws_gst.max_row + 1):
                ws_gst.cell(
                    row=r, column=gst_last_col + 1,
                    value=f'=IFERROR(VLOOKUP(A{r},\'Sales register\'!A:A,1,FALSE),"Not Found")'
                )

            wb.save(gstr_path)

            with open(gstr_path, "rb") as f:
                outputs["GSTR-1 Workbook.xlsx"] = f.read()

        st.session_state.outputs = outputs
        st.session_state.processed = True
        st.success("Processing completed successfully")

    except Exception as e:
        st.error(str(e))

# ---------------- Downloads ---------------- #

if st.session_state.processed:
    for filename, data in st.session_state.outputs.items():
        st.download_button(
            label=f"Download {filename}",
            data=data,
            file_name=filename,
            key=filename
        )




