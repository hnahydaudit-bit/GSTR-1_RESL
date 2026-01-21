import streamlit as st
import pandas as pd
import os
import tempfile

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

            # ---------- SD + SR CONSOLIDATION ---------- #

            df_sd = normalize_columns(pd.read_excel(paths["sd.xlsx"]))
            df_sr = normalize_columns(pd.read_excel(paths["sr.xlsx"]))

            df_sales = pd.concat([df_sd, df_sr], ignore_index=True)

            # ---------- GL PROCESSING ---------- #

            df_gl = normalize_columns(pd.read_excel(paths["gl.xlsx"]))

            gl_text_col = find_column_by_keywords(
                df_gl, ["g/l", "account", "long", "text"], "GL Long Text"
            )
            gl_account_col = find_column_by_keywords(
                df_gl, ["g/l", "account"], "GL Account"
            )
            value_col = find_column_by_keywords(
                df_gl, ["value"], "GL Amount"
            )
            doc_col_gl = find_column_by_keywords(
                df_gl, ["document"], "GL Document Number"
            )

            gst_accounts = [
                "Central GST Payable",
                "Integrated GST Payable",
                "State GST Payable",
            ]

            df_gst = df_gl[df_gl[gl_text_col].isin(gst_accounts)].copy()
            df_revenue = df_gl[
                df_gl[gl_account_col].astype(str).str.startswith("3")
            ].copy()

            # ---------- LOOKUPS ---------- #

            generic_field_col = find_column_by_keywords(
                df_sales, ["generic", "8"], "Generic Field 8"
            )

            revenue_doc_set = set(df_revenue[doc_col_gl].astype(str))
            gst_doc_set = set(df_gst[doc_col_gl].astype(str))

            df_sales["Match with Revenue (Doc No)"] = (
                df_sales[generic_field_col].astype(str)
                .isin(revenue_doc_set)
                .map({True: "Yes", False: "No"})
            )

            df_sales["Match with GST Payable (Doc No)"] = (
                df_sales[generic_field_col].astype(str)
                .isin(gst_doc_set)
                .map({True: "Yes", False: "No"})
            )

            # ---------- GSTR-1 WORKBOOK ---------- #

            gstr_path = os.path.join(
                tmpdir, f"{company_code}_GSTR-1_Workbook.xlsx"
            )

            with pd.ExcelWriter(gstr_path, engine="openpyxl") as writer:
                df_sales.to_excel(
                    writer, sheet_name="Sales register", index=False
                )
                df_revenue.to_excel(
                    writer, sheet_name="Revenue", index=False
                )
                df_gst.to_excel(
                    writer, sheet_name="GST Payable", index=False
                )

            with open(gstr_path, "rb") as f:
                outputs["GSTR-1 Workbook.xlsx"] = f.read()

        st.session_state.outputs = outputs
        st.session_state.processed = True
        st.success("Processing completed successfully")

    except Exception as e:
        st.error(str(e))

# ---------------- Downloads ---------------- #

if st.session_state.processed:
    st.subheader("Download Outputs")

    for filename, data in st.session_state.outputs.items():
        st.download_button(
            label=f"Download {filename}",
            data=data,
            file_name=filename,
            key=filename
        )



