import streamlit as st
import pandas as pd
import os
import tempfile

st.set_page_config(page_title="GSTR-1 Processor", layout="centered")

st.title("Excel File Processor (GSTR-1)")

xxxx = st.text_input("Company Code (XXXX)")

sd_file = st.file_uploader("Upload SD File", type=["xlsx"])
sr_file = st.file_uploader("Upload SR File", type=["xlsx"])
tb_file = st.file_uploader("Upload TB File", type=["xlsx"])
gl_file = st.file_uploader("Upload GL Dump File", type=["xlsx"])

if st.button("Process Files"):
    if not xxxx:
        st.error("Company Code is required")
        st.stop()

    if not all([sd_file, sr_file, tb_file, gl_file]):
        st.error("All files must be uploaded")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            sd_path = os.path.join(tmpdir, "sd.xlsx")
            sr_path = os.path.join(tmpdir, "sr.xlsx")
            tb_path = os.path.join(tmpdir, "tb.xlsx")
            gl_path = os.path.join(tmpdir, "gl.xlsx")

            for file, path in [
                (sd_file, sd_path),
                (sr_file, sr_path),
                (tb_file, tb_path),
                (gl_file, gl_path),
            ]:
                with open(path, "wb") as f:
                    f.write(file.getbuffer())

            # Step 1: Consolidate SD & SR
            df_sd = pd.read_excel(sd_path)
            df_sr = pd.read_excel(sr_path)
            df_consolidated = pd.concat([df_sd, df_sr.iloc[1:]], ignore_index=True)

            consolidated_path = os.path.join(tmpdir, f"{xxxx}_SD_SR_Consolidated.xlsx")
            df_consolidated.to_excel(consolidated_path, index=False)

            # Step 2: GL Processing
            df_gl = pd.read_excel(gl_path)

            gst_accounts = [
                "Central GST Payable",
                "Integrated GST Payable",
                "State GST Payable",
            ]

            df_gst = df_gl[df_gl["G/L Account: Long Text"].isin(gst_accounts)]
            df_revenue = df_gl[df_gl["G/L Account"].astype(str).str.startswith("3")]

            gstr_path = os.path.join(tmpdir, f"{xxxx}_GSTR-1_Workbook.xlsx")
            with pd.ExcelWriter(gstr_path) as writer:
                df_gst.to_excel(writer, sheet_name="GST Payable", index=False)
                df_revenue.to_excel(writer, sheet_name="Revenue", index=False)

            # Step 3: Summary
            gst_summary = (
                df_gst.groupby("G/L Account: Long Text")["Company Code Currency Value"]
                .sum()
            )

            df_tb = pd.read_excel(tb_path)
            df_tb_gst = df_tb[df_tb["G/L Account: Long Text"].isin(gst_accounts)]
            df_tb_gst["Difference"] = df_tb_gst["Period 09 C"] - df_tb_gst["Period 09 D"]

            tb_summary = (
                df_tb_gst.groupby("G/L Account: Long Text")["Difference"].sum()
            )

            summary_df = pd.DataFrame({
                "GST Type": gst_summary.index,
                "GST Payable Amount": gst_summary.values,
                "TB Difference": tb_summary.reindex(gst_summary.index).values,
            })

            summary_df["Net Difference"] = (
                summary_df["GST Payable Amount"] - summary_df["TB Difference"]
            )

            summary_path = os.path.join(tmpdir, f"{xxxx}_Summary.xlsx")
            summary_df.to_excel(summary_path, index=False)

            st.success("Processing completed")

            for label, path in {
                "SD-SR Consolidated": consolidated_path,
                "GSTR-1 Workbook": gstr_path,
                "Summary": summary_path,
            }.items():
                with open(path, "rb") as f:
                    st.download_button(label, f, file_name=os.path.basename(path))

        except Exception as e:
            st.error(str(e))
