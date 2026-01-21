import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# Global variables for file paths
sd_file = None
sr_file = None
tb_file = None
gl_file = None

# Function to select SD file
def select_sd():
    global sd_file
    sd_file = filedialog.askopenfilename(title="Select SD File", filetypes=[("Excel files", "*.xlsx")])
    if sd_file:
        label_sd.config(text=f"SD File: {os.path.basename(sd_file)}")

# Function to select SR file
def select_sr():
    global sr_file
    sr_file = filedialog.askopenfilename(title="Select SR File", filetypes=[("Excel files", "*.xlsx")])
    if sr_file:
        label_sr.config(text=f"SR File: {os.path.basename(sr_file)}")

# Function to select TB file
def select_tb():
    global tb_file
    tb_file = filedialog.askopenfilename(title="Select TB File", filetypes=[("Excel files", "*.xlsx")])
    if tb_file:
        label_tb.config(text=f"TB File: {os.path.basename(tb_file)}")

# Function to select GL Dump file
def select_gl():
    global gl_file
    gl_file = filedialog.askopenfilename(title="Select GL Dump File", filetypes=[("Excel files", "*.xlsx")])
    if gl_file:
        label_gl.config(text=f"GL Dump File: {os.path.basename(gl_file)}")

# Function to process the files
def process():
    xxxx = entry_xxxx.get().strip()
    if not xxxx:
        messagebox.showerror("Error", "Please enter the Company Code (XXXX).")
        return
    if not all([sd_file, sr_file, tb_file, gl_file]):
        messagebox.showerror("Error", "Please select all files.")
        return
    
    try:
        # Get the directory from the first file (assuming all are in the same directory)
        dir_path = os.path.dirname(sd_file)
        
        # Step 1: Consolidate SD and SR files
        df_sd = pd.read_excel(sd_file)
        df_sr = pd.read_excel(sr_file)
        # Concatenate: Keep header from SD, append data from SR (excluding its header)
        df_consolidated = pd.concat([df_sd, df_sr.iloc[1:]], ignore_index=True)
        consolidated_file = os.path.join(dir_path, f"{xxxx}_SD_SR_Consolidated.xlsx")
        df_consolidated.to_excel(consolidated_file, index=False)
        
        # Step 2: Process GL Dump file
        df_gl = pd.read_excel(gl_file)
        
        # Filter for GST Payable
        gst_conditions = df_gl['G/L Account: Long Text'].isin(['Central GST Payable', 'Integrated GST Payable', 'State GST Payable'])
        df_gst = df_gl[gst_conditions]
        
        # Filter for Revenue (G/L Account starting with '3')
        revenue_conditions = df_gl['G/L Account'].astype(str).str.startswith('3')
        df_revenue = df_gl[revenue_conditions]
        
        # Create new workbook with GST Payable and Revenue sheets
        gstr_workbook = os.path.join(dir_path, f"{xxxx}_GSTR-1_Workbook.xlsx")
        with pd.ExcelWriter(gstr_workbook) as writer:
            df_gst.to_excel(writer, sheet_name='GST Payable', index=False)
            df_revenue.to_excel(writer, sheet_name='Revenue', index=False)
        
        # Step 3: Create summary
        # Summary from GST Payable sheet
        gst_summary = df_gst.groupby('G/L Account: Long Text')['Company Code Currency Value'].sum()
        
        # Summary from TB file
        df_tb = pd.read_excel(tb_file)
        tb_conditions = df_tb['G/L Account: Long Text'].isin(['Central GST Payable', 'Integrated GST Payable', 'State GST Payable'])
        df_tb_gst = df_tb[tb_conditions]
        df_tb_gst['Difference'] = df_tb_gst['Period 09 C'] - df_tb_gst['Period 09 D']
        tb_summary = df_tb_gst.groupby('G/L Account: Long Text')['Difference'].sum()
        
        # Create summary DataFrame with differences
        summary_df = pd.DataFrame({
            'GST Type': gst_summary.index,
            'GST Payable Amount': gst_summary.values,
            'TB Difference': tb_summary.reindex(gst_summary.index).values,
            'Net Difference': gst_summary.values - tb_summary.reindex(gst_summary.index).values
        })
        summary_file = os.path.join(dir_path, f"{xxxx}_Summary.xlsx")
        summary_df.to_excel(summary_file, index=False)
        
        messagebox.showinfo("Success", f"Processing complete. Files saved in {dir_path}")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main window
root = tk.Tk()
root.title("Excel File Processor")
root.geometry("400x400")

# Company Code input
tk.Label(root, text="Company Code (XXXX):").pack(pady=5)
entry_xxxx = tk.Entry(root)
entry_xxxx.pack(pady=5)

# File selection buttons and labels
tk.Button(root, text="Select SD File", command=select_sd).pack(pady=5)
label_sd = tk.Label(root, text="No SD file selected")
label_sd.pack()

tk.Button(root, text="Select SR File", command=select_sr).pack(pady=5)
label_sr = tk.Label(root, text="No SR file selected")
label_sr.pack()

tk.Button(root, text="Select TB File", command=select_tb).pack(pady=5)
label_tb = tk.Label(root, text="No TB file selected")
label_tb.pack()

tk.Button(root, text="Select GL Dump File", command=select_gl).pack(pady=5)
label_gl = tk.Label(root, text="No GL Dump file selected")
label_gl.pack()

# Process button
tk.Button(root, text="Process Files", command=process).pack(pady=20)

# Run the GUI
root.mainloop()
