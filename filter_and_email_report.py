import pandas as pd
import glob
import os
import sys
import subprocess

try:
    import win32com.client as win32
except ImportError:
    win32com = None
    
from pathlib import Path

def get_download_folder():
    return str(Path.home() / "Downloads")

# ======= CONFIGURATION =======
DOWNLOAD_DIR = get_download_folder()
print(f"Using download folder: {DOWNLOAD_DIR}")
FILE_PATTERN = "xmlRpt*.xls" #Pattern to find your report file
OUTPUT_FILE = os.path.join(DOWNLOAD_DIR, "filtered_report.xlsx")
# =============================

def convert_xls_to_xlsx(xls_path):
    xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"
    excel = win32.Dispatch('Excel.Application')
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(xls_path)
    wb.SaveAs(xlsx_path, FileFormat=51)  # 51 is xlOpenXMLWorkbook = .xlsx
    wb.Close()
    excel.Quit()
    print(f"Converted {xls_path} to {xlsx_path}")
    return xlsx_path

def find_latest_report(download_dir, pattern):
    import os, glob
    file_pattern = os.path.join(download_dir, download_dir, pattern)
    files = glob.glob(file_pattern)
    if not files:
        print("No report files found.")
        sys.exit(1)
    latest_file = max(files, key=os.path.getmtime)
    print(f"Latest report found: {latest_file}")
    return latest_file

def prompt_filters():
    dispatch = input("DispatchZone to filter (leave blank for all): ").strip()
    r_blank = input("Show only blank R? (yes/no): ").strip()
    signed_blank = input("Show only blank SignedBy? (yes/no): ").strip()
    return dispatch, r_blank, signed_blank

def apply_filters(df, dispatch, r_blank, signed_blank):
    if dispatch:
        df = df[df["DispatchZone"].astype(str).str.contains(dispatch, case=False, na=False)]
    if r_blank.lower() == "yes":
        df = df[df["R"].isna() | (df["R"] == "")]
    if signed_blank.lower() == "yes":
        df = df[df["SignedBy"].isna() | (df["SignedBy"] == "")]
    return df

def save_filtered_excel(df, output_file):
    df.to_excel(output_file, index=False)
    print(f"Filtered file saved as: {output_file}")

def create_outlook_email(output_file):
    if win32 is None:
        print("win32com.client not installed. Outlook email automation is unavailable.")
        return
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "Filtered Delivery Report"
    mail.Body = "Please find the filtered report attached."
    mail.Attachments.Add(output_file)
    mail.To = input("Enter recipient emails (comma separated): ").strip()
    mail.Display()

def main():
    latest_file = find_latest_report(DOWNLOAD_DIR, FILE_PATTERN)
    print(f"Latest report found: {latest_file}")

    ext = os.path.splitext(latest_file)[1].lower()
    if ext == ".xls":
        latest_file = convert_xls_to_xlsx(latest_file)
        df = pd.read_excel(latest_file, engine="openpyxl")
    elif ext == ".xlsx":
        df = pd.read_excel(latest_file, engine="openpyxl")
    else:
        raise ValueError(f"Unsupported file extension: {ext}")
        
    df = df.drop_duplicates(subset=['OrderNumber']) #Remove duplicated OrderNumbers

    print(f"\nColumns found: {', '.join(df.columns)}")

    dispatch, r_blank, signed_blank = prompt_filters()
    df_filtered = apply_filters(df, dispatch, r_blank, signed_blank)
    
    #Sort by Driver ascending
    df_filtered = df_filtered.sort_values(by="Driver", ascending=True)

    if df_filtered.empty:
        print("No records matched your filters.")
    else:
        save_filtered_excel(df_filtered, OUTPUT_FILE)
        send_mail = input("Send via Outlook? (yes/no): ").strip()
        if send_mail.lower() == "yes":
            create_outlook_email(OUTPUT_FILE)
        else:
            try:
                subprocess.Popen(['start', OUTPUT_FILE], shell=True)
                print(f"Opened {OUTPUT_FILE} in Excel.")
            except Exception as e:
                print(f"Could not open the file automatically: {e}")

if __name__ == "__main__":
    main()
