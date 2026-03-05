APP_VERSION = "2.3"
import pandas as pd
import glob
import os
import sys
import subprocess
import requests
import webbrowser

from colorama import init, Fore, Style
init(autoreset=True)

try:
    import win32com.client as win32
except ImportError:
    win32com = None
    
from pathlib import Path

def check_for_updates():
    """Checks for a version mismatch and offers to auto-download and restart the script."""
    VERSION_URL = "https://raw.githubusercontent.com/jboyer0000/report-automation/master/version.txt"
    # Ensure this URL points to the raw content of the main script file
    SCRIPT_URL = "https://raw.githubusercontent.com/jboyer0000/report-automation/master/filter_and_email_report.py"
    
    try:
        response = requests.get(VERSION_URL)
        response.raise_for_status()
        latest_version = response.text.strip()
        
        if latest_version != APP_VERSION:
            print(Fore.CYAN + Style.BRIGHT + f"A new version ({latest_version}) is available!")
            choice = input(Fore.YELLOW + "Would you like to auto-update and restart now? (yes/no): ").strip().lower()
            
            if choice == "yes":
                print(Fore.WHITE + "Downloading update from GitHub...")
                new_script = requests.get(SCRIPT_URL)
                new_script.raise_for_status()

                # Determine the path of the script currently being executed
                current_script_path = os.path.realpath(sys.argv[0])
                
                # Overwrite the current file with the new content
                with open(current_script_path, "wb") as f:
                    f.write(new_script.content)
                
                print(Fore.GREEN + Style.BRIGHT + "Update installed successfully! Restarting...")
                
                # Re-launch the script using the current Python interpreter
                os.execv(sys.executable, ['python'] + sys.argv)
    except Exception as e:
        print(Fore.RED + f"Could not complete update. Continuing with current version... Error: {e}")

def get_download_folder():
    return str(Path.home() / "Downloads")

# ======= CONFIGURATION =======
DOWNLOAD_DIR = get_download_folder()
FILE_PATTERN = "xmlRpt*.xls" 
OUTPUT_FILE = os.path.join(DOWNLOAD_DIR, "filtered_report.xlsx")
# =============================

def convert_xls_to_xlsx(xls_path):
    xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"
    excel = win32.Dispatch('Excel.Application')
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(xls_path)
    wb.SaveAs(xlsx_path, FileFormat=51)  # 51 is .xlsx
    wb.Close()
    excel.Quit()
    print(Fore.CYAN + f"Converted {xls_path} to {xlsx_path}")
    return xlsx_path

def find_latest_report(download_dir, pattern):
    file_pattern = os.path.join(download_dir, pattern)
    files = glob.glob(file_pattern)
    if not files:
        print(Fore.YELLOW + "No report files found.")
        return None
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def prompt_filters():
    print(Fore.CYAN + Style.BRIGHT + "\n=== FILTER REPORT PROMPTS ===")
    dispatch = input(Fore.YELLOW + "DispatchZone to filter (leave blank for all): ").strip()

    if dispatch:
        user_defaults = input(Fore.YELLOW + "Use default 'yes' for other filters? (yes/no): ").strip().lower()
        if user_defaults == 'yes':
            hide_blank_r = 'yes'
            hide_driver_data = 'yes'
            signed_blank = 'yes'
            print(Fore.GREEN + "Using default filters 'yes'.")
        else:
            print(Fore.CYAN + "Customizing filters:")
            hide_blank_r = input(Fore.YELLOW + "Hide rows with blank receive scans? (yes/no): ").strip()
            hide_driver_data = input(Fore.YELLOW + "Hide rows with data in Driver? (yes/no): ").strip()
            signed_blank = input(Fore.YELLOW + "Show only blank SignedBy? (yes/no): ").strip()
    else:
        print(Fore.RED + "No DispatchZone entered, please answer manually:")
        hide_blank_r = input(Fore.YELLOW + "Hide rows with blank receive scans? (yes/no): ").strip()
        hide_driver_data = input(Fore.YELLOW + "Hide rows with data in Driver? (yes/no): ").strip()
        signed_blank = input(Fore.YELLOW + "Show only blank SignedBy? (yes/no): ").strip()

    return dispatch, hide_blank_r, hide_driver_data, signed_blank

def apply_filters(df, dispatch, hide_blank_r, hide_driver_data, signed_blank):
    if dispatch:
        df = df[df["DispatchZone"].astype(str).str.contains(dispatch, case=False, na=False)]
    if hide_blank_r.lower() == "yes":
        df = df[~(df["R"].isna() | (df["R"] == ""))]
    if hide_driver_data.lower() == "yes":
        df = df[df["Driver"].isna() | (df["Driver"] == "")]
    if signed_blank.lower() == "yes":
        df = df[df["SignedBy"].isna() | (df["SignedBy"] == "")]
    return df

def create_outlook_email(output_file):
    if win32 is None:
        print(Fore.RED + "Outlook automation unavailable (win32com missing).")
        return
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "Filtered Delivery Report"
    mail.Body = "Please find the filtered report attached."
    mail.Attachments.Add(output_file)
    mail.To = input(Fore.YELLOW + "Enter recipient emails (comma separated): ").strip()
    mail.Display()

def main():
    check_for_updates()
    
    print(Fore.WHITE + f"Using download folder: {DOWNLOAD_DIR}")
    latest_file = find_latest_report(DOWNLOAD_DIR, FILE_PATTERN)
    
    if not latest_file:
        sys.exit(1)

    print(Fore.GREEN + f"Latest report found: {latest_file}")

    ext = os.path.splitext(latest_file)[1].lower()
    if ext == ".xls":
        latest_file = convert_xls_to_xlsx(latest_file)
    
    df = pd.read_excel(latest_file, engine="openpyxl")
    df = df.drop_duplicates(subset=['OrderNumber'])
    
    print(Fore.GREEN + f"Columns found: {', '.join(df.columns)}")

    dispatch, hb_r, hd_d, sb = prompt_filters()
    df_filtered = apply_filters(df, dispatch, hb_r, hd_d, sb)
    
    # Fill NA to prevent sorting issues
    df_filtered["Driver"] = df_filtered["Driver"].fillna("").astype(str)
    df_filtered = df_filtered.sort_values(by="Driver", ascending=True)

    if df_filtered.empty:
        print(Fore.RED + "No records matched your filters.")
    else:
        df_filtered.to_excel(OUTPUT_FILE, index=False)
        print(Fore.GREEN + f"Filtered file saved: {OUTPUT_FILE}")
        
        send_mail = input(Fore.YELLOW + "Send via Outlook? (yes/no): ").strip().lower()
        if send_mail == "yes":
            create_outlook_email(OUTPUT_FILE)
        else:
            os.startfile(OUTPUT_FILE) # Simpler Windows open command
            print(Fore.GREEN + "Opened file in Excel.")

if __name__ == "__main__":
    main()