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
    VERSION_URL = "https://raw.githubusercontent.com/jboyer0000/report-automation/master/version.txt"
    try:
        response = requests.get(VERSION_URL)
        response.raise_for_status()
        latest_version = response.text.strip()
        if latest_version != APP_VERSION:
            print(f"A new version ({latest_version}) is available! Please update.")
            choice = input("Open download page? (yes/no): ").strip().lower()
            if choice == "yes":
                webbrowser.open("https://github.com/jboyer0000/report-automation/releases/latest")
    except Exception as e:
        print(Fore.RED + Style.BRIGHT + f"Could not check for updates. Continuing... Error: {e}")

def get_download_folder():
    return str(Path.home() / "Downloads")

# ======= CONFIGURATION =======
DOWNLOAD_DIR = get_download_folder()
print(Fore.WHITE + f"Using download folder: {DOWNLOAD_DIR}")
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
    print(Fore.CYAN + f"Converted {xls_path} to {xlsx_path}")
    return xlsx_path

def find_latest_report(download_dir, pattern):
    import os, glob
    file_pattern = os.path.join(download_dir, download_dir, pattern)
    files = glob.glob(file_pattern)
    if not files:
        print(Fore.YELLOW + "No report files found.")
        sys.exit(1)
    latest_file = max(files, key=os.path.getmtime)
    print(Fore.GREEN + f"Latest report found: {latest_file}")
    return latest_file

def prompt_filters():
    print(Fore.CYAN + Style.BRIGHT + "=== FILTER REPORT PROMPTS ===")
    dispatch = input(Fore.YELLOW + "DispatchZone to filter (leave blank for all): ").strip()

    if dispatch:
        user_defaults = input(Fore.YELLOW + "Use default 'yes' for other filters? (yes/no): ").strip().lower()
        if user_defaults == 'yes':
            hide_blank_r = 'yes'
            hide_driver_data = 'yes'
            signed_blank = 'yes'
            print(Fore.GREEN + "Using default filters 'yes.")
        else:
            print(Fore.CYAN + "Customizing filters:")
            hide_blank_r = input(Fore.YELLOW + "Hide rows with blank receive scans? (yes or no): ").strip()
            hide_driver_data = input(Fore.YELLOW + "Hide rows with data in Driver? (yes/no): ").strip()
            signed_blank = input(Fore.YELLOW + "Show only blank SignedBy? (yes/no): ").strip()
    else:
									  
																						 
																					 
																			
			 
										 
																							 
																						 
																				
		 
        print(Fore.RED + "No DispatchZone entered, please answer the following filter questions.")
        hide_blank_r = input(Fore.YELLOW + "Hide rows with blank receive scans? (yes or no): ").strip()
        hide_driver_data = input(Fore.YELLOW + "Hide rows with data in Driver? ").strip()
        signed_blank = input(Fore.YELLOW + "Show only blank SignedBy? ").strip()

    return dispatch, hide_blank_r, hide_driver_data, signed_blank


def apply_filters(df, dispatch, hide_blank_r, hide_driver_data, signed_blank):
    if dispatch:
        df = df[df["DispatchZone"].astype(str).str.contains(dispatch, case=False, na=False)]
    if hide_blank_r.lower() == "yes":
        df = df[~(df["R"].isna() | (df["R"] == ""))]  # exclude blank R rows
    if hide_driver_data.lower() == "yes":
        # Keep only rows where Driver is blank
        df = df[df["Driver"].isna() | (df["Driver"] == "")]
    if signed_blank.lower() == "yes":
        df = df[df["SignedBy"].isna() | (df["SignedBy"] == "")]
    return df

def save_filtered_excel(df, output_file):
    df.to_excel(output_file, index=False)
    print(Fore.GREEN + f"Filtered file saved as: {output_file}")

def create_outlook_email(output_file):
    if win32 is None:
        print(Fore.RED + "win32com.client not installed. Outlook email automation is unavailable.")
        return
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "Filtered Delivery Report"
    mail.Body = "Please find the filtered report attached."
    mail.Attachments.Add(output_file)
    mail.To = input(Fore.YELLOW + "Enter recipient emails (comma separated): ").strip()
    mail.Display()

def main():
    should_exit = False

    while not should_exit:
        latest_file = find_latest_report(DOWNLOAD_DIR, FILE_PATTERN)
        print(Fore.GREEN + f"Latest report found: {latest_file}")

        ext = os.path.splitext(latest_file)[1].lower()
        if ext == ".xls":
            latest_file = convert_xls_to_xlsx(latest_file)
            df = pd.read_excel(latest_file, engine="openpyxl")
        elif ext == ".xlsx":
            df = pd.read_excel(latest_file, engine="openpyxl")
        else:
            raise ValueError(f"Unsupported file extension: {ext}")

        df = df.drop_duplicates(subset=['OrderNumber'])  # Remove duplicates
        print(Fore.GREEN + f"\nColumns found: {', '.join(df.columns)}")

        if should_exit:
            break

        dispatch, hide_blank_r, hide_driver_data, signed_blank = prompt_filters()

        df_filtered = apply_filters(df, dispatch, hide_blank_r, hide_driver_data, signed_blank)
        df_filtered["Driver"] = df_filtered["Driver"].fillna("").astype(str)
        df_filtered = df_filtered.sort_values(by="Driver", ascending=True)

        if df_filtered.empty:
            print(Fore.RED + "No records matched your filters.")
            retry = input(Fore.YELLOW + "Please download the new report file and press Enter to try again or type 'exit' to quit: ").strip().lower()
            if retry == 'exit':
                print(Fore.RED + "Exiting.")
                should_exit = True
                break
            else:
                print(Fore.YELLOW + "Retrying with new report file...")
                continue
        else:
            save_filtered_excel(df_filtered, OUTPUT_FILE)
            send_mail = input(Fore.YELLOW + "Send via Outlook? (yes/no): ").strip()
            if send_mail.lower() == "yes":
                create_outlook_email(OUTPUT_FILE)
            else:
                try:
                    subprocess.Popen(['start', OUTPUT_FILE], shell=True)
                    print(Fore.GREEN + f"Opened {OUTPUT_FILE} in Excel.")
                except Exception as e:
                    print(Fore.RED + f"Could not open the file automatically: {e}")
            should_exit = True

if __name__ == "__main__":
    check_for_updates()
    main()