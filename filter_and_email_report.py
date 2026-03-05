import pandas as pd
import glob
import os
import sys
import subprocess
import requests
import webbrowser
from colorama import init, Fore, Style
from pathlib import Path

# Initialize colorama for Windows terminal colors
init(autoreset=True)

try:
    import win32com.client as win32
except ImportError:
    win32com = None

def check_for_updates():
    """Checks for updates and redirects the user to the GitHub Release page."""
    VERSION_URL = "https://raw.githubusercontent.com/jboyer0000/report-automation/master/version.txt"
    RELEASE_URL = "https://github.com/jboyer0000/report-automation/releases/latest"
    
    try:
        response = requests.get(VERSION_URL, timeout=5)
        response.raise_for_status()
        latest_version = response.text.strip()
        
        if latest_version != APP_VERSION:
            print(Fore.CYAN + Style.BRIGHT + f"\n[UPDATE] A new version ({latest_version}) is available!")
            print(Fore.WHITE + "To update, please download the latest version from GitHub.")
            choice = input(Fore.YELLOW + "Open the download page now? (yes/no): ").strip().lower()
            
            if choice == "yes":
                print(Fore.GREEN + "Opening browser...")
                webbrowser.open(RELEASE_URL)
                print(Fore.WHITE + "Please close this program before installing the new version.")
                input(Fore.WHITE + "Press Enter to continue with the current version or close the window...")
    except Exception as e:
        print(Fore.RED + f"Update check skipped. Error: {e}")

def get_download_folder():
    return str(Path.home() / "Downloads")

# ======= CONFIGURATION =======
DOWNLOAD_DIR = get_download_folder()
FILE_PATTERN = "xmlRpt*.xls*" 
OUTPUT_FILE = os.path.join(DOWNLOAD_DIR, "filtered_report.xlsx")
# =============================

def convert_xls_to_xlsx(xls_path):
    xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"
    excel = win32.Dispatch('Excel.Application')
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(xls_path)
    wb.SaveAs(xlsx_path, FileFormat=51) 
    wb.Close()
    excel.Quit()
    return xlsx_path

def find_latest_report(download_dir, pattern):
    files = glob.glob(os.path.join(download_dir, pattern))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def cleanup_old_reports(download_dir, pattern, output_file):
    print(Fore.CYAN + "\n=== CLEANUP ===")
    confirm = input(Fore.YELLOW + "Delete ALL 'xmlRpt' files and the filtered report? (yes/no): ").strip().lower()
    
    if confirm == 'yes':
        files_to_remove = glob.glob(os.path.join(download_dir, pattern))
        if os.path.exists(output_file):
            files_to_remove.append(output_file)
            
        for f in files_to_remove:
            try:
                os.remove(f)
                print(Fore.WHITE + f"Deleted: {os.path.basename(f)}")
            except Exception as e:
                print(Fore.RED + f"Could not delete {f}: {e}")
        print(Fore.GREEN + "Downloads folder cleaned.")

def prompt_filters():
    print(Fore.CYAN + Style.BRIGHT + "\n=== FILTER REPORT PROMPTS ===")
    dispatch = input(Fore.YELLOW + "DispatchZone to filter (leave blank for all): ").strip()

    if dispatch:
        if input(Fore.YELLOW + "Use default 'yes' for other filters? (yes/no): ").strip().lower() == 'yes':
            return dispatch, 'yes', 'yes', 'yes'
        
    hbr = input(Fore.YELLOW + "Hide blank receive scans? (yes/no): ").strip()
    hdd = input(Fore.YELLOW + "Hide rows with Driver data? (yes/no): ").strip()
    sb = input(Fore.YELLOW + "Show only blank SignedBy? (yes/no): ").strip()

    return dispatch, hbr, hdd, sb

def apply_filters(df, dispatch, hbr, hdd, sb):
    if dispatch:
        df = df[df["DispatchZone"].astype(str).str.contains(dispatch, case=False, na=False)]
    if hbr.lower() == "yes":
        df = df[~(df["R"].isna() | (df["R"] == ""))]
    if hdd.lower() == "yes":
        df = df[df["Driver"].isna() | (df["Driver"] == "")]
    if sb.lower() == "yes":
        df = df[df["SignedBy"].isna() | (df["SignedBy"] == "")]
    return df

def create_outlook_email(output_file):
    if win32 is None:
        print(Fore.RED + "Outlook automation unavailable.")
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
    
    while True:
        print(Fore.CYAN + Style.BRIGHT + "\n" + "="*45)
        print(Fore.WHITE + f"Monitoring Downloads: {DOWNLOAD_DIR}")
        
        latest_file = find_latest_report(DOWNLOAD_DIR, FILE_PATTERN)
        
        if not latest_file:
            print(Fore.RED + "No reports found (xmlRpt*.xls).")
            if input(Fore.YELLOW + "Search again? (Enter for Yes, 'exit' to quit): ").lower() == 'exit':
                break
            continue

        print(Fore.GREEN + f"Found: {os.path.basename(latest_file)}")

        try:
            if latest_file.lower().endswith(".xls"):
                latest_file = convert_xls_to_xlsx(latest_file)
            
            df = pd.read_excel(latest_file, engine="openpyxl")
            df = df.drop_duplicates(subset=['OrderNumber'])
            
            d, h, dr, s = prompt_filters()
            df_filtered = apply_filters(df, d, h, dr, s)
            df_filtered["Driver"] = df_filtered["Driver"].fillna("").astype(str)
            df_filtered = df_filtered.sort_values(by="Driver", ascending=True)

            if df_filtered.empty:
                print(Fore.RED + "No records matched your filters.")
            else:
                df_filtered.to_excel(OUTPUT_FILE, index=False)
                print(Fore.GREEN + f"Filtered report saved: {OUTPUT_FILE}")
                
                if input(Fore.YELLOW + "Send via Outlook? (yes/no): ").lower() == "yes":
                    create_outlook_email(OUTPUT_FILE)
                else:
                    os.startfile(OUTPUT_FILE)
                    print(Fore.GREEN + "Opening in Excel...")

            cleanup_old_reports(DOWNLOAD_DIR, FILE_PATTERN, OUTPUT_FILE)

        except Exception as e:
            print(Fore.RED + f"Error processing report: {e}")

        print(Fore.CYAN + "\n" + "-"*45)
        if input(Fore.YELLOW + "Process another report? (Enter for Yes, 'exit' to quit): ").lower() == 'exit':
            break

if __name__ == "__main__":
    main()