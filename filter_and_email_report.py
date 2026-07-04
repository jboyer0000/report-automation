APP_VERSION = "2.9"
import zipfile
import io
import time
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
    
def clean_old_updates():
    """Silently deletes old executable versions left behind by the updater."""
    if not getattr(sys, 'frozen', False):
        return # Skip if running as a standard .py script in development
        
    current_dir = os.path.dirname(sys.executable)
    for file in os.listdir(current_dir):
        if file.endswith("_OLD.exe"):
            old_file_path = os.path.join(current_dir, file)
            
            # Retry loop to overcome OS file lock race conditions
            for _ in range(5):
                try:
                    os.remove(old_file_path)
                    break # Success, break the loop
                except PermissionError:
                    time.sleep(0.5) # Wait for Windows to kill the old process
                except Exception:
                    pass

def check_for_updates():
    VERSION_URL = "https://raw.githubusercontent.com/jboyer0000/report-automation/master/version.txt"
    API_URL = "https://api.github.com/repos/jboyer0000/report-automation/releases/latest"
    
    try:
        response = requests.get(VERSION_URL, timeout=5)
        response.raise_for_status()
        latest_version = response.text.strip()
        
        if float(latest_version) > float(APP_VERSION):
            print(Fore.CYAN + Style.BRIGHT + f"\n[UPDATE] Version {latest_version} is available!")
            choice = input(Fore.YELLOW + "Would you like to auto-update now? (yes/no): ").strip().lower()
            
            if choice == "yes":
                if getattr(sys, 'frozen', False):
                    print(Fore.GREEN + "Initializing automatic update...")
                    
                    # 1. Fetch download URL via GitHub API
                    api_response = requests.get(API_URL, timeout=5).json()
                    
                    download_url = None
                    for asset in api_response.get('assets', []):
                        if asset['name'].endswith('.zip'):
                            download_url = asset['browser_download_url']
                            break
                            
                    if not download_url:
                        raise Exception("Critical: No .zip archive found in the latest release assets.")
                    
                    # 2. Download zip into memory
                    print(Fore.WHITE + "Downloading update...")
                    zip_response = requests.get(download_url)
                    zip_data = zipfile.ZipFile(io.BytesIO(zip_response.content))
                    
                    # 3. Environment mapping
                    current_exe = sys.executable
                    current_dir = os.path.dirname(current_exe)
                    exe_name = os.path.basename(current_exe)
                    
                    # 4. Release locks by killing AHK monitor
                    subprocess.run(["taskkill", "/F", "/IM", "AutoClickSave.exe"], capture_output=True)
                    time.sleep(0.5)
                    
                    # 5. The Rename-and-Replace bypass
                    old_exe = os.path.join(current_dir, exe_name.replace(".exe", "_OLD.exe"))
                    if os.path.exists(old_exe):
                        os.remove(old_exe) 
                    os.rename(current_exe, old_exe)
                    
                    # 6. Extract new files
                    print(Fore.WHITE + "Extracting new files...")
                    zip_data.extractall(current_dir)
                    
                    # 7. Handoff
                    print(Fore.GREEN + "Update complete. Rebooting terminal...")
                    time.sleep(1)
                    subprocess.Popen([os.path.join(current_dir, "filter_and_email_report.exe")])
                    sys.exit() 
                else:
                    print(Fore.RED + "Auto-update disabled in uncompiled development environment.")
                    
    except Exception as e:
        print(Fore.RED + f"Auto-update sequence failed: {e}")
        
        # 1. Attempt to revert the executable rename if it occurred
        try:
            if 'old_exe' in locals() and 'current_exe' in locals():
                if os.path.exists(old_exe) and not os.path.exists(current_exe):
                    os.rename(old_exe, current_exe)
        except Exception:
            pass # If reverting fails, pass silently
            
        # 2. Force termination to prevent zombie execution
        input(Fore.YELLOW + "Press Enter to close the program and try again...")
        sys.exit()

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

    # The shortcut is now decoupled and offered universally
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
    
def launch_ahk_monitor():
    ahk_exe_path = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), "AutoClickSave.exe")
    
    if os.path.exists(ahk_exe_path):
        process = subprocess.Popen([ahk_exe_path])
        print(Fore.GREEN + "Background terminal monitor initialized.")
        return process # Return the process object so we can control it later
    else:
        print(Fore.RED + f"Execution Failed: {ahk_exe_path} not found.")
        return None

def main():
    clean_old_updates()
    check_for_updates()
    
    # Store the process variable
    ahk_process = launch_ahk_monitor()
    
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
            # kill the background process before breaking the loop
            if ahk_process:
                ahk_process.terminate()
                print(Fore.YELLOW + "Background monitor terminated.")
            break

if __name__ == "__main__":
    main()