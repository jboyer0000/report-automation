# Report Automation Tool

**Automated Excel Report Filtering & Email Tool** for processing daily delivery reports.

[![Version](https://img.shields.io/badge/version-2.9-blue.svg)](https://github.com/jboyer0000/report-automation/releases)

---

## Overview
This tool automates the daily processing of Ecourier delivery reports. It detects the latest `xmlRpt*.xls` file in your Downloads folder, applies operational filters, and packages the final `.xlsx` file for immediate Outlook distribution.

---

## Installation & Setup
** Version 2.9 introduces an autonomous auto-update engine. This is the last time you will need to manually download an update.**

1. Navigate to the [**Releases page**](https://github.com/jboyer0000/report-automation/releases).
2. Download `Report_Automation_V*.zip` from the Assets section.
3. **Crucial:** Extract the `.zip` file into a dedicated folder on your Desktop. 
4. Double-click `filter_and_email_report.exe` to launch the tool.

*Note: Windows Defender may flag the executable the first time it runs. Click **More info**, then **Run anyway**. Once authorized, it will execute normally moving forward.*

---

## Standard Workflow
1. Run your web query on Ecourier. Click **Save** (do not click Open) so the file lands in your standard Downloads folder.
2. Launch the tool.
3. **Answer the prompts**:
   - Enter a DispatchZone (or leave blank for all).
   - Use the default 'yes' shortcut, or manually configure the remaining filters.
4. **Choose output**:
   - Send directly via Outlook (auto-attaches the Excel file).
   - Open directly in Excel for manual review.

---

## Filter Definitions
| Filter | Purpose |
|--------|---------|
| **DispatchZone** | Isolates rows matching a specific zone (e.g., "100", "700"). |
| **Hide blank receive scans** | Shows only confirmed physical receive scans. |
| **Hide Driver data** | Shows only unassigned freight. |
| **Show blank SignedBy** | Keeps only rows missing a signature. |