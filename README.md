# 📊 Report Automation Tool

**Automated Excel Report Filtering & Email Tool** for processing delivery reports with customizable filters and Outlook integration.

[![Version](https://img.shields.io/badge/version-2.3-blue.svg)](https://github.com/jboyer0000/report-automation/releases)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()
[![License](https://img.shields.io/badge/license-MIT-green.svg)]()

---

## 🎯 Overview

This tool automates the process of filtering delivery reports by automatically detecting the latest report file in your Downloads folder, applying custom filters, and optionally sending the filtered results via Outlook email.

Perfect for teams that regularly process `xmlRpt*.xls` reports and need to quickly filter by dispatch zones, signed deliveries, driver assignments, and receive scans.

---

## ✨ Features

### 🔍 **Smart Report Detection**
- Automatically finds the latest `xmlRpt*.xls` file in your Downloads folder
- Converts legacy `.xls` files to `.xlsx` format automatically
- Removes duplicate orders for clean data

### ⚙️ **Flexible Filtering Options**
- **Filter by DispatchZone** - Focus on specific zones
- **Hide blank receive scans** - Show only received items
- **Filter out Driver data** - Remove assigned deliveries
- **Show unsigned deliveries** - Display only blank SignedBy fields
- **Quick defaults** - One-click filtering with common settings

### 📧 **Outlook Integration**
- Generate polished Excel reports
- Auto-create Outlook emails with attachments
- Customizable recipient lists
- Fallback: Open report directly in Excel

### 🔄 **Auto-Update System**
- Checks for new versions on startup
- Notifies users when updates are available
- One-click access to latest releases

---

## 📦 Installation

### **Option 1: Download Executable (Recommended for Users)**

1. Go to the [**Releases page**](https://github.com/jboyer0000/report-automation/releases)
2. Download the latest `filter_and_email_report.exe`
3. Run the executable - **no installation required!**

### **Option 2: Run from Source (For Developers)**

```bash
# Clone the repository
git clone https://github.com/jboyer0000/report-automation.git
cd report-automation

# Install dependencies
pip install -r requirements.txt

# Run the script
python filter_and_email_report.py
```

---

## 🚀 Usage

### **Quick Start**

1. **Download your report** - Ensure `xmlRpt*.xls` file is in your Downloads folder
2. **Run the tool** - Double-click `filter_and_email_report.exe`
3. **Answer the prompts**:
   - Enter a DispatchZone to filter (or leave blank for all)
   - Choose filter options (or use defaults)
4. **Choose output method**:
   - Send via Outlook email, or
   - Open directly in Excel

### **Example Workflow**

```
📂 Downloads folder contains: xmlRpt_2024-11-18.xls

🔄 Tool automatically:
   ✓ Finds latest report
   ✓ Converts to .xlsx
   ✓ Removes duplicates

❓ User chooses filters:
   • DispatchZone: "100, 200, etc"
   • Hide blank receive scans: yes
   • Hide Driver data: yes
   • Show blank SignedBy: yes

📊 Result:
   ✓ Filtered report saved as "filtered_report.xlsx"
   ✓ Outlook email created with attachment
   ✓ Ready to send!
```

---

## 🎨 Filter Options Explained

| Filter | What It Does | When to Use |
|--------|-------------|-------------|
| **DispatchZone** | Shows only rows matching a specific zone (e.g., "100", "700") | Focus on your hub |
| **Hide blank receive scans** | Removes rows where the "R" (receive) column is empty | Show only confirmed deliveries |
| **Hide Driver data** | Removes rows where Driver field has data | Show only unassigned deliveries |
| **Show blank SignedBy** | Keeps only rows where SignedBy is empty | Find unsigned/incomplete deliveries |

---

## 💻 System Requirements

- **OS:** Windows 10/11
- **Excel:** Microsoft Excel (for .xls conversion and viewing)
- **Outlook:** Microsoft Outlook (optional - for email features)
- **Internet:** Required for auto-update checks

---

## 🛠️ For Developers

### **Project Structure**

```
report-automation/
├── filter_and_email_report.py   # Main script
├── filter_and_email_report.spec # PyInstaller build config
├── version.txt                  # Version tracking
├── tests/                       # Unit tests
│   └── test_filters.py
└── dist/                        # Compiled executable
    └── filter_and_email_report.exe
```

### **Build Executable**

```bash
# Using the spec file (recommended)
pyinstaller filter_and_email_report.spec --clean

# Output: dist/filter_and_email_report.exe
```

### **Dependencies**

- `pandas` - Data processing
- `openpyxl` - Excel file handling
- `pywin32` - Windows COM automation (Excel, Outlook)
- `requests` - Update checking
- `colorama` - Console colors

---

## 🐛 Troubleshooting

### **"No report files found"**
- Ensure `xmlRpt*.xls` files are in your Downloads folder
- Check file naming matches the pattern

### **"Excel conversion failed"**
- Make sure Microsoft Excel is installed
- Try opening the file manually in Excel first

### **"Outlook email failed"**
- Verify Outlook is installed and configured
- Use the "Open in Excel" option as alternative

### **Update check fails**
- Check your internet connection
- The tool will continue without update check if offline

---

## 📝 Changelog

### **Version 2.3** (Current)
- 🐛 Fixed auto-updater 404 error
- 🔧 Updated version check to use GitHub raw content
- 🎨 Improved error messaging with colors

### **Version 2.2**
- Added auto-update functionality
- Enhanced filter prompts
- Bug fixes and improvements

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

---

## 💬 Support

Having issues? [Open an issue](https://github.com/jboyer0000/report-automation/issues) on GitHub.

---

## 👨‍💻 Author

**jboyer0000**

- GitHub: [@jboyer0000](https://github.com/jboyer0000)
- Repository: [report-automation](https://github.com/jboyer0000/report-automation)

---

<div align="center">
</div>

