# Sales Summary Report Generator 📊

This is a Python-based desktop application designed to generate sales summary reports for multiple sites by connecting to a SQL Server database. The app uses **Tkinter** for the GUI, supports parallel processing for faster report generation, and allows exporting reports as `.zip` files containing individual `.txt` files for each site.

---

## Features ✨

- **Database Connectivity**: Connects to SQL Server databases using customizable IP addresses or server series (`10.16.x.x` or `10.28.x.x`). 🔗
- **Multi-Site Support**: Generate reports for multiple sites via manual input or bulk upload (Excel, CSV, or text files). 📂
- **Customizable Date Range**: Select a date range for generating sales transaction summaries. 📅
- **Parallel Processing**: Processes multiple sites in parallel for faster report generation. ⚡
- **Export Reports**: Automatically saves reports as `.zip` files containing `.txt` files for each site. 📥
- **Live Logs**: Real-time logging of application activity for better transparency. 📝
- **Error Handling**: Gracefully handles connection errors and invalid inputs with clear error messages. ❌➡️✅

---
D 💡 E 🚀 M 🎯 O 🔍



Uploading Untitled video - Made with Clipchamp.mp4…



## Prerequisites ⚙️

Before running the application, ensure you have the following installed:

- **Python 3.7+** 🐍
- Required Python packages:
  - `tkinter`
  - `tkcalendar`
  - `pyodbc`
  - `pandas`
  - `openpyxl` (for Excel file support)
  - `concurrent.futures` (built-in)
- **ODBC Driver 17 for SQL Server** 🖥️
- Access to the SQL Server database with valid credentials 🔑

Install dependencies using the following command:

```bash
pip install pyodbc pandas openpyxl tkcalendar
```
# Installation 💻
1. Clone the repository to your local machine:
```bash
git clone https://github.com/akashsg247777/Apollo-Sqles-Report.git
```
2. Navigate to the project directory:
```bash
cd sales-summary-report-generator
```
3. Install the required dependencies:
```bash
pip install -r requirements.txt
```
4. Run the application:
```bash
python main.py
```

# Usage 🚀
**Step-by-Step Guide:**

1. Launch the Application : Run the script, and the GUI will appear. 🖥️
2. Enter Database Credentials :
   * Provide the username, password, and database name. 🛠️
   * Optionally, enter a custom IP address for manual overrides. 🌐
3. Select Site IDs :
   * Manually input a single site ID or upload a file containing multiple site IDs. 📋
4. Set Date Range :
   * Use the calendar widget to select the "From Date" and "To Date". 📅
5 Generate Reports :
   * Click the "Generate Report" button to start processing. ⚙️
   * Monitor the live log for progress updates. 📊
6. Download Reports :
   * Once completed, click the "Download Report" button to save the .zip file. 📥

Feel free to reach out if you have any questions or suggestions! 

📧akashsg247@gmail.com

📞+91 8618041675
