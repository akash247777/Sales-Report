# Sales Summary Report Generator 📊

This is a Python-based desktop application designed to generate sales summary reports for multiple sites by connecting to a SQL Server database. The app uses **Tkinter** for the GUI, supports parallel processing for faster report generation, and allows exporting reports as `.zip` files containing individual `.txt` files for each site.

---

## Table of Contents 📑

1. [Features ✨](#features-✨)
2. [Prerequisites ⚙️](#prerequisites-⚙️)
3. [Installation 💻](#installation-💻)
4. [Usage 🚀](#usage-🚀)
5. [Contributing 🤝](#contributing-🤝)
6. [License 📜](#license-📜)

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
