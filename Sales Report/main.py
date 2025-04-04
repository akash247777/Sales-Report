import tkinter as tk  # Import the tkinter module for creating the GUI
from tkinter import ttk, messagebox, filedialog  # Import additional tkinter modules for widgets and dialogs
from tkinter.scrolledtext import ScrolledText  # Import ScrolledText for a scrollable text area
from tkcalendar import DateEntry  # Import DateEntry for date selection (requires tkcalendar package)
import pyodbc  # Import pyodbc for database connection
from datetime import datetime  # Import datetime for date and time operations
from decimal import Decimal  # Import Decimal for precise decimal calculations
import io  # Import io for in-memory file operations
import zipfile  # Import zipfile for creating ZIP archives
import pandas as pd  # Import pandas for data manipulation
import concurrent.futures  # Import concurrent.futures for parallel processing
import os  # Import os for file operations
import threading  # Import threading for running tasks in separate threads

def format_currency(value):
    """Format a numeric value as a currency string with 2 decimals."""
    return f"{value:,.2f}"  # Format the value as a string with commas and 2 decimal places

def format_report(result, site_id, site_name, from_date, to_date):
    """
    Process the query result (a list of tuples) and produce a text report
    in which every line is fixed to 180 characters.
    """
    PAGE_WIDTH = 180  # Define the width of each line in the report

    def fix_line(line, width=PAGE_WIDTH):
        clean = line.rstrip("\n")  # Remove trailing newline characters
        if len(clean) < width:
            return clean + " " * (width - len(clean))  # Pad the line with spaces to reach the desired width
        else:
            return clean[:width]  # Truncate the line to the desired width

    now = datetime.now()  # Get the current date and time
    header_date = now.strftime("%d/%m/%Y")  # Format the date as DD/MM/YYYY
    header_time = now.strftime("%I:%M %p")  # Format the time as HH:MM AM/PM

    lines = []  # Initialize a list to store the lines of the report
    # Header Section.
    lines.append(f"DATE: {header_date}".rjust(PAGE_WIDTH))  # Add the date to the report, right-aligned
    lines.append(f"TIME: {header_time}".rjust(PAGE_WIDTH))  # Add the time to the report, right-aligned
    lines.append("")  # Add a blank line
    lines.append("APOLLO PHARMACIES LIMITED".center(PAGE_WIDTH))  # Add the company name, centered
    lines.append(f"{site_id} - {site_name}".center(PAGE_WIDTH))  # Add the site ID and name, centered
    lines.append("")  # Add a blank line
    lines.append("Sales Transaction Summary Report".center(PAGE_WIDTH))  # Add the report title, centered
    lines.append(f"From Date : {from_date}    To Date : {to_date}".center(PAGE_WIDTH))  # Add the date range, centered
    lines.append("-" * PAGE_WIDTH)  # Add a separator line

    header_groups = (
        "|" +
        " SALES ".center(55) +
        "|" +
        " RETURNS ".center(55) +
        "|" +
        " NET ".center(55) +
        "|"
    )  # Define the header groups for the report
    lines.append(header_groups)  # Add the header groups to the report
    lines.append("-" * PAGE_WIDTH)  # Add a separator line

    header_cols = (
        f"{'BILLTYPE':<17} |"
        f"{'NO':>8} |"
        f"{'AMT':>12} |"
        f"{'DISC':>12} |"
        f"{'NET':>12} |"
        f"{'NO':>6} |"
        f"{'AMT':>12} |"
        f"{'DISC':>12} |"
        f"{'NET':>12} |"
        f"{'NO':>6} |"
        f"{'AMT':>12} |"
        f"{'DISC':>12} |"
        f"{'NET':>12} |"
    )  # Define the header columns for the report
    lines.append(header_cols)  # Add the header columns to the report
    lines.append("-" * PAGE_WIDTH)  # Add a separator line

    # Process data rows.
    sales_data = []  # Initialize a list to store sales data
    partner_data = []  # Initialize a list to store partner data

    for row in result:
        isheader = row[0]  # Get the value of the first column
        if isheader in (1, 3):
            sale_net = row[3]  # Get the net sales amount
            sale_disc = row[4]  # Get the sales discount
            ret_net = row[5]  # Get the net returns amount
            ret_disc = row[6]  # Get the returns discount
            sales_data.append({
                "BILLTYPE": row[2],  # Get the bill type
                "SALECOUNT": row[7],  # Get the sales count
                "SALE_NET": row[3],  # Get the net sales amount
                "SALE_DISC": row[4],  # Get the sales discount
                "SALE_AMT": sale_net + sale_disc,  # Calculate the total sales amount
                "RETCOUNT": row[8],  # Get the returns count
                "RET_NET": ret_net,  # Get the net returns amount
                "RET_DISC": ret_disc,  # Get the returns discount
                "RET_AMT": ret_net + ret_disc  # Calculate the total returns amount
            })  # Add the sales data to the list
        elif isheader == 0:
            partner_data.append({
                "NAME": row[2],  # Get the partner name
                "BILLCNT": row[7],  # Get the bill count
                "AMOUNT": row[3]  # Get the amount
            })  # Add the partner data to the list

    tot_sale_count = tot_sale_amt = tot_sale_disc = tot_sale_net = 0  # Initialize totals for sales
    tot_ret_count = tot_ret_amt = tot_ret_disc = tot_ret_net = 0  # Initialize totals for returns
    net_cash_sales = 0  # Initialize net cash sales

    for s in sales_data:
        if s["BILLTYPE"].upper() != "GIFT":
            tot_sale_count   += s["SALECOUNT"]  # Add the sales count to the total
            tot_sale_amt     += float(s["SALE_AMT"])  # Add the sales amount to the total
            tot_sale_disc    += float(s["SALE_DISC"])  # Add the sales discount to the total
            tot_sale_net     += float(s["SALE_NET"])  # Add the net sales amount to the total
            tot_ret_count    += s["RETCOUNT"]  # Add the returns count to the total
            tot_ret_amt      += float(s["RET_AMT"])  # Add the returns amount to the total
            tot_ret_disc     += float(s["RET_DISC"])  # Add the returns discount to the total
            tot_ret_net      += float(s["RET_NET"])  # Add the net returns amount to the total
        if s["BILLTYPE"].upper() == "CASH":
            net_cash_sales = s["SALE_NET"] + s["RET_NET"]  # Calculate the net cash sales

    tot_overall_count = tot_sale_count + tot_ret_count  # Calculate the total overall count
    tot_overall_amt   = tot_sale_amt + tot_ret_amt  # Calculate the total overall amount
    tot_overall_disc  = tot_sale_disc + tot_ret_disc  # Calculate the total overall discount
    tot_overall_net   = tot_sale_net + tot_ret_net  # Calculate the total overall net amount

    for s in sales_data:
        overall_count = s["SALECOUNT"] + s["RETCOUNT"]  # Calculate the overall count for the row
        overall_amt   = s["SALE_AMT"] + s["RET_AMT"]  # Calculate the overall amount for the row
        overall_disc  = s["SALE_DISC"] + s["RET_DISC"]  # Calculate the overall discount for the row
        overall_net   = s["SALE_NET"] + s["RET_NET"]  # Calculate the overall net amount for the row
        row_line = (
            f"{s['BILLTYPE']:<17} |"
            f"{s['SALECOUNT']:8d} |"
            f"{format_currency(s['SALE_AMT']):>12} |"
            f"{format_currency(s['SALE_DISC']):>12} |"
            f"{format_currency(s['SALE_NET']):>12} |"
            f"{s['RETCOUNT']:6d} |"
            f"{format_currency(s['RET_AMT']):>12} |"
            f"{format_currency(s['RET_DISC']):>12} |"
            f"{format_currency(s['RET_NET']):>12} |"
            f"{overall_count:6d} |"
            f"{format_currency(overall_amt):>12} |"
            f"{format_currency(overall_disc):>12} |"
            f"{format_currency(overall_net):>12} |"
        )  # Format the row data as a string
        lines.append(row_line)  # Add the row data to the report
    lines.append("-" * PAGE_WIDTH)  # Add a separator line

    totals_line = (
        f"{'TOTALAMOUNT   :':<17} |"
        f"{tot_sale_count:8d} |"
        f"{format_currency(tot_sale_amt):>12} |"
        f"{format_currency(tot_sale_disc):>12} |"
        f"{format_currency(tot_sale_net):>12} |"
        f"{int(tot_ret_count):6d} |"
        f"{format_currency(tot_ret_amt):>12} |"
        f"{format_currency(tot_ret_disc):>12} |"
        f"{format_currency(tot_ret_net):>12} |"
        f"{int(tot_overall_count):6d} |"
        f"{format_currency(tot_overall_amt):>12} |"
        f"{format_currency(tot_overall_disc):>12} |"
        f"{format_currency(tot_overall_net):>12} |"
    )  # Format the totals as a string
    lines.append(totals_line)  # Add the totals to the report
    lines.append("-" * PAGE_WIDTH)  # Add a separator line

    lines.extend([
        "\nSALES :-",
        f"\n       Net Cash Sales        : {format_currency(net_cash_sales)}",
        "       Total Paid In         :       0.00",
        "       Total Paid out        :       0.00"
    ])  # Add additional sales information to the report
    total_paid_in = Decimal('0.0')  # Initialize total paid in
    total_paid_out = Decimal('0.0')  # Initialize total paid out
    total_sales = Decimal(net_cash_sales) + total_paid_in + total_paid_out  # Calculate total sales
    lines.append(f"       Total Sales           : {format_currency(total_sales)}\n")  # Add total sales to the report

    lines.extend([
        "HealingCard Collections:",
        f"     Cash Collections        : {'0':>9}",
        f"     Credit Card Collections : {'0':>9}",
        f"     Total Collection        : {'0':>9}\n",
        f"Total Cash Amount            : {format_currency(total_sales)} ",
        "\n" + "-" * 180 + "\n"
    ])  # Add healing card collections to the report

    lines.append("\nPartner Program Summary  :\n")  # Add a header for the partner program summary
    partner_header = " slno| Name                                     |     NoInv        |    Amount    |"  # Define the partner header
    lines.append(partner_header)  # Add the partner header to the report
    lines.append("-" * PAGE_WIDTH)  # Add a separator line

    tot_partner_inv = tot_partner_amt = 0  # Initialize totals for partner invoices and amounts
    for idx, p in enumerate(partner_data, start=1):
        tot_partner_inv += p["BILLCNT"]  # Add the bill count to the total
        tot_partner_amt += float(p["AMOUNT"])  # Add the amount to the total
        part_line = (
            f"{idx:6d} | {p['NAME']:<38} |     {p['BILLCNT']:12d} | {format_currency(p['AMOUNT']):>12} |"
        )  # Format the partner data as a string
        lines.append(part_line)  # Add the partner data to the report
    lines.append("-" * (PAGE_WIDTH - 50))  # Add a separator line
    partner_totals_line = (
        f"      TOTAL AMOUNT:                    {tot_partner_inv:27d} | {format_currency(tot_partner_amt):>9} |"
    )  # Format the partner totals as a string
    lines.append(partner_totals_line)  # Add the partner totals to the report
    lines.append("-" * (PAGE_WIDTH - 50))  # Add a separator line

    fixed_lines = [fix_line(line) for line in lines]  # Fix the width of each line in the report
    return "\n".join(fixed_lines)  # Return the report as a string

# =============================================================================
# Modified Connection Functions to Allow Manual IP Override
# =============================================================================

def try_connection(series, formatted_site_id, username, password, database):
    """Attempt to connect using the given IP series."""
    host = f"{series}{formatted_site_id}"  # Construct the host address
    try:
        connection = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={host};"
            f"UID={username};"
            f"PWD={password};"
            f"DATABASE={database};"
        )  # Attempt to connect to the database
        return connection  # Return the connection object if successful
    except pyodbc.Error:
        return None  # Return None if the connection fails

def connect_to_database(site_id, username, password, database, ip_series_choice, custom_ip=None):
    """
    Attempt to connect to the server.

    If a custom_ip is provided (i.e. the user clicked EDIT/ALERT IP), then use it directly.
    Otherwise, use the IP series ('16' or '28') to build the host from the site_id.
    """
    if custom_ip:  # Use the manually entered IP address
        host = custom_ip  # Set the host to the custom IP
        try:
            connection = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={host};"
                f"UID={username};"
                f"PWD={password};"
                f"DATABASE={database};"
            )  # Attempt to connect to the database
            return connection  # Return the connection object if successful
        except pyodbc.Error:
            return None  # Return None if the connection fails
    else:
        # Format the site id as before (for example: '13100' becomes '131.00')
        try:
            formatted_site_id = f"{site_id[:3]}.{int(site_id[3:])}"  # Format the site ID
        except Exception as e:
            raise ValueError(f"Error formatting Site ID: {e}")  # Raise an error if formatting fails
        if ip_series_choice == "16":
            ip_series = ["10.16."]  # Set the IP series to '10.16.'
        elif ip_series_choice == "28":
            ip_series = ["10.28."]  # Set the IP series to '10.28.'
        else:
            ip_series = []  # Set the IP series to an empty list if the choice is invalid

        with concurrent.futures.ThreadPoolExecutor(max_workers=len(ip_series)) as executor:
            futures = {executor.submit(try_connection, series, formatted_site_id, username, password, database): series
                       for series in ip_series}  # Submit connection attempts for each IP series
            for future in concurrent.futures.as_completed(futures):
                connection = future.result()  # Get the result of each connection attempt
                if connection:
                    return connection  # Return the connection object if successful
        return None  # Return None if all connection attempts fail

def get_report_data(site_id, from_date, to_date, username, password, database, ip_series_choice, custom_ip=None):
    """
    Connect to the database, run the query, and return (result, site_name).
    The custom_ip (if provided) overrides the IP series.
    """
    connection = connect_to_database(site_id, username, password, database, ip_series_choice, custom_ip)  # Connect to the database
    if not connection:
        raise ConnectionError(f"Could not connect to the server for site {site_id}.")  # Raise an error if the connection fails

    try:
        cursor = connection.cursor()  # Create a cursor object
        cursor.execute("SELECT name FROM ax.inventsite WHERE siteid = ?", site_id)  # Execute a query to get the site name
        site_row = cursor.fetchone()  # Fetch the first row of the result
        site_name = site_row[0] if site_row else "Unknown Site"  # Get the site name from the row

        query = """
        select ISHEADER = CASE WHEN BILLTYPE ='GIFT' THEN 3 ELSE 1 END,
               ACXCORPCODE = -1,
               upper(BILLTYPE) BILLTYPE,
               sum(saleamt) NETSALEAMT,
               Cast(sum(discamt) as decimal(12,2)) DISCAMT,
               Cast(sum(RETAMT) as decimal(12,2)) NETRETAMT,
               Cast(sum(RETDISC) as decimal(12,2)) RETDISC,
               sum(isscnt) SALECOUNT,
               sum(retcnt) RETCNT
        from (
            select name billtype,
                   Cast(sum(AMOUNTTENDERED) as decimal(12,2)) saleamt,
                   sum(DISCAMOUNT) DISCAMT,
                   0 RETAMT,
                   0 RETDISC,
                   count(distinct rt.receiptid) isscnt,
                   0 retcnt
            from ax.retailtransactiontable rt
            join ax.RETAILTRANSACTIONPAYMENTTRANS rpt
              on rt.TRANSACTIONID = rpt.TRANSACTIONID and rt.RECEIPTID = rpt.RECEIPTID
            join RETAILTENDERTYPETABLE rtt
              on rpt.TENDERTYPE = rtt.TENDERTYPEID
            where ENTRYSTATUS = 0
              and acxtranstype = 0
              and rpt.TRANSACTIONSTATUS = 0
              and rpt.BUSINESSDATE between ? and ?
              and rpt.RECEIPTID not in (
                    Select IQ.receiptid
                    from ax.RETAILTRANSACTIONPAYMENTTRANS as IQ
                    where IQ.receiptid like 'IP%'
                      and IQ.tendertype in (1,2)
                      and IQ.BUSINESSDATE between ? and ?
              )
            group by name, DISCAMOUNT
            union
            select name,
                   0 saleamt,
                   0 DISCAMT,
                   Sum(AMOUNTTENDERED) AMOUNTTENDERED,
                   sum(-1*DISCAMOUNT) DISCAMT,
                   0 isscnt,
                   count(distinct rt.receiptid) retcnt
            from ax.retailtransactiontable rt
            join ax.RETAILTRANSACTIONPAYMENTTRANS rpt
              on rt.TRANSACTIONID = rpt.TRANSACTIONID and rt.RECEIPTID = rpt.RECEIPTID
            join RETAILTENDERTYPETABLE rtt
              on rpt.TENDERTYPE = rtt.TENDERTYPEID
            where ENTRYSTATUS = 0
              and acxtranstype <> 0
              and rpt.TRANSACTIONSTATUS = 0
              and rpt.BUSINESSDATE between ? and ?
            group by name
        ) a
        group by billtype
        union all
        select ISHEADER = 0,
               ACXCORPCODE,
               ax.getcorporatename(acxcorpcode) CORPORATE,
               (cast(sum(CASE WHEN ACXCORPCODE ='172' AND ACXCREDIT = 0 THEN 0 ELSE -1*GROSSAMOUNT END)
                - sum(case when ACXTRANSTYPE = 0 then discamount
                           when ACXTRANSTYPE <> 0 then -1*discamount end) as decimal(18,2)) - sum(ACXLOYALTY)) NETAMT,
               0, 0, 0,
               count(distinct CASE WHEN ACXCORPCODE='172' AND ACXCREDIT = 0 THEN NULL ELSE receiptid END) BILLCNT,
               0
        from ax.retailtransactiontable
        where ENTRYSTATUS = 0
          and BUSINESSDATE between ? and ?
        group by acxcorpcode
        union all
        Select ISHEADER = 2,
               PAYMENTCODE,
               'HEALINGCARD-' + PAYMENTTYPE,
               sum(TRANSAMT) Amount,
               0, 0, 0, 0, 0
        from HEALING_CARD_TRANSACTION
        where ACTIONID in (0,1)
          and cast(TRANSACTIONDATE as date) between ? and ?
        group by PAYMENTCODE, PAYMENTTYPE
        union all
        Select ISHEADER = 4,
               0,
               'OMS CASH COLLECTION',
               isnull(SUM(COLLECTEDAMT),0) as COLLECTEDAMT,
               0, 0, 0, 0, 0
        from ax.ACXSETTLEMENTDETAILS
        where cast(SETTLEMENTDATE as date) between ? and ?
        union all
        select ISHEADER = 5,
               tendertype,
               'IP COLLECTION',
               isnull(SUM(AMOUNTTENDERED), 0) as COLLECTEDAMT,
               0, 0, 0, 0, 0
        from ax.retailtransactionpaymenttrans
        where tendertype in (1,2)
          and receiptid like 'IP%'
          and BUSINESSDATE between ? and ?
        group by tendertype
        """  # Define the query to retrieve report data
        params = (
            from_date, to_date,
            from_date, to_date,
            from_date, to_date,
            from_date, to_date,
            from_date, to_date,
            from_date, to_date,
            from_date, to_date
        )  # Define the parameters for the query
        cursor.execute(query, params)  # Execute the query with the parameters
        result = cursor.fetchall()  # Fetch all rows of the result
        return result, site_name  # Return the result and site name
    finally:
        connection.close()  # Close the connection

# =============================================================================
# Tkinter Application with Continuous Log Output
# =============================================================================

class SalesSummaryReportApp(tk.Tk):
    def __init__(self):
        super().__init__()  # Initialize the parent class
        self.title("Sales Summary Report")  # Set the window title
        self.geometry("1200x800")  # Set the window size
        self.configure(bg="#FFFFFF")  # Set the background color to white

        # Variables to store ZIP data and file path.
        self.zip_buffer = None  # Initialize the ZIP buffer
        self.file_path = None  # Initialize the file path

        # Set up ttk style with Times New Roman fonts.
        self.style = ttk.Style(self)  # Create a style object
        self.style.theme_use("clam")  # Set the theme to 'clam'
        self.style.configure("TFrame", background="#FFFFFF")  # Configure the frame style
        self.style.configure("TLabel", background="#FFFFFF", foreground="#333333", font=("Times New Roman", 12, "bold"))  # Configure the label style
        self.style.configure("TButton", background="#CCCCCC", foreground="#333333", font=("Times New Roman", 12, "bold"))  # Configure the button style
        self.style.map("TButton", background=[('active', '#AAAAAA')], foreground=[('active', '#000000')])  # Configure the button hover style
        self.style.configure("Input.TLabelframe", background="#F0F0F0", foreground="#333333",
                             font=("Times New Roman", 12, "bold"), borderwidth=2)  # Configure the labelframe style
        self.style.configure("Input.TLabelframe.Label", background="#F0F0F0", foreground="#333333")  # Configure the labelframe label style
        self.style.configure("TRadiobutton", font=("Times New Roman", 12, "bold"), background="#F0F0F0", foreground="#333333")  # Configure the radiobutton style

        self.create_widgets()  # Create the widgets

    def create_widgets(self):
        # Top frame for input controls.
        controls_frame = ttk.Frame(self, style="TFrame")  # Create a frame for input controls
        controls_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)  # Pack the frame with padding

        # ----------------------------
        # Credentials & Site ID Frame
        # ----------------------------
        creds_frame = ttk.Labelframe(controls_frame, text="DB Credentials", style="Input.TLabelframe")  # Create a labelframe for credentials
        creds_frame.pack(fill=tk.X, padx=5, pady=5)  # Pack the labelframe with padding
        ttk.Label(creds_frame, text="Username:").grid(row=0, column=0, padx=5, pady=4, sticky="w")  # Create a label for the username
        self.username_entry = ttk.Entry(creds_frame, width=20, font=("Times New Roman", 12, "bold"))  # Create an entry for the username
        self.username_entry.grid(row=0, column=1, padx=5, pady=4)  # Pack the entry with padding
        ttk.Label(creds_frame, text="Password:").grid(row=0, column=2, padx=5, pady=4, sticky="w")  # Create a label for the password
        self.password_entry = ttk.Entry(creds_frame, width=20, show="*", font=("Times New Roman", 12, "bold"))  # Create an entry for the password
        self.password_entry.grid(row=0, column=3, padx=5, pady=4)  # Pack the entry with padding
        ttk.Label(creds_frame, text="Database:").grid(row=0, column=4, padx=5, pady=4, sticky="w")  # Create a label for the database
        self.database_entry = ttk.Entry(creds_frame, width=20, font=("Times New Roman", 12, "bold"))  # Create an entry for the database
        self.database_entry.grid(row=0, column=5, padx=5, pady=4)  # Pack the entry with padding
        ttk.Label(creds_frame, text="Site ID:").grid(row=0, column=6, padx=5, pady=4, sticky="w")  # Create a label for the site ID
        self.siteid_entry = ttk.Entry(creds_frame, width=15, font=("Times New Roman", 12, "bold"))  # Create an entry for the site ID
        self.siteid_entry.grid(row=0, column=7, padx=5, pady=4)  # Pack the entry with padding

        # ----------------------------
        # Manual IP Override (Edit IP) Frame
        # ----------------------------
        manual_ip_frame = ttk.Labelframe(controls_frame, text="Manual IP Override (Optional)", style="Input.TLabelframe")  # Create a labelframe for manual IP override
        manual_ip_frame.pack(fill=tk.X, padx=5, pady=5)  # Pack the labelframe with padding
        ttk.Label(manual_ip_frame, text="Custom IP:").pack(side=tk.LEFT, padx=5, pady=4)  # Create a label for the custom IP
        self.custom_ip_entry = ttk.Entry(manual_ip_frame, width=20, font=("Times New Roman", 12, "bold"))  # Create an entry for the custom IP
        self.custom_ip_entry.pack(side=tk.LEFT, padx=5, pady=4)  # Pack the entry with padding
        # (If left blank, the connection will use the IP Series options below.)

        # ----------------------------
        # File Upload Frame
        # ----------------------------
        file_frame = ttk.Labelframe(controls_frame, text="Upload File for Multiple Site IDs (Optional)", style="Input.TLabelframe")  # Create a labelframe for file upload
        file_frame.pack(fill=tk.X, padx=5, pady=5)  # Pack the labelframe with padding
        self.file_label = ttk.Label(file_frame, text="No file selected", style="TLabel")  # Create a label for the file path
        self.file_label.pack(side=tk.LEFT, padx=5, pady=4)  # Pack the label with padding
        ttk.Button(file_frame, text="Select File", command=self.select_file).pack(side=tk.LEFT, padx=5, pady=4)  # Create a button to select a file
        ttk.Button(file_frame, text="Remove File", command=self.remove_file).pack(side=tk.LEFT, padx=5, pady=4)  # Create a button to remove the selected file

        # ----------------------------
        # Server Series Frame
        # ----------------------------
        series_frame = ttk.Labelframe(controls_frame, text="Server Series", style="Input.TLabelframe")  # Create a labelframe for server series
        series_frame.pack(fill=tk.X, padx=5, pady=5)  # Pack the labelframe with padding
        self.ip_series = tk.StringVar(value="16")  # Create a string variable for the IP series
        ttk.Radiobutton(series_frame, text="10.16.x.x", variable=self.ip_series, value="16", style="TRadiobutton").pack(side=tk.LEFT, padx=10, pady=4)  # Create a radiobutton for the '10.16.x.x' series
        ttk.Radiobutton(series_frame, text="10.28.x.x", variable=self.ip_series, value="28", style="TRadiobutton").pack(side=tk.LEFT, padx=10, pady=4)  # Create a radiobutton for the '10.28.x.x' series

        # ----------------------------
        # Date Range Frame
        # ----------------------------
        date_frame = ttk.Labelframe(controls_frame, text="Date Range", style="Input.TLabelframe")  # Create a labelframe for the date range
        date_frame.pack(fill=tk.X, padx=5, pady=5)  # Pack the labelframe with padding
        ttk.Label(date_frame, text="From Date:").pack(side=tk.LEFT, padx=5, pady=4)  # Create a label for the from date
        self.from_date = DateEntry(date_frame, width=20, date_pattern='yyyy-mm-dd',
                                   background="white", foreground="black", font=("Times New Roman", 12, "bold"))  # Create a date entry for the from date
        self.from_date.set_date(datetime.today())  # Set the default value to today's date
        self.from_date.pack(side=tk.LEFT, padx=5, pady=4)  # Pack the date entry with padding
        ttk.Label(date_frame, text="To Date:").pack(side=tk.LEFT, padx=5, pady=4)  # Create a label for the to date
        self.to_date = DateEntry(date_frame, width=20, date_pattern='yyyy-mm-dd',
                                 background="white", foreground="black", font=("Times New Roman", 12, "bold"))  # Create a date entry for the to date
        self.to_date.set_date(datetime.today())  # Set the default value to today's date
        self.to_date.pack(side=tk.LEFT, padx=5, pady=4)  # Pack the date entry with padding

        # ----------------------------
        # Action Buttons Frame
        # ----------------------------
        action_frame = ttk.Frame(controls_frame, style="TFrame")  # Create a frame for action buttons
        action_frame.pack(fill=tk.X, padx=5, pady=5)  # Pack the frame with padding
        ttk.Button(action_frame, text="Generate Report", command=self.generate_reports).pack(side=tk.LEFT, padx=10, pady=4, expand=True, fill=tk.X)  # Create a button to generate reports
        self.download_button = ttk.Button(action_frame, text="Download Report", command=self.download_reports, state="disabled")  # Create a button to download reports
        self.download_button.pack(side=tk.LEFT, padx=10, pady=4, expand=True, fill=tk.X)  # Pack the button with padding

        # ----------------------------
        # Log Output Area
        # ----------------------------
        log_frame = ttk.Frame(self, style="TFrame")  # Create a frame for the log output area
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))  # Pack the frame with padding
        log_top_frame = ttk.Frame(log_frame, style="TFrame")  # Create a top frame for the log output area
        log_top_frame.pack(side=tk.TOP, fill=tk.X)  # Pack the top frame with padding
        log_label = ttk.Label(log_top_frame, text="Log Output:", style="TLabel")  # Create a label for the log output area
        log_label.pack(side=tk.LEFT, padx=5, pady=5)  # Pack the label with padding
        ttk.Button(log_top_frame, text="Clear Log", command=self.clear_log).pack(side=tk.RIGHT, padx=5, pady=5)  # Create a button to clear the log
        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, background="#FFFFFF", foreground="#333333", font=("Times New Roman", 12))  # Create a scrollable text area for the log output
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)  # Pack the text area with padding

    def select_file(self):
        filename = filedialog.askopenfilename(title="Select File", filetypes=[("Excel files", "*.xlsx"), ("Text files", "*.txt"), ("CSV files", "*.csv")])  # Open a file dialog to select a file
        if filename:
            self.file_label.config(text=os.path.basename(filename))  # Set the label text to the selected file name
            self.file_path = filename  # Set the file path to the selected file
            self.safe_log(f"Selected file: {filename}")  # Log the selected file

    def remove_file(self):
        """Clears the file selection."""
        self.file_path = None  # Clear the file path
        self.file_label.config(text="No file selected")  # Set the label text to 'No file selected'
        self.safe_log("File selection cleared.")  # Log the file selection cleared

    def clear_log(self):
        """Clears the log output area."""
        self.log_text.delete("1.0", tk.END)  # Delete all text in the log output area

    def log(self, message):
        """Append a message to the log area."""
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")  # Insert the message with a timestamp to the log output area
        self.log_text.see(tk.END)  # Scroll to the end of the log output area

    def safe_log(self, message):
        """
        Thread-safe logging: schedule log updates on the main thread.
        """
        self.after(0, self.log, message)  # Schedule the log update on the main thread

    # -------------------------------------------------------------------------
    # Report Generation: Run in a worker thread so that logs are updated live.
    # -------------------------------------------------------------------------
    def generate_reports(self):
        # Clear previous ZIP buffer and disable download button.
        self.zip_buffer = None  # Clear the ZIP buffer
        self.download_button.config(state="disabled")  # Disable the download button
        self.safe_log("Started generating report...")  # Log the start of report generation
        # Start the report generation in a separate thread.
        threading.Thread(target=self.run_reports, daemon=True).start()  # Start a new thread to run the reports

    def run_reports(self):
        """
        This method (running in a worker thread) collects input values,
        validates credentials, and processes the sites in parallel.
        All key steps are logged immediately.
        """
        try:
            username = self.username_entry.get().strip()  # Get the username from the entry
            password = self.password_entry.get().strip()  # Get the password from the entry
            database = self.database_entry.get().strip()  # Get the database from the entry
            site_id_manual = self.siteid_entry.get().strip()  # Get the site ID from the entry
            ip_series_choice = self.ip_series.get()  # Get the selected IP series
            from_date_str = self.from_date.get_date().strftime("%Y-%m-%d")  # Get the from date as a string
            to_date_str = self.to_date.get_date().strftime("%Y-%m-%d")  # Get the to date as a string
            custom_ip = self.custom_ip_entry.get().strip()  # Get the manually entered IP (if any)

            # Determine whether to use file input or manual site ID.
            if self.file_label.cget("text") != "No file selected":
                self.safe_log("File upload mode detected. Reading Site IDs from file...")  # Log the file upload mode
                try:
                    ext = os.path.splitext(self.file_path)[1].lower()  # Get the file extension
                    if ext == '.xlsx':
                        df = pd.read_excel(self.file_path)  # Read the Excel file
                        cols = [col.lower() for col in df.columns]  # Get the column names
                        if "siteid" in cols:
                            site_ids = df["siteid"].astype(str).tolist()  # Get the site IDs from the 'siteid' column
                        else:
                            raise Exception("Excel file must contain a column named 'siteid'.")  # Raise an error if the 'siteid' column is not found
                    else:
                        with open(self.file_path, "r", encoding="utf-8") as f:
                            file_text = f.read()  # Read the file content
                        if "," in file_text:
                            site_ids = [s.strip() for s in file_text.split(",") if s.strip()]  # Split the content by commas
                        else:
                            site_ids = [s.strip() for s in file_text.splitlines() if s.strip()]  # Split the content by lines
                    self.safe_log(f"Found {len(site_ids)} site IDs in file.")  # Log the number of site IDs found
                except Exception as e:
                    self.after(0, messagebox.showerror, "Error", f"Error reading file: {e}")  # Show an error message if reading the file fails
                    return
            else:
                if not site_id_manual:
                    self.after(0, messagebox.showerror, "Error", "Please enter a Site ID or upload a file.")  # Show an error message if no site ID is entered
                    return
                site_ids = [site_id_manual]  # Use the manually entered site ID
                self.safe_log("Manual input mode selected.")  # Log the manual input mode

            # Validate credentials using the first site ID.
            self.safe_log("Validating connection with test connection...")  # Log the start of connection validation
            test_conn = connect_to_database(site_ids[0], username, password, database, ip_series_choice, custom_ip)  # Connect to the database using the first site ID
            if not test_conn:
                raise Exception("Test connection failed.")  # Raise an error if the test connection fails
            test_conn.close()  # Close the test connection
            self.safe_log("Test connection successful.")  # Log the successful test connection

            successful_reports = {}  # Initialize a dictionary to store successful reports
            failed_sites = {}  # Initialize a dictionary to store failed sites
            total_sites = len(site_ids)  # Get the total number of sites

            # Define helper function to process each site.
            def process_site(sid, index):
                try:
                    self.safe_log(f"Processing site {sid} ({index}/{total_sites})...")  # Log the start of processing for the site
                    result, site_name = get_report_data(sid, from_date_str, to_date_str, username, password, database, ip_series_choice, custom_ip)  # Get the report data for the site
                    report_text = format_report(result, sid, site_name, from_date_str, to_date_str)  # Format the report data as a string
                    self.safe_log(f"Completed site {sid}.")  # Log the completion of processing for the site
                    return (sid, report_text, None)  # Return the site ID, report text, and no error
                except Exception as e:
                    self.safe_log(f"Error processing site {sid}: {e}")  # Log the error processing the site
                    return (sid, None, str(e))  # Return the site ID, no report text, and the error message

            # Process sites in parallel.
            with concurrent.futures.ThreadPoolExecutor(max_workers=min(total_sites, 10)) as executor:
                futures = {executor.submit(process_site, sid, i): sid for i, sid in enumerate(site_ids, start=1)}  # Submit processing for each site
                for future in concurrent.futures.as_completed(futures):
                    sid, report_text, error = future.result()  # Get the result of processing for each site
                    if report_text is not None:
                        successful_reports[sid] = report_text  # Add the successful report to the dictionary
                    else:
                        failed_sites[sid] = error  # Add the failed site to the dictionary

            self.safe_log("Report generation completed.")  # Log the completion of report generation
            if successful_reports:
                self.zip_buffer = io.BytesIO()  # Create a buffer for the ZIP file
                with zipfile.ZipFile(self.zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for sid, report in successful_reports.items():
                        zip_file.writestr(f"{sid}.txt", report)  # Write each report to the ZIP file
                self.zip_buffer.seek(0)  # Seek to the beginning of the buffer
                self.safe_log("Reports generated successfully. Click on Download Report.")  # Log the successful generation of reports
                self.after(0, lambda: self.download_button.config(state="normal"))  # Enable the download button
            else:
                self.after(0, messagebox.showerror, "Error", "No successful reports to save.")  # Show an error message if no reports were generated
                return

            if failed_sites:
                errors = "\n".join([f"{sid}: {err}" for sid, err in failed_sites.items()])  # Get the errors for the failed sites
                self.safe_log("Failed Sites:\n" + errors)  # Log the failed sites
                self.after(0, messagebox.showwarning, "Warning", f"Some sites could not be processed:\n{errors}")  # Show a warning message for the failed sites

        except Exception as e:
            self.safe_log(f"Error: {e}")  # Log the error
            self.after(0, messagebox.showerror, "Error", str(e))  # Show an error message

    def download_reports(self):
        """
        Automatically save the ZIP file to the user's Downloads folder as 'SiteReports.zip'.
        If the file already exists, append a counter to the file name.
        """
        try:
            downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")  # Get the Downloads folder path
            base_filename = "SiteReports"  # Set the base file name
            extension = ".zip"  # Set the file extension
            file_path = os.path.join(downloads_folder, base_filename + extension)  # Get the file path
            counter = 1  # Initialize the counter
            while os.path.exists(file_path):
                file_path = os.path.join(downloads_folder, f"{base_filename}{counter}{extension}")  # Append the counter to the file name
                counter += 1  # Increment the counter
            with open(file_path, "wb") as f:
                f.write(self.zip_buffer.getvalue())  # Write the ZIP file to the file path
            self.safe_log(f"ZIP file auto-saved to {file_path}")  # Log the file path
            messagebox.showinfo("Success", f"Report auto-saved to:\n{file_path}")  # Show a success message with the file path
        except Exception as e:
            messagebox.showerror("Error", f"Error saving ZIP file: {e}")  # Show an error message if saving the ZIP file fails

if __name__ == "__main__":
    app = SalesSummaryReportApp()  # Create an instance of the application
    app.mainloop()  # Start the main event loop
