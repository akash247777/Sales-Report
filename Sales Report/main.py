import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
from tkcalendar import DateEntry  # pip install tkcalendar
import pyodbc
from datetime import datetime
from decimal import Decimal
import io
import zipfile
import pandas as pd
import concurrent.futures
import os
import threading


def format_currency(value):
    """Format a numeric value as a currency string with 2 decimals."""
    return f"{value:,.2f}"

def format_report(result, site_id, site_name, from_date, to_date):
    """
    Process the query result (a list of tuples) and produce a text report
    in which every line is fixed to 180 characters.
    """
    PAGE_WIDTH = 180

    def fix_line(line, width=PAGE_WIDTH):
        clean = line.rstrip("\n")
        if len(clean) < width:
            return clean + " " * (width - len(clean))
        else:
            return clean[:width]

    now = datetime.now()
    header_date = now.strftime("%d/%m/%Y")
    header_time = now.strftime("%I:%M %p")

    lines = []
    # Header Section.
    lines.append(f"DATE: {header_date}".rjust(PAGE_WIDTH))
    lines.append(f"TIME: {header_time}".rjust(PAGE_WIDTH))
    lines.append("")
    lines.append("APOLLO PHARMACIES LIMITED".center(PAGE_WIDTH))
    lines.append(f"{site_id} - {site_name}".center(PAGE_WIDTH))
    lines.append("")
    lines.append("Sales Transaction Summary Report".center(PAGE_WIDTH))
    lines.append(f"From Date : {from_date}    To Date : {to_date}".center(PAGE_WIDTH))
    lines.append("-" * PAGE_WIDTH)

    header_groups = (
        "|" +
        " SALES ".center(55) +
        "|" +
        " RETURNS ".center(55) +
        "|" +
        " NET ".center(55) +
        "|"
    )
    lines.append(header_groups)
    lines.append("-" * PAGE_WIDTH)

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
    )
    lines.append(header_cols)
    lines.append("-" * PAGE_WIDTH)

    # Process data rows.
    sales_data = []
    partner_data = []

    for row in result:
        isheader = row[0]
        if isheader in (1, 3):
            sale_net = row[3]
            sale_disc = row[4]
            ret_net = row[5]
            ret_disc = row[6]
            sales_data.append({
                "BILLTYPE": row[2],
                "SALECOUNT": row[7],
                "SALE_NET": row[3],
                "SALE_DISC": row[4],
                "SALE_AMT": sale_net + sale_disc,
                "RETCOUNT": row[8],
                "RET_NET": ret_net,
                "RET_DISC": ret_disc,
                "RET_AMT": ret_net + ret_disc
            })
        elif isheader == 0:
            partner_data.append({
                "NAME": row[2],
                "BILLCNT": row[7],
                "AMOUNT": row[3]
            })

    tot_sale_count = tot_sale_amt = tot_sale_disc = tot_sale_net = 0
    tot_ret_count = tot_ret_amt = tot_ret_disc = tot_ret_net = 0
    net_cash_sales = 0

    for s in sales_data:
        if s["BILLTYPE"].upper() != "GIFT":
            tot_sale_count   += s["SALECOUNT"]
            tot_sale_amt     += float(s["SALE_AMT"])
            tot_sale_disc    += float(s["SALE_DISC"])
            tot_sale_net     += float(s["SALE_NET"])
            tot_ret_count    += s["RETCOUNT"]
            tot_ret_amt      += float(s["RET_AMT"])
            tot_ret_disc     += float(s["RET_DISC"])
            tot_ret_net      += float(s["RET_NET"])
        if s["BILLTYPE"].upper() == "CASH":
            net_cash_sales = s["SALE_NET"] + s["RET_NET"]

    tot_overall_count = tot_sale_count + tot_ret_count
    tot_overall_amt   = tot_sale_amt + tot_ret_amt
    tot_overall_disc  = tot_sale_disc + tot_ret_disc
    tot_overall_net   = tot_sale_net + tot_ret_net

    for s in sales_data:
        overall_count = s["SALECOUNT"] + s["RETCOUNT"]
        overall_amt   = s["SALE_AMT"] + s["RET_AMT"]
        overall_disc  = s["SALE_DISC"] + s["RET_DISC"]
        overall_net   = s["SALE_NET"] + s["RET_NET"]
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
        )
        lines.append(row_line)
    lines.append("-" * PAGE_WIDTH)

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
    )
    lines.append(totals_line)
    lines.append("-" * PAGE_WIDTH)

    lines.extend([
        "\nSALES :-",
        f"\n       Net Cash Sales        : {format_currency(net_cash_sales)}",
        "       Total Paid In         :       0.00",
        "       Total Paid out        :       0.00"
    ])
    total_paid_in = Decimal('0.0')
    total_paid_out = Decimal('0.0')
    total_sales = Decimal(net_cash_sales) + total_paid_in + total_paid_out
    lines.append(f"       Total Sales           : {format_currency(total_sales)}\n")
    
    lines.extend([
        "HealingCard Collections:",
        f"     Cash Collections        : {'0':>9}",
        f"     Credit Card Collections : {'0':>9}",
        f"     Total Collection        : {'0':>9}\n",
        f"Total Cash Amount            : {format_currency(total_sales)} ",
        "\n" + "-" * 180 + "\n"
    ])

    lines.append("\nPartner Program Summary  :\n")
    partner_header = " slno| Name                                     |     NoInv        |    Amount    |"
    lines.append(partner_header)
    lines.append("-" * PAGE_WIDTH)
    
    tot_partner_inv = tot_partner_amt = 0
    for idx, p in enumerate(partner_data, start=1):
        tot_partner_inv += p["BILLCNT"]
        tot_partner_amt += float(p["AMOUNT"])
        part_line = (
            f"{idx:6d} | {p['NAME']:<38} |     {p['BILLCNT']:12d} | {format_currency(p['AMOUNT']):>12} |"
        )
        lines.append(part_line)
    lines.append("-" * (PAGE_WIDTH - 50))
    partner_totals_line = (
        f"      TOTAL AMOUNT:                    {tot_partner_inv:27d} | {format_currency(tot_partner_amt):>9} |"
    )
    lines.append(partner_totals_line)
    lines.append("-" * (PAGE_WIDTH - 50))
    
    fixed_lines = [fix_line(line) for line in lines]
    return "\n".join(fixed_lines)

# =============================================================================
# Modified Connection Functions to Allow Manual IP Override
# =============================================================================

def try_connection(series, formatted_site_id, username, password, database):
    """Attempt to connect using the given IP series."""
    host = f"{series}{formatted_site_id}"
    try:
        connection = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={host};"
            f"UID={username};"
            f"PWD={password};"
            f"DATABASE={database};"
        )
        return connection
    except pyodbc.Error:
        return None

def connect_to_database(site_id, username, password, database, ip_series_choice, custom_ip=None):
    """
    Attempt to connect to the server.
    
    If a custom_ip is provided (i.e. the user clicked EDIT/ALERT IP), then use it directly.
    Otherwise, use the IP series ('16' or '28') to build the host from the site_id.
    """
    if custom_ip:  # Use the manually entered IP address
        host = custom_ip
        try:
            connection = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={host};"
                f"UID={username};"
                f"PWD={password};"
                f"DATABASE={database};"
            )
            return connection
        except pyodbc.Error:
            return None
    else:
        # Format the site id as before (for example: '13100' becomes '131.00')
        try:
            formatted_site_id = f"{site_id[:3]}.{int(site_id[3:])}"
        except Exception as e:
            raise ValueError(f"Error formatting Site ID: {e}")
        if ip_series_choice == "16":
            ip_series = ["10.16."]
        elif ip_series_choice == "28":
            ip_series = ["10.28."]
        else:
            ip_series = []
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=len(ip_series)) as executor:
            futures = {executor.submit(try_connection, series, formatted_site_id, username, password, database): series 
                       for series in ip_series}
            for future in concurrent.futures.as_completed(futures):
                connection = future.result()
                if connection:
                    return connection
        return None

def get_report_data(site_id, from_date, to_date, username, password, database, ip_series_choice, custom_ip=None):
    """
    Connect to the database, run the query, and return (result, site_name).
    The custom_ip (if provided) overrides the IP series.
    """
    connection = connect_to_database(site_id, username, password, database, ip_series_choice, custom_ip)
    if not connection:
        raise ConnectionError(f"Could not connect to the server for site {site_id}.")
    
    try:
        cursor = connection.cursor()
        cursor.execute("SELECT name FROM ax.inventsite WHERE siteid = ?", site_id)
        site_row = cursor.fetchone()
        site_name = site_row[0] if site_row else "Unknown Site"
        
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
        """
        params = (
            from_date, to_date,
            from_date, to_date,
            from_date, to_date,
            from_date, to_date,
            from_date, to_date,
            from_date, to_date,
            from_date, to_date
        )
        cursor.execute(query, params)
        result = cursor.fetchall()
        return result, site_name
    finally:
        connection.close()

# =============================================================================
# Tkinter Application with Continuous Log Output
# =============================================================================

class SalesSummaryReportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sales Summary Report")
        self.geometry("1200x800")
        self.configure(bg="#FFFFFF")  # White background

        # Variables to store ZIP data and file path.
        self.zip_buffer = None
        self.file_path = None  # For file upload mode

        # Set up ttk style with Times New Roman fonts.
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.style.configure("TFrame", background="#FFFFFF")
        self.style.configure("TLabel", background="#FFFFFF", foreground="#333333", font=("Times New Roman", 12, "bold"))
        self.style.configure("TButton", background="#CCCCCC", foreground="#333333", font=("Times New Roman", 12, "bold"))
        self.style.map("TButton", background=[('active', '#AAAAAA')], foreground=[('active', '#000000')])
        self.style.configure("Input.TLabelframe", background="#F0F0F0", foreground="#333333",
                             font=("Times New Roman", 12, "bold"), borderwidth=2)
        self.style.configure("Input.TLabelframe.Label", background="#F0F0F0", foreground="#333333")
        self.style.configure("TRadiobutton", font=("Times New Roman", 12, "bold"), background="#F0F0F0", foreground="#333333")

        self.create_widgets()

    def create_widgets(self):
        # Top frame for input controls.
        controls_frame = ttk.Frame(self, style="TFrame")
        controls_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        # ----------------------------
        # Credentials & Site ID Frame
        # ----------------------------
        creds_frame = ttk.Labelframe(controls_frame, text="DB Credentials", style="Input.TLabelframe")
        creds_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(creds_frame, text="Username:").grid(row=0, column=0, padx=5, pady=4, sticky="w")
        self.username_entry = ttk.Entry(creds_frame, width=20, font=("Times New Roman", 12, "bold"))
        self.username_entry.grid(row=0, column=1, padx=5, pady=4)
        ttk.Label(creds_frame, text="Password:").grid(row=0, column=2, padx=5, pady=4, sticky="w")
        self.password_entry = ttk.Entry(creds_frame, width=20, show="*", font=("Times New Roman", 12, "bold"))
        self.password_entry.grid(row=0, column=3, padx=5, pady=4)
        ttk.Label(creds_frame, text="Database:").grid(row=0, column=4, padx=5, pady=4, sticky="w")
        self.database_entry = ttk.Entry(creds_frame, width=20, font=("Times New Roman", 12, "bold"))
        self.database_entry.grid(row=0, column=5, padx=5, pady=4)
        ttk.Label(creds_frame, text="Site ID:").grid(row=0, column=6, padx=5, pady=4, sticky="w")
        self.siteid_entry = ttk.Entry(creds_frame, width=15, font=("Times New Roman", 12, "bold"))
        self.siteid_entry.grid(row=0, column=7, padx=5, pady=4)

        # ----------------------------
        # Manual IP Override (Edit IP) Frame
        # ----------------------------
        manual_ip_frame = ttk.Labelframe(controls_frame, text="Manual IP Override (Optional)", style="Input.TLabelframe")
        manual_ip_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(manual_ip_frame, text="Custom IP:").pack(side=tk.LEFT, padx=5, pady=4)
        self.custom_ip_entry = ttk.Entry(manual_ip_frame, width=20, font=("Times New Roman", 12, "bold"))
        self.custom_ip_entry.pack(side=tk.LEFT, padx=5, pady=4)
        # (If left blank, the connection will use the IP Series options below.)

        # ----------------------------
        # File Upload Frame
        # ----------------------------
        file_frame = ttk.Labelframe(controls_frame, text="Upload File for Multiple Site IDs (Optional)", style="Input.TLabelframe")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        self.file_label = ttk.Label(file_frame, text="No file selected", style="TLabel")
        self.file_label.pack(side=tk.LEFT, padx=5, pady=4)
        ttk.Button(file_frame, text="Select File", command=self.select_file).pack(side=tk.LEFT, padx=5, pady=4)
        ttk.Button(file_frame, text="Remove File", command=self.remove_file).pack(side=tk.LEFT, padx=5, pady=4)

        # ----------------------------
        # Server Series Frame
        # ----------------------------
        series_frame = ttk.Labelframe(controls_frame, text="Server Series", style="Input.TLabelframe")
        series_frame.pack(fill=tk.X, padx=5, pady=5)
        self.ip_series = tk.StringVar(value="16")
        ttk.Radiobutton(series_frame, text="10.16.x.x", variable=self.ip_series, value="16", style="TRadiobutton").pack(side=tk.LEFT, padx=10, pady=4)
        ttk.Radiobutton(series_frame, text="10.28.x.x", variable=self.ip_series, value="28", style="TRadiobutton").pack(side=tk.LEFT, padx=10, pady=4)

        # ----------------------------
        # Date Range Frame
        # ----------------------------
        date_frame = ttk.Labelframe(controls_frame, text="Date Range", style="Input.TLabelframe")
        date_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(date_frame, text="From Date:").pack(side=tk.LEFT, padx=5, pady=4)
        self.from_date = DateEntry(date_frame, width=20, date_pattern='yyyy-mm-dd',
                                   background="white", foreground="black", font=("Times New Roman", 12, "bold"))
        self.from_date.set_date(datetime.today())
        self.from_date.pack(side=tk.LEFT, padx=5, pady=4)
        ttk.Label(date_frame, text="To Date:").pack(side=tk.LEFT, padx=5, pady=4)
        self.to_date = DateEntry(date_frame, width=20, date_pattern='yyyy-mm-dd',
                                 background="white", foreground="black", font=("Times New Roman", 12, "bold"))
        self.to_date.set_date(datetime.today())
        self.to_date.pack(side=tk.LEFT, padx=5, pady=4)

        # ----------------------------
        # Action Buttons Frame
        # ----------------------------
        action_frame = ttk.Frame(controls_frame, style="TFrame")
        action_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(action_frame, text="Generate Report", command=self.generate_reports).pack(side=tk.LEFT, padx=10, pady=4, expand=True, fill=tk.X)
        self.download_button = ttk.Button(action_frame, text="Download Report", command=self.download_reports, state="disabled")
        self.download_button.pack(side=tk.LEFT, padx=10, pady=4, expand=True, fill=tk.X)

        # ----------------------------
        # Log Output Area
        # ----------------------------
        log_frame = ttk.Frame(self, style="TFrame")
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        log_top_frame = ttk.Frame(log_frame, style="TFrame")
        log_top_frame.pack(side=tk.TOP, fill=tk.X)
        log_label = ttk.Label(log_top_frame, text="Log Output:", style="TLabel")
        log_label.pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(log_top_frame, text="Clear Log", command=self.clear_log).pack(side=tk.RIGHT, padx=5, pady=5)
        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, background="#FFFFFF", foreground="#333333", font=("Times New Roman", 12))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def select_file(self):
        filename = filedialog.askopenfilename(title="Select File", filetypes=[("Excel files", "*.xlsx"), ("Text files", "*.txt"), ("CSV files", "*.csv")])
        if filename:
            self.file_label.config(text=os.path.basename(filename))
            self.file_path = filename
            self.safe_log(f"Selected file: {filename}")

    def remove_file(self):
        """Clears the file selection."""
        self.file_path = None
        self.file_label.config(text="No file selected")
        self.safe_log("File selection cleared.")

    def clear_log(self):
        """Clears the log output area."""
        self.log_text.delete("1.0", tk.END)

    def log(self, message):
        """Append a message to the log area."""
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)

    def safe_log(self, message):
        """
        Thread-safe logging: schedule log updates on the main thread.
        """
        self.after(0, self.log, message)

    # -------------------------------------------------------------------------
    # Report Generation: Run in a worker thread so that logs are updated live.
    # -------------------------------------------------------------------------
    def generate_reports(self):
        # Clear previous ZIP buffer and disable download button.
        self.zip_buffer = None
        self.download_button.config(state="disabled")
        self.safe_log("Started generating report...")
        # Start the report generation in a separate thread.
        threading.Thread(target=self.run_reports, daemon=True).start()

    def run_reports(self):
        """
        This method (running in a worker thread) collects input values,
        validates credentials, and processes the sites in parallel.
        All key steps are logged immediately.
        """
        try:
            username = self.username_entry.get().strip()
            password = self.password_entry.get().strip()
            database = self.database_entry.get().strip()
            site_id_manual = self.siteid_entry.get().strip()
            ip_series_choice = self.ip_series.get()
            from_date_str = self.from_date.get_date().strftime("%Y-%m-%d")
            to_date_str = self.to_date.get_date().strftime("%Y-%m-%d")
            custom_ip = self.custom_ip_entry.get().strip()  # Get the manually entered IP (if any)

            # Determine whether to use file input or manual site ID.
            if self.file_label.cget("text") != "No file selected":
                self.safe_log("File upload mode detected. Reading Site IDs from file...")
                try:
                    ext = os.path.splitext(self.file_path)[1].lower()
                    if ext == '.xlsx':
                        df = pd.read_excel(self.file_path)
                        cols = [col.lower() for col in df.columns]
                        if "siteid" in cols:
                            site_ids = df["siteid"].astype(str).tolist()
                        else:
                            raise Exception("Excel file must contain a column named 'siteid'.")
                    else:
                        with open(self.file_path, "r", encoding="utf-8") as f:
                            file_text = f.read()
                        if "," in file_text:
                            site_ids = [s.strip() for s in file_text.split(",") if s.strip()]
                        else:
                            site_ids = [s.strip() for s in file_text.splitlines() if s.strip()]
                    self.safe_log(f"Found {len(site_ids)} site IDs in file.")
                except Exception as e:
                    self.after(0, messagebox.showerror, "Error", f"Error reading file: {e}")
                    return
            else:
                if not site_id_manual:
                    self.after(0, messagebox.showerror, "Error", "Please enter a Site ID or upload a file.")
                    return
                site_ids = [site_id_manual]
                self.safe_log("Manual input mode selected.")

            # Validate credentials using the first site ID.
            self.safe_log("Validating connection with test connection...")
            test_conn = connect_to_database(site_ids[0], username, password, database, ip_series_choice, custom_ip)
            if not test_conn:
                raise Exception("Test connection failed.")
            test_conn.close()
            self.safe_log("Test connection successful.")

            successful_reports = {}
            failed_sites = {}
            total_sites = len(site_ids)

            # Define helper function to process each site.
            def process_site(sid, index):
                try:
                    self.safe_log(f"Processing site {sid} ({index}/{total_sites})...")
                    result, site_name = get_report_data(sid, from_date_str, to_date_str, username, password, database, ip_series_choice, custom_ip)
                    report_text = format_report(result, sid, site_name, from_date_str, to_date_str)
                    self.safe_log(f"Completed site {sid}.")
                    return (sid, report_text, None)
                except Exception as e:
                    self.safe_log(f"Error processing site {sid}: {e}")
                    return (sid, None, str(e))

            # Process sites in parallel.
            with concurrent.futures.ThreadPoolExecutor(max_workers=min(total_sites, 10)) as executor:
                futures = {executor.submit(process_site, sid, i): sid for i, sid in enumerate(site_ids, start=1)}
                for future in concurrent.futures.as_completed(futures):
                    sid, report_text, error = future.result()
                    if report_text is not None:
                        successful_reports[sid] = report_text
                    else:
                        failed_sites[sid] = error

            self.safe_log("Report generation completed.")
            if successful_reports:
                self.zip_buffer = io.BytesIO()
                with zipfile.ZipFile(self.zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for sid, report in successful_reports.items():
                        zip_file.writestr(f"{sid}.txt", report)
                self.zip_buffer.seek(0)
                self.safe_log("Reports generated successfully. Click on Download Report.")
                self.after(0, lambda: self.download_button.config(state="normal"))
            else:
                self.after(0, messagebox.showerror, "Error", "No successful reports to save.")
                return

            if failed_sites:
                errors = "\n".join([f"{sid}: {err}" for sid, err in failed_sites.items()])
                self.safe_log("Failed Sites:\n" + errors)
                self.after(0, messagebox.showwarning, "Warning", f"Some sites could not be processed:\n{errors}")

        except Exception as e:
            self.safe_log(f"Error: {e}")
            self.after(0, messagebox.showerror, "Error", str(e))

    def download_reports(self):
        """
        Automatically save the ZIP file to the user's Downloads folder as 'SiteReports.zip'.
        If the file already exists, append a counter to the file name.
        """
        try:
            downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            base_filename = "SiteReports"
            extension = ".zip"
            file_path = os.path.join(downloads_folder, base_filename + extension)
            counter = 1
            while os.path.exists(file_path):
                file_path = os.path.join(downloads_folder, f"{base_filename}{counter}{extension}")
                counter += 1
            with open(file_path, "wb") as f:
                f.write(self.zip_buffer.getvalue())
            self.safe_log(f"ZIP file auto-saved to {file_path}")
            messagebox.showinfo("Success", f"Report auto-saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving ZIP file: {e}")

if __name__ == "__main__":
    app = SalesSummaryReportApp()
    app.mainloop()
