import threading
import time
import random
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import schedule
import datetime
import pytz

# --- BLOOMBERG TERMINAL COLOR PALETTE ---
BG_DARK = "#121212"      # Deep terminal background
BG_PANEL = "#1E1E1E"     # Panel background
TEXT_AMBER = "#FFB000"   # Bloomberg amber headers/accents
TEXT_WHITE = "#E0E0E0"   # Primary text
TEXT_DIM = "#888888"     # Secondary/dim text
ACCENT_GREEN = "#00FF66" # Positive indicators / highlights / Bullish candle
BORDER_COLOR = "#333333" # Subtle borders
BTN_BG = "#2A2A2A"       # Button background

class BloombergScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BLOOMBERG PROFESSIONAL - MULTI-TIMEFRAME QUANTITATIVE SCREENER [NSE]")
        self.root.geometry("1450x900")
        self.root.configure(bg=BG_DARK)
        
        self.file_path = None
        self.df_stocks = None
        self.scan_results = []
        
        # Threading & Control Flags
        self.scan_thread = None
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.stop_event = threading.Event()
        
        # Scheduler Configuration Flags
        self.scheduler_active = False
        self.ntfy_topic_default = "my_bloomberg_screener"
        
        self.setup_styles()
        self.build_ui()
        
        # Start background scheduler runner & IST clock ticker
        self.start_scheduler_thread()
        self.update_ist_clock()
        
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Treeview",
                        background=BG_PANEL,
                        foreground=TEXT_WHITE,
                        fieldbackground=BG_PANEL,
                        borderwidth=0,
                        rowheight=26,
                        font=("Consolas", 10))
        style.configure("Treeview.Heading",
                        background=BG_DARK,
                        foreground=TEXT_AMBER,
                        font=("Consolas", 11, "bold"),
                        borderwidth=1)
        style.map("Treeview", background=[('selected', '#333333')], foreground=[('selected', TEXT_AMBER)])
        
        style.configure("TProgressbar",
                        thickness=12,
                        troughcolor=BG_DARK,
                        background=TEXT_AMBER,
                        bordercolor=BORDER_COLOR)

    def build_ui(self):
        header_frame = tk.Frame(self.root, bg=BG_DARK, bd=1, relief="solid")
        header_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        lbl_title = tk.Label(header_frame, text="BBG > MULTI-TIMEFRAME QUANTITATIVE SCREENER [NSE]", 
                             fg=TEXT_AMBER, bg=BG_DARK, font=("Consolas", 13, "bold"))
        lbl_title.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Right Header Widgets (IST Clock + Progress Bar + Status)
        right_header_frame = tk.Frame(header_frame, bg=BG_DARK)
        right_header_frame.pack(side=tk.RIGHT, padx=5, pady=5)
        
        self.lbl_clock = tk.Label(right_header_frame, text="IST: 00:00:00", fg=ACCENT_GREEN, bg=BG_DARK, font=("Consolas", 11, "bold"))
        self.lbl_clock.pack(side=tk.LEFT, padx=(0, 15))
        
        self.progress_bar = ttk.Progressbar(right_header_frame, orient="horizontal", length=160, mode="determinate", style="TProgressbar")
        self.progress_bar.pack(side=tk.LEFT, padx=(0, 15))
        
        self.lbl_status = tk.Label(right_header_frame, text="STATUS: IDLE", fg=TEXT_DIM, bg=BG_DARK, font=("Consolas", 9))
        self.lbl_status.pack(side=tk.LEFT)
        
        main_container = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bg=BORDER_COLOR, bd=0, sashwidth=4)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # --- LEFT PANEL: CONTROLS, TIMEFRAME FILTERS & SCHEDULER ---
        left_panel = tk.Frame(main_container, bg=BG_PANEL, padx=10, pady=10)
        main_container.add(left_panel, width=440)
        
        tk.Label(left_panel, text="1. INPUT DATA (.XLSX)", fg=TEXT_AMBER, bg=BG_PANEL, font=("Consolas", 11, "bold")).pack(anchor="w", pady=(0, 5))
        
        file_frame = tk.Frame(left_panel, bg=BG_PANEL)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_load = tk.Button(file_frame, text="SELECT EXCEL FILE", command=self.load_excel, 
                                  bg=BTN_BG, fg=TEXT_AMBER, font=("Consolas", 9, "bold"), relief="groove", bd=1)
        self.btn_load.pack(side=tk.LEFT, padx=(0, 5))
        
        self.lbl_filename = tk.Label(file_frame, text="No file selected", fg=TEXT_DIM, bg=BG_PANEL, font=("Consolas", 9), wraplength=250, justify="left")
        self.lbl_filename.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Label(left_panel, text="2. MULTI-TIMEFRAME RSI & ADX FILTERS", fg=TEXT_AMBER, bg=BG_PANEL, font=("Consolas", 11, "bold")).pack(anchor="w", pady=(5, 5))
        
        tf_container = tk.Frame(left_panel, bg=BG_PANEL)
        tf_container.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.canvas_tf = tk.Canvas(tf_container, bg=BG_PANEL, highlightthickness=0)
        v_scroll_tf = ttk.Scrollbar(tf_container, orient="vertical", command=self.canvas_tf.yview)
        h_scroll_tf = ttk.Scrollbar(tf_container, orient="horizontal", command=self.canvas_tf.xview)
        
        self.scrollable_tf_frame = tk.Frame(self.canvas_tf, bg=BG_PANEL)
        self.scrollable_tf_frame.bind(
            "<Configure>",
            lambda e: self.canvas_tf.configure(scrollregion=self.canvas_tf.bbox("all"))
        )
        
        self.canvas_tf.create_window((0, 0), window=self.scrollable_tf_frame, anchor="nw")
        self.canvas_tf.configure(yscrollcommand=v_scroll_tf.set, xscrollcommand=h_scroll_tf.set)
        
        v_scroll_tf.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll_tf.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas_tf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.timeframes = ["4h", "Daily", "Weekly", "Monthly", "Quarterly", "Half Yearly"]
        self.tf_vars = {}
        
        for tf in self.timeframes:
            frame_box = tk.LabelFrame(self.scrollable_tf_frame, text=f" {tf} ", fg=TEXT_WHITE, bg=BG_PANEL, font=("Consolas", 9, "bold"), bd=1, relief="solid")
            frame_box.pack(fill=tk.X, pady=4, padx=2)
            
            enable_var = tk.BooleanVar(value=True if tf in ["Daily", "Weekly"] else False)
            rsi_min_var = tk.StringVar(value="40")
            rsi_max_var = tk.StringVar(value="60")
            adx_min_var = tk.StringVar(value="20")
            adx_max_var = tk.StringVar(value="50")
            
            tk.Checkbutton(frame_box, text="Enable", variable=enable_var, bg=BG_PANEL, fg=TEXT_WHITE, selectcolor=BG_DARK, activebackground=BG_PANEL, font=("Consolas", 8)).grid(row=0, column=0, columnspan=4, sticky="w", padx=5)
            
            tk.Label(frame_box, text="RSI Min/Max:", fg=TEXT_DIM, bg=BG_PANEL, font=("Consolas", 8)).grid(row=1, column=0, sticky="w", padx=5)
            tk.Entry(frame_box, textvariable=rsi_min_var, width=5, bg=BG_DARK, fg=TEXT_WHITE, insertbackground=TEXT_WHITE, font=("Consolas", 8)).grid(row=1, column=1, padx=2)
            tk.Entry(frame_box, textvariable=rsi_max_var, width=5, bg=BG_DARK, fg=TEXT_WHITE, insertbackground=TEXT_WHITE, font=("Consolas", 8)).grid(row=1, column=2, padx=2)
            
            tk.Label(frame_box, text="ADX Min/Max:", fg=TEXT_DIM, bg=BG_PANEL, font=("Consolas", 8)).grid(row=2, column=0, sticky="w", padx=5, pady=2)
            tk.Entry(frame_box, textvariable=adx_min_var, width=5, bg=BG_DARK, fg=TEXT_WHITE, insertbackground=TEXT_WHITE, font=("Consolas", 8)).grid(row=2, column=1, padx=2, pady=2)
            tk.Entry(frame_box, textvariable=adx_max_var, width=5, bg=BG_DARK, fg=TEXT_WHITE, insertbackground=TEXT_WHITE, font=("Consolas", 8)).grid(row=2, column=2, padx=2, pady=2)
            
            self.tf_vars[tf] = {
                "enabled": enable_var,
                "rsi_min": rsi_min_var,
                "rsi_max": rsi_max_var,
                "adx_min": adx_min_var,
                "adx_max": adx_max_var
            }
            
        # --- SCHEDULER CONFIGURATION FRAME ---
        sched_box = tk.LabelFrame(left_panel, text=" AUTOMATED MARKET CLOSE SCAN ", fg=TEXT_AMBER, bg=BG_PANEL, font=("Consolas", 9, "bold"), bd=1, relief="solid")
        sched_box.pack(fill=tk.X, pady=(8, 5), padx=2)
        
        tk.Label(sched_box, text="Trigger Time (HH:MM):", fg=TEXT_DIM, bg=BG_PANEL, font=("Consolas", 8)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.sched_time_var = tk.StringVar(value="15:30")
        tk.Entry(sched_box, textvariable=self.sched_time_var, width=8, bg=BG_DARK, fg=TEXT_WHITE, insertbackground=TEXT_WHITE, font=("Consolas", 9)).grid(row=0, column=1, padx=5, pady=5)
        
        self.btn_toggle_sched = tk.Button(sched_box, text="ENABLE SCHEDULE", command=self.toggle_scheduler, 
                                          bg=BTN_BG, fg=ACCENT_GREEN, font=("Consolas", 8, "bold"), relief="groove", bd=1, cursor="hand2")
        self.btn_toggle_sched.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        
        # --- EXECUTION CONTROL BUTTONS (START, PAUSE, RESUME, STOP) ---
        control_frame = tk.Frame(left_panel, bg=BG_PANEL)
        control_frame.pack(fill=tk.X, pady=(5, 5))
        
        self.btn_start = tk.Button(control_frame, text="START", command=self.start_scan, bg=ACCENT_GREEN, fg=BG_DARK, font=("Consolas", 9, "bold"), width=8, relief="flat", cursor="hand2")
        self.btn_start.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        self.btn_pause = tk.Button(control_frame, text="PAUSE", command=self.pause_scan, bg=TEXT_AMBER, fg=BG_DARK, font=("Consolas", 9, "bold"), width=8, relief="flat", cursor="hand2", state=tk.DISABLED)
        self.btn_pause.grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        
        self.btn_resume = tk.Button(control_frame, text="RESUME", command=self.resume_scan, bg="#00AAFF", fg=BG_DARK, font=("Consolas", 9, "bold"), width=8, relief="flat", cursor="hand2", state=tk.DISABLED)
        self.btn_resume.grid(row=0, column=2, padx=2, pady=2, sticky="ew")
        
        self.btn_stop = tk.Button(control_frame, text="STOP", command=self.stop_scan, bg="#FF4444", fg=TEXT_WHITE, font=("Consolas", 9, "bold"), width=8, relief="flat", cursor="hand2", state=tk.DISABLED)
        self.btn_stop.grid(row=0, column=3, padx=2, pady=2, sticky="ew")
        
        control_frame.columnconfigure((0, 1, 2, 3), weight=1)

        # --- RIGHT PANEL: RESULTS TABLE & NTFY/EXCEL BUTTONS ---
        right_panel = tk.Frame(main_container, bg=BG_PANEL, padx=10, pady=10)
        main_container.add(right_panel, width=980)
        
        results_header_frame = tk.Frame(right_panel, bg=BG_PANEL)
        results_header_frame.pack(fill=tk.X, pady=(0, 5))
        
        tk.Label(results_header_frame, text="3. SCAN RESULTS", fg=TEXT_AMBER, bg=BG_PANEL, font=("Consolas", 11, "bold")).pack(side=tk.LEFT)
        
        btn_action_frame = tk.Frame(results_header_frame, bg=BG_PANEL)
        btn_action_frame.pack(side=tk.RIGHT)
        
        self.btn_test_ntfy = tk.Button(btn_action_frame, text="TEST NOTIFICATION", command=self.test_ntfy_connection, 
                                       bg=BTN_BG, fg="#00AAFF", font=("Consolas", 9, "bold"), relief="groove", bd=1, cursor="hand2")
        self.btn_test_ntfy.pack(side=tk.LEFT, padx=(0, 5))
        
        self.btn_ntfy = tk.Button(btn_action_frame, text="PUSH VIA NTFY", command=self.send_ntfy_alert, 
                                  bg=BTN_BG, fg=TEXT_AMBER, font=("Consolas", 9, "bold"), relief="groove", bd=1, cursor="hand2")
        self.btn_ntfy.pack(side=tk.LEFT, padx=(0, 5))
        
        self.btn_export = tk.Button(btn_action_frame, text="EXPORT TO EXCEL", command=self.export_to_excel, 
                                    bg=BTN_BG, fg=ACCENT_GREEN, font=("Consolas", 9, "bold"), relief="groove", bd=1, cursor="hand2")
        self.btn_export.pack(side=tk.LEFT)
        
        table_frame = tk.Frame(right_panel, bg=BG_PANEL)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("Name", "NSE Code", "Timeframe", "LTP", "RSI", "ADX", "Confidence Level")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("Name", text="Name")
        self.tree.column("Name", width=160, anchor="w", stretch=True)
        
        self.tree.heading("NSE Code", text="NSE Code")
        self.tree.column("NSE Code", width=100, anchor="w")
        
        self.tree.heading("Timeframe", text="Timeframe")
        self.tree.column("Timeframe", width=90, anchor="center")
        
        self.tree.heading("LTP", text="LTP")
        self.tree.column("LTP", width=80, anchor="e")
        
        self.tree.heading("RSI", text="RSI")
        self.tree.column("RSI", width=70, anchor="center")
        
        self.tree.heading("ADX", text="ADX")
        self.tree.column("ADX", width=70, anchor="center")
        
        self.tree.heading("Confidence Level", text="Confidence Level")
        self.tree.column("Confidence Level", width=130, anchor="center")
        
        v_scroll_tree = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        h_scroll_tree = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll_tree.set, xscrollcommand=h_scroll_tree.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scroll_tree.grid(row=0, column=1, sticky="ns")
        h_scroll_tree.grid(row=1, column=0, sticky="ew")
        
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        
        self.tree.bind("<Double-1>", self.on_stock_double_click)

    def update_ist_clock(self):
        try:
            ist_tz = pytz.timezone('Asia/Kolkata')
            current_ist = datetime.datetime.now(ist_tz).strftime('%H:%M:%S IST')
            self.lbl_clock.config(text=f"IST: {current_ist}")
        except Exception:
            pass
        self.root.after(1000, self.update_ist_clock)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                if not {'SYMBOL', 'NAME OF COMPANY'}.issubset(df.columns):
                    messagebox.showerror("Error", "Excel file must contain 'SYMBOL' and 'NAME OF COMPANY' columns.")
                    return
                self.file_path = file_path
                self.df_stocks = df
                filename_short = file_path.split("/")[-1]
                self.lbl_filename.config(text=filename_short, fg=ACCENT_GREEN)
                self.lbl_status.config(text=f"LOADED {len(df)} STOCKS", fg=ACCENT_GREEN)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")

    def export_to_excel(self):
        if not self.scan_results:
            messagebox.showwarning("Warning", "No scan results available to export.")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                df_export = pd.DataFrame(self.scan_results)
                df_export.to_excel(file_path, index=False)
                messagebox.showinfo("Success", f"Scan results successfully exported to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export file:\n{str(e)}")

    def test_ntfy_connection(self):
        topic = simpledialog.askstring("ntfy Connection Test", "Enter your ntfy.sh topic name:", initialvalue=self.ntfy_topic_default)
        if not topic:
            return
        self.ntfy_topic_default = topic.strip()
        
        try:
            url = f"https://ntfy.sh/{self.ntfy_topic_default}"
            data = "Bloomberg Terminal Screener: Connection test successful!".encode('utf-8')
            req = urllib.request.Request(url, data=data, headers={"Title": "Connection Test", "Priority": "default"})
            with urllib.request.urlopen(req, timeout=5) as response:
                if response.status == 200:
                    messagebox.showinfo("Success", f"Test notification successfully sent to ntfy.sh/{self.ntfy_topic_default}!")
                else:
                    messagebox.showwarning("Warning", f"Received HTTP status: {response.status}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to connect to ntfy.sh:\n{str(e)}")

    def send_ntfy_alert(self, auto_triggered=False):
        if not self.scan_results:
            if not auto_triggered:
                messagebox.showwarning("Warning", "No scan results available to push via ntfy.")
            return
            
        if not auto_triggered:
            topic = simpledialog.askstring("ntfy Alert", "Enter your ntfy.sh topic name:", initialvalue=self.ntfy_topic_default)
            if not topic:
                return
            self.ntfy_topic_default = topic.strip()
            
        try:
            url = f"https://ntfy.sh/{self.ntfy_topic_default}"
            
            msg_lines = [f"Found {len(self.scan_results)} matching setups:\n"]
            for res in self.scan_results[:15]:
                msg_lines.append(f"{res['Name']} ({res['NSE Code']}) | TF: {res['Timeframe']} | LTP: {res['LTP']} | RSI: {res['RSI']} | ADX: {res['ADX']} | Conf: {res['Confidence Level']}")
            
            if len(self.scan_results) > 15:
                msg_lines.append(f"\n...and {len(self.scan_results) - 15} more stocks.")
                
            payload = "\n".join(msg_lines).encode('utf-8')
            req = urllib.request.Request(url, data=payload, headers={
                "Title": f"BBG Market Close Alert ({len(self.scan_results)} Matches)",
                "Priority": "high",
                "Tags": "chart,bar_chart"
            })
            
            with urllib.request.urlopen(req, timeout=8) as response:
                if response.status == 200 and not auto_triggered:
                    messagebox.showinfo("Success", f"Successfully pushed scan results to ntfy.sh/{self.ntfy_topic_default}!")
        except Exception as e:
            print(f"Failed to push alert via ntfy: {e}")

    def toggle_scheduler(self):
        if not self.scheduler_active:
            target_time = self.sched_time_var.get().strip()
            try:
                datetime.datetime.strptime(target_time, "%H:%M")
                schedule.clear('market_scan')
                schedule.every().day.at(target_time).do(self.scheduled_scan_job).tag('market_scan')
                
                self.scheduler_active = True
                self.btn_toggle_sched.config(text="DISABLE SCHEDULE", fg="#FF4444")
                messagebox.showinfo("Scheduler Enabled", f"Automated background scan scheduled daily at {target_time} IST.")
            except ValueError:
                messagebox.showerror("Error", "Invalid time format! Please use HH:MM (24-hour format, e.g., 15:30).")
        else:
            schedule.clear('market_scan')
            self.scheduler_active = False
            self.btn_toggle_sched.config(text="ENABLE SCHEDULE", fg=ACCENT_GREEN)
            messagebox.showinfo("Scheduler Disabled", "Automated background schedule has been turned off.")

    def start_scheduler_thread(self):
        def run_scheduler_loop():
            while True:
                schedule.run_pending()
                time.sleep(1)
        t = threading.Thread(target=run_scheduler_loop, daemon=True)
        t.start()

    def scheduled_scan_job(self):
        if self.df_stocks is None:
            print("Scheduled scan skipped: No Excel stock list loaded.")
            return
        print("Executing scheduled background market close scan...")
        self.run_scan_background_sync()
        self.send_ntfy_alert(auto_triggered=True)

    def run_scan_background_sync(self):
        self.scan_results = []
        enabled_tfs = {tf: vals for tf, vals in self.tf_vars.items() if vals["enabled"].get()}
        if not enabled_tfs:
            return
            
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(self.scan_single_stock, row, enabled_tfs) for _, row in self.df_stocks.iterrows()]
            for future in as_completed(futures):
                res_list = future.result()
                if res_list:
                    for r in res_list:
                        self.scan_results.append(r)

    def start_scan(self):
        if self.df_stocks is None:
            messagebox.showwarning("Warning", "Please load an Excel file first.")
            return
        
        self.stop_event.clear()
        self.pause_event.set()
        
        self.btn_start.config(state=tk.DISABLED, bg=TEXT_DIM)
        self.btn_pause.config(state=tk.NORMAL)
        self.btn_resume.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.btn_load.config(state=tk.DISABLED)
        
        self.progress_bar['value'] = 0
        self.lbl_status.config(text="SCANNING IN PROGRESS...", fg=TEXT_AMBER)
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.scan_results = []
            
        self.scan_thread = threading.Thread(target=self.run_scan_worker)
        self.scan_thread.daemon = True
        self.scan_thread.start()

    def pause_scan(self):
        self.pause_event.clear()
        self.btn_pause.config(state=tk.DISABLED, bg=TEXT_DIM)
        self.btn_resume.config(state=tk.NORMAL, bg="#00AAFF")
        self.lbl_status.config(text="SCAN PAUSED", fg=TEXT_AMBER)

    def resume_scan(self):
        self.pause_event.set()
        self.btn_pause.config(state=tk.NORMAL, bg=TEXT_AMBER)
        self.btn_resume.config(state=tk.DISABLED, bg=BTN_BG)
        self.lbl_status.config(text="SCANNING RESUMED...", fg=TEXT_AMBER)

    def stop_scan(self):
        self.stop_event.set()
        self.pause_event.set()
        self.lbl_status.config(text="SCAN STOPPED", fg="#FF4444")
        self.progress_bar['value'] = 0
        self.reset_controls()

    def reset_controls(self):
        self.btn_start.config(state=tk.NORMAL, bg=ACCENT_GREEN)
        self.btn_pause.config(state=tk.DISABLED, bg=TEXT_AMBER)
        self.btn_resume.config(state=tk.DISABLED, bg=BTN_BG)
        self.btn_stop.config(state=tk.DISABLED, bg="#FF4444")
        self.btn_load.config(state=tk.NORMAL)

    def calculate_rsi(self, series, period=14):
        delta = series.diff()
        gain = delta.clip(lower=0)
        loss = -1 * delta.clip(upper=0)
        avg_gain = gain.ewm(alpha=1/period, min_periods=period, adjust=False).mean()
        avg_loss = loss.ewm(alpha=1/period, min_periods=period, adjust=False).mean()
        rs = avg_gain / avg_loss
        rsi = 100 - (100 / (1 + rs))
        return rsi

    def calculate_adx_and_dmi(self, df, period=14):
        high = df['High']
        low = df['Low']
        close = df['Close']
        
        plus_dm = high.diff()
        minus_dm = low.diff()
        plus_dm[plus_dm < 0] = 0
        minus_dm[minus_dm > 0] = 0
        minus_dm = abs(minus_dm)
        
        tr1 = high - low
        tr2 = abs(high - close.shift(1))
        tr3 = abs(low - close.shift(1))
        tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
        
        tr_smoothed = tr.rolling(period).sum()
        plus_di = 100 * (plus_dm.rolling(period).sum() / tr_smoothed)
        minus_di = 100 * (minus_dm.rolling(period).sum() / tr_smoothed)
        
        dx = 100 * abs(plus_di - minus_di) / (plus_di + minus_di)
        adx = dx.rolling(period).mean()
        return adx, plus_di, minus_di

    def get_timeframe_params(self, tf_name):
        mapping = {
            "4h": {"interval": "60m", "period": "60d", "resample": "4H"},
            "Daily": {"interval": "1d", "period": "1y", "resample": None},
            "Weekly": {"interval": "1wk", "period": "2y", "resample": None},
            "Monthly": {"interval": "1mo", "period": "5y", "resample": None},
            "Quarterly": {"interval": "3mo", "period": "10y", "resample": None},
            "Half Yearly": {"interval": "6mo", "period": "15y", "resample": None}
        }
        return mapping.get(tf_name, mapping["Daily"])

    def fetch_timeframe_data(self, ticker_symbol, tf_name, retries=3):
        params = self.get_timeframe_params(tf_name)
        for attempt in range(retries):
            try:
                time.sleep(random.uniform(0.1, 0.3))
                ticker = yf.Ticker(ticker_symbol)
                hist = ticker.history(period=params["period"], interval=params["interval"])
                if not hist.empty and len(hist) >= 20:
                    if tf_name == "4h":
                        hist = hist.resample('4H').agg({'Open': 'first', 'High': 'max', 'Low': 'min', 'Close': 'last', 'Volume': 'sum'}).dropna()
                    return hist
                break
            except Exception as e:
                if "Too Many Requests" in str(e) or "Rate limited" in str(e):
                    sleep_time = (2 ** attempt) + random.uniform(1, 3)
                    time.sleep(sleep_time)
                else:
                    break
        return pd.DataFrame()

    def scan_single_stock(self, row, enabled_tfs):
        if self.stop_event.is_set():
            return []
        
        self.pause_event.wait()
        if self.stop_event.is_set():
            return []

        raw_symbol = str(row['SYMBOL']).strip().upper()
        comp_name = str(row['NAME OF COMPANY']).strip()
        ticker_symbol = raw_symbol if raw_symbol.endswith(".NS") else f"{raw_symbol}.NS"
        
        stock_results = []
        try:
            stock_pass = True
            total_criteria = 0
            passed_criteria = 0
            
            for tf_name, cfg in enabled_tfs.items():
                if self.stop_event.is_set():
                    break
                
                hist = self.fetch_timeframe_data(ticker_symbol, tf_name)
                if hist.empty or len(hist) < 20:
                    stock_pass = False
                    break
                    
                rsi_series = self.calculate_rsi(hist['Close'])
                adx_series, _, _ = self.calculate_adx_and_dmi(hist)
                
                current_rsi = float(rsi_series.iloc[-1])
                current_adx = float(adx_series.iloc[-1])
                
                r_min = float(cfg["rsi_min"].get())
                r_max = float(cfg["rsi_max"].get())
                a_min = float(cfg["adx_min"].get())
                a_max = float(cfg["adx_max"].get())
                
                total_criteria += 2
                if r_min <= current_rsi <= r_max:
                    passed_criteria += 1
                if a_min <= current_adx <= a_max:
                    passed_criteria += 1
                    
                if not (r_min <= current_rsi <= r_max and a_min <= current_adx <= a_max):
                    stock_pass = False
            
            if stock_pass and total_criteria > 0 and not self.stop_event.is_set():
                confidence = round((passed_criteria / total_criteria) * 100, 1)
                
                for tf_name in enabled_tfs.keys():
                    hist = self.fetch_timeframe_data(ticker_symbol, tf_name)
                    if not hist.empty:
                        ltp = float(hist['Close'].iloc[-1])
                        rsi_val = float(self.calculate_rsi(hist['Close']).iloc[-1])
                        adx_val, _, _ = self.calculate_adx_and_dmi(hist)
                        adx_val = float(adx_val.iloc[-1])
                        
                        res_dict = {
                            "Name": comp_name,
                            "NSE Code": ticker_symbol,
                            "Timeframe": tf_name,
                            "LTP": f"{ltp:.2f}",
                            "RSI": f"{rsi_val:.1f}",
                            "ADX": f"{adx_val:.1f}",
                            "Confidence Level": f"{confidence}%"
                        }
                        stock_results.append(res_dict)
        except Exception as e:
            print(f"Error scanning {ticker_symbol}: {e}")
            
        return stock_results

    def run_scan_worker(self):
        enabled_tfs = {tf: vals for tf, vals in self.tf_vars.items() if vals["enabled"].get()}
        
        if not enabled_tfs:
            self.root.after(0, lambda: messagebox.showwarning("Warning", "Please enable at least one timeframe."))
            self.root.after(0, self.reset_controls)
            return

        total_stocks = len(self.df_stocks)
        completed_count = 0
        results = []

        with ThreadPoolExecutor(max_workers=4) as executor:
            future_to_stock = {
                executor.submit(self.scan_single_stock, row, enabled_tfs): row 
                for _, row in self.df_stocks.iterrows()
            }
            
            for future in as_completed(future_to_stock):
                if self.stop_event.is_set():
                    break
                
                completed_count += 1
                progress_pct = (completed_count / total_stocks) * 100
                
                stock_res = future.result()
                if stock_res:
                    for res_dict in stock_res:
                        results.append(res_dict)
                        self.scan_results = results
                        self.root.after(0, lambda r=res_dict: self.tree.insert("", tk.END, values=(
                            r["Name"], r["NSE Code"], r["Timeframe"], r["LTP"], r["RSI"], r["ADX"], r["Confidence Level"])))
                
                self.root.after(0, lambda p=progress_pct, c=completed_count, t=total_stocks: [
                    self.lbl_status.config(text=f"PROGRESS: {c}/{t}", fg=TEXT_AMBER),
                    self.progress_bar.config(value=p)
                ])
                
        if not self.stop_event.is_set():
            count = len(results)
            self.root.after(0, lambda: [
                self.lbl_status.config(text=f"COMPLETE ({count} MATCHES)", fg=ACCENT_GREEN),
                self.progress_bar.config(value=100)
            ])
            self.root.after(0, self.reset_controls)

    def on_stock_double_click(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        item_data = self.tree.item(selected_item)
        values = item_data['values']
        if not values:
            return
        
        comp_name = values[0]
        ticker_code = values[1]
        timeframe = values[2]
        self.open_chart_window(ticker_code, comp_name, timeframe)

    def open_chart_window(self, ticker_code, comp_name, timeframe):
        chart_win = tk.Toplevel(self.root)
        chart_win.title(f"BBG CHART ANALYSIS: {comp_name} ({ticker_code}) [{timeframe}]")
        chart_win.geometry("980x880")
        chart_win.configure(bg=BG_DARK)
        
        tk.Label(chart_win, text=f"TECHNICAL ANALYSIS - {comp_name} [{ticker_code}] | TIMEFRAME: {timeframe}", fg=TEXT_AMBER, bg=BG_DARK, font=("Consolas", 12, "bold")).pack(pady=8)
        
        try:
            df = self.fetch_timeframe_data(ticker_code, timeframe)
            
            if df.empty or len(df) < 20:
                tk.Label(chart_win, text=f"Error: Insufficient data for {timeframe} timeframe.", fg="red", bg=BG_DARK).pack(pady=20)
                return
                
            rsi = self.calculate_rsi(df['Close'])
            adx, plus_di, minus_di = self.calculate_adx_and_dmi(df)
            
            fig, (ax1, ax2, ax3, ax4) = plt.subplots(4, 1, figsize=(9.5, 8.2), sharex=True, gridspec_kw={'height_ratios': [3, 1.3, 1.3, 1.3]})
            fig.patch.set_facecolor(BG_DARK)
            
            for ax in [ax1, ax2, ax3, ax4]:
                ax.set_facecolor(BG_PANEL)
                ax.tick_params(colors=TEXT_WHITE, labelsize=8)
                ax.xaxis.label.set_color(TEXT_WHITE)
                ax.yaxis.label.set_color(TEXT_WHITE)
                for spine in ax.spines.values():
                    spine.set_edgecolor(BORDER_COLOR)
            
            # --- 1. CANDLESTICK CHART SUBPLOT ---
            x_vals = range(len(df))
            for i, (idx, row) in enumerate(df.iterrows()):
                op = row['Open']
                cl = row['Close']
                hi = row['High']
                lo = row['Low']
                
                color = ACCENT_GREEN if cl >= op else "#FF4444"
                ax1.vlines(i, lo, hi, color=TEXT_WHITE, linewidth=0.8, alpha=0.8)
                bottom = min(op, cl)
                height = abs(cl - op)
                if height == 0:
                    height = 0.01
                rect = plt.Rectangle((i - 0.35, bottom), 0.7, height, facecolor=color, edgecolor=color)
                ax1.add_patch(rect)
                
            ax1.set_xlim(-1, len(df))
            ax1.set_title(f"Candlestick Price Action ({timeframe})", color=TEXT_WHITE, fontsize=9, fontname="Consolas")
            ax1.grid(True, color=BORDER_COLOR, linestyle="--", alpha=0.5)
            
            # --- 2. RSI SUBPLOT ---
            ax2.plot(x_vals, rsi, color="#00AAFF", linewidth=1.2, label="RSI (14)")
            ax2.axhline(60, color="#FF4444", linestyle="--", alpha=0.7, label="Overbought (60)")
            ax2.axhline(40, color=ACCENT_GREEN, linestyle="--", alpha=0.7, label="Oversold (40)")
            ax2.set_title("Relative Strength Index (RSI 60/40)", color=TEXT_WHITE, fontsize=9, fontname="Consolas")
            ax2.legend(loc="upper left", facecolor=BG_PANEL, edgecolor=BORDER_COLOR, labelcolor=TEXT_WHITE, fontsize=8)
            ax2.set_ylim(0, 100)
            ax2.grid(True, color=BORDER_COLOR, linestyle="--", alpha=0.5)
            
            # --- 3. ADX SUBPLOT ---
            ax3.plot(x_vals, adx, color="#FFB000", linewidth=1.2, label="ADX (14)")
            ax3.axhline(20, color=TEXT_DIM, linestyle="--", alpha=0.7, label="Weak Trend (20)")
            ax3.axhline(50, color="#FF4444", linestyle="--", alpha=0.7, label="Strong Trend (50)")
            ax3.set_title("Average Directional Index (ADX 20/50)", color=TEXT_WHITE, fontsize=9, fontname="Consolas")
            ax3.legend(loc="upper left", facecolor=BG_PANEL, edgecolor=BORDER_COLOR, labelcolor=TEXT_WHITE, fontsize=8)
            ax3.set_ylim(0, 100)
            ax3.grid(True, color=BORDER_COLOR, linestyle="--", alpha=0.5)
            
            # --- 4. DMI SUBPLOT ---
            ax4.plot(x_vals, plus_di, color=ACCENT_GREEN, linewidth=1.2, label="+DI (Positive)")
            ax4.plot(x_vals, minus_di, color="#FF4444", linewidth=1.2, label="-DI (Negative)")
            ax4.set_title("Directional Movement Index (DMI)", color=TEXT_WHITE, fontsize=9, fontname="Consolas")
            ax4.legend(loc="upper left", facecolor=BG_PANEL, edgecolor=BORDER_COLOR, labelcolor=TEXT_WHITE, fontsize=8)
            ax4.grid(True, color=BORDER_COLOR, linestyle="--", alpha=0.5)
            
            date_labels = [d.strftime('%Y-%m-%d') for d in df.index]
            n_ticks = min(6, len(date_labels))
            step = max(1, len(date_labels) // n_ticks)
            tick_indices = range(0, len(date_labels), step)
            ax4.set_xticks(tick_indices)
            ax4.set_xticklabels([date_labels[i] for i in tick_indices], rotation=15, fontsize=8)
            
            plt.tight_layout()
            
            canvas = FigureCanvasTkAgg(fig, master=chart_win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
        except Exception as e:
            tk.Label(chart_win, text=f"Failed to load timeframe charts & indicators: {str(e)}", fg="red", bg=BG_DARK).pack(pady=20)

if __name__ == "__main__":
    root = tk.Tk()
    app = BloombergScannerApp(root)
    root.mainloop()
