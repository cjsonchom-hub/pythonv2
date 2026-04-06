import customtkinter as ctk  
from tkinter import messagebox  
import tkinter.ttk as ttk  
import tkinter as tk  
import random  
from datetime import datetime, timedelta  
from config import *  

ctk.set_appearance_mode("dark")  
ctk.set_default_color_theme("dark-blue")  

ICONS = {  
    "dashboard": "▣",  
    "import":    "↓",  
    "report":    "↗",  
    "inventory": "▤",  
    "settings":  "✦",  
    "logout":    "⏻",  
    "logo":      "◈  ERP SYSTEM",  
}  

DEPT_REPORTS = {  
    "ADMIN": {  
        "All Departments": [  
            "Monthly Sales Report",  
            "Quarterly Summary",  
            "Annual Financial Report",  
            "User Activity Log",  
            "Purchase Order Summary",  
            "Inventory Stock Report",  
            "HR Headcount Report",  
            "Finance P&L Statement",  
        ],  
        "Purchase": [  
            "Purchase Order Summary",  
            "Vendor Performance Report",  
            "Purchase vs Budget",  
        ],  
        "Inventory": [  
            "Inventory Stock Report",  
            "Low Stock Alert Report",  
            "Stock Movement History",  
        ],  
        "HR": [  
            "HR Headcount Report",  
            "Attendance Summary",  
            "Payroll Summary",  
        ],  
        "Finance": [  
            "Finance P&L Statement",  
            "Monthly Sales Report",  
            "Quarterly Summary",  
            "Annual Financial Report",  
        ],  
        "Sales": [  
            "Monthly Sales Report",  
            "Quarterly Summary",  
            "Customer Report",  
        ],  
    },  
    "FINANCE": {  
        "Finance": [  
            "Finance P&L Statement",  
            "Monthly Sales Report",  
            "Quarterly Summary",  
            "Annual Financial Report",  
        ],  
    },  
    "PURCHASE": {  
        "Purchase": [  
            "Purchase Order Summary",  
            "Vendor Performance Report",  
            "Purchase vs Budget",  
        ],  
    },  
    "HR": {  
        "HR": [  
            "HR Headcount Report",  
            "Attendance Summary",  
            "Payroll Summary",  
        ],  
    },  
    "SALES": {  
        "Sales": [  
            "Monthly Sales Report",  
            "Quarterly Summary",  
            "Customer Report",  
        ],  
    },  
    "INVENTORY": {  
        "Inventory": [  
            "Inventory Stock Report",  
            "Low Stock Alert Report",  
            "Stock Movement History",  
        ],  
    },  
    "USER": {  
        "General": [  
            "Monthly Sales Report",  
            "User Activity Log",  
        ],  
    },  
}  

RAINBOW = [  
    "#FF6B6B", "#FF9F43", "#FECA57",  
    "#48DBFB", "#FF9FF3", "#54A0FF", "#5F27CD",  
]  

TAG_STYLE = {  
    "UPDATE": {"bg": "#1A2A3A", "fg": "#3498DB"},  
    "NEW":    {"bg": "#1A3A1A", "fg": "#2ECC71"},  
    "FIX":    {"bg": "#3A1A1A", "fg": "#E74C3C"},  
    "INFO":   {"bg": "#2A2A1A", "fg": "#F39C12"},  
}  

# neon config: (color_on, color_off, bg)  
NEON_STATUS = {  
    "Running":   ("#39FF14", "#0D3300", "#0A1A0A"),  
    "Connected": ("#39FF14", "#0D3300", "#0A1A0A"),  
    "Standby":   ("#FFD700", "#332B00", "#1A1600"),  
    "Stopped":   ("#FF3131", "#330A0A", "#1A0A0A"),  
}  


class MainMenuWindow:  

    def __init__(self, root, user_data, logout_callback):  
        self.root            = root  
        self.user_data       = user_data  
        self.logout_callback = logout_callback  
        self.current_frame   = None  
        self.active_btn      = None  
        self._clock_job      = None  
        self._neon_jobs      = []  

        self.setup_window()  
        self.build_layout()  
        self.show_dashboard()  

    # ────────────────────────────────────────────────────────────  
    def setup_window(self):  
        self.root.title("ERP System")  
        w, h = 1280, 720  
        sw   = self.root.winfo_screenwidth()  
        sh   = self.root.winfo_screenheight()  
        x    = (sw - w) // 2  
        y    = (sh - h) // 2  
        self.root.geometry(f"{w}x{h}+{x}+{y}")  
        self.root.minsize(1100, 650)  
        self.root.configure(bg="#0D0D0D")  

    # ────────────────────────────────────────────────────────────  
    def build_layout(self):  
        self.root.grid_columnconfigure(1, weight=1)  
        self.root.grid_rowconfigure(0, weight=1)  
        self._build_sidebar()  
        self._build_content_area()  

    # ────────────────────────────────────────────────────────────  
    # SIDEBAR  
    # ────────────────────────────────────────────────────────────  
    def _build_sidebar(self):  
        self.sidebar = ctk.CTkFrame(  
            self.root,  
            width=240,  
            corner_radius=0,  
            fg_color="#111111",  
        )  
        self.sidebar.grid(row=0, column=0, sticky="nsew")  
        self.sidebar.grid_propagate(False)  

        logo_frame = ctk.CTkFrame(  
            self.sidebar, fg_color="#1A1A1A",  
            corner_radius=0, height=70,  
        )  
        logo_frame.pack(fill="x")  
        logo_frame.pack_propagate(False)  

        ctk.CTkLabel(  
            logo_frame,  
            text=ICONS["logo"],  
            font=ctk.CTkFont(size=17, weight="bold"),  
            text_color="#FFFFFF",  
        ).place(relx=0.5, rely=0.5, anchor="center")  

        ctk.CTkFrame(self.sidebar, height=1, fg_color="#2A2A2A").pack(fill="x")  

        user_box = ctk.CTkFrame(  
            self.sidebar, fg_color="#1A1A1A", corner_radius=10,  
        )  
        user_box.pack(fill="x", padx=12, pady=12)  

        ctk.CTkLabel(  
            user_box,  
            text=f"  {self.user_data[0]}",  
            font=ctk.CTkFont(size=13, weight="bold"),  
            text_color="#FFFFFF", anchor="w",  
        ).pack(fill="x", padx=10, pady=(10, 2))  

        ctk.CTkLabel(  
            user_box,  
            text=f"  {self.user_data[1].upper()}",  
            font=ctk.CTkFont(size=11),  
            text_color="#666666", anchor="w",  
        ).pack(fill="x", padx=10, pady=(0, 10))  

        ctk.CTkFrame(self.sidebar, height=1, fg_color="#2A2A2A").pack(fill="x", pady=(0, 8))  

        ctk.CTkLabel(  
            self.sidebar,  
            text="  NAVIGATION",  
            font=ctk.CTkFont(size=10),  
            text_color="#444444", anchor="w",  
        ).pack(fill="x", padx=15, pady=(4, 6))  

        menu_scroll = ctk.CTkScrollableFrame(  
            self.sidebar,  
            fg_color="transparent",  
            scrollbar_button_color="#2A2A2A",  
            scrollbar_button_hover_color="#3A3A3A",  
        )  
        menu_scroll.pack(fill="both", expand=True, padx=8)  

        role     = self.user_data[1].upper()  
        is_admin = role == "ADMIN"  

        menu_items = [  
            ("▣   News & Updates",    self.show_dashboard,       True,                          False),  
            ("↗   Report & Query",    self.show_export,           True,                          False),  
            ("▤   Inventory",         self.show_inventory,        True,                          False),  
         #   ("💻  IT Equipment",      self.show_it_equipment,     is_admin,                      False),
            ("◈   Finance",           self.show_finance,          role in ["ADMIN", "FINANCE"],  False),  
            ("✦   Settings",          self.show_settings,         is_admin,                      False),  
        ]  

        self.menu_buttons = {}  
        for text, callback, visible, is_sub in menu_items:  
            if not visible:  
                continue  
            btn = ctk.CTkButton(  
                menu_scroll,  
                text=f"  {text}",  
                font=ctk.CTkFont(size=12 if is_sub else 13),  
                anchor="w",  
                height=38 if is_sub else 42,  
                corner_radius=8,  
                fg_color="transparent",  
                text_color="#666666" if is_sub else "#AAAAAA",  
                hover_color="#1E1E1E",  
                command=lambda cb=callback, lb=text.strip(): self._nav(cb, lb),  
            )  
            btn.pack(fill="x", pady=1)  
            self.menu_buttons[text.strip()] = btn  

        ctk.CTkFrame(self.sidebar, height=1, fg_color="#2A2A2A").pack(  
            fill="x", padx=12, pady=8  
        )  

        ctk.CTkButton(  
            self.sidebar,  
            text=f"  {ICONS['logout']}    Logout",  
            font=ctk.CTkFont(size=13),  
            anchor="w",  
            height=42,  
            corner_radius=8,  
            fg_color="transparent",  
            text_color="#E74C3C",  
            hover_color="#2A1A1A",  
            command=self.logout,  
        ).pack(fill="x", padx=8, pady=(0, 12))  

    # ────────────────────────────────────────────────────────────  
    def _nav(self, callback, label):  
        self._cancel_neon()  
        if self._clock_job:  
            self.root.after_cancel(self._clock_job)  
            self._clock_job = None  
        for btn in self.menu_buttons.values():  
            btn.configure(fg_color="transparent", text_color="#AAAAAA")  
        if label in self.menu_buttons:  
            self.menu_buttons[label].configure(  
                fg_color="#1E1E1E", text_color="#FFFFFF"  
            )  
        callback()  

    # ────────────────────────────────────────────────────────────  
    # CONTENT AREA  
    # ────────────────────────────────────────────────────────────  
    def _build_content_area(self):  
        self.content_area = ctk.CTkFrame(  
            self.root, corner_radius=0, fg_color="#0D0D0D",  
        )  
        self.content_area.grid(row=0, column=1, sticky="nsew")  
        self.content_area.grid_columnconfigure(0, weight=1)  
        self.content_area.grid_rowconfigure(0, weight=1)  
        self._draw_noise_overlay()  

    def _draw_noise_overlay(self):  
        canvas = tk.Canvas(  
            self.content_area, bg="#0D0D0D", highlightthickness=0,  
        )  
        canvas.place(x=0, y=0, relwidth=1, relheight=1)  
        for _ in range(4000):  
            x = random.randint(0, 1920)  
            y = random.randint(0, 1080)  
            g = random.randint(18, 32)  
            c = f"#{g:02x}{g:02x}{g:02x}"  
            canvas.create_rectangle(x, y, x + 1, y + 1, fill=c, outline="")  
        canvas.lower()  

    # ────────────────────────────────────────────────────────────  
    # NEON ENGINE  
    # ────────────────────────────────────────────────────────────  
    def _cancel_neon(self):  
        for job in self._neon_jobs:  
            try:  
                self.root.after_cancel(job)  
            except Exception:  
                pass  
        self._neon_jobs.clear()  

    def _neon_blink(self, widget, color_on, color_off, ms_on=900, ms_off=250):  
        alive = {"v": True}  

        def glow():  
            if not alive["v"]:  
                return  
            try:  
                widget.configure(text_color=color_on)  
            except Exception:  
                alive["v"] = False  
                return  
            j = self.root.after(ms_on, fade)  
            self._neon_jobs.append(j)  

        def fade():  
            if not alive["v"]:  
                return  
            try:  
                widget.configure(text_color=color_off)  
            except Exception:  
                alive["v"] = False  
                return  
            j = self.root.after(ms_off, glow)  
            self._neon_jobs.append(j)  

        widget._neon_alive = alive  
        glow()  

    # ────────────────────────────────────────────────────────────  
    # BASE HELPERS  
    # ────────────────────────────────────────────────────────────  
    def _clear(self):  
        self._cancel_neon()  
        if self._clock_job:  
            self.root.after_cancel(self._clock_job)  
            self._clock_job = None  
        if self.current_frame:  
            self.current_frame.destroy()  

    def _base_frame(self, title: str):  
        self._clear()  
        frame = ctk.CTkScrollableFrame(  
            self.content_area,  
            fg_color="#0D0D0D",  
            scrollbar_button_color="#2A2A2A",  
            scrollbar_button_hover_color="#3A3A3A",  
        )  
        frame.grid(row=0, column=0, sticky="nsew", padx=30, pady=30)  
        frame.grid_columnconfigure(0, weight=1)  

        ctk.CTkLabel(  
            frame,  
            text=title,  
            font=ctk.CTkFont(size=26, weight="bold"),  
            text_color="#FFFFFF",  
            anchor="w",  
        ).grid(row=0, column=0, sticky="w", pady=(0, 20))  

        self.current_frame = frame  
        return frame  

    def _stat_row(self, parent, stats: list, row=1):  
        row_frame = ctk.CTkFrame(parent, fg_color="transparent")  
        row_frame.grid(row=row, column=0, sticky="ew", pady=(0, 20))  
        for i, (title, value, icon) in enumerate(stats):  
            card = ctk.CTkFrame(  
                row_frame, corner_radius=14,  
                fg_color="#161616", border_width=1, border_color="#2A2A2A",  
            )  
            card.grid(row=0, column=i, sticky="ew", padx=6, pady=4)  
            row_frame.grid_columnconfigure(i, weight=1)  
            ctk.CTkLabel(card, text=icon,  
                         font=ctk.CTkFont(size=26), text_color="#555555").pack(pady=(18, 4))  
            ctk.CTkLabel(card, text=title,  
                         font=ctk.CTkFont(size=11), text_color="#555555").pack()  
            ctk.CTkLabel(card, text=value,  
                         font=ctk.CTkFont(size=20, weight="bold"), text_color="#FFFFFF").pack(pady=(4, 18))  

    def _make_table(self, parent, columns: dict, rows: list, row=1):  
        frame = ctk.CTkFrame(  
            parent, fg_color="#111111",  
            corner_radius=12, border_width=1, border_color="#2A2A2A",  
        )  
        frame.grid(row=row, column=0, sticky="nsew", pady=10)  

        tree_frame = ctk.CTkFrame(frame, fg_color="transparent")  
        tree_frame.pack(fill="both", expand=True, padx=12, pady=12)  

        style = ttk.Style()  
        style.theme_use("clam")  
        style.configure("Modern.Treeview",  
            background="#111111", foreground="#CCCCCC",  
            rowheight=36, fieldbackground="#111111",  
            borderwidth=0, font=("Segoe UI", 11))  
        style.configure("Modern.Treeview.Heading",  
            background="#0D0D0D", foreground="#555555",  
            font=("Segoe UI", 11, "bold"), relief="flat", borderwidth=0)  
        style.map("Modern.Treeview",  
            background=[("selected", "#1E1E1E")],  
            foreground=[("selected", "#FFFFFF")])  

        col_keys = list(columns.keys())  
        tree = ttk.Treeview(  
            tree_frame, columns=col_keys,  
            show="headings", style="Modern.Treeview",  
        )  
        for col, width in columns.items():  
            tree.heading(col, text=col)  
            tree.column(col, width=width, anchor=tk.CENTER)  

        tree.tag_configure("odd",  background="#111111")  
        tree.tag_configure("even", background="#161616")  

        for i, row_data in enumerate(rows):  
            tag = "even" if i % 2 == 0 else "odd"  
            tree.insert("", tk.END, values=row_data, tags=(tag,))  

        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)  
        tree.configure(yscrollcommand=sb.set)  
        tree.pack(side="left", fill="both", expand=True)  
        sb.pack(side="right", fill="y")  
        return tree  

    def _make_rainbow_label(self, parent, text: str, font_size=12):  
        frame = ctk.CTkFrame(parent, fg_color="transparent")  
        frame.pack(side="left", padx=(0, 6))  
        idx = 0  
        for char in text:  
            if char == " ":  
                ctk.CTkLabel(frame, text=" ",  
                             font=ctk.CTkFont(size=font_size, weight="bold"),  
                             text_color="#FFFFFF", width=4).pack(side="left")  
            else:  
                ctk.CTkLabel(frame, text=char,  
                             font=ctk.CTkFont(size=font_size, weight="bold"),  
                             text_color=RAINBOW[idx % len(RAINBOW)],  
                             width=10).pack(side="left")  
                idx += 1  
        return frame  

    # ────────────────────────────────────────────────────────────  
    # WORK WEEK  
    # ────────────────────────────────────────────────────────────  
    def _get_work_week_info(self, dt: datetime) -> dict:  
        iso_year, iso_week, iso_weekday = dt.isocalendar()  
        monday = dt - timedelta(days=iso_weekday - 1)  
        friday = monday + timedelta(days=4)  
        return {  
            "week_num":    iso_week,  
            "monday":      monday.strftime("%d %b"),  
            "friday":      friday.strftime("%d %b %Y"),  
            "day_of_week": iso_weekday,  
        }  

    # ────────────────────────────────────────────────────────────  
    # LIVE CLOCK  
    # ────────────────────────────────────────────────────────────  
    def _tick_clock(self):  
        try:  
            if not self._clock_label.winfo_exists():  
                return  
        except Exception:  
            return  

        now  = datetime.now()  
        info = self._get_work_week_info(now)  

        self._clock_label.configure(text=now.strftime("%H : %M : %S"))  
        self._date_label.configure(text=now.strftime("%A,  %d  %B  %Y"))  
        self._week_label.configure(  
            text=f"Work Week  {info['week_num']}  |  {info['monday']} – {info['friday']}"  
        )  

        for dow, (cell, lbl) in self._day_labels.items():  
            if dow == info["day_of_week"]:  
                cell.configure(fg_color="#1E2A1E", border_width=1, border_color="#2ECC71")  
                lbl.configure(text_color="#2ECC71",  
                              font=ctk.CTkFont(size=11, weight="bold"))  
            else:  
                cell.configure(fg_color="#1A1A1A", border_width=0)  
                lbl.configure(text_color="#444444",  
                              font=ctk.CTkFont(size=11))  

        self._clock_job = self.root.after(1000, self._tick_clock)  

    # ════════════════════════════════════════════════════════════  
    # PAGE : DASHBOARD  
    # ════════════════════════════════════════════════════════════  
    def show_dashboard(self):  
        f = self._base_frame("News & Updates")  

        # ── DateTime / WorkWeek bar ───────────────────────────  
        dt_card = ctk.CTkFrame(  
            f, fg_color="#111111", corner_radius=14,  
            border_width=1, border_color="#2A2A2A",  
        )  
        dt_card.grid(row=1, column=0, sticky="ew", pady=(0, 12))  

        dt_inner = ctk.CTkFrame(dt_card, fg_color="transparent")  
        dt_inner.pack(fill="x", padx=20, pady=14)  

        self._clock_label = ctk.CTkLabel(  
            dt_inner, text="",  
            font=ctk.CTkFont(size=28, weight="bold"),  
            text_color="#FFFFFF",  
        )  
        self._clock_label.pack(side="left")  

        dt_right = ctk.CTkFrame(dt_inner, fg_color="transparent")  
        dt_right.pack(side="right", anchor="e")  

        self._date_label = ctk.CTkLabel(  
            dt_right, text="",  
            font=ctk.CTkFont(size=13),  
            text_color="#AAAAAA", anchor="e",  
        )  
        self._date_label.pack(anchor="e")  

        self._week_label = ctk.CTkLabel(  
            dt_right, text="",  
            font=ctk.CTkFont(size=12),  
            text_color="#555555", anchor="e",  
        )  
        self._week_label.pack(anchor="e", pady=(2, 0))  

        # Day bar  
        day_bar = ctk.CTkFrame(dt_card, fg_color="transparent")  
        day_bar.pack(fill="x", padx=20, pady=(0, 14))  

        self._day_labels = {}  
        for i, day in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):  
            cell = ctk.CTkFrame(  
                day_bar, corner_radius=8,  
                fg_color="#1A1A1A", width=62, height=36,  
            )  
            cell.grid(row=0, column=i, padx=4)  
            cell.grid_propagate(False)  
            lbl = ctk.CTkLabel(cell, text=day,  
                               font=ctk.CTkFont(size=11), text_color="#444444")  
            lbl.place(relx=0.5, rely=0.5, anchor="center")  
            self._day_labels[i + 1] = (cell, lbl)  

        self._tick_clock()  

        # ── Status + Online row ───────────────────────────────  
        top_row = ctk.CTkFrame(f, fg_color="transparent")  
        top_row.grid(row=2, column=0, sticky="ew", pady=(0, 16))  
        top_row.grid_columnconfigure(0, weight=3)  
        top_row.grid_columnconfigure(1, weight=1)  

        # Program Status  
        status_card = ctk.CTkFrame(  
            top_row, fg_color="#111111", corner_radius=14,  
            border_width=1, border_color="#2A2A2A",  
        )  
        status_card.grid(row=0, column=0, sticky="ew", padx=(0, 10))  

        ctk.CTkLabel(  
            status_card, text="Program Status",  
            font=ctk.CTkFont(size=12, weight="bold"),  
            text_color="#444444", anchor="w",  
        ).pack(anchor="w", padx=18, pady=(14, 6))  

        ctk.CTkFrame(status_card, height=1, fg_color="#1E1E1E").pack(fill="x", padx=18)  

        program_status = [  
            ("Core System",    "Running",   "Running"),  
            ("Database",       "Connected", "Connected"),  
            ("Import Service", "Running",   "Running"),  
            ("Report Engine",  "Running",   "Running"),  
            ("Backup Service", "Standby",   "Standby"),  
            ("Sync Service",   "Stopped",   "Stopped"),  
        ]  

        status_inner = ctk.CTkFrame(status_card, fg_color="transparent")  
        status_inner.pack(fill="x", padx=18, pady=10)  

        for i, (svc, state, key) in enumerate(program_status):  
            col_i = i % 3  
            row_i = i // 3  
            neon  = NEON_STATUS.get(key, NEON_STATUS["Stopped"])  
            c_on  = neon[0]  
            c_off = neon[1]  
            bg    = neon[2]  

            item = ctk.CTkFrame(status_inner, fg_color=bg, corner_radius=8)  
            item.grid(row=row_i, column=col_i, sticky="ew", padx=6, pady=5)  
            status_inner.grid_columnconfigure(col_i, weight=1)  

            dot = ctk.CTkLabel(item, text="●",  
                               font=ctk.CTkFont(size=11),  
                               text_color=c_on, width=18)  
            dot.pack(side="left", padx=(10, 4), pady=10)  

            ctk.CTkLabel(item, text=svc,  
                         font=ctk.CTkFont(size=12),  
                         text_color="#888888").pack(side="left", padx=(0, 8))  

            state_lbl = ctk.CTkLabel(item, text=state,  
                                     font=ctk.CTkFont(size=12, weight="bold"),  
                                     text_color=c_on)  
            state_lbl.pack(side="right", padx=(0, 12))  

            # Neon blink  
            if key in ("Running", "Connected"):  
                self._neon_blink(dot,       c_on, c_off, 1000, 220)  
                self._neon_blink(state_lbl, c_on, c_off, 1000, 220)  
            elif key == "Standby":  
                self._neon_blink(dot,       c_on, c_off, 1500, 600)  
                self._neon_blink(state_lbl, c_on, c_off, 1500, 600)  
            # Stopped ไม่ blink (สีคงที่)  

        ctk.CTkFrame(status_card, height=10, fg_color="transparent").pack()  

        # ── Online Users — Total Only ─────────────────────────  
        total_online = 4  

        online_card = ctk.CTkFrame(  
            top_row, fg_color="#111111", corner_radius=14,  
            border_width=1, border_color="#2A2A2A",  
        )  
        online_card.grid(row=0, column=1, sticky="nsew")  

        ctk.CTkLabel(  
            online_card, text="Online Users",  
            font=ctk.CTkFont(size=12, weight="bold"),  
            text_color="#444444", anchor="w",  
        ).pack(anchor="w", padx=18, pady=(14, 6))  

        ctk.CTkFrame(online_card, height=1, fg_color="#1E1E1E").pack(fill="x", padx=18)  

        # big neon number  
        count_lbl = ctk.CTkLabel(  
            online_card,  
            text=str(total_online),  
            font=ctk.CTkFont(size=54, weight="bold"),  
            text_color="#39FF14",  
        )  
        count_lbl.pack(pady=(22, 0))  
        self._neon_blink(count_lbl, "#39FF14", "#0D3300", 1200, 300)  

        ctk.CTkLabel(  
            online_card, text="users online",  
            font=ctk.CTkFont(size=12), text_color="#444444",  
        ).pack(pady=(4, 0))  

        ctk.CTkFrame(online_card, height=1, fg_color="#1E1E1E").pack(  
            fill="x", padx=18, pady=12,  
        )  

        # dot + System Active  
        dot_row = ctk.CTkFrame(online_card, fg_color="transparent")  
        dot_row.pack(pady=(0, 16))  

        dot_sys = ctk.CTkLabel(dot_row, text="●",  
                               font=ctk.CTkFont(size=10),  
                               text_color="#39FF14", width=16)  
        dot_sys.pack(side="left")  
        self._neon_blink(dot_sys, "#39FF14", "#0D3300", 1200, 300)  

        ctk.CTkLabel(dot_row, text="  System Active",  
                     font=ctk.CTkFont(size=12),  
                     text_color="#555555").pack(side="left")  

        # ── News Feed ─────────────────────────────────────────  
        ctk.CTkLabel(  
            f, text="Latest News",  
            font=ctk.CTkFont(size=14, weight="bold"),  
            text_color="#FFFFFF", anchor="w",  
        ).grid(row=3, column=0, sticky="w", pady=(8, 6))  

        news_items = [
            {
                "tag":          "UPDATE",
                "date":         "06 Apr 2026  •  08:30",
                "title":        "Import Module — รองรับไฟล์ขนาดใหญ่ได้แล้ว",
                "body":         (
                    "ตอนนี้คุณสามารถ Import ไฟล์ Excel ขนาดสูงสุด 200 MB ได้แล้ว "
                    "ระบบจะแบ่ง batch อัตโนมัติเพื่อป้องกัน timeout "
                    "พร้อม progress bar แสดงสถานะ realtime"
                ),
                "border_color": "#1E2A3A",
            },
            {
                "tag":          "UPDATE",
                "date":         "05 Apr 2026  •  14:00",
                "title":        "Report & Query — เพิ่ม Filter ตามแผนกแล้ว",
                "body":         (
                    "แต่ละ Role จะเห็นเฉพาะ Report ของแผนกตัวเอง "
                    "Admin สามารถดูได้ทุกแผนก รวมถึง Export PDF ได้ทันที"
                ),
                "border_color": "#1E2A3A",
            },
            {
                "tag":          "NEW",
                "date":         "04 Apr 2026  •  10:15",
                "title":        "Inventory Module — ระบบแจ้งเตือน Low Stock",
                "body":         (
                    "เพิ่มระบบแจ้งเตือนอัตโนมัติเมื่อสินค้าใกล้หมด "
                    "สามารถกำหนด threshold ได้ในหน้า Settings"
                ),
                "border_color": "#1E3A1E",
            },
            {
                "tag":          "FIX",
                "date":         "03 Apr 2026  •  09:00",
                "title":        "แก้ไข Bug — Login หมดเวลา Session ไม่ Logout",
                "body":         (
                    "แก้ปัญหา Session ที่ค้างอยู่หลัง timeout "
                    "ตอนนี้ระบบจะ Logout อัตโนมัติและล้าง token ทันที"
                ),
                "border_color": "#3A1E2A",
            },
            {
                "tag":          "INFO",
                "date":         "01 Apr 2026  •  08:00",
                "title":        "ประกาศ — Maintenance วันที่ 08 เมษายน 2026",
                "body":         (
                    "ระบบจะหยุดให้บริการชั่วคราวในวันที่ 08 เม.ย. เวลา 02:00–04:00 น. "
                    "เพื่ออัปเดต Database และ Server Configuration"
                ),
                "border_color": "#3A3A1E",
            },
        ]

        for idx, news in enumerate(news_items):
            tag   = news["tag"]
            style = TAG_STYLE.get(tag, TAG_STYLE["INFO"])

            card = ctk.CTkFrame(
                f,
                fg_color="#111111",
                corner_radius=14,
                border_width=1,
                border_color=news["border_color"],
            )
            card.grid(row=idx + 4, column=0, sticky="ew", pady=6)

            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="x", padx=20, pady=14)

            # meta row
            meta_row = ctk.CTkFrame(inner, fg_color="transparent")
            meta_row.pack(fill="x", pady=(0, 6))

            ctk.CTkLabel(
                meta_row,
                text=f"  {tag}  ",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color=style["bg"],
                text_color=style["fg"],
                corner_radius=6,
            ).pack(side="left", padx=(0, 8))

            if tag in ("UPDATE", "NEW"):
                rainbow_text = "Updated ! ! !" if tag == "UPDATE" else "New Feature ! ! !"
                self._make_rainbow_label(meta_row, rainbow_text, font_size=12)

            ctk.CTkLabel(
                meta_row,
                text=news["date"],
                font=ctk.CTkFont(size=11),
                text_color="#333333",
            ).pack(side="right")

            ctk.CTkLabel(
                inner,
                text=news["title"],
                font=ctk.CTkFont(size=14, weight="bold"),
                text_color="#FFFFFF",
                anchor="w",
                wraplength=700,
            ).pack(anchor="w", pady=(2, 4))

            ctk.CTkFrame(inner, height=1, fg_color="#1E1E1E").pack(fill="x", pady=4)

            ctk.CTkLabel(
                inner,
                text=news["body"],
                font=ctk.CTkFont(size=12),
                text_color="#555555",
                anchor="w",
                justify="left",
                wraplength=700,
            ).pack(anchor="w", pady=(2, 0))

    # ──────────────────────────────────────────────────────────────
    # LIVE CLOCK TICKER
    # ──────────────────────────────────────────────────────────────
    def _tick_clock(self):
        """อัปเดต clock ทุก 1 วินาที"""
        # ตรวจว่า label ยังอยู่
        try:
            self._clock_label.winfo_exists()
        except Exception:
            return

        if not self._clock_label.winfo_exists():
            return

        now  = datetime.now()
        info = self._get_work_week_info(now)

        # clock
        self._clock_label.configure(text=now.strftime("%H : %M : %S"))

        # date
        self._date_label.configure(
            text=now.strftime("%A,  %d  %B  %Y")
        )

        # work week
        self._week_label.configure(
            text=(
                f"Work Week  {info['week_num']}  |  "
                f"{info['monday']} – {info['friday']}"
            )
        )

        # highlight วันปัจจุบันใน day bar
        for dow, (cell, lbl) in self._day_labels.items():
            if dow == info["day_of_week"]:
                cell.configure(fg_color="#1E2A1E", border_width=1, border_color="#2ECC71")
                lbl.configure(text_color="#2ECC71", font=ctk.CTkFont(size=11, weight="bold"))
            else:
                cell.configure(fg_color="#1A1A1A", border_width=0)
                lbl.configure(text_color="#444444", font=ctk.CTkFont(size=11))

        self._clock_job = self.root.after(1000, self._tick_clock)

    # ──────────────────────────────────────────────────────────────
    # PAGE: IMPORT EXCEL
    # ──────────────────────────────────────────────────────────────
    def show_import(self):
        f = self._base_frame("Import Excel")

        drop_zone = ctk.CTkFrame(
            f,
            corner_radius=16,
            height=220,
            fg_color="#111111",
            border_width=1,
            border_color="#2A2A2A",
        )
        drop_zone.grid(row=1, column=0, sticky="ew", pady=20)
        drop_zone.grid_propagate(False)

        ctk.CTkLabel(
            drop_zone, text="↓",
            font=ctk.CTkFont(size=36), text_color="#333333",
        ).place(relx=0.5, rely=0.28, anchor="center")

        ctk.CTkLabel(
            drop_zone,
            text="Drag & Drop Excel file here",
            font=ctk.CTkFont(size=15),
            text_color="#444444",
        ).place(relx=0.5, rely=0.48, anchor="center")

        ctk.CTkButton(
            drop_zone,
            text="Browse Files",
            width=160,
            height=38,
            corner_radius=8,
            fg_color="#1E1E1E",
            hover_color="#2A2A2A",
            border_width=1,
            border_color="#333333",
            text_color="#FFFFFF",
            font=ctk.CTkFont(size=13),
        ).place(relx=0.5, rely=0.74, anchor="center")

        ctk.CTkLabel(
            f,
            text="Supported: .xlsx  .xls          Max size: 50 MB",
            font=ctk.CTkFont(size=12),
            text_color="#333333",
        ).grid(row=2, column=0, pady=8)

    # ──────────────────────────────────────────────────────────────
    # PAGE: REPORT & QUERY
    # ──────────────────────────────────────────────────────────────
    def show_export(self):
        f = self._base_frame("Report & Query")

        role     = self.user_data[1].upper()
        dept_map = DEPT_REPORTS.get(role, DEPT_REPORTS["USER"])
        dept_names = list(dept_map.keys())

        # ── Department selector ───────────────────────────────
        select_frame = ctk.CTkFrame(
            f, fg_color="#111111", corner_radius=14,
            border_width=1, border_color="#2A2A2A",
        )
        select_frame.grid(row=1, column=0, sticky="ew", pady=(0, 12))

        ctk.CTkLabel(
            select_frame, text="Select Department",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#555555", anchor="w",
        ).pack(anchor="w", padx=20, pady=(16, 8))

        ctk.CTkFrame(select_frame, height=1, fg_color="#2A2A2A").pack(fill="x", padx=20)

        dept_btn_frame = ctk.CTkFrame(select_frame, fg_color="transparent")
        dept_btn_frame.pack(fill="x", padx=16, pady=14)

        # ── Report list ───────────────────────────────────────
        report_frame = ctk.CTkFrame(
            f, fg_color="#111111", corner_radius=14,
            border_width=1, border_color="#2A2A2A",
        )
        report_frame.grid(row=2, column=0, sticky="ew", pady=(0, 12))

        report_header = ctk.CTkLabel(
            report_frame, text="Available Reports",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#555555", anchor="w",
        )
        report_header.pack(anchor="w", padx=20, pady=(16, 8))

        ctk.CTkFrame(report_frame, height=1, fg_color="#2A2A2A").pack(fill="x", padx=20)

        report_list_frame = ctk.CTkFrame(report_frame, fg_color="transparent")
        report_list_frame.pack(fill="x", padx=16, pady=10)

        export_btn = ctk.CTkButton(
            f,
            text="↓   Export as PDF",
            width=200, height=44, corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#1A3A2A", hover_color="#1E4A34",
            border_width=1, border_color="#2ECC71",
            text_color="#2ECC71",
        )
        export_btn.grid(row=3, column=0, pady=16)

        self._report_vars  = {}
        self._active_dept  = ctk.StringVar(value="")
        self._dept_buttons = {}

        def render_reports(dept_name):
            for w in report_list_frame.winfo_children():
                w.destroy()
            self._report_vars.clear()
            reports = dept_map.get(dept_name, [])
            if not reports:
                ctk.CTkLabel(
                    report_list_frame,
                    text="No reports available",
                    font=ctk.CTkFont(size=12),
                    text_color="#444444",
                ).pack(anchor="w", padx=4, pady=12)
                return
            for rname in reports:
                row_f = ctk.CTkFrame(report_list_frame, fg_color="transparent")
                row_f.pack(fill="x", pady=6)
                var = ctk.BooleanVar()
                self._report_vars[rname] = var
                ctk.CTkCheckBox(
                    row_f, text=rname, variable=var,
                    font=ctk.CTkFont(size=13), text_color="#CCCCCC",
                    checkbox_width=20, checkbox_height=20, corner_radius=4,
                    fg_color="#1E1E1E", hover_color="#2A2A2A",
                    border_color="#333333", checkmark_color="#FFFFFF",
                ).pack(side="left", padx=4)

        def on_dept_click(dept_name):
            self._active_dept.set(dept_name)
            for d, b in self._dept_buttons.items():
                b.configure(fg_color="#1A1A1A", border_color="#2A2A2A", text_color="#666666")
            self._dept_buttons[dept_name].configure(
                fg_color="#1E2A1E", border_color="#2ECC71", text_color="#2ECC71"
            )
            report_header.configure(text=f"Available Reports  —  {dept_name}")
            render_reports(dept_name)

        for dept in dept_names:
            btn = ctk.CTkButton(
                dept_btn_frame, text=dept,
                width=130, height=36, corner_radius=8,
                font=ctk.CTkFont(size=12),
                fg_color="#1A1A1A", hover_color="#1E2A1E",
                border_width=1, border_color="#2A2A2A",
                text_color="#666666",
                command=lambda d=dept: on_dept_click(d),
            )
            btn.pack(side="left", padx=6, pady=4)
            self._dept_buttons[dept] = btn

        if dept_names:
            on_dept_click(dept_names[0])

        def do_export():
            selected = [n for n, v in self._report_vars.items() if v.get()]
            if not selected:
                messagebox.showwarning("No Selection", "Please select at least one report")
                return
            dept = self._active_dept.get()
            report_list = "\n  •  ".join(selected)
            messagebox.showinfo(
                "Export",
                f"Exporting {len(selected)} report(s)\nDepartment: {dept}\n\n  •  {report_list}",
            )

        export_btn.configure(command=do_export)

    # ──────────────────────────────────────────────────────────────
    # PAGE: INVENTORY
    # ──────────────────────────────────────────────────────────────
    def show_inventory(self):
        f = self._base_frame("Inventory")

        self._stat_row(f, [
            ("Total Items",    "3,240", "▤"),
            ("Low Stock",      "12",    "▲"),
            ("Out of Stock",   "5",     "✕"),
            ("Reorder Needed", "8",     "↺"),
        ], row=1)

        self._make_table(f,
            columns={"ID": 60, "Item Name": 200, "Category": 150,
                     "Stock": 80, "Unit": 80, "Status": 110},
            rows=[
                (1, "Product A", "Electronics", 150, "pcs", "In Stock"),
                (2, "Product B", "Furniture",    8,  "pcs", "Low Stock"),
                (3, "Product C", "Clothing",     0,  "pcs", "Out of Stock"),
                (4, "Product D", "Electronics", 320, "pcs", "In Stock"),
                (5, "Product E", "Stationery",   45, "box", "In Stock"),
            ], row=2,
        )


    
    # ──────────────────────────────────────────────────────────────
    # PAGE: FINANCE
    # ──────────────────────────────────────────────────────────────
    def show_finance(self):
        if self.user_data[1].upper() not in ["ADMIN", "FINANCE"]:
            messagebox.showerror("Access Denied", "ไม่มีสิทธิ์เข้าถึง")
            return

        f = self._base_frame("Finance")

        self._stat_row(f, [
            ("Revenue",  "฿2,500,000", "▲"),
            ("Expenses", "฿980,000",   "▼"),
            ("Profit",   "฿1,520,000", "◈"),
        ], row=1)

        ctk.CTkLabel(
            f, text="Recent Transactions",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#FFFFFF", anchor="w",
        ).grid(row=2, column=0, sticky="w", pady=(10, 4))

        self._make_table(f,
            columns={"Date": 120, "Description": 250, "Type": 100,
                     "Amount": 120, "Balance": 130},
            rows=[
                ("2026-04-01", "Sales Revenue Q1", "Income",  "+฿500,000", "฿2,500,000"),
                ("2026-04-02", "Office Supplies",   "Expense", "-฿15,000",  "฿2,485,000"),
                ("2026-04-03", "Software License",  "Expense", "-฿50,000",  "฿2,435,000"),
                ("2026-04-04", "Client Payment",    "Income",  "+฿200,000", "฿2,635,000"),
                ("2026-04-05", "Staff Salary",      "Expense", "-฿300,000", "฿2,335,000"),
            ], row=3,
        )

    # ──────────────────────────────────────────────────────────────
    # PAGE: USER MANAGEMENT  (under Settings)
    # ──────────────────────────────────────────────────────────────
    def show_user_management(self):
        if self.user_data[1].upper() != "ADMIN":
            messagebox.showerror("Access Denied", "Admins only")
            return

        f = self._base_frame("Settings  —  User Management")

        btn_f = ctk.CTkFrame(f, fg_color="transparent")
        btn_f.grid(row=1, column=0, sticky="w", pady=(0, 15))

        ctk.CTkButton(
            btn_f, text="+ Add User", width=130, height=38,
            corner_radius=8, fg_color="#1A3A2A", hover_color="#1E4A34",
            border_width=1, border_color="#2ECC71", text_color="#2ECC71",
            font=ctk.CTkFont(size=13),
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_f, text="- Delete", width=130, height=38,
            corner_radius=8, fg_color="#3A1A1A", hover_color="#4A1E1E",
            border_width=1, border_color="#E74C3C", text_color="#E74C3C",
            font=ctk.CTkFont(size=13),
        ).pack(side="left")

        self._make_table(f,
            columns={"ID": 50, "Username": 130, "Full Name": 160,
                     "Role": 100, "Status": 80, "Last Login": 160},
            rows=[
                (1, "admin",    "Administrator", "ADMIN",   "Active",   "2026-04-06 08:00"),
                (2, "john",     "John Doe",      "USER",    "Active",   "2026-04-06 09:15"),
                (3, "finance1", "Jane Smith",    "FINANCE", "Active",   "2026-04-05 14:30"),
                (4, "viewer",   "Bob Wilson",    "USER",    "Inactive", "2026-03-28 11:00"),
            ], row=2,
        )

    # ──────────────────────────────────────────────────────────────
    # PAGE: SCHEDULE  (under Settings)
    # ──────────────────────────────────────────────────────────────
    def show_schedule(self):
        if self.user_data[1].upper() != "ADMIN":
            messagebox.showerror("Access Denied", "Admins only")
            return

        f = self._base_frame("Settings  —  Schedule")

        now = datetime.now().strftime("%A, %d %B %Y")
        ctk.CTkLabel(
            f, text=f"Today  —  {now}",
            font=ctk.CTkFont(size=13),
            text_color="#444444", anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(0, 15))

        self._make_table(f,
            columns={"Date": 120, "Time": 80, "Event": 200,
                     "Assigned To": 150, "Status": 100},
            rows=[
                ("2026-04-06", "09:00", "Team Meeting",       "All Staff",  "Pending"),
                ("2026-04-07", "10:00", "Q1 Report Deadline", "Finance",    "Pending"),
                ("2026-04-08", "14:00", "System Maintenance", "IT Admin",   "Scheduled"),
                ("2026-04-10", "09:00", "Monthly Review",     "Management", "Scheduled"),
                ("2026-04-15", "11:00", "Inventory Audit",    "Warehouse",  "Scheduled"),
            ], row=2,
        )

    # ──────────────────────────────────────────────────────────────
    # PAGE: AUDIT LOG  (under Settings)
    # ──────────────────────────────────────────────────────────────
    def show_audit_log(self):
        if self.user_data[1].upper() != "ADMIN":
            messagebox.showerror("Access Denied", "Admins only")
            return

        f = self._base_frame("Settings  —  Audit Log")

        filter_f = ctk.CTkFrame(f, fg_color="transparent")
        filter_f.grid(row=1, column=0, sticky="w", pady=(0, 15))

        ctk.CTkLabel(
            filter_f, text="Filter by User:",
            font=ctk.CTkFont(size=13), text_color="#666666",
        ).pack(side="left", padx=(0, 8))

        ctk.CTkEntry(
            filter_f, width=200, height=38, corner_radius=8,
            fg_color="#111111", border_color="#2A2A2A",
            text_color="#FFFFFF",
            placeholder_text="Search user...",
            placeholder_text_color="#444444",
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            filter_f, text="Search", width=100, height=38,
            corner_radius=8, fg_color="#1E1E1E", hover_color="#2A2A2A",
            border_width=1, border_color="#333333",
            text_color="#FFFFFF", font=ctk.CTkFont(size=13),
        ).pack(side="left")

        self._make_table(f,
            columns={"Timestamp": 160, "User": 100, "Action": 110,
                     "Module": 120, "Detail": 280},
            rows=[
                ("2026-04-06 08:00", "admin",    "LOGIN",       "Auth",      "Logged in successfully"),
                ("2026-04-06 08:05", "john",     "IMPORT",      "Excel",     "Imported Q1_Sales.xlsx"),
                ("2026-04-06 08:30", "finance1", "EXPORT",      "Report",    "Exported Monthly Report"),
                ("2026-04-06 09:00", "admin",    "USER_UPDATE", "User Mgmt", "Updated user: john"),
                ("2026-04-06 09:30", "john",     "VIEW",        "Inventory", "Viewed inventory list"),
                ("2026-04-06 10:00", "admin",    "SETTINGS",    "Settings",  "Changed DB config"),
            ], row=2,
        )

    # ──────────────────────────────────────────────────────────────
    # PAGE: SETTINGS
    # ──────────────────────────────────────────────────────────────
    def show_settings(self):
        if self.user_data[1].lower() != "admin":
            messagebox.showerror("Access Denied", "Admins only")
            return

        f = self._base_frame("Settings")

        sections = [
            ("▤   General Settings",       self.show_settings),
            ("◉   User Management",        self.show_user_management),
            ("▦   Schedule",               self.show_schedule),
            ("≡   Audit Log",              self.show_audit_log),
            ("↓   Import Excel",           self.show_import),
        ]

        for i, (name, callback) in enumerate(sections):
            btn = ctk.CTkButton(
                f, text=name,
                font=ctk.CTkFont(size=13), anchor="w",
                height=54, corner_radius=12,
                fg_color="#111111", hover_color="#161616",
                border_width=1, border_color="#2A2A2A",
                text_color="#AAAAAA",
                command=callback,
            )
            btn.grid(row=i + 1, column=0, sticky="ew", pady=6, padx=4)

    # ──────────────────────────────────────────────────────────────
    # LOGOUT
    # ──────────────────────────────────────────────────────────────
    def logout(self):
        if messagebox.askyesno("Logout", "Are you sure you want to logout?"):
            # หยุด clock
            if self._clock_job:
                self.root.after_cancel(self._clock_job)
                self._clock_job = None
            # ล้าง widgets ทั้งหมด
            for widget in self.root.winfo_children():
                widget.destroy()
            # กลับไปหน้า Login
            self.logout_callback()