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
        "Purchase":  ["Purchase Order Summary", "Vendor Performance Report", "Purchase vs Budget"],
        "Inventory": ["Inventory Stock Report", "Low Stock Alert Report", "Stock Movement History"],
        "HR":        ["HR Headcount Report", "Attendance Summary", "Payroll Summary"],
        "Finance":   ["Finance P&L Statement", "Monthly Sales Report", "Quarterly Summary", "Annual Financial Report"],
        "Sales":     ["Monthly Sales Report", "Quarterly Summary", "Customer Report"],
    },
    "FINANCE":   {"Finance":   ["Finance P&L Statement", "Monthly Sales Report", "Quarterly Summary", "Annual Financial Report"]},
    "PURCHASE":  {"Purchase":  ["Purchase Order Summary", "Vendor Performance Report", "Purchase vs Budget"]},
    "HR":        {"HR":        ["HR Headcount Report", "Attendance Summary", "Payroll Summary"]},
    "SALES":     {"Sales":     ["Monthly Sales Report", "Quarterly Summary", "Customer Report"]},
    "INVENTORY": {"Inventory": ["Inventory Stock Report", "Low Stock Alert Report", "Stock Movement History"]},
    "USER":      {"General":   ["Monthly Sales Report", "User Activity Log"]},
}

RAINBOW = ["#FF6B6B", "#FF9F43", "#FECA57", "#48DBFB", "#FF9FF3", "#54A0FF", "#5F27CD"]

TAG_STYLE = {
    "UPDATE": {"bg": "#1A2A3A", "fg": "#3498DB"},
    "NEW":    {"bg": "#1A3A1A", "fg": "#2ECC71"},
    "FIX":    {"bg": "#3A1A1A", "fg": "#E74C3C"},
    "INFO":   {"bg": "#2A2A1A", "fg": "#F39C12"},
}

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

    def build_layout(self):
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_content_area()

    def _build_sidebar(self):
        self.sidebar = ctk.CTkFrame(
            self.root, width=240, corner_radius=0, fg_color="#111111",
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        logo_frame = ctk.CTkFrame(
            self.sidebar, fg_color="#1A1A1A", corner_radius=0, height=70,
        )
        logo_frame.pack(fill="x")
        logo_frame.pack_propagate(False)
        ctk.CTkLabel(
            logo_frame, text=ICONS["logo"],
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color="#FFFFFF",
        ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkFrame(self.sidebar, height=1, fg_color="#2A2A2A").pack(fill="x")

        user_box = ctk.CTkFrame(self.sidebar, fg_color="#1A1A1A", corner_radius=10)
        user_box.pack(fill="x", padx=12, pady=12)
        ctk.CTkLabel(
            user_box, text=f"  {self.user_data[0]}",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#FFFFFF", anchor="w",
        ).pack(fill="x", padx=10, pady=(10, 2))
        ctk.CTkLabel(
            user_box, text=f"  {self.user_data[1].upper()}",
            font=ctk.CTkFont(size=11),
            text_color="#666666", anchor="w",
        ).pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkFrame(self.sidebar, height=1, fg_color="#2A2A2A").pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            self.sidebar, text="  NAVIGATION",
            font=ctk.CTkFont(size=10),
            text_color="#444444", anchor="w",
        ).pack(fill="x", padx=15, pady=(4, 6))

        menu_scroll = ctk.CTkScrollableFrame(
            self.sidebar, fg_color="transparent",
            scrollbar_button_color="#2A2A2A",
            scrollbar_button_hover_color="#3A3A3A",
        )
        menu_scroll.pack(fill="both", expand=True, padx=8)

        role     = self.user_data[1].upper()
        is_admin = role == "ADMIN"

        menu_items = [
            ("▣   News & Updates",    self.show_dashboard,      True,                         False),
            ("↗   Report & Query",    self.show_export,          True,                         False),
            ("▤   Inventory",         self.show_inventory,       True,                         False),
            ("💻  IT Equipment",      self.show_it_equipment,    True,                         False),
            ("◈   Finance",           self.show_finance,         role in ["ADMIN", "FINANCE"], False),
            ("✦   Settings",          self.show_settings,        is_admin,                     False),
            ("    ├ User Management", self.show_user_management, is_admin,                     True),
            ("    ├ Schedule",        self.show_schedule,        is_admin,                     True),
            ("    ├ Audit Log",       self.show_audit_log,       is_admin,                     True),
            ("    └ Import Excel",    self.show_import,          is_admin,                     True),
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

        ctk.CTkFrame(self.sidebar, height=1, fg_color="#2A2A2A").pack(fill="x", padx=12, pady=8)

        ctk.CTkButton(
            self.sidebar,
            text=f"  {ICONS['logout']}    Logout",
            font=ctk.CTkFont(size=13), anchor="w",
            height=42, corner_radius=8,
            fg_color="transparent", text_color="#E74C3C",
            hover_color="#2A1A1A", command=self.logout,
        ).pack(fill="x", padx=8, pady=(0, 12))

    def _nav(self, callback, label):
        self._cancel_neon()
        if self._clock_job:
            self.root.after_cancel(self._clock_job)
            self._clock_job = None
        for btn in self.menu_buttons.values():
            btn.configure(fg_color="transparent", text_color="#AAAAAA")
        if label in self.menu_buttons:
            self.menu_buttons[label].configure(fg_color="#1E1E1E", text_color="#FFFFFF")
        callback()

    def _build_content_area(self):
        self.content_area = ctk.CTkFrame(
            self.root, corner_radius=0, fg_color="#0D0D0D",
        )
        self.content_area.grid(row=0, column=1, sticky="nsew")
        self.content_area.grid_columnconfigure(0, weight=1)
        self.content_area.grid_rowconfigure(0, weight=1)
        self._draw_noise_overlay()

    def _draw_noise_overlay(self):
        canvas = tk.Canvas(self.content_area, bg="#0D0D0D", highlightthickness=0)
        canvas.place(x=0, y=0, relwidth=1, relheight=1)
        for _ in range(4000):
            x = random.randint(0, 1920)
            y = random.randint(0, 1080)
            g = random.randint(18, 32)
            c = f"#{g:02x}{g:02x}{g:02x}"
            canvas.create_rectangle(x, y, x + 1, y + 1, fill=c, outline="")
        canvas.lower()

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
            frame, text=title,
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color="#FFFFFF", anchor="w",
        ).grid(row=0, column=0, sticky="w", pady=(0, 20))
        self.current_frame = frame
        return frame

    def _stat_row(self, parent, stats: list, row=1):
        row_frame = ctk.CTkFrame(parent, fg_color="transparent")
        row_frame.grid(row=row, column=0, sticky="ew", pady=(0, 20))
        for i, (title, value, icon) in enumerate(stats):
            card = ctk.CTkFrame(
                row_frame, corner_radius=14, fg_color="#161616",
                border_width=1, border_color="#2A2A2A",
            )
            card.grid(row=0, column=i, sticky="ew", padx=6, pady=4)
            row_frame.grid_columnconfigure(i, weight=1)
            ctk.CTkLabel(card, text=icon, font=ctk.CTkFont(size=26),
                         text_color="#555555").pack(pady=(18, 4))
            ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=11),
                         text_color="#555555").pack()
            ctk.CTkLabel(card, text=value, font=ctk.CTkFont(size=20, weight="bold"),
                         text_color="#FFFFFF").pack(pady=(4, 18))

    def _make_table(self, parent, columns: dict, rows: list, row=1):
        frame = ctk.CTkFrame(
            parent, fg_color="#111111", corner_radius=12,
            border_width=1, border_color="#2A2A2A",
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
                lbl.configure(text_color="#2ECC71", font=ctk.CTkFont(size=11, weight="bold"))
            else:
                cell.configure(fg_color="#1A1A1A", border_width=0)
                lbl.configure(text_color="#444444", font=ctk.CTkFont(size=11))
        self._clock_job = self.root.after(1000, self._tick_clock)

    # ════════════════════════════════════════════════════════
    # PAGE: DASHBOARD
    # ════════════════════════════════════════════════════════
    def show_dashboard(self):
        f = self._base_frame("News & Updates")

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

        top_row = ctk.CTkFrame(f, fg_color="transparent")
        top_row.grid(row=2, column=0, sticky="ew", pady=(0, 16))
        top_row.grid_columnconfigure(0, weight=3)
        top_row.grid_columnconfigure(1, weight=1)

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

            if key in ("Running", "Connected"):
                self._neon_blink(dot,       c_on, c_off, 1000, 220)
                self._neon_blink(state_lbl, c_on, c_off, 1000, 220)
            elif key == "Standby":
                self._neon_blink(dot,       c_on, c_off, 1500, 600)
                self._neon_blink(state_lbl, c_on, c_off, 1500, 600)

        ctk.CTkFrame(status_card, height=10, fg_color="transparent").pack()

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

        count_lbl = ctk.CTkLabel(
            online_card, text="4",
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

        ctk.CTkLabel(
            f, text="Latest News",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#FFFFFF", anchor="w",
        ).grid(row=3, column=0, sticky="w", pady=(8, 6))

        news_items = [
            {
                "tag":   "UPDATE",
                "date":  "06 Apr 2026  •  08:30",
                "title": "Import Module — รองรับไฟล์ขนาดใหญ่ได้แล้ว",
                "body":  "ตอนนี้คุณสามารถ Import ไฟล์ Excel ขนาดสูงสุด 200 MB ได้แล้ว "
                         "ระบบจะแบ่ง batch อัตโนมัติเพื่อป้องกัน timeout",
            },
            {
                "tag":   "UPDATE",
                "date":  "05 Apr 2026  •  14:00",
                "title": "Report & Query — เพิ่ม Filter ตามแผนกแล้ว",
                "body":  "แต่ละ Role จะเห็นเฉพาะ Report ของแผนกตัวเอง "
                         "Admin สามารถดูได้ทุกแผนก รวมถึง Export PDF ได้ทันที",
            },
            {
                "tag":   "NEW",
                "date":  "04 Apr 2026  •  10:15",
                "title": "IT Equipment Module — เพิ่มระบบจัดการครุภัณฑ์ IT",
                "body":  "เพิ่มหน้า IT Equipment สำหรับจัดการครุภัณฑ์ IT ทั้งหมด "
                         "รองรับการ Add / Edit / Delete และแสดงสถานะอุปกรณ์",
            },
            {
                "tag":   "NEW",
                "date":  "03 Apr 2026  •  09:00",
                "title": "Inventory — ระบบแจ้งเตือน Low Stock",
                "body":  "เพิ่มระบบแจ้งเตือนอัตโนมัติเมื่อสินค้าใกล้หมด "
                         "สามารถกำหนด threshold ได้ในหน้า Settings",
            },
            {
                "tag":   "FIX",
                "date":  "02 Apr 2026  •  16:45",
                "title": "Fix — แก้ไขปัญหา Login timeout",
                "body":  "แก้ไขปัญหาที่ Session หมดอายุเร็วกว่ากำหนด "
                         "ตอนนี้ระบบจะ refresh token อัตโนมัติ",
            },
        ]

        for idx, news in enumerate(news_items):
            tag_key  = news["tag"]
            ts       = TAG_STYLE.get(tag_key, TAG_STYLE["INFO"])

            card = ctk.CTkFrame(
                f, fg_color="#111111", corner_radius=12,
                border_width=1, border_color="#2A2A2A",
            )
            card.grid(row=4 + idx, column=0, sticky="ew", pady=5)

            top = ctk.CTkFrame(card, fg_color="transparent")
            top.pack(fill="x", padx=18, pady=(14, 6))

            tag_badge = ctk.CTkFrame(
                top, fg_color=ts["bg"], corner_radius=6,
            )
            tag_badge.pack(side="left")
            ctk.CTkLabel(
                tag_badge, text=f"  {tag_key}  ",
                font=ctk.CTkFont(size=10, weight="bold"),
                text_color=ts["fg"],
            ).pack(padx=2, pady=3)

            ctk.CTkLabel(
                top, text=news["date"],
                font=ctk.CTkFont(size=11),
                text_color="#444444",
            ).pack(side="right")

            ctk.CTkLabel(
                card, text=news["title"],
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color="#FFFFFF", anchor="w",
            ).pack(anchor="w", padx=18, pady=(0, 4))

            ctk.CTkLabel(
                card, text=news["body"],
                font=ctk.CTkFont(size=12),
                text_color="#666666", anchor="w",
                wraplength=800, justify="left",
            ).pack(anchor="w", padx=18, pady=(0, 14))

    # ════════════════════════════════════════════════════════
    # PAGE: REPORT & QUERY
    # ════════════════════════════════════════════════════════
    def show_export(self):
        f = self._base_frame("Report & Query")

        role       = self.user_data[1].upper()
        dept_map   = DEPT_REPORTS.get(role, DEPT_REPORTS["USER"])
        dept_names = list(dept_map.keys())

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
            f, text="↓   Export as PDF",
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

    # ════════════════════════════════════════════════════════
    # PAGE: INVENTORY
    # ════════════════════════════════════════════════════════
    def show_inventory(self):
        f = self._base_frame("Inventory")

        self._stat_row(f, [
            ("Total Items",    "3,240", "▤"),
            ("Low Stock",      "12",    "▲"),
            ("Out of Stock",   "5",     "✕"),
            ("Reorder Needed", "8",     "↺"),
        ], row=1)

        self._make_table(f,
            columns={
                "ID": 60, "Item Name": 200, "Category": 150,
                "Stock": 80, "Unit": 80, "Status": 110,
            },
            rows=[
                (1, "Product A", "Electronics", 150, "pcs", "In Stock"),
                (2, "Product B", "Furniture",     8, "pcs", "Low Stock"),
                (3, "Product C", "Clothing",       0, "pcs", "Out of Stock"),
                (4, "Product D", "Electronics",  320, "pcs", "In Stock"),
                (5, "Product E", "Stationery",    45, "box", "In Stock"),
            ], row=2,
        )

    # ════════════════════════════════════════════════════════
    # PAGE: IT EQUIPMENT
    # ════════════════════════════════════════════════════════
    def show_it_equipment(self):
        self._clear()
        frame = ctk.CTkFrame(
            self.content_area, corner_radius=0, fg_color="#0D0D0D",
        )
        frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        self.current_frame = frame

        try:
            from EquipmentIT import EquipmentITWindow
            EquipmentITWindow(frame, self.user_data)
        except ImportError as e:
            ctk.CTkLabel(
                frame,
                text="⚠  ไม่พบไฟล์ EquipmentIT.py",
                font=ctk.CTkFont(size=20, weight="bold"),
                text_color="#E74C3C",
            ).place(relx=0.5, rely=0.45, anchor="center")
            ctk.CTkLabel(
                frame, text=str(e),
                font=ctk.CTkFont(size=13),
                text_color="#555555",
            ).place(relx=0.5, rely=0.52, anchor="center")

    # ════════════════════════════════════════════════════════
    # PAGE: FINANCE
    # ════════════════════════════════════════════════════════
    def show_finance(self):
        f = self._base_frame("Finance")

        self._stat_row(f, [
            ("Revenue",  "฿ 4,200,000", "◈"),
            ("Expense",  "฿ 2,800,000", "↓"),
            ("Profit",   "฿ 1,400,000", "↑"),
            ("Pending",  "฿   320,000", "⏳"),
        ], row=1)

        self._make_table(f,
            columns={
                "ID": 60, "Date": 130, "Description": 220,
                "Type": 100, "Amount": 120, "Status": 110,
            },
            rows=[
                (1, "2026-04-01", "Sales Revenue Q1",    "Income",  "฿1,200,000", "Approved"),
                (2, "2026-04-02", "Office Supplies",      "Expense", "฿  15,000",  "Approved"),
                (3, "2026-04-03", "Server Maintenance",   "Expense", "฿  80,000",  "Pending"),
                (4, "2026-04-04", "Product Sales",        "Income",  "฿  450,000", "Approved"),
                (5, "2026-04-05", "Staff Training",       "Expense", "฿  35,000",  "Pending"),
            ], row=2,
        )

    # ════════════════════════════════════════════════════════
    # PAGE: SETTINGS
    # ════════════════════════════════════════════════════════
    def show_settings(self):
        f = self._base_frame("Settings")

        settings_groups = [
            ("System", [
                ("System Name",    "ERP System v2.0"),
                ("Language",       "Thai / English"),
                ("Timezone",       "Asia/Bangkok (UTC+7)"),
                ("Date Format",    "DD/MM/YYYY"),
            ]),
            ("Security", [
                ("Session Timeout", "30 minutes"),
                ("Max Login Attempts", "5"),
                ("Password Policy", "Min 8 chars, 1 uppercase, 1 number"),
            ]),
            ("Notification", [
                ("Low Stock Alert",  "Enabled"),
                ("Email Report",     "Disabled"),
                ("Backup Reminder",  "Enabled"),
            ]),
        ]

        for grp_idx, (group_name, items) in enumerate(settings_groups):
            card = ctk.CTkFrame(
                f, fg_color="#111111", corner_radius=14,
                border_width=1, border_color="#2A2A2A",
            )
            card.grid(row=grp_idx + 1, column=0, sticky="ew", pady=8)

            ctk.CTkLabel(
                card, text=group_name,
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color="#555555", anchor="w",
            ).pack(anchor="w", padx=20, pady=(16, 8))
            ctk.CTkFrame(card, height=1, fg_color="#2A2A2A").pack(fill="x", padx=20)

            for key, val in items:
                row_f = ctk.CTkFrame(card, fg_color="transparent")
                row_f.pack(fill="x", padx=20, pady=8)
                ctk.CTkLabel(row_f, text=key,
                             font=ctk.CTkFont(size=12),
                             text_color="#666666", anchor="w",
                             width=220).pack(side="left")
                ctk.CTkLabel(row_f, text=val,
                             font=ctk.CTkFont(size=12),
                             text_color="#CCCCCC", anchor="w").pack(side="left")

            ctk.CTkFrame(card, height=8, fg_color="transparent").pack()

    # ════════════════════════════════════════════════════════
    # PAGE: USER MANAGEMENT
    # ════════════════════════════════════════════════════════
    def show_user_management(self):
        f = self._base_frame("User Management")

        self._stat_row(f, [
            ("Total Users",  "24",  "◉"),
            ("Active",       "18",  "●"),
            ("Inactive",     "6",   "○"),
            ("Roles",        "7",   "✦"),
        ], row=1)

        self._make_table(f,
            columns={
                "ID": 50, "Username": 140, "Full Name": 180,
                "Role": 110, "Department": 130,
                "Status": 90, "Last Login": 150,
            },
            rows=[
                (1, "admin",    "Administrator",   "ADMIN",    "IT",       "Active",   "2026-04-06 08:30"),
                (2, "john",     "John Doe",        "USER",     "Sales",    "Active",   "2026-04-06 09:10"),
                (3, "finance1", "Finance Officer", "FINANCE",  "Finance",  "Active",   "2026-04-05 17:00"),
                (4, "jane",     "Jane Smith",      "HR",       "HR",       "Active",   "2026-04-06 08:55"),
                (5, "purchase1","Buyer One",       "PURCHASE", "Purchase", "Inactive", "2026-03-28 10:00"),
            ], row=2,
        )

    # ════════════════════════════════════════════════════════
    # PAGE: SCHEDULE
    # ════════════════════════════════════════════════════════
    def show_schedule(self):
        f = self._base_frame("Schedule")

        self._make_table(f,
            columns={
                "ID": 50, "Task": 220, "Assigned To": 150,
                "Due Date": 130, "Priority": 100, "Status": 110,
            },
            rows=[
                (1, "Monthly Report Generation", "Finance Team",  "2026-04-10", "High",   "Pending"),
                (2, "Server Backup",              "IT Admin",      "2026-04-07", "High",   "Scheduled"),
                (3, "Inventory Count",            "Warehouse",     "2026-04-15", "Medium", "Pending"),
                (4, "HR Payroll Processing",      "HR Team",       "2026-04-08", "High",   "In Progress"),
                (5, "System Update",              "IT Admin",      "2026-04-20", "Low",    "Scheduled"),
            ], row=1,
        )

    # ════════════════════════════════════════════════════════
    # PAGE: AUDIT LOG
    # ════════════════════════════════════════════════════════
    def show_audit_log(self):
        f = self._base_frame("Audit Log")

        self._make_table(f,
            columns={
                "ID": 50, "Timestamp": 170, "User": 120,
                "Action": 160, "Module": 120,
                "Detail": 250, "IP": 130,
            },
            rows=[
                (1, "2026-04-06 08:30:01", "admin",    "LOGIN",        "Auth",      "Successful login",            "192.168.1.10"),
                (2, "2026-04-06 08:35:22", "admin",    "VIEW",         "Dashboard", "Viewed dashboard",            "192.168.1.10"),
                (3, "2026-04-06 09:10:05", "john",     "LOGIN",        "Auth",      "Successful login",            "192.168.1.25"),
                (4, "2026-04-06 09:15:44", "john",     "EXPORT",       "Report",    "Exported Monthly Sales PDF",  "192.168.1.25"),
                (5, "2026-04-06 09:45:18", "finance1", "LOGIN",        "Auth",      "Successful login",            "192.168.1.31"),
                (6, "2026-04-06 10:02:33", "admin",    "DELETE",       "User Mgmt", "Deleted user: olduser",       "192.168.1.10"),
                (7, "2026-04-06 10:30:00", "jane",     "LOGIN",        "Auth",      "Successful login",            "192.168.1.42"),
            ], row=1,
        )

    # ════════════════════════════════════════════════════════
    # PAGE: IMPORT EXCEL
    # ════════════════════════════════════════════════════════
    def show_import(self):
        f = self._base_frame("Import Excel")

        info_card = ctk.CTkFrame(
            f, fg_color="#111111", corner_radius=14,
            border_width=1, border_color="#2A2A2A",
        )
        info_card.grid(row=1, column=0, sticky="ew", pady=(0, 16))

        ctk.CTkLabel(
            info_card, text="Import Data from Excel",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#555555", anchor="w",
        ).pack(anchor="w", padx=20, pady=(16, 8))
        ctk.CTkFrame(info_card, height=1, fg_color="#2A2A2A").pack(fill="x", padx=20)

        ctk.CTkLabel(
            info_card,
            text="รองรับไฟล์ .xlsx / .xls  ขนาดสูงสุด 200 MB\n"
                 "ระบบจะ validate ข้อมูลก่อน import และแสดงผลสรุป",
            font=ctk.CTkFont(size=12),
            text_color="#666666", anchor="w", justify="left",
        ).pack(anchor="w", padx=20, pady=12)

        btn_row = ctk.CTkFrame(info_card, fg_color="transparent")
        btn_row.pack(anchor="w", padx=20, pady=(0, 16))

        ctk.CTkButton(
            btn_row, text="  Choose File",
            width=160, height=42, corner_radius=8,
            fg_color="#1A2A3A", hover_color="#1E3A4A",
            border_width=1, border_color="#3498DB",
            text_color="#3498DB",
            font=ctk.CTkFont(size=13),
            command=lambda: messagebox.showinfo("Import", "เลือกไฟล์ Excel"),
        ).pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            btn_row, text="↓  Download Template",
            width=180, height=42, corner_radius=8,
            fg_color="#1A1A2A", hover_color="#1E1E3A",
            border_width=1, border_color="#5F27CD",
            text_color="#5F27CD",
            font=ctk.CTkFont(size=13),
            command=lambda: messagebox.showinfo("Template", "ดาวน์โหลด Template แล้ว"),
        ).pack(side="left")

        ctk.CTkFrame(info_card, height=8, fg_color="transparent").pack()

        self._make_table(f,
            columns={
                "File Name": 200, "Module": 130, "Records": 90,
                "Status": 100, "Imported By": 130, "Date": 160,
            },
            rows=[
                ("inventory_apr.xlsx", "Inventory", "320", "Success", "admin", "2026-04-05 14:30"),
                ("users_q1.xlsx",      "Users",     "45",  "Success", "admin", "2026-04-01 09:00"),
                ("finance_mar.xlsx",   "Finance",   "128", "Failed",  "admin", "2026-03-31 16:20"),
                ("purchase_apr.xlsx",  "Purchase",  "67",  "Success", "admin", "2026-04-03 11:15"),
            ], row=2,
        )

    # ════════════════════════════════════════════════════════
    # LOGOUT
    # ════════════════════════════════════════════════════════
    def logout(self):
        if messagebox.askyesno("Logout", "ต้องการออกจากระบบ?"):
            self._cancel_neon()
            if self._clock_job:
                self.root.after_cancel(self._clock_job)
                self._clock_job = None
            self.logout_callback()
