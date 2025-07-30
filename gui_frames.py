# gui_frames.py

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.pyplot as plt
from datetime import datetime
import os
import time
import webbrowser
import re

try:
    import openpyxl
except ImportError:
    openpyxl = None


# --- 使用正则表达式解析选择的通道号 ---
def parse_channel_selection(text):
    channels = set()
    if not text: return []
    tokens = re.split(r',\s*(?![^()]*\))', text.strip())
    range_pattern = re.compile(r'\(\s*(\d+)\s*,\s*(\d+)\s*\)')
    for token in tokens:
        token = token.strip()
        if not token: continue
        match = range_pattern.fullmatch(token)
        if match:
            start, end = int(match.group(1)), int(match.group(2))
            if start > end: start, end = end, start
            for i in range(start, end + 1): channels.add(i)
            continue
        if token.isdigit():
            channels.add(int(token))
        else:
            print(f"警告: 无法解析的通道输入 '{token}'")
    return sorted([ch - 1 for ch in channels])


# --- 1. ConnectionFrame (Telnet连接) ---
class ConnectionFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.columnconfigure(0, weight=1)
        self.columnconfigure(2, weight=1)

        # IP地址和端口输入
        ttk.Label(self, text="Instrument IP Address:", font=("Helvetica", 12)).grid(row=0, column=1, pady=(20, 5))
        self.ip_entry = ttk.Entry(self, width=30)
        self.ip_entry.insert(0, "169.254.223.159")
        self.ip_entry.grid(row=1, column=1, pady=5, ipady=4)

        ttk.Label(self, text="Port:", font=("Helvetica", 12)).grid(row=2, column=1, pady=(5, 5))
        self.port_entry = ttk.Entry(self, width=10)
        self.port_entry.insert(0, "5023")
        self.port_entry.grid(row=3, column=1, pady=5)

        # 连接按钮
        button_frame = ttk.Frame(self)
        button_frame.grid(row=4, column=1, pady=10)
        self.connect_button = ttk.Button(button_frame, text="Connect", command=self.connect_device)
        self.connect_button.pack(side="left", padx=5)
        self.disconnect_button = ttk.Button(button_frame, text="Disconnect", command=self.disconnect_device,
                                            state="disabled")
        self.disconnect_button.pack(side="left", padx=5)

        # 状态显示
        self.status_label = ttk.Label(self, text="Status: Not Connected", foreground="red")
        self.status_label.grid(row=5, column=1, pady=5)
        self.device_id_label = ttk.Label(self, text="Instrument: N/A")
        self.device_id_label.grid(row=6, column=1, pady=5)

        # 继续按钮
        self.next_button = ttk.Button(self, text="Continue to Test", state="disabled",
                                      command=lambda: controller.show_frame("RunningFrame"))
        self.next_button.grid(row=7, column=1, pady=20)

        # About按钮
        about_button = ttk.Button(self, text="About", command=self.open_github_link)
        about_button.grid(row=8, column=1, pady=10)

    def open_github_link(self):
        url = "https://github.com/Ashes2Ashes233/WaveScan"
        try:
            webbrowser.open_new(url)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open the link.\nError: {e}")

    def connect_device(self):
        ip_address = self.ip_entry.get()
        port = self.port_entry.get()

        if not ip_address:
            messagebox.showerror("Error", "No IP Address")
            return

        try:
            port = int(port)
        except ValueError:
            messagebox.showerror("Error", "Invalid port number")
            return

        if self.controller.connect_instrument(ip_address, port):
            self.status_label.config(text="Status: Connected", foreground="green")
            self.device_id_label.config(text=f"Instrument: {self.controller.get_device_id()}")
            self.connect_button.config(state="disabled")
            self.disconnect_button.config(state="normal")
            self.next_button.config(state="normal")
        else:
            self.status_label.config(text="Status: Connection Failed", foreground="red")
            self.device_id_label.config(text="Instrument: N/A")
            messagebox.showerror("Connection Failed",
                                 f"Cannot connect to {ip_address}:{port}.\nPlease check IP address or instrument.")

    def disconnect_device(self):
        self.controller.disconnect_instrument()
        self.status_label.config(text="Status: Not Connected", foreground="red")
        self.device_id_label.config(text="Instrument: N/A")
        self.connect_button.config(state="normal")
        self.disconnect_button.config(state="disabled")
        self.next_button.config(state="disabled")


# --- 2. 报告设置界面 ---
class SettingsFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.entries = {}

        ttk.Label(self, text="Report Configuration", font=("Helvetica", 16, "bold")).pack(pady=10)

        frame = ttk.Frame(self)
        frame.pack(pady=20, padx=20)

        # 简化的字段
        fields = [
            ("Test name", "Entry", ""),
            ("Test type", "Combobox", ["Normal", "Abnormal"]),
            ("Sample number", "Entry", ""),
            ("Model number", "Entry", ""),
            ("Lab request", "Entry", ""),
            ("Tester", "Entry", ""),
            ("Equipment", "Entry", "Keysight X-Series Analyzer"),
        ]

        for i, (label_text, widget_type, default_value) in enumerate(fields):
            label = ttk.Label(frame, text=label_text + ":")
            label.grid(row=i, column=0, sticky="w", padx=5, pady=5)

            if widget_type == "Entry":
                widget = ttk.Entry(frame, width=40)
                widget.insert(0, default_value)
            elif widget_type == "Combobox":
                widget = ttk.Combobox(frame, values=default_value, width=38)
                if default_value:
                    widget.current(0)

            widget.grid(row=i, column=1, sticky="ew", padx=5, pady=5)
            self.entries[label_text] = widget

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Clear Info", command=self.clear_info).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Generate Report", command=self.confirm_and_generate_report).pack(
            side="left", padx=10)
        ttk.Button(button_frame, text="Back to Test", command=lambda: self.controller.show_frame("RunningFrame")).pack(
            side="left", padx=10)

    def clear_info(self):
        for widget in self.entries.values():
            if isinstance(widget, ttk.Entry):
                widget.delete(0, tk.END)
            elif isinstance(widget, ttk.Combobox):
                widget.set('')

    def confirm_and_generate_report(self):
        settings = {}
        for key, widget in self.entries.items():
            if isinstance(widget, (ttk.Entry, ttk.Combobox)):
                settings[key] = widget.get()

        self.controller.settings = settings
        self.controller.generate_final_report()


# --- 3. 运行界面 (微波频率统计) ---
class RunningFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(main_pane)
        main_pane.add(left_frame, weight=1)

        # 控制面板
        control_frame = ttk.Frame(left_frame)
        control_frame.pack(fill=tk.X, pady=5, padx=5)

        self.start_button = ttk.Button(control_frame, text="Start", command=self.start_test)
        self.start_button.pack(side="left", padx=5)

        self.stop_button = ttk.Button(control_frame, text="Stop", command=self.stop_test, state="disabled")
        self.stop_button.pack(side="left", padx=5)

        # 测试时间设置
        ttk.Label(control_frame, text="Test Time (min):", font=("Helvetica", 10)).pack(side="left", padx=(10, 5))
        self.test_time_entry = ttk.Entry(control_frame, width=6)
        self.test_time_entry.insert(0, "")
        self.test_time_entry.pack(side="left", padx=5)

        self.rate_label = ttk.Label(control_frame, text="Rate: 0.0 Hz", font=("Helvetica", 10))
        self.rate_label.pack(side="left", padx=(10, 5))

        # 频率表格
        table_frame = ttk.LabelFrame(left_frame, text="Top Frequencies")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        cols = ("Rank", "Frequency (MHz)", "Hits")
        self.tree = ttk.Treeview(table_frame, columns=cols, show='headings', height=8)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80, anchor='center')
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 右侧面板
        right_frame = ttk.Frame(main_pane)
        main_pane.add(right_frame, weight=2)
        plot_frame = ttk.LabelFrame(right_frame, text="Frequency Distribution (2400-2500 MHz)")
        plot_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)

        # 初始化图表，固定范围
        self.ax.set_xlim(2400, 2500)
        self.ax.set_title("Frequency Distribution (2400-2500 MHz)")
        self.ax.set_xlabel("Frequency (MHz)")
        self.ax.set_ylabel("Count")
        self.ax.grid(True, linestyle='--', alpha=0.7)

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.toolbar = NavigationToolbar2Tk(self.canvas, plot_frame)
        self.toolbar.update()

        # 移除范围调整控件

        report_frame = ttk.LabelFrame(right_frame, text="Report Notes")
        report_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(report_frame, text="Observations:").pack(anchor="w", padx=5, pady=2)
        self.observations_text = scrolledtext.ScrolledText(report_frame, height=3)
        self.observations_text.pack(fill=tk.X, padx=5, pady=5)
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(pady=10)
        self.create_report_button = ttk.Button(button_frame, text="Generate Report", command=self.proceed_to_report)
        self.create_report_button.pack(side="left", padx=10)

    def start_test(self):
        """启动测试，考虑测试时间设置"""
        try:
            test_time = int(self.test_time_entry.get())
            if test_time <= 0:
                raise ValueError("Test time must be positive")
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid positive number for test time (minutes)")
            return

        self.controller.start_data_acquisition(duration_minutes=test_time)
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.rate_label.config(text="Rate: ... Hz")
        self.test_time_entry.config(state="disabled")  # 测试开始后禁用时间设置

    def stop_test(self):
        self.controller.stop_data_acquisition()
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.rate_label.config(text="Rate: 0.0 Hz")
        self.test_time_entry.config(state="normal")  # 测试结束后启用时间设置

    def update_ui(self, top_frequencies, rate):
        """更新UI，现在接收采集速率作为参数"""
        self.rate_label.config(text=f"Rate: {rate:.1f} Hz")
        self.tree.delete(*self.tree.get_children())
        for rank, (freq, count) in enumerate(top_frequencies, 1):
            self.tree.insert("", "end", values=(rank, f"{freq}", count))  # 频率已经是整数
        # 自动更新图表
        self.controller.update_frequency_plot()

    def proceed_to_report(self):
        observations = self.observations_text.get("1.0", tk.END).strip()
        self.controller.prepare_for_report(observations)
        self.controller.show_frame("SettingsFrame")