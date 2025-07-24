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

try:
    import openpyxl
except ImportError:
    openpyxl = None


# --- 1. ConnectionFrame (已修改以支持GPIB) ---
class ConnectionFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=0)  # 主内容列
        self.columnconfigure(2, weight=1)

        # --- 仪器选择 ---
        ttk.Label(self, text="Select Instrument:", font=("Helvetica", 12)).grid(row=0, column=1, pady=(20, 5))
        self.instrument_var = tk.StringVar()
        self.instrument_selector = ttk.Combobox(self, textvariable=self.instrument_var, state="readonly", width=35)
        self.instrument_selector['values'] = ('Keithley 2701 (TCP/IP)', 'Keithley 2700 (GPIB)')
        self.instrument_selector.current(0)
        self.instrument_selector.grid(row=1, column=1, pady=5, ipady=4)
        self.instrument_selector.bind("<<ComboboxSelected>>", self.on_instrument_select)

        # --- 地址输入 ---
        self.address_label = ttk.Label(self, text="IP Address:", font=("Helvetica", 12))
        self.address_label.grid(row=2, column=1, pady=(10, 5))

        self.address_entry = ttk.Entry(self, width=38)
        self.address_entry.grid(row=3, column=1, pady=5, ipady=4)

        # --- 初始状态 ---
        self.on_instrument_select(None)  # 初始化标签和默认地址

        # --- 通道范围选择 (保持不变) ---
        self.device_type_var = tk.StringVar(value="1-80")
        device_selector_frame = ttk.Frame(self)
        device_selector_frame.grid(row=4, column=1, pady=5)
        ttk.Label(device_selector_frame, text="Channel Range:").pack(side="left", padx=5)
        ttk.Radiobutton(device_selector_frame, text="1-80", variable=self.device_type_var, value="1-80").pack(
            side="left")
        ttk.Radiobutton(device_selector_frame, text="81-160", variable=self.device_type_var, value="81-160").pack(
            side="left")

        # --- 按钮和状态 (布局行号已调整) ---
        button_frame = ttk.Frame(self)
        button_frame.grid(row=5, column=1, pady=10)
        self.connect_button = ttk.Button(button_frame, text="Connect", command=self.connect_device)
        self.connect_button.pack(side="left", padx=5)
        self.disconnect_button = ttk.Button(button_frame, text="Disable Connect", command=self.disconnect_device,
                                            state="disabled")
        self.disconnect_button.pack(side="left", padx=5)

        self.status_label = ttk.Label(self, text="Status: Not Connected", foreground="red")
        self.status_label.grid(row=6, column=1, pady=5)
        self.device_id_label = ttk.Label(self, text="Instrument: N/A")
        self.device_id_label.grid(row=7, column=1, pady=5)
        self.next_button = ttk.Button(self, text="Continue to Test", state="disabled",
                                      command=lambda: controller.show_frame("RunningFrame"))
        self.next_button.grid(row=8, column=1, pady=20)

        about_button = ttk.Button(self, text="About", command=self.open_github_link)
        about_button.grid(row=9, column=1, pady=10)

    def on_instrument_select(self, event):
        """当用户切换仪器选项时，更新地址标签和默认值。"""
        selection = self.instrument_var.get()
        self.address_entry.delete(0, tk.END)
        if "TCP/IP" in selection:
            self.address_label.config(text="IP Address:")
            self.address_entry.insert(0, "192.168.1.100")
        elif "GPIB" in selection:
            self.address_label.config(text="GPIB Address (e.g., 22):")
            self.address_entry.insert(0, "22")

    def open_github_link(self):
        url = "https://github.com/Ashes2Ashes233/TemPyScan"
        webbrowser.open_new(url)

    def connect_device(self):
        address = self.address_entry.get()
        device_type = self.device_type_var.get()
        instrument_selection = self.instrument_var.get()

        if not address:
            messagebox.showerror("Error", "Address cannot be empty.")
            return

        conn_type = 'TCPIP' if 'TCP/IP' in instrument_selection else 'GPIB'

        if self.controller.connect_instrument(conn_type, address, device_type):
            self.status_label.config(text="Status: Connected", foreground="green")
            self.device_id_label.config(text=f"Instrument: {self.controller.get_device_id()}")
            self.connect_button.config(state="disabled")
            self.disconnect_button.config(state="normal")
            self.next_button.config(state="normal")
            self.instrument_selector.config(state="disabled")
        else:
            self.status_label.config(text="Status: Connection Failed", foreground="red")
            self.device_id_label.config(text="Instrument: N/A")
            messagebox.showerror("Connection Failed", f"Could not connect to device at {address} via {conn_type}.")

    def disconnect_device(self):
        self.controller.disconnect_instrument()
        self.status_label.config(text="Status: Not Connected", foreground="red")
        self.device_id_label.config(text="Instrument: N/A")
        self.connect_button.config(state="normal")
        self.disconnect_button.config(state="disabled")
        self.next_button.config(state="disabled")
        self.instrument_selector.config(state="normal")


# --- SettingsFrame (保持不变) ---
# ... (此处省略 SettingsFrame 的代码, 无需修改) ...
class SettingsFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.entries = {}
        frame = ttk.Frame(self)
        frame.pack(pady=20, padx=20)
        ttk.Label(self, text="Report Configuration", font=("Helvetica", 16, "bold")).pack(pady=10)
        fields = [
            ("Test type", "Combobox", ["Normal", "Abnormal"]), ("Test name", "Entry", ""),
            ("Operating Voltage", "Entry", ""), ("Operating Frequency", "Entry", ""),
            ("Operating Duration", "Entry", ""),
            ("Rating Voltage", "Entry", ""), ("Rating Frequency", "Entry", ""),
            ("Ambient Channel", "Entry", ""),
            ("Temperature contrast with limit", "Entry", ""), ("Lab request", "Entry", ""),
            ("Sample number", "Entry", ""), ("Model number", "Entry", ""),
            ("Sample Characteristics", "Text", ""), ("Tester", "Entry", ""),
            ("Equipment", "Entry", ""),
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
            elif widget_type == "Text":
                widget = scrolledtext.ScrolledText(frame, width=30, height=4, wrap=tk.WORD)
            widget.grid(row=i, column=1, sticky="ew", padx=5, pady=5)
            self.entries[label_text] = widget

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Clear Info", command=self.clear_info).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Generate Report & Data Table", command=self.confirm_and_generate_report).pack(
            side="left", padx=10)
        ttk.Button(button_frame, text="Back to Test", command=lambda: self.controller.show_frame("RunningFrame")).pack(
            side="left", padx=10)

    def clear_info(self):
        for widget in self.entries.values():
            if isinstance(widget, ttk.Entry):
                widget.delete(0, tk.END)
            elif isinstance(widget, ttk.Combobox):
                widget.set('')
            elif isinstance(widget, scrolledtext.ScrolledText):
                widget.delete('1.0', tk.END)

    def confirm_and_generate_report(self):
        settings = {}
        for key, widget in self.entries.items():
            if isinstance(widget, (ttk.Entry, ttk.Combobox)):
                settings[key] = widget.get()
            elif isinstance(widget, scrolledtext.ScrolledText):
                settings[key] = widget.get('1.0', tk.END).strip()
        self.controller.settings = settings
        print("Final report settings confirmed.")
        self.controller.generate_final_report()


# --- RunningFrame (保持不变) ---
# ... (此处省略 RunningFrame 的代码, 无需修改) ...
class RunningFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)
        left_frame = ttk.Frame(main_pane)
        main_pane.add(left_frame, weight=2)
        control_frame = ttk.Frame(left_frame)
        control_frame.pack(fill=tk.X, pady=5, padx=5)

        self.start_button = ttk.Button(control_frame, text="Start", command=self.start_test)
        self.start_button.pack(side="left", padx=5)
        self.stop_button = ttk.Button(control_frame, text="Stop", command=self.stop_test, state="disabled")
        self.stop_button.pack(side="left", padx=5)
        ttk.Label(control_frame, text="Cycle/s(Scan time not included):").pack(side="left", padx=(10, 0))
        self.interval_entry = ttk.Entry(control_frame, width=5)
        self.interval_entry.insert(0, "2")
        self.interval_entry.pack(side="left", padx=5)
        # 添加环境通道设置
        ambient_frame = ttk.Frame(control_frame)
        ambient_frame.pack(side="left", padx=(10, 0))
        ttk.Label(ambient_frame, text="Ambient Channel:").pack(side="left")
        self.ambient_channel_entry = ttk.Entry(ambient_frame, width=5)
        self.ambient_channel_entry.pack(side="left", padx=5)
        self.ambient_channel_entry.insert(0, "0")  # 默认值
        table_frame = ttk.Frame(left_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        cols = ("Channel", "Location", "Current Temp (°C)", "Max Temp (°C)", "Threshold (°C)")
        self.tree = ttk.Treeview(table_frame, columns=cols, show='headings', height=25)
        for col in cols: self.tree.heading(col, text=col)
        self.tree.column("Channel", width=60, anchor='center')
        self.tree.column("Location", width=150, anchor='w')
        self.tree.column("Current Temp (°C)", width=120, anchor='center')
        self.tree.column("Max Temp (°C)", width=120, anchor='center')
        self.tree.column("Threshold (°C)", width=120, anchor='center')
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.tag_configure('over_threshold', foreground='white', background='red', font=('Consolas', 10, 'bold'))
        self.tree.bind("<Double-1>", self.on_double_click)
        right_pane = ttk.PanedWindow(main_pane, orient=tk.VERTICAL)
        main_pane.add(right_pane, weight=3)
        plot_frame = ttk.LabelFrame(right_pane, text="Graph Panel - Use toolbar to Pan/Zoom")
        right_pane.add(plot_frame, weight=2)
        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.toolbar = NavigationToolbar2Tk(self.canvas, plot_frame)
        self.toolbar.update()
        ch_select_frame = ttk.Frame(plot_frame)
        ch_select_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(ch_select_frame, text="Channels:").pack(side=tk.LEFT, padx=5)
        self.plot_channels_entry = ttk.Entry(ch_select_frame)
        self.plot_channels_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        time_frame = ttk.Frame(plot_frame)
        time_frame.pack(fill=tk.X, pady=(2, 5))
        ttk.Label(time_frame, text="Time Range (s):").pack(side=tk.LEFT, padx=5)
        ttk.Label(time_frame, text="Start:").pack(side=tk.LEFT)
        self.start_time_entry = ttk.Entry(time_frame, width=8)
        self.start_time_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(time_frame, text="End:").pack(side=tk.LEFT)
        self.end_time_entry = ttk.Entry(time_frame, width=8)
        self.end_time_entry.pack(side=tk.LEFT)
        self.update_plot_button = ttk.Button(time_frame, text="Update Plot", command=self.redraw_historical_plot)
        self.update_plot_button.pack(side=tk.RIGHT, padx=5)
        y_frame = ttk.Frame(plot_frame)
        y_frame.pack(fill=tk.X, pady=(2, 5))
        ttk.Label(y_frame, text="Y Range (°C):").pack(side=tk.LEFT, padx=5)
        ttk.Label(y_frame, text="Min:").pack(side=tk.LEFT)
        self.y_min_entry = ttk.Entry(y_frame, width=8)
        self.y_min_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(y_frame, text="Max:").pack(side=tk.LEFT)
        self.y_max_entry = ttk.Entry(y_frame, width=8)
        self.y_max_entry.pack(side=tk.LEFT)
        report_build_frame = ttk.Frame(right_pane)
        right_pane.add(report_build_frame, weight=1)
        report_info_frame = ttk.LabelFrame(report_build_frame, text="Report Notes")
        report_info_frame.pack(fill="both", expand=True, padx=5, pady=5)
        report_info_frame.columnconfigure(1, weight=1)
        ttk.Label(report_info_frame, text="Phenomena&Result:").grid(row=0, column=0, sticky='nw', padx=5, pady=2)
        self.phenomena_text = scrolledtext.ScrolledText(report_info_frame, height=3)
        self.phenomena_text.grid(row=0, column=1, sticky='nsew', padx=5, pady=2)
        ttk.Label(report_info_frame, text="Notes:").grid(row=1, column=0, sticky='nw', padx=5, pady=2)
        self.notes_text = scrolledtext.ScrolledText(report_info_frame, height=3)
        self.notes_text.grid(row=1, column=1, sticky='nsew', padx=5, pady=2)
        report_info_frame.rowconfigure(0, weight=1)
        report_info_frame.rowconfigure(1, weight=1)
        export_button_frame = ttk.Frame(report_build_frame)
        export_button_frame.pack(pady=10)
        self.create_report_button = ttk.Button(export_button_frame, text="Proceed to Report Settings",
                                               command=self.proceed_to_report)
        self.create_report_button.pack(side="left", padx=10)

    def start_test(self):
        # 获取环境通道设置
        ambient_channel = self.ambient_channel_entry.get().strip()

        # 清除树状视图
        self.tree.delete(*self.tree.get_children())
        configs = self.controller.get_channel_configs()
        for i, config in enumerate(configs):
            self.tree.insert("", "end", iid=i, values=(i + 1, config['location'], "N/A", "N/A", config['threshold']))

        # 启动数据采集，传递环境通道设置
        self.controller.start_data_acquisition(
            int(self.interval_entry.get()),
            ambient_channel
        )

        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")

    def on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        column_id = self.tree.identify_column(event.x)
        column_index = int(column_id.replace('#', '')) - 1
        if column_index not in [1, 4]: return
        row_id = self.tree.focus()
        x, y, width, height = self.tree.bbox(row_id, column_id)
        entry = ttk.Entry(self.tree)
        entry.place(x=x, y=y, width=width, height=height)
        current_value = self.tree.item(row_id, "values")[column_index]
        entry.insert(0, current_value)
        entry.focus()
        entry.bind("<FocusOut>", lambda e: entry.destroy())
        entry.bind("<Return>", lambda e: self.save_edit(entry, row_id, column_index))

    def save_edit(self, entry, row_id, column_index):
        new_value = entry.get()
        values = list(self.tree.item(row_id, "values"))
        values[column_index] = new_value
        self.tree.item(row_id, values=values)
        self.controller.update_channel_config(int(row_id), column_index, new_value)
        entry.destroy()

    def stop_test(self):
        self.controller.stop_data_acquisition()
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")

    def redraw_historical_plot(self, **kwargs):
        self.ax.clear()
        self.ax.grid(True)
        self.ax.set_title(kwargs.get('title', "Temperature History"))
        self.ax.set_xlabel("Time (seconds)")
        self.ax.set_ylabel("Temperature (°C)")
        channels_to_plot = kwargs.get('channels_to_plot',
                                      self.controller.parse_channel_selection(self.plot_channels_entry.get()))
        start_time_str = kwargs.get('start_time', self.start_time_entry.get())
        end_time_str = kwargs.get('end_time', self.end_time_entry.get())
        sliced_data = self.controller.get_sliced_data(channels_to_plot, start_time_str, end_time_str)
        if not sliced_data: self.canvas.draw(); return
        sliced_history = sliced_data['history']
        slice_start_ts = sliced_data['actual_start_ts']
        colors = plt.get_cmap('tab20').colors
        plotted_something = False
        for i in channels_to_plot:
            if i in sliced_history and sliced_history[i]:
                timestamps, temps_y = zip(*sliced_history[i])
                elapsed_time = [ts - slice_start_ts for ts in timestamps]
                self.ax.plot(elapsed_time, temps_y, label=f"Ch {i + 1}", color=colors[i % len(colors)])
                plotted_something = True
        y_min = kwargs.get('y_min', self.y_min_entry.get())
        y_max = kwargs.get('y_max', self.y_max_entry.get())
        try:
            if y_min and y_min.strip() != '': self.ax.set_ylim(bottom=float(y_min))
            if y_max and y_max.strip() != '': self.ax.set_ylim(top=float(y_max))
        except ValueError:
            pass
        if plotted_something: self.ax.legend(loc='upper left', fontsize='small')
        self.canvas.draw()

    def update_ui(self, temps, max_temps):
        if temps is None: return
        channel_configs = self.controller.get_channel_configs()
        for i in range(160):
            temp = temps[i]
            if np.isnan(temp): continue
            try:
                threshold = float(channel_configs[i]['threshold'])
            except (ValueError, KeyError):
                threshold = float('inf')
            max_temp_val = max_temps[i]
            max_temp_str = f"{max_temp_val:.2f}" if not np.isinf(max_temp_val) else "N/A"
            tag = 'over_threshold' if temp > threshold else ''
            self.tree.item(i, values=(
                i + 1, channel_configs[i]['location'], f"{temp:.2f}", max_temp_str, channel_configs[i]['threshold']),
                           tags=(tag,))
        self.redraw_historical_plot(title="Live Temperature View", start_time="", end_time="",
                                    y_min=self.y_min_entry.get(), y_max=self.y_max_entry.get())

    def proceed_to_report(self):
        notes = {'phenomena': self.phenomena_text.get("1.0", tk.END).strip(),
                 'notes': self.notes_text.get("1.0", tk.END).strip()}
        channels_str = self.plot_channels_entry.get()
        time_range = {'start': self.start_time_entry.get(), 'end': self.end_time_entry.get()}
        self.controller.prepare_for_report(notes, channels_str, time_range)
        self.controller.show_frame("SettingsFrame")