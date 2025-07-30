# main_app.py (修改后)

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import queue
import numpy as np
from datetime import datetime, timedelta
from collections import defaultdict

from gui_frames import ConnectionFrame, SettingsFrame, RunningFrame
# 使用FakeSignalAnalyzer进行测试
from instrument_controller import FakeSignalAnalyzer as InstrumentController
#from instrument_controller import SignalAnalyzerController as InstrumentController
from report_generator import generate_pdf_report
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg


class SpectrumAnalyzerApp(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Microwave Spectrum Analyzer")
        self.geometry("1200x800")

        self.instrument = None
        # 为防止高速下内存溢出，给队列一个最大尺寸
        self.data_queue = queue.Queue(maxsize=10000)
        self.data_thread = None
        self.stop_thread = threading.Event()
        self.is_running = False
        self.test_timer = None  # 测试计时器
        self.test_end_time = None  # 测试结束时间

        self.frequency_counts = defaultdict(int)
        self.top_frequencies = []

        # 用于计算和显示采集速率
        self.reads_since_last_update = 0
        self.last_update_time = time.time()

        # 固定为2400-2500MHz
        self.plot_fig = None
        self.settings = {}
        self.report_observations = ""

        container = ttk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (ConnectionFrame, RunningFrame, SettingsFrame):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("ConnectionFrame")
        # UI以200ms的间隔更新 (5 FPS)，与数据采集速率解耦
        self.after(200, self.process_queue)

    def show_frame(self, page_name):
        """显示指定名称的Frame"""
        frame = self.frames[page_name]
        frame.tkraise()

    def connect_instrument(self, ip_address, port):
        self.instrument = InstrumentController(ip_address, port)
        return self.instrument.connect()

    def disconnect_instrument(self):
        if self.instrument:
            self.instrument.close()
            self.instrument = None

    def get_device_id(self):
        if self.instrument and self.instrument.connected:
            return f"Keysight X-Series @ {self.instrument.address}:{self.instrument.port}"
        return "N/A"

    def start_data_acquisition(self, duration_minutes=None):
        """启动数据采集 (不需要interval参数)"""
        self.is_running = True
        self.stop_thread.clear()

        # 重置统计数据和速率计算器
        self.frequency_counts.clear()
        self.top_frequencies = []
        self.reads_since_last_update = 0
        self.last_update_time = time.time()
        # 清空队列中可能残留的旧数据
        while not self.data_queue.empty():
            self.data_queue.get()

        # 设置测试结束时间（如果提供了持续时间）
        if duration_minutes is not None and duration_minutes > 0:
            self.test_end_time = time.time() + duration_minutes * 60
            # 每分钟检查一次是否到达结束时间
            self.test_timer = threading.Timer(60.0, self.check_test_time)
            self.test_timer.daemon = True
            self.test_timer.start()
        else:
            self.test_end_time = None

        self.data_thread = threading.Thread(target=self._data_acquisition_loop, daemon=True)
        self.data_thread.start()

    def check_test_time(self):
        """检查测试时间是否结束"""
        if self.is_running and self.test_end_time and time.time() >= self.test_end_time:
            # 时间到，停止测试
            self.stop_data_acquisition()
            # 更新UI状态
            running_frame = self.frames["RunningFrame"]
            if running_frame:
                running_frame.stop_button.config(state="disabled")
                running_frame.start_button.config(state="normal")
                running_frame.test_time_entry.config(state="normal")
                running_frame.rate_label.config(text="Rate: 0.0 Hz")
            messagebox.showinfo("Test Completed", "Test duration has been reached.")
        elif self.is_running:
            # 每分钟继续检查
            self.test_timer = threading.Timer(60.0, self.check_test_time)
            self.test_timer.daemon = True
            self.test_timer.start()

    def stop_data_acquisition(self):
        self.is_running = False
        self.stop_thread.set()

        # 取消测试计时器
        if self.test_timer:
            self.test_timer.cancel()
            self.test_timer = None
        self.test_end_time = None

    def _data_acquisition_loop(self):
        """数据采集线程，全速运行"""
        while not self.stop_thread.is_set():
            if self.instrument and self.instrument.connected:
                try:
                    peak_freq = self.instrument.read_peak_frequency()
                    if peak_freq is not None:
                        # 只将单个频率值放入队列
                        # 四舍五入为整数
                        rounded_freq = round(peak_freq)
                        if not self.data_queue.full():
                            self.data_queue.put(rounded_freq)
                        else:
                            # 队列满了，丢弃数据并打印警告，避免刷屏
                            print("Warning: Data queue is full, dropping data.", end='\r')
                except Exception as e:
                    print(f"数据读取时发生严重错误: {e}")
                    # 可以选择在这里自动停止采集
                    self.stop_data_acquisition()
                    break
        print("\n数据采集线程已停止。")

    def process_queue(self):
        """以固定的时间间隔处理队列中的所有数据，并更新UI"""
        # 1. 批量处理数据
        reads_this_batch = 0
        while not self.data_queue.empty():
            try:
                freq = self.data_queue.get_nowait()
                # 注意：这里不再需要四舍五入，因为数据在放入队列时已经处理
                self.frequency_counts[freq] += 1
                reads_this_batch += 1
            except queue.Empty:
                break

        self.reads_since_last_update += reads_this_batch

        # 2. 计算实际采集速率
        now = time.time()
        elapsed = now - self.last_update_time
        rate = 0
        if elapsed > 0.1:  # 避免刚开始时elapsed过小导致速率异常
            rate = self.reads_since_last_update / elapsed
            self.last_update_time = now
            self.reads_since_last_update = 0

        # 3. 如果有新数据，更新统计和UI
        if reads_this_batch > 0:
            self.top_frequencies = sorted(
                self.frequency_counts.items(),
                key=lambda x: x[1],
                reverse=True
            )[:5]

        running_frame = self.frames["RunningFrame"]
        if self.is_running:
            # 始终更新UI，即使没有新数据也要更新速率显示
            running_frame.update_ui(self.top_frequencies, rate)

        # 调度下一次UI更新
        self.after(200, self.process_queue)

    def update_frequency_plot(self):
        """更新频率分布图 (固定范围为2400-2500MHz)"""
        running_frame = self.frames["RunningFrame"]
        ax = running_frame.ax
        ax.clear()

        # 固定频率范围
        freq_min = 2400
        freq_max = 2500

        in_range = [(freq, count) for freq, count in self.frequency_counts.items()
                    if freq_min <= freq <= freq_max]

        if not in_range:
            ax.set_title(f"Frequency Distribution ({freq_min}-{freq_max} MHz)")
            ax.set_xlabel("Frequency (MHz)")
            ax.set_ylabel("Count")
            ax.grid(True, linestyle='--', alpha=0.7)
            ax.set_xlim(freq_min, freq_max)  # 设置固定范围
            running_frame.canvas.draw()
            return

        bin_size = 1.0  # 使用1MHz的bin大小
        bins = {}
        for freq, count in in_range:
            # 频率已经是整数，直接使用
            bins[freq] = bins.get(freq, 0) + count

        sorted_bins = sorted(bins.items())
        bin_centers = [bc for bc, count in sorted_bins]
        counts = [count for bc, count in sorted_bins]

        # 使用条形图显示
        ax.bar(bin_centers, counts, width=bin_size * 0.8, color='steelblue')

        # 标注前5个频率点
        for rank, (freq, count) in enumerate(self.top_frequencies[:5], 1):
            if freq_min <= freq <= freq_max:
                ax.annotate(f"#{rank}", (freq, count), textcoords="offset points",
                            xytext=(0, 10), ha='center', fontsize=9,
                            arrowprops=dict(arrowstyle="->", color='red'))

        ax.set_title(f"Frequency Distribution ({freq_min}-{freq_max} MHz)")
        ax.set_xlabel("Frequency (MHz)")
        ax.set_ylabel("Count")
        ax.grid(True, linestyle='--', alpha=0.7)
        ax.set_xlim(freq_min, freq_max)  # 设置固定范围

        running_frame.canvas.draw()
        self.plot_fig = running_frame.fig

    def prepare_for_report(self, observations):
        self.report_observations = observations

    def generate_final_report(self):
        if not self.frequency_counts:
            messagebox.showwarning("Warning", "No frequency data collected.")
            return

        report_data = self.settings.copy()
        report_data['Start time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        report_data['Observations'] = self.report_observations

        table_data = [["Rank", "Frequency (MHz)", "Count"]]
        for rank, (freq, count) in enumerate(self.top_frequencies, 1):
            table_data.append([str(rank), f"{freq}", str(count)])  # 频率已经是整数

        report_data['frequency_data'] = table_data

        fig = self.plot_fig
        if fig is None:
            self.update_frequency_plot()
            fig = self.plot_fig

        temp_image_path = "temp_frequency_plot.png"
        if fig:
            fig.savefig(temp_image_path, dpi=150, bbox_inches='tight')
            plot_data = [{'title': "Frequency Distribution", 'path': temp_image_path}]
        else:
            plot_data = []

        filepath_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Documents", "*.pdf")],
            title="Save Report As"
        )

        if not filepath_pdf:
            if os.path.exists(temp_image_path):
                os.remove(temp_image_path)
            return

        success = generate_pdf_report(filepath_pdf, report_data, plot_data)

        if os.path.exists(temp_image_path):
            os.remove(temp_image_path)

        if success:
            messagebox.showinfo("Success", f"Report saved to:\n{filepath_pdf}")
        else:
            messagebox.showerror("Error", "Failed to generate report")

    def on_closing(self):
        self.stop_thread.set()
        if self.data_thread and self.data_thread.is_alive():
            self.data_thread.join(timeout=1.0)
        if self.is_running:
            self.stop_data_acquisition()
        if self.instrument:
            self.instrument.close()
        if hasattr(self, 'plot_fig'):
            plt.close(self.plot_fig)
        self.destroy()


if __name__ == "__main__":
    app = SpectrumAnalyzerApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()