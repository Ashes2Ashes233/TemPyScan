# main_app.py

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import time
import queue
import numpy as np
from datetime import datetime
from collections import deque
from gui_frames import ConnectionFrame, SettingsFrame, RunningFrame
# from instrument_controller import FakeKeithley2701 as InstrumentController
from instrument_controller import KeithleyController as InstrumentController
from report_generator import generate_pdf_report
import os
import re
import openpyxl


# parse_channel_selection
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


class ThermoApp(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("TemPyScan")
        self.instrument = None
        self.settings = {}
        self.data_queue = queue.Queue()
        self.data_thread = None
        self.stop_thread = threading.Event()
        self.is_running = False
        self.channel_offset = 0
        self.max_temps = np.full(160, -np.inf)
        self.history = {i: deque() for i in range(160)}
        self.start_time = None
        self.stop_time = None
        self.start_timestamp = 0
        self.channel_configs = [{'location': '', 'threshold': ''} for i in range(160)]
        self.report_notes = {}
        self.report_channels_str = ""
        self.report_time_range = {}
        self.init = False
        self.ambient_channel = None
        self.ambient_start_temp = "N/A"
        self.ambient_end_temp = "N/A"
        container = ttk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.frames = {}
        for F in (ConnectionFrame, SettingsFrame, RunningFrame):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame("ConnectionFrame")
        self.after(100, self.process_queue)

    # --- connect_instrument ---
    def connect_instrument(self, conn_type, address, device_type):
        """
        连接到指定的仪器.
        :param conn_type: 'TCPIP' 或 'GPIB'
        :param address: IP 地址或 GPIB 地址
        :param device_type: "1-80" 或 "81-160"
        """
        self.channel_offset = 0 if device_type == "1-80" else 80
        # 实例化通用的控制器
        self.instrument = InstrumentController(conn_type=conn_type, address=address)
        return self.instrument.connect()

    def parse_channel_selection(self, text):
        return parse_channel_selection(text)

    def generate_final_report(self):
        running_frame = self.frames['RunningFrame']
        channels_for_report = self.parse_channel_selection(self.report_channels_str)
        if not channels_for_report: messagebox.showwarning("Warning", "No channels specified for the report."); return
        sliced_data = self.get_sliced_data(channels_for_report, self.report_time_range.get('start'),
                                           self.report_time_range.get('end'))
        if not sliced_data or not sliced_data['history']: messagebox.showwarning("Warning",
                                                                                 "No data found for the selected channels and time range."); return
        filepath_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Documents", "*.pdf")],
                                                    title="Save Report As")
        if not filepath_pdf: return
        filepath_excel = os.path.splitext(filepath_pdf)[0] + '.xlsx'
        temp_image_files = []
        plot_data_for_report = []
        try:
            # 实现动态分组逻辑
            group_size_str = self.settings.get('Channels per Graph', '')

            # 如果输入为空，则将所有通道画在一张图里
            if group_size_str == '':
                # 设置一个比最大通道数还大的数，确保所有通道都在一个组里
                group_size = 161
            else:
                try:
                    group_size = int(group_size_str)
                    if group_size <= 0:
                        group_size = 161  # 如果输入无效（如0或负数），也视为画一张图
                except ValueError:
                    messagebox.showwarning("Warning",
                                           "Invalid 'Channels per Graph' value. Plotting all channels in one graph.")
                    group_size = 161

            valid_channels_for_report = sorted(sliced_data['history'].keys())
            # 使用新的 group_size
            for i in range(0, len(valid_channels_for_report), group_size):
                channel_group = valid_channels_for_report[i: i + group_size]

                # 动态生成标题
                if len(channel_group) == 1:
                    group_title = f"Channel: {channel_group[0] + 1}"
                else:
                    group_title = f"Channels: {channel_group[0] + 1} to {channel_group[-1] + 1}"

                running_frame.redraw_historical_plot(
                    channels_to_plot=channel_group,
                    title=group_title,
                    start_time=self.report_time_range.get('start'),
                    end_time=self.report_time_range.get('end'),
                    y_min=running_frame.y_min_entry.get(),
                    y_max=running_frame.y_max_entry.get()
                )
                group_path = f"temp_report_group_{i}.png"
                self.frames['RunningFrame'].fig.savefig(group_path, dpi=300)
                temp_image_files.append(group_path)
                plot_data_for_report.append({'title': group_title, 'path': group_path})

            report_data = self.settings.copy()
            report_data['Phenomena And Result'] = self.report_notes.get('phenomena', '')
            report_data['Notes'] = self.report_notes.get('notes', '')
            report_data['Start time'] = time.asctime(
                time.localtime(self.start_timestamp + float(self.report_time_range.get('start'))))
            report_data['Stop time'] = time.asctime(
                time.localtime((self.start_timestamp + float(self.report_time_range.get('end')))))
            if self.ambient_channel is not None:
                report_data['Ambient Channel'] = str(self.ambient_channel + 1)
            else:
                report_data['Ambient Channel'] = "N/A"
            report_data['Ambient Temp Start'] = self.ambient_start_temp
            report_data['Ambient Temp Stop'] = self.ambient_end_temp
            table_data = []
            sliced_max_temps = sliced_data['max_temps']
            for i, ch_index in enumerate(valid_channels_for_report):
                config = self.channel_configs[ch_index]
                max_temp = sliced_max_temps[ch_index]
                threshold_str = config['threshold'].strip()
                if threshold_str == '':
                    status = "N/A"
                else:
                    try:
                        threshold_val = float(threshold_str)
                        status = "F" if max_temp > threshold_val else "P"
                    except ValueError:
                        status = "N/A"
                table_data.append(
                    [str(i + 1), config['location'], str(ch_index + 1),
                     f"{max_temp:.2f}", config['threshold'], status])
            report_data['test_data'] = table_data
            success_pdf = generate_pdf_report(filepath_pdf, report_data, plot_data_for_report)
            success_excel = False
            try:
                headers, data_rows = self.get_formatted_excel_data(channels_for_report,
                                                                   self.report_time_range.get('start'),
                                                                   self.report_time_range.get('end'))
                if headers and data_rows:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Temperature Data"
                    ws.append(headers)
                    for row_data in data_rows: ws.append(row_data)
                    wb.save(filepath_excel)
                    success_excel = True
            except Exception as e:
                print(f"Failed to save Excel data: {e}")
            if success_pdf and success_excel:
                messagebox.showinfo("Success", f"Report and data saved to:\n{filepath_pdf}\n{filepath_excel}")
            elif success_pdf:
                messagebox.showwarning("Partly success", f"Report saved, data saving error.\nPDF: {filepath_pdf}")
            else:
                messagebox.showerror("Failed", "An error occurred.")
        finally:
            for img_path in temp_image_files:
                if os.path.exists(img_path): os.remove(img_path)
            running_frame.redraw_historical_plot()

    def show_frame(self, page_name):
        self.frames[page_name].tkraise()

    def update_channel_config(self, channel_index, field_index, value):
        key_map = {1: 'location', 4: 'threshold'}
        key = key_map.get(field_index)
        if key and 0 <= channel_index < len(self.channel_configs): self.channel_configs[channel_index][key] = value

    def disconnect_instrument(self):
        if self.instrument: self.instrument.close(); self.instrument = None

    def get_device_id(self):
        if self.instrument and self.instrument.connected: return self.instrument.query("*IDN?")
        return "N/A"

    def start_data_acquisition(self, interval, ambient_channel_str):
        self.is_running = True
        self.stop_thread.clear()
        self.max_temps.fill(-np.inf)

        # 清除历史数据
        for i in range(160):
            self.history[i].clear()

        self.start_time = datetime.now()
        self.start_timestamp = time.time()
        self.stop_time = None

        # 设置环境通道
        try:
            # 用户输入的是通道号（1-160），转换为索引（0-159）
            self.ambient_channel = int(ambient_channel_str) - 1
        except ValueError:
            self.ambient_channel = None
        print(self.ambient_channel)
        # 重置环境温度记录
        self.ambient_start_temp = "N/A"
        self.ambient_end_temp = "N/A"

        # 启动数据采集线程
        self.data_thread = threading.Thread(
            target=self._data_acquisition_loop,
            args=(interval,),
            daemon=True
        )
        self.data_thread.start()

        # 初始化仪器
        self.instrument.init_temperature_scan()
        time.sleep(0.1)
        self.init = True

    def stop_data_acquisition(self):
        self.is_running = False
        self.stop_thread.set()
        self.stop_time = datetime.now()
        self.stop_timestamp = time.time()
        self.init = False

        # 测试停止后，检查数据量并提供降采样选项
        # 寻找一个有数据的通道来判断数据点数量
        first_active_channel = next((ch for ch, data in self.history.items() if data), None)
        if first_active_channel is not None:
            num_points = len(self.history[first_active_channel])
            # 当数据点过多时（例如超过5000个），提示用户
            if num_points > 5000:
                user_choice = messagebox.askyesno(
                    "High Data Volume",
                    f"The test recorded {num_points} data points, which may slow down plotting and reporting.\n\n"
                    "Do you want to downsample the data (keep every 2nd point) to improve performance?",
                    detail="This will not affect the Max Temp values."
                )
                if user_choice:
                    print("Performing downsampling...")
                    for i in range(160):
                        if self.history[i]:
                            original_deque = self.history[i]
                            # 每两个点保留一个
                            self.history[i] = deque(list(original_deque)[::2])
                    print("Downsampling complete.")

    def _data_acquisition_loop(self, interval):
        while not self.stop_thread.is_set():
            if self.instrument and self.instrument.connected and self.init == True:
                try:
                    read_time = time.time()
                    raw_data = self.instrument.get_data('READ?')
                    if raw_data:
                        temps_80ch = np.array([t for t in raw_data])
                        if self.instrument.opt == "@101:140,201:240":
                            temps_160ch = np.full(160, np.nan)
                            temps_160ch[self.channel_offset: self.channel_offset + 80] = temps_80ch
                            self.data_queue.put((read_time, temps_160ch))
                        elif self.instrument.opt == "@101:140":
                            temps_160ch = np.full(160, np.nan)
                            temps_160ch[self.channel_offset: self.channel_offset + 40] = temps_80ch
                            self.data_queue.put((read_time, temps_160ch))
                        elif self.instrument.opt == "@201:240":
                            temps_160ch = np.full(160, np.nan)
                            temps_160ch[self.channel_offset + 40: self.channel_offset + 80] = temps_80ch
                            self.data_queue.put((read_time, temps_160ch))
                        #if len(temps_80ch) == 80:
                        #    temps_160ch = np.full(160, np.nan)
                        #    temps_160ch[self.channel_offset: self.channel_offset + 80] = temps_80ch
                        #    self.data_queue.put((read_time, temps_160ch))
                except Exception as e:
                    print(f"数据读取错误: {e}")
            time.sleep(interval)

    def process_queue(self):
        try:
            if not self.data_queue.empty():
                read_time, temps = self.data_queue.get_nowait()
                valid_temps_mask = ~np.isnan(temps)
                self.max_temps[valid_temps_mask] = np.maximum(self.max_temps[valid_temps_mask],
                                                              temps[valid_temps_mask])
                for i in range(160):
                    # 将环境温度的记录逻辑，移动到有效数据的判断之内
                    if not np.isnan(temps[i]):
                        self.history[i].append((read_time, temps[i]))

                        # 只有当 temps[i] 是一个有效数字时，才执行这里的逻辑
                        if self.ambient_channel is not None and i == self.ambient_channel:
                            # 如果是环境通道且是第一个有效数据点，记录开始温度
                            if self.ambient_start_temp == "N/A":
                                print('1')
                                self.ambient_start_temp = f"{temps[i]:.2f}"
                                print(self.ambient_start_temp)
                            # 总是更新结束温度为最后一个有效值
                            print('2')
                            self.ambient_end_temp = f"{temps[i]:.2f}"
                            print(self.ambient_end_temp)

                running_frame = self.frames["RunningFrame"]
                if self.is_running:
                    running_frame.update_ui(temps, self.max_temps)
        finally:
            self.after(200, self.process_queue)

    def on_closing(self):
        self.stop_thread.set()
        if self.instrument: self.instrument.close()
        self.destroy()

    def get_channel_configs(self):
        return self.channel_configs

    def get_history(self):
        return self.history

    def get_start_timestamp(self):
        return self.start_timestamp

    def prepare_for_report(self, notes, channels_str, time_range):
        self.report_notes = notes
        self.report_channels_str = channels_str
        self.report_time_range = time_range
        """
        ambient_channel_str = self.settings.get('Ambient Channel', '').strip()
        if ambient_channel_str:
            try:
                # 输入的是通道号（1-160），转换为索引（0-159）
                self.ambient_channel = int(ambient_channel_str) - 1
            except ValueError:
                self.ambient_channel = None
        else:
            self.ambient_channel = None
        """

    def find_closest_timestamp_index(self, timestamps, target_ts):
        return np.argmin(np.abs(np.array(timestamps) - target_ts))

    def get_sliced_data(self, channels_to_slice, start_str, end_str):
        if self.start_timestamp == 0: return None
        history = self.get_history()

        # 寻找一个在所选通道中且有数据的通道，以确定完整时间轴
        first_channel_with_data = next((ch for ch in channels_to_slice if ch in history and history[ch]), None)

        # 如果所有选中的通道都没有数据，则直接返回None
        if first_channel_with_data is None:
            # 进一步检查：也许是所有通道都没有数据
            any_data_exists = any(data for data in history.values())
            if not any_data_exists:
                print("No data recorded at all.")
                return None
            # 如果有数据但不在选中通道里，仍然需要一个时间轴来处理空时间范围的情况
            # 随便找一个有数据的通道
            fallback_channel = next(ch for ch, data in history.items() if data)
            full_timestamps = [ts for ts, temp in history[fallback_channel]]
        else:
            full_timestamps = [ts for ts, temp in history[first_channel_with_data]]

        # 如果即使回退也找不到任何时间戳（例如测试刚开始就停止）
        if not full_timestamps:
            return None

        try:
            # 如果输入框为空，则使用完整范围
            start_offset = float(start_str) if start_str and start_str.strip() else 0
            end_offset = float(end_str) if end_str and end_str.strip() else (full_timestamps[-1] - self.start_timestamp)
        except ValueError:
            messagebox.showerror("Error", "Invalid time range. Please enter numbers only.")
            return None

        target_start_ts = self.start_timestamp + start_offset
        target_end_ts = self.start_timestamp + end_offset

        start_idx = self.find_closest_timestamp_index(full_timestamps, target_start_ts)
        end_idx = self.find_closest_timestamp_index(full_timestamps, target_end_ts)

        if start_idx > end_idx: start_idx, end_idx = end_idx, start_idx

        actual_slice_start_ts = full_timestamps[start_idx]
        sliced_history = {}
        sliced_max_temps = np.full(160, -np.inf)

        for ch in channels_to_slice:
            if ch in history and history[ch]:
                # 对该通道的完整历史进行切片
                sliced_history[ch] = list(history[ch])[start_idx: end_idx + 1]
                if sliced_history[ch]:
                    temps_in_slice = [temp for ts, temp in sliced_history[ch]]
                    sliced_max_temps[ch] = max(temps_in_slice)

        return {'history': sliced_history, 'max_temps': sliced_max_temps, 'actual_start_ts': actual_slice_start_ts}

    def get_formatted_excel_data(self, channels_to_export, start_str, end_str):
        sliced_data = self.get_sliced_data(channels_to_export, start_str, end_str)
        if not sliced_data or not sliced_data['history']: return None, None
        valid_channels = sorted(sliced_data['history'].keys())
        headers = ["Date", "Time (s)"] + [f"Channel {i + 1}" for i in valid_channels]
        first_valid_channel_history = sliced_data['history'][valid_channels[0]]
        slice_start_ts = sliced_data['actual_start_ts']
        data_map = {
            ts: [datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S'), f"{ts - slice_start_ts:.2f}"] + [None] * len(
                valid_channels) for ts, temp in first_valid_channel_history}
        for i, ch_index in enumerate(valid_channels):
            col_index = i + 2
            for timestamp, temperature in sliced_data['history'][ch_index]:
                if timestamp in data_map: data_map[timestamp][col_index] = f"{temperature:.4f}"
        sorted_timestamps = sorted(data_map.keys())
        data_rows = [data_map[ts] for ts in sorted_timestamps]
        return headers, data_rows


if __name__ == "__main__":
    app = ThermoApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()