#VM7000/PW3335 Data Collection 0_5_0
#-------------------------------------------------------------------------------
#VM7000 info: ohkura VM7000A Paperless Recorder
#       document : WXPVM70mnA0002E March, 2014(Rev.5) 
#PW3335 info : GW Instek PW3335 Programmable DC Power Meter
#       document : PW_Communicator_zh / 2018 年1月出版 (改定1.60版)
#-------------------------------------------------------------------------------
#Rev 0_1 2025/3/19 紀錄VM7000與PW3335數據
#Rev 0_2 2025/3/20 增加設備紀錄開關選項
#Rev 0_2_3 2025/3/24 6個devices各自獨立有開始,停止紀錄的功能
#Rev 0_3_7 2025/3/25 改為單獨執行一個工位記錄,但可監看圖形化介面
#Rev 0_3_8 2025/3/26 增加計算區間平均溫度功能,以功能鍵[ctrl]+上下鍵調整日期與時間
#Rev 0_4_0 2025/3/27 圖表改為內崁canvas
#Rev 0_5_0 2025/4/9 增加多工位的功能,每個工位獨立顯示與紀錄
#-------------------------------------------------------------------------------


import socket
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox  # 修正：添加 messagebox 的導入
import csv
from datetime import datetime, timedelta  # 修正：添加 timedelta 的導入
import pandas as pd  # 修正：添加 pandas 的導入
import threading
import os,sys
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from matplotlib import rcParams
from matplotlib.ticker import FuncFormatter
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
from decimal import Decimal


# 設定 matplotlib 使用的字體
rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 使用微軟正黑體
rcParams['axes.unicode_minus'] = False  # 解決負號無法顯示的問題

class VM7000:
    def __init__(self, ip_address, port=502):
        self.ip_address = ip_address
        self.port = port
        self.sock = None

    def connect(self):
        """Establish a TCP connection to the device."""
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.connect((self.ip_address, self.port))

    def disconnect(self):
        """Close the TCP connection."""
        if self.sock:
            self.sock.close()
            self.sock = None

    def send_command(self, command):
        """Send a command to the device and return the response."""
        if not self.sock:
            raise ConnectionError("Socket is not connected to the device.")
        self.sock.sendall(command)
        time.sleep(0.1)
        response = self.sock.recv(1024)
        return response

    def get_value(self, n_addr, n_func, s_bit_pos, s_count):
        """Construct and send a command to retrieve data."""
        data_length = 6  # Modbus TCP header固定長度
        command = bytearray(12)
        command[0:2] = (0).to_bytes(2, byteorder='big')  # Transaction ID
        command[2:4] = (0).to_bytes(2, byteorder='big')  # Protocol ID
        command[4:6] = (data_length).to_bytes(2, byteorder='big')  # 設定長度
        command[6] = n_addr
        command[7] = n_func
        command[8] = int(s_bit_pos[:2], 16)
        command[9] = int(s_bit_pos[2:], 16)
        command[10] = int(s_count[:2], 16)
        command[11] = int(s_count[2:], 16)
        return self.send_command(command)

    def hex_to_decimal(self, response_bytes):
        """Convert a byte response to a list of decimal values."""
        if len(response_bytes) < 9:
            raise ValueError("Invalid response length: too short")

        data_bytes = response_bytes[9:]  # 忽略 Modbus TCP Header

        decimal_values = []
        for i in range(0, len(data_bytes), 2):
            decimal_values.append((data_bytes[i] << 8) | data_bytes[i + 1])
        
        return decimal_values
        
    def decode_temperature(self, response_bytes):
        """Decode response data into temperatures."""
        if len(response_bytes) < 9:
            raise ValueError("Invalid response length: too short")

        # **刪除 Modbus TCP Header (前 9 Bytes)**
        data_bytes = response_bytes[9:]  

        if len(data_bytes) % 2 != 0:
            raise ValueError("Invalid data length: must be even")

        temperatures = []
        for i in range(0, len(data_bytes), 2):
            raw_value = (data_bytes[i] << 8) | data_bytes[i+1]
            
            # 轉換為有號 16-bit 整數
            if raw_value >= 0x8000:
                raw_value -= 0x10000

            temperatures.append(raw_value / 10.0)  # 1 unit = 0.1°C
        
        return temperatures

class PW3335:
    def __init__(self, ip_address, port=3300):
        self.ip_address = ip_address
        self.port = port
        self.sock = None

    def connect(self):
        """Establish a TCP connection to the power meter."""
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.connect((self.ip_address, self.port))

    def disconnect(self):
        """Close the TCP connection."""
        if self.sock:
            self.sock.close()
            self.sock = None

    def query_data(self):
        """Query voltage, current, power, and accumulated power."""
        if not self.sock:
            raise ConnectionError("Socket is not connected to the power meter.")
        self.sock.sendall(b':MEAS? U,I,P,WH\n')
        response = self.sock.recv(1024).decode('ascii').strip()
        try:
            # Parse the response format: "U +110.14E+0;I +0.0000E+0;P +000.00E+0;WP +00.0000E+0"
            data = response.split(';')
            if len(data) == 4:
                parsed_data = [float(item.split(' ')[-1].replace('E+0', '')) for item in data]
                return parsed_data
            else:
                raise ValueError(f"Unexpected response format: {response}")
        except Exception as e:
            print(f"Error parsing response: {response}, Exception: {e}")
            raise ValueError(f"Failed to parse response: {response}")

class App:
    def __init__(self, root):
        self.root = root
        self.file_path = ""  # 初始化 file_path 屬性
        self.collecting = {}  # 初始化 collecting 屬性，用於跟蹤正在收集數據的設備
        self.vm7000_instances = {}
        self.pw3335_instances = {}
        self.time_data = []
        self.temperature_data = []
        self.power_data = []
        self.pause_plot = False  # 新增變數，用於控制圖表更新的暫停/恢復
        self.original_text = ""

        # 為每個工位創建獨立的數據存儲
        self.station_data = {
            f"工位{i}": {
                "time_data": [],
                "temperature_data": [],
                "power_data": [],
            }
            for i in range(1, 7)
        }
        
        # 初始化 Notebook（頁面容器）
        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, columnspan=12, padx=5, pady=5)

        # 創建 6 個頁面（工位 1 到工位 6）
        self.frames = {}
        for i in range(1, 7):
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=f" 工位{i} ", padding=5)
            self.frames[f"工位{i}"] = frame
            # 在每個頁面中添加控件
            self.setup_station_page(frame, f"工位{i}")

        # 綁定窗口關閉事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)


    def setup_station_page(self, frame, station_name):
        """設置每個工位頁面的控件"""
        station_name = station_name.replace(" ", "")  # 去除空格，統一名稱格式

        # 初始化日期和時間變數
        start_date = tk.StringVar()
        start_time = tk.StringVar()
        end_date = tk.StringVar()
        end_time = tk.StringVar()

        # 設置初始值
        now = datetime.now()
        start_date.set(now.strftime('%Y-%m-%d'))
        start_time.set(now.strftime('%H:%M'))
        end_date.set(now.strftime('%Y-%m-%d'))
        end_time.set(now.strftime('%H:%M'))

        # File path selection
        ttk.Label(frame, text="儲存路徑:").grid(row=0, column=0, padx=5, pady=5)
        file_path_var = tk.StringVar()
        file_path_entry = ttk.Entry(frame, textvariable=file_path_var, width=30)
        file_path_entry.grid(row=0, column=1, padx=5, pady=5)
        browse_button = ttk.Button(frame, text="Browse", command=lambda: self.browse_file(file_path_var))
        browse_button.grid(row=0, column=2, padx=5, pady=5)

        # Frequency selection
        ttk.Label(frame, text="記錄頻率(sec):").grid(row=1, column=0, padx=5, pady=5)
        frequency_var = tk.IntVar(value=3)  # 預設為 60 秒, 3秒用來debug
        frequency_menu = ttk.Combobox(frame, textvariable=frequency_var, state="readonly")
        frequency_menu['values'] = [60, 180, 300]
        frequency_menu.grid(row=1, column=1, padx=5, pady=5)

        # VM7000 頻道設定文字框
        ttk.Label(frame, text="溫度CH 設定:").grid(row=2, column=0, padx=5, pady=5)
        vm7000_channels_var = tk.StringVar(value="1-3")
        vm7000_channels_entry = ttk.Entry(frame, textvariable=vm7000_channels_var, width=20)
        vm7000_channels_entry.grid(row=2, column=1, padx=5, pady=5)

        # X 軸範圍選擇下拉式選單
        ttk.Label(frame, text="時間區間:").grid(row=3, column=0, padx=5, pady=5)
        x_axis_range_var = tk.StringVar(value="30min")
        x_axis_range_menu = ttk.Combobox(frame, textvariable=x_axis_range_var, state="readonly")
        x_axis_range_menu['values'] = ["30min", "3hrs", "12hrs", "24hrs"]
        x_axis_range_menu.grid(row=3, column=1, padx=5, pady=5)

        # Start, Stop, and Pause/Resume buttons
        start_button = ttk.Button(frame, text="Start", command=lambda: self.start_collection(station_name), state="normal")
        start_button.grid(row=1, column=2, padx=5, pady=5)
        stop_button = ttk.Button(frame, text="Stop", command=lambda: self.stop_collection(station_name), state="disabled")
        stop_button.grid(row=2, column=2, padx=5, pady=5)
        pause_button = ttk.Button(frame, text="暫停", command=lambda: self.toggle_pause_plot(station_name), state="disabled")
        pause_button.grid(row=3, column=2, padx=5, pady=5)

        # Temperature data display
        ttk.Label(frame, text="溫度:").grid(row=1, column=3,columnspan=9, padx=5, pady=5)
        temperature_labels = []
        for i in range(18):
            label = ttk.Label(frame, text="--", width=5, relief="solid", anchor="center")
            label.grid(row=2 + i // 9, column=(i % 9)+3, padx=2, pady=2)
            temperature_labels.append(label)

        # 開始日期與時間
        calculate_avg_button = ttk.Button(frame, text="計算區間", command=self.calculate_avg_temp)
        calculate_avg_button.grid(row=4, column=2, padx=5, pady=5)
        def increment_date_time(var, increment, unit):
            try:
                current_value = pd.to_datetime(var.get())
                if unit == "day":
                    new_value = current_value + pd.Timedelta(days=increment)
                elif unit == "hour":
                    new_value = current_value + pd.Timedelta(hours=increment)
                var.set(new_value.strftime('%Y-%m-%d' if unit == "day" else '%H:%M'))
            except Exception:
                messagebox.showerror("錯誤", "無效的日期或時間格式！")

        def bind_increment(widget, var, unit):
            def on_key(event):
                if event.state & 0x4:  # 檢查是否按下 CTRL 鍵
                    if event.keysym == "Up":
                        increment_date_time(var, 1, unit)
                    elif event.keysym == "Down":
                        increment_date_time(var, -1, unit)
            widget.bind("<KeyPress-Up>", on_key)
            widget.bind("<KeyPress-Down>", on_key)

        start_date_entry = ttk.Entry(frame, textvariable=start_date, width=10)
        start_date_entry.grid(row=4, column=0, padx=5, pady=5, sticky="w")
        bind_increment(start_date_entry, start_date, "day")

        start_time_entry = ttk.Entry(frame, textvariable=start_time, width=10)
        start_time_entry.grid(row=5, column=0, padx=5, pady=5, sticky="w")
        bind_increment(start_time_entry, start_time, "hour")

        end_date_entry = ttk.Entry(frame, textvariable=end_date, width=10)
        end_date_entry.grid(row=6, column=0, padx=5, pady=5, sticky="w")
        bind_increment(end_date_entry, end_date, "day")

        end_time_entry = ttk.Entry(frame, textvariable=end_time, width=10)
        end_time_entry.grid(row=7, column=0, padx=5, pady=5, sticky="w")
        bind_increment(end_time_entry, end_time, "hour")

        # Multi-line text box for displaying results
        avg_temp_text = tk.Text(frame, height=5, width=50, state="disabled", wrap="word")
        avg_temp_text.grid(row=4, rowspan= 4, column=1, columnspan=2, padx=5, pady=5)

        # Add a canvas to embed the Matplotlib figure
        figure = plt.Figure(figsize=(10, 4), dpi=100)
        canvas = FigureCanvasTkAgg(figure, master=frame)  # 將 canvas 綁定到當前 frame
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.grid(row=8, column=0, columnspan=12, padx=5, pady=5)

        # Create a frame for the toolbar
        toolbar_frame = tk.Frame(frame)
        toolbar_frame.grid(row=7, column=3, columnspan=7, padx=5, pady=5)

        # Add the Navigation Toolbar to the frame
        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
        toolbar.update()

        # Create subplots for temperature and power
        ax_temp = figure.add_subplot(211, facecolor='lightgray')
        ax_power = figure.add_subplot(212, sharex=ax_temp, facecolor='lightgray')

        #ax_temp.set_title("Temperature Data")
        ax_temp.set_ylabel("Temperature (°C)")
        ax_temp.grid(True)

        #ax_power.set_title("Power Data")
        ax_power.set_xlabel("Time")
        ax_power.set_ylabel("Power (W)")
        ax_power.grid(True)

        # Initialize lines for temperature and power
        self.temp_lines = []
        self.power_line, = ax_power.plot([], [], label="Power (W)", color="orange")
        #ax_power.legend()

        # 保存控件到工位的屬性中
        setattr(self, f"{station_name}_figure", figure)
        setattr(self, f"{station_name}_canvas", canvas)
        setattr(self, f"{station_name}_ax_temp", ax_temp)
        setattr(self, f"{station_name}_ax_power", ax_power)
        setattr(self, f"{station_name}_start_button", start_button)
        setattr(self, f"{station_name}_stop_button", stop_button)
        setattr(self, f"{station_name}_pause_button", pause_button)
        setattr(self, f"{station_name}_Browse_button", browse_button)
        setattr(self, f"{station_name}_temperature_labels", temperature_labels)
        setattr(self, f"{station_name}_file_path_var", file_path_var)
        setattr(self, f"{station_name}_file_path_entry", file_path_entry)
        setattr(self, f"{station_name}_vm7000_channels_var", vm7000_channels_var)
        setattr(self, f"{station_name}_vm7000_channels_entry", vm7000_channels_entry)
        setattr(self, f"{station_name}_frequency_var", frequency_var)
        setattr(self, f"{station_name}_frequency_menu", frequency_menu)
        setattr(self, f"{station_name}_x_axis_range_var", x_axis_range_var)
        setattr(self, f"{station_name}_start_date_entry", start_date_entry)
        setattr(self, f"{station_name}_start_time_entry", start_time_entry)
        setattr(self, f"{station_name}_end_date_entry", end_date_entry)
        setattr(self, f"{station_name}_end_time_entry", end_time_entry)
        setattr(self, f"{station_name}_avg_temp_text", avg_temp_text)
        
        # 初始化 collecting 狀態
        self.collecting[station_name] = False
        # Update the canvas periodically
        self.update_canvas(station_name)

    def calculate_avg_temp(self):
        """計算指定時間範圍內的平均溫度，基於圖表數據"""
        try:
            # 獲取當前選中的工位名稱
            selected_station = self.notebook.tab(self.notebook.select(), "text").replace("🔴", "").strip()

            # 動態獲取對應工位的日期和時間輸入框
            start_date_entry = getattr(self, f"{selected_station}_start_date_entry", None)
            start_time_entry = getattr(self, f"{selected_station}_start_time_entry", None)
            end_date_entry = getattr(self, f"{selected_station}_end_date_entry", None)
            end_time_entry = getattr(self, f"{selected_station}_end_time_entry", None)
            avg_temp_text = getattr(self, f"{selected_station}_avg_temp_text", None)

            if not all([start_date_entry, start_time_entry, end_date_entry, end_time_entry, avg_temp_text]):
                raise AttributeError(f"One or more required widgets for {selected_station} are not defined.")

            # 獲取開始和結束時間
            start_datetime = pd.to_datetime(f"{start_date_entry.get()} {start_time_entry.get()}")
            end_datetime = pd.to_datetime(f"{end_date_entry.get()} {end_time_entry.get()}")

            if start_datetime >= end_datetime:
                tk.messagebox.showerror("Error", f"開始時間: {start_datetime}, 結束時間: {end_datetime} \n開始時間必須早於結束時間")
                return

            # 篩選在指定範圍內的溫度數據
            station_data = self.station_data[selected_station]
            filtered_temps = [
                temps for time, temps in zip(station_data["time_data"], station_data["temperature_data"])
                if start_datetime <= time <= end_datetime
            ]

            if not filtered_temps:
                tk.messagebox.showinfo("Info", "指定範圍內沒有溫度數據")
                return

            # 計算每個頻道的平均溫度
            avg_temps = []
            for i in range(len(filtered_temps[0])):
                channel_temps = [temps[i] for temps in filtered_temps if temps[i] is not None]
                if channel_temps:
                    avg_temps.append(sum(channel_temps) / len(channel_temps))
                else:
                    avg_temps.append(None)

            # 顯示結果
            avg_temp_text.config(state="normal")
            avg_temp_text.delete("1.0", tk.END)
            for i in range(0, len(avg_temps), 3):
                line = "".join(
                    f"CH{j + 1:02}: {avg_temps[j]:>4.1f}°C".ljust(15) if avg_temps[j] is not None else f"CH{j + 1:02}: {'--':>4}".ljust(15)
                    for j in range(i, min(i + 3, len(avg_temps)))
                )
                avg_temp_text.insert(tk.END, line + "\n")
            avg_temp_text.config(state="disabled")

        except Exception as e:
            tk.messagebox.showerror("Error", f"計算平均溫度時發生錯誤: {e}")

    def update_temperature_display(self, station_name, temperatures):
        """更新溫度數據顯示"""
        temperature_labels = getattr(self, f"{station_name}_temperature_labels", [])
        for i, label in enumerate(temperature_labels):
            if i < len(temperatures) and temperatures[i] is not None:
                label.config(text=f"{temperatures[i]:.1f}")  # 格式化為小數點後一位
                if temperatures[i] > 999:
                    label.config(text="--")
            else:
                label.config(text="--")  # 如果數據為 None 或超出範圍，顯示占位符


    def toggle_pause_plot(self, station_name):
        """暫停或恢復圖表更新"""
        self.pause_plot = not self.pause_plot
        pause_button = getattr(self, f"{station_name}_pause_button", None)

        if self.pause_plot:
            if pause_button:
                pause_button.config(text="繼續更新")

            # 填入目前 X 軸的資料到 start_date, start_time, end_date, end_time
            x_start, x_end = self.get_x_axis_range(station_name)
            start_date_entry = getattr(self, f"{station_name}_start_date_entry", None)
            start_time_entry = getattr(self, f"{station_name}_start_time_entry", None)
            end_date_entry = getattr(self, f"{station_name}_end_date_entry", None)
            end_time_entry = getattr(self, f"{station_name}_end_time_entry", None)

            if start_date_entry and start_time_entry and end_date_entry and end_time_entry:
                start_date_entry.delete(0, tk.END)
                start_date_entry.insert(0, x_start.strftime('%Y-%m-%d'))
                start_time_entry.delete(0, tk.END)
                start_time_entry.insert(0, x_start.strftime('%H:%M'))
                end_date_entry.delete(0, tk.END)
                end_date_entry.insert(0, x_end.strftime('%Y-%m-%d'))
                end_time_entry.delete(0, tk.END)
                end_time_entry.insert(0, x_end.strftime('%H:%M'))
        else:
            if pause_button:
                pause_button.config(text="暫停")

    def browse_file(self, file_path_var):
        file_path = filedialog.askdirectory()
        file_path_var.set(file_path)
        self.file_path = file_path  # 將選擇的路徑保存到 self.file_path

    def parse_channels(self, channel_str: str) -> list[int]:
        """Parse channel configuration string, supporting ranges and comma-separated values."""
        channels = set()
        parts = channel_str.split(",")
        for part in parts:
            if "-" in part:
                try:
                    start, end = map(int, part.split("-"))
                    if start > end:
                        raise ValueError(f"Invalid range: {part}")
                    channels.update(range(start, end + 1))
                except ValueError:
                    raise ValueError(f"Invalid range format: {part}")
            else:
                try:
                    channels.add(int(part))
                except ValueError:
                    raise ValueError(f"Invalid channel number: {part}")
        return sorted(channels)

    def start_collection(self, station_name):
        """啟動數據收集"""

        if not self.file_path:
            tk.messagebox.showerror("Error", "先選擇儲存路徑")
            return

        # 獲取當前工位的 IP 地址
        station_index = int(station_name.replace("工位", "")) - 1
        vm_ip = f"192.168.1.{station_index + 1}"
        pw_ip = f"192.168.1.{station_index + 7}"

        # 解析頻道設定
        try:
            channels = self.parse_channels(getattr(self, f"{station_name}_vm7000_channels_var").get())
        except ValueError:
            tk.messagebox.showerror("Error", "頻道設定格式錯誤，請使用 '1-3' 或 '1,2,3'")
            return

        try:
            self.vm7000_instances[vm_ip] = VM7000(vm_ip)
            self.vm7000_instances[vm_ip].connect()
            self.pw3335_instances[pw_ip] = PW3335(pw_ip)
            self.pw3335_instances[pw_ip].connect()
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to connect to devices: {e}")
            return

        self.collecting[vm_ip] = True

        # 禁用其他控件
        getattr(self, f"{station_name}_start_button").config(state="disabled")
        getattr(self, f"{station_name}_stop_button").config(state="normal")
        getattr(self, f"{station_name}_pause_button").config(state="normal")  # 啟用暫停按鈕
        getattr(self, f"{station_name}_Browse_button").config(state="disabled")
        getattr(self, f"{station_name}_frequency_menu").config(state="disabled")
        getattr(self, f"{station_name}_vm7000_channels_entry").config(state="disabled")
        getattr(self, f"{station_name}_file_path_entry").config(state="disabled")

        # 啟動數據收集
        threading.Thread(target=self.collect_data, args=(vm_ip, pw_ip, channels, station_name), daemon=True).start()

        # 在主執行緒中啟動即時監看圖表
        self.show_live_plot(station_name)

        # 啟動資料收集時，設定該工位頁籤加上紅點
        tab_index = int(station_name.replace("工位", "")) - 1
        original_text = self.notebook.tab(tab_index, option="text")
        if "🔴" not in original_text:
            self.notebook.tab(tab_index, text=f"🔴工位{tab_index + 1}")


    def stop_collection(self, station_name):
        """停止指定工位的數據收集"""
        # 獲取對應工位的 IP 地址
        station_index = int(station_name.replace("工位", "")) - 1
        vm_ip = f"192.168.1.{station_index + 1}"
        pw_ip = f"192.168.1.{station_index + 7}"

        # 停止該工位的數據收集
        if vm_ip in self.collecting:
            self.collecting[vm_ip] = False
            if vm_ip in self.vm7000_instances:
                self.vm7000_instances[vm_ip].disconnect()
                del self.vm7000_instances[vm_ip]
            if pw_ip in self.pw3335_instances:
                self.pw3335_instances[pw_ip].disconnect()
                del self.pw3335_instances[pw_ip]

        # 清除該工位的圖表資料
        station_data = self.station_data[station_name]
        station_data["time_data"].clear()
        station_data["temperature_data"].clear()
        station_data["power_data"].clear()

        # 清空圖表
        ax_temp = getattr(self, f"{station_name}_ax_temp", None)
        ax_power = getattr(self, f"{station_name}_ax_power", None)
        if ax_temp and ax_power:
            ax_temp.clear()
            ax_power.clear()

            # 重設圖表標題與軸標籤
            ax_temp.set_title("Temperature Data")
            ax_temp.set_ylabel("Temperature (°C)")
            ax_temp.grid(True)

            ax_power.set_title("Power Data")
            ax_power.set_xlabel("Time")
            ax_power.set_ylabel("Power (W)")
            ax_power.grid(True)

            # 更新畫布
            canvas = getattr(self, f"{station_name}_canvas", None)
            if canvas:
                canvas.draw()

        # 啟用其他控件
        getattr(self, f"{station_name}_start_button").config(state="normal")
        getattr(self, f"{station_name}_stop_button").config(state="disabled")
        getattr(self, f"{station_name}_pause_button").config(state="disabled")  # 啟用暫停按鈕
        getattr(self, f"{station_name}_Browse_button").config(state="normal")
        getattr(self, f"{station_name}_frequency_menu").config(state="readonly")
        getattr(self, f"{station_name}_vm7000_channels_entry").config(state="normal")
        getattr(self, f"{station_name}_file_path_entry").config(state="normal")

        # 停止資料收集時，將紅點移除
        tab_index = int(station_name.replace("工位", "")) - 1
        self.original_text = self.notebook.tab(tab_index, option="text")
        self.notebook.tab(tab_index, text=f"工位{tab_index + 1}")


    def collect_data(self, vm_ip, pw_ip, channels, station_name):
        """收集數據並保存到 CSV 文件"""
        vm = self.vm7000_instances.get(vm_ip)
        pw = self.pw3335_instances.get(pw_ip)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"{self.file_path}/{timestamp}_Station_{vm_ip.split('.')[-1]}.csv"

        with open(file_name, mode="a", newline="", buffering=1) as file:
            writer = csv.writer(file)
            writer.writerow(["Date", "Time"] + [f"Temp{ch}" for ch in channels] + ["U(V)", "I(A)", "P(W)", "WP(Wh)"])

            try:
                frequency_var = getattr(self, f"{station_name}_frequency_var", None)
                if not frequency_var:
                    raise AttributeError(f"Frequency variable for {station_name} is not defined.")

                while self.collecting.get(vm_ip, False):
                    temperatures = [None] * len(channels)
                    if vm:
                        try:
                            response = vm.get_value(1, 4, "0064", "0012")
                            all_temperatures = vm.decode_temperature(response)
                            temperatures = [all_temperatures[ch - 1] for ch in channels if ch <= len(all_temperatures)]
                        except Exception as e:
                            print(f"Error collecting VM7000 data for {vm_ip}: {e}")
                            tk.messagebox.showerror("Error", f"Error collecting VM7000 data for {vm_ip}: {e}")
                    power_data = [0, 0, 0]
                    if pw:
                        try:
                            power_data = pw.query_data()[:4]
                        except Exception as e:
                            print(f"Error collecting PW3335 data for {pw_ip}: {e}")
                            tk.messagebox.showerror("Error", f"Error collecting PW3335 data for {pw_ip}: {e}")
                    
                    now = datetime.now()
                    writer.writerow([now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")] + temperatures + power_data)
                    file.flush()  # 確保數據寫入磁盤

                    # 更新即時監看數據
                    station_data = self.station_data[station_name]
                    station_data["time_data"].append(now)
                    station_data["temperature_data"].append(temperatures)
                    station_data["power_data"].append(power_data[2])  # 只取功率 (P)

                    # 更新溫度數據顯示
                    self.root.after(0, self.update_temperature_display, station_name, temperatures)

                    # 保留 X 軸範圍內的數據
                    max_range_start = datetime.now() - timedelta(hours=189)
                    while station_data["time_data"] and station_data["time_data"][0] < max_range_start:
                        station_data["time_data"].pop(0)
                        station_data["temperature_data"].pop(0)
                        station_data["power_data"].pop(0)

                    time.sleep(frequency_var.get())
            except Exception as e:
                print(f"Data collection error: {e}")
                tk.messagebox.showerror("Error", f"Data collection error: {e}")
                self.stop_collection(station_name)  # 傳遞 station_name 以便停止正確的工位

    def show_live_plot(self, station_name):
        """顯示即時監看圖表"""
        def plot():
            # 獲取當前工位的 figure
            figure = getattr(self, f"{station_name}_figure", None)
            if not figure:
                raise AttributeError(f"Figure for {station_name} is not defined.")

            # 清空舊的圖表
            figure.clear()

            # 創建新的子圖
            ax_temp = figure.add_subplot(211)
            ax_power = figure.add_subplot(212, sharex=ax_temp)

            # 隱藏 ax_temp 的 x 軸,用來顯示圖例
            ax_temp.tick_params(labelbottom=False)

            figure.suptitle(f"SAMPO RD2 冰箱測試 - {station_name}")
            figure.set_facecolor("lightgray")
            ax_temp.set_facecolor("lightcyan")
            ax_power.set_facecolor("lightyellow")

            # 顯示 Y 軸格線
            ax_temp.grid(axis='y', linestyle='--', alpha=0.7)
            ax_power.grid(axis='y', linestyle='--', alpha=0.7)

            ax_temp.yaxis.set_major_locator(plt.MultipleLocator(5))  # 每隔 5 度顯示一條格線

            # 溫度子圖
            channels = self.parse_channels(getattr(self, f"{station_name}_vm7000_channels_var").get())
            temp_lines = [ax_temp.plot([], [], label=f"Temp {ch}")[0] for ch in channels]
            ax_temp.set_ylabel("Temperature (°C)")
            ax_temp.legend([f"CH{ch}" for ch in channels], loc='upper center', bbox_to_anchor=(0.5, 0), ncol=len(channels), fontsize='small')

            # 電力子圖
            power_line, = ax_power.plot([], [], label="Power (W)", color="orange")
            ax_power.set_ylabel("Power (W)")

            # 設置 Y 軸刻度格式為小數點後 1 位
            ax_power.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{x:.1f}"))

            def update(frame):
                if self.pause_plot:  # 如果圖表更新被暫停，直接返回
                    return temp_lines + [power_line]

                # 從對應工位的數據結構中讀取數據
                station_data = self.station_data[station_name]
                time_data = station_data["time_data"]
                temperature_data = station_data["temperature_data"]
                power_data = station_data["power_data"]

                # 更新溫度數據
                for i, line in enumerate(temp_lines):
                    if len(temperature_data) > 0:
                        temp_data = [temps[i] if len(temps) > i else None for temps in temperature_data]
                        line.set_data(time_data, temp_data)

                # 更新電力數據
                if len(power_data) > 0:
                    power_line.set_data(time_data, power_data)

                # 更新 X 軸範圍
                x_range = self.get_x_axis_range(station_name)
                ax_temp.set_xlim(x_range)
                ax_power.set_xlim(x_range)

                # 以最大最小值自動調整 Y 軸格線
                if len(temperature_data) > 0:
                    all_temps = [temp for temps in temperature_data for temp in temps if temp is not None]
                    if all_temps:
                        min_temp, max_temp = min(all_temps), max(all_temps)
                        #print({max_temp, min_temp})
                        if (max_temp - min_temp) > 30:
                            ax_temp.yaxis.set_major_locator(plt.MultipleLocator(10))
                            ax_temp.yaxis.set_minor_locator(plt.MultipleLocator(5))
                            ax_temp.set_ylim(min_temp - 5, max_temp + 5)
                        elif (max_temp - min_temp) < 5:
                            ax_temp.yaxis.set_major_locator(plt.MultipleLocator(1))
                            ax_temp.yaxis.set_minor_locator(plt.MultipleLocator(0.5))
                            ax_temp.set_ylim(min_temp - 1, max_temp + 1)
                        else:
                            ax_temp.yaxis.set_major_locator(plt.MultipleLocator(5))
                            ax_temp.yaxis.set_minor_locator(plt.MultipleLocator(2.5))
                            ax_temp.set_ylim(min_temp - 2, max_temp + 2)
                        

                ax_temp.relim()
                ax_temp.autoscale_view()
                ax_power.relim()
                ax_power.autoscale_view()

                return temp_lines + [power_line]

            # 使用 FuncAnimation 更新圖表
            self.ani = FuncAnimation(figure, update, interval=10000, blit=False, cache_frame_data=False)

            # 更新嵌入的 canvas
            canvas = getattr(self, f"{station_name}_canvas", None)
            if canvas:
                canvas.draw()

        # 在主執行緒中啟動 plot 函數
        self.root.after(0, plot)

    def get_x_axis_range(self, station_name):
        """根據選擇的 X 軸範圍返回時間範圍"""
        now = datetime.now()
        range_mapping = {
            "30min": timedelta(minutes=30),
            "3hrs": timedelta(hours=3),
            "12hrs": timedelta(hours=12),
            "24hrs": timedelta(hours=24),
        }
        x_axis_range_var = getattr(self, f"{station_name}_x_axis_range_var", None)
        if x_axis_range_var is None:
            raise AttributeError(f"x_axis_range_var for {station_name} is not defined.")
        selected_range = range_mapping.get(x_axis_range_var.get(), timedelta(minutes=30))
        return now - selected_range, now
    
    def update_canvas(self, station_name):
        """Update the Matplotlib canvas for the specified station."""
        canvas = getattr(self, f"{station_name}_canvas", None)
        if canvas:
            canvas.draw()
        else:
            raise ValueError(f"Canvas for {station_name} is not defined.")

    def on_closing(self):
        """檢查是否有工位正在啟動，若有則跳出警告訊息"""
        active_stations = [station for station, is_collecting in self.collecting.items() if is_collecting]
        if active_stations:
            tk.messagebox.showwarning(
                "警告", 
                f"以下工位正在收集數據，請先停止數據收集再退出程序：\n{', '.join(active_stations)}"
            )
        else:
            self.root.destroy()  # 正常退出程序


if __name__ == "__main__":
    root = tk.Tk()
    # 初始化 tk.StringVar() 變數
    start_date = tk.StringVar()
    start_time = tk.StringVar()
    end_date = tk.StringVar()
    end_time = tk.StringVar()
    root.title("VM7000/PW3335 Data Collection 0_5_0")
    app = App(root)
    root.mainloop()
