# instrument_controller.py (可选依赖处理版)

import time
import pyvisa
import pyvisa_py
import numpy as np

# --- 将 pygpib 设为可选依赖 ---
try:
    import pygpib as gpib

    GPIB_ENABLED = True
except ImportError:
    GPIB_ENABLED = False
    print("警告：'pygpib' 库未安装。GPIB连接功能将不可用。")
    print("若需使用GPIB，请运行: pip install pygpib")


# --- 仿真器 ---
class FakeKeithley2701:
    def __init__(self, address, connection_type='TCPIP'):
        self.address = address
        self.connected = False
        self._base_temps = np.random.uniform(20.0, 25.0, 80)
        self._drifting_channels = np.random.choice(80, size=5, replace=False)
        print(f"模拟设备：初始化于地址 {self.address}")

    def connect(self):
        print(f"模拟设备：正在尝试连接到 {self.address}...")
        time.sleep(1)
        self.connected = True
        print("模拟设备：连接成功。")
        return True

    def query(self, command):
        if not self.connected:
            raise ConnectionError("模拟设备未连接")
        if command == '*IDN?':
            return "KEITHLEY INSTRUMENTS INC.,MODEL 2701,DEV001,A09/A02"
        elif command.upper().startswith(('READ?', 'FETC?')):
            noise = np.random.normal(0, 0.1, 80)
            self._base_temps[self._drifting_channels] += np.random.uniform(0, 0.5)
            temperatures = self._base_temps + noise
            return ','.join([f'{temp:.4f}' for temp in temperatures])
        else:
            return ""

    def init_temperature_scan(self, thermocouple_type='K', nplc=1):
        print("模拟设备：已接收初始化指令。")
        return True

    def get_data(self, command):
        noise = np.random.normal(0, 0.1, 80)
        temperatures = self._base_temps + noise
        open_circuit_indices = np.random.choice(80, 3, replace=False)
        for i in open_circuit_indices:
            temperatures[i] = 1.0e+6
        return list(temperatures)

    def close(self):
        self.connected = False
        print("模拟设备：连接已断开。")


# --- 这是真实设备控制器 ---
class Keithley27xxController:
    def __init__(self, address, connection_type='TCPIP'):
        self.rm = pyvisa.ResourceManager('@py')
        self.instrument = None
        self.connected = False
        self.scan_list = "(@101:140,201:240)"
        self.connection_type = connection_type

        # 根据连接类型构建不同的VISA资源地址字符串
        if self.connection_type == 'TCPIP':
            self.visa_address = f"TCPIP0::{address}::1394::SOCKET"
        elif self.connection_type == 'GPIB':
            # --- 核心改动：在使用GPIB前检查库是否存在 ---
            if not GPIB_ENABLED:
                # 如果库不存在，则抛出一个明确的错误，这个错误会被 connect 方法捕获
                raise ImportError("'pygpib' library is not installed, cannot use GPIB connection.")
            self.visa_address = address
        else:
            raise ValueError(f"不支持的连接类型: {self.connection_type}")

    def connect(self):
        """
        建立与设备的连接
        """
        try:
            print(f"正在以 {self.connection_type} 模式连接到: {self.visa_address}")
            self.instrument = self.rm.open_resource(self.visa_address)

            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self.instrument.timeout = 20000
            time.sleep(0.2)

            print("连接已建立，正在验证设备身份...")
            idn = self.query('*IDN?')
            print(f"收到设备ID: {idn}")

            if 'KEITHLEY' in idn.upper():
                self.connected = True
                print("设备验证成功，连接已就绪。")
                return True
            else:
                print("设备验证失败，IDN不匹配或无响应。")
                self.close()
                return False

        # --- 捕获 ImportError ---
        except ImportError as e:
            # 这个错误只会在用户选择GPIB但未安装pygpib时触发
            print(f"连接失败：缺少必要的库。")
            # 我们可以通过 `messagebox` 在UI层面给出更友好的提示，
            # 但为了保持控制器类的独立性，这里只打印错误。
            # UI的 `connect_device` 方法会捕获所有异常并弹出通用失败窗口。
            print(f"详细错误: {e}")
            return False
        except pyvisa.errors.VisaIOError as e:
            print(f"连接失败：发生VISA I/O错误。请检查地址、物理连接和VISA后端。")
            print(f"详细错误: {e}")
            return False
        except Exception as e:
            print(f"连接失败：发生未知错误 - {e}")
            return False

    # ... (write, init_temperature_scan, query, get_data, close 方法) ...
    def write(self, command):
        if not self.connected or self.instrument is None:
            raise ConnectionError("仪器未连接，无法发送指令")
        try:
            self.instrument.write(command)
            time.sleep(0.2)
        except pyvisa.errors.VisaIOError as e:
            print(f"指令 '{command}' 写入失败: {e}")
            self.connected = False
            raise

    def init_temperature_scan(self, thermocouple_type='K', nplc=5):
        if not self.connected:
            print("必须先连接设备才能进行初始化。")
            return False
        print("\n开始配置温度扫描参数...")
        try:
            self.instrument.write('*CLS')
            self.write(f"ROUT:SCAN {self.scan_list}")
            self.write(f"SENS:TEMP:TRAN TC {self.scan_list}")
            self.write(f"SENS:TC:TYPE K {self.scan_list}")
            self.write(f"SENS:FUNC 'TEMP' {self.scan_list}")
            self.write("SENS:UNIT:TEMP C")
            self.write(f"SENS:TEMP:NPLC {nplc}")
            self.write("TRAC:CLE")
            self.write("TRIG:SOUR IMM")
            self.write("SAMP:COUN 80")
            self.write("ROUT:SCAN:TSO IMM")
            self.write("INIT:CONT OFF")
            self.write("FORM:ELEM READ")
            self.write("TRIG:COUN 1")
            self.write("SENS:TC:CJON:STATE ON")
            print("101配置:", self.query("CONF? (@101)"))
            print(f"配置完成：扫描列表 {self.scan_list}")
            print(f"测量参数：{thermocouple_type}型热电偶, 摄氏度, NPLC={nplc}")
            print("初始化成功，已准备好采集温度数据。")
            return True
        except Exception as e:
            print(f"初始化过程中发生错误: {e}")
            return False

    def query(self, command):
        if self.instrument is None:
            raise ConnectionError("仪器对象未初始化，无法查询")
        try:
            return self.instrument.query(command).strip()
        except pyvisa.errors.VisaIOError as e:
            print(f"指令 '{command}' 执行失败: {e}")
            self.connected = False
            return "Error: VISA IO Error"

    def get_data(self, command):
        if command.upper().startswith(('READ?', 'FETC?')):
            real_temp = []
            self.write("ROUT:SCAN:LSEL INT")
            temperatures = self.instrument.query_ascii_values('READ?', container=list)
            self.write("ROUT:SCAN:LSEL NONE")
            for temp in temperatures:
                if 1000000 > temp:
                    real_temp.append(float(temp))
                else:
                    real_temp.append(None)
            return real_temp
        else:
            return []

    def close(self):
        if self.instrument:
            try:
                self.instrument.close()
                print("设备连接已成功关闭。")
            except pyvisa.errors.VisaIOError:
                pass
        self.instrument = None
        self.connected = False