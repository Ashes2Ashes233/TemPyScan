# instrument_controller.py

import time
import pyvisa
import pyvisa_py
import numpy as np


# --- 仿真器代码 (保持不变) ---
class FakeKeithley2701:
    def __init__(self, address):
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
            # 根据连接的设备类型返回不同IDN
            if "GPIB" in self.address:
                return "KEITHLEY INSTRUMENTS INC.,MODEL 2700,DEV002,D07/A02"
            return "KEITHLEY INSTRUMENTS INC.,MODEL 2701,DEV001,A09/A02"
        elif command.upper().startswith(('READ?', 'FETC?')):
            import numpy as np
            noise = np.random.normal(0, 0.1, 80)
            self._base_temps[self._drifting_channels] += np.random.uniform(0, 0.5)
            temperatures = self._base_temps + noise
            return ','.join([f'{temp:.4f}' for temp in temperatures])
        else:
            return ""

    def close(self):
        self.connected = False
        print("模拟设备：连接已断开。")


# --- 重构后的真实设备控制器 (支持TCPIP和GPIB) ---
class KeithleyController:
    def __init__(self, conn_type, address):
        """
        初始化控制器.
        :param conn_type: 连接类型, 'TCPIP' 或 'GPIB'.
        :param address: IP地址或GPIB地址.
        """
        self.conn_type = conn_type.upper()
        self.address_str = address
        self.resource_string = self._build_resource_string()
        self.model = None  # 新增模型标识

        # 推荐使用NI-VISA后端来兼容GPIB和TCPIP。'@py'主要用于pyvisa-py，对GPIB支持有限。
        # 如果安装了NI-VISA，可以传入空字符串''或不传参数。
        # 为保持原代码逻辑，此处仍保留'@py'，但建议为GPIB连接安装和使用NI-VISA。
        self.rm = pyvisa.ResourceManager()  # 使用默认后端，更具兼容性
        self.instrument = None
        self.connected = False
        self.sample_count = 80
        self.scan_list = "(@101:140,201:240)"
        self.opt = "@101:140,201:240"

    def _build_resource_string(self):
        """
        根据连接类型构建PyVISA资源字符串.
        """
        if self.conn_type == 'TCPIP':
            # TCPIP连接字符串 (Keithley 2701)
            return f"TCPIP0::{self.address_str}::1394::SOCKET"
        elif self.conn_type == 'GPIB':
            # GPIB连接字符串 (Keithley 2700)
            # 假设GPIB板卡号为0，地址由用户输入
            return f"GPIB0::{self.address_str}::INSTR"
        else:
            raise ValueError(f"不支持的连接类型: {self.conn_type}")

    def connect(self):
        """
        建立与设备的连接
        """
        try:
            print(f"正在尝试连接到: {self.resource_string}")
            self.instrument = self.rm.open_resource(self.resource_string)

            # 设置通信参数
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self.instrument.timeout = 20000  # 20秒判定超时

            time.sleep(0.2)

            print("连接已建立，正在验证设备身份...")
            idn = self.query('*IDN?')
            installed_modules = self.query('*OPT?')
            print(f"已安装模块: {installed_modules}")
            if installed_modules == "7708,7708":
                self.scan_list = "(@101:140,201:240)"
                self.sample_count = 80
                self.opt = "@101:140,201:240"
            elif installed_modules == "7708,NONE":
                self.scan_list = "(@101:140)"
                self.sample_count = 40
                self.opt = "@101:140"
            elif installed_modules == "NONE,7708":
                self.scan_list = "(@201:240)"
                self.sample_count = 40
                self.opt = "@201:240"

            # 验证逻辑可以更通用
            if 'KEITHLEY' in idn.upper() and ('2701' in idn or '2700' in idn):
                self.connected = True
                print("设备验证成功，连接已就绪。")
                return True
            else:
                print("设备验证失败，IDN不匹配或无响应。")
                self.close()
                return False

        except pyvisa.errors.VisaIOError as e:
            print(f"连接失败：发生VISA I/O错误。")
            print("请检查:")
            if self.conn_type == 'TCPIP':
                print("- IP地址是否正确且设备在线？")
                print("- 网络连接和防火墙设置是否正确？")
            elif self.conn_type == 'GPIB':
                print("- GPIB地址是否正确？")
                print("- GPIB卡驱动 (如NI-VISA) 是否已正确安装？")
                print("- 设备是否已开机并连接到GPIB总线？")
            print(f"详细错误: {e}")
            return False
        except Exception as e:
            print(f"连接失败：发生未知错误 - {e}")
            return False

    def write(self, command):
        """
        一个专用的写入方法，用于发送无需响应的指令
        """
        if not self.connected or self.instrument is None:
            raise ConnectionError("仪器未连接，无法发送指令")
        try:
            self.instrument.write(command)
            time.sleep(0.1)
        except pyvisa.errors.VisaIOError as e:
            print(f"指令 '{command}' 写入失败: {e}")
            self.connected = False  # 更新状态
            raise  # 重新抛出异常，让上层知道操作失败

    def init_temperature_scan(self, thermocouple_type='K', nplc=5):
        """
        初始化仪器进行多通道温度扫描
        配置仪器进行80个通道的热电偶温度测量
        """
        if not self.connected:
            print("必须先连接设备才能进行初始化。")
            return False

        print("\n开始配置温度扫描参数...")
        try:
            self.write('*CLS')  # 清除状态
            print(self.scan_list)
            # 1. 配置全局参数
            self.write(f"SENS:FUNC 'TEMP', {self.scan_list}")
            self.write(f"SENS:TEMP:NPLC {nplc}")  # 设置积分时间

            # 2. 配置通道级参数
            self.write(f"SENS:TEMP:TRAN TC, {self.scan_list}")
            self.write(f"SENS:TC:TYPE {thermocouple_type}, {self.scan_list}")
            self.write(f"SENS:TC:CJON:STATE ON, {self.scan_list}")
            self.write(f"SENS:TEMP:UNIT C, {self.scan_list}")


            # 3. 配置触发和采样
            self.write("TRAC:CLE")
            self.write("INIT:CONT OFF")
            self.write("TRIG:SOUR IMM")
            self.write("TRIG:COUN 1")
            self.write(f"SAMP:COUN {self.sample_count}")
            self.write(f"ROUT:SCAN {self.scan_list}")
            self.write("ROUT:SCAN:TSO IMM")
            self.write("FORM:ELEM READ")
            self.write("ROUT:SCAN:LSEL INT")  # 扫描打开


            # 正确查询通道配置（添加空格）
            #print("测量函数:", self.query("SENS:FUNC? (@101)"))
            #print("通道101的热电偶类型:", self.query("SENS:TC:TYPE? (@101)"))  # 应返回 "K"
            #print("当前扫描列表:", self.query("ROUT:SCAN?"))  # 应返回您的扫描列表
            # 查询扫描列表
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
            # 如果在查询过程中发生错误，说明连接可能已断开
            print(f"指令 '{command}' 执行失败: {e}")
            self.connected = False  # 更新状态
            return "Error: VISA IO Error"

    def get_data(self,command):
        if command.upper().startswith(('READ?', 'FETC?')):
            real_temp=[]
            temperatures = self.instrument.query_ascii_values('READ?', container=list)  # 直接获取列表
            #print(temperatures)
            #self.write("ROUT:SCAN:LSEL NONE") #扫描关闭
            """
            这里需要说明一下，由于80个通道扫描耗时较高，最终获得的数据间隔可能与设置的不一致
            """
            for temp in temperatures:
                if 1000000 > temp:
                    real_temp.append(float(temp))
                else:
                    real_temp.append(None)
            return  real_temp
            #return ','.join([f'{temp:.4f}' for temp in temperatures])
        else:
            return []


    def close(self):
        if self.instrument:
            try:
                self.write("ROUT:SCAN:LSEL NONE") #扫描关闭
                self.instrument.close()
                print("设备连接已成功关闭。")
            except pyvisa.errors.VisaIOError:
                pass  # 关闭时可能出错，忽略即可
        self.instrument = None
        self.connected = False