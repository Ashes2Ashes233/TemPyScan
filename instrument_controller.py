# instrument_controller.py

import time
import pyvisa
import pyvisa_py
import numpy as np

# --- 仿真器代码 ---
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


# --- 真实设备控制器 ---
class Keithley2701Controller:
    def __init__(self, address):
        # 恢复原始的、正确的SOCKET连接地址格式
        self.address = f"TCPIP0::{address}::1394::SOCKET"
        # --- 关键修正：明确使用 pyvisa-py 后端 ---
        # 这能解决很多跨平台和环境的依赖问题
        self.rm = pyvisa.ResourceManager('@py')
        self.instrument = None
        self.connected = False
        self.scan_list = "(@101:140,201:240)"

    def connect(self):
        """
        建立与设备的连接
        """
        try:
            print(f"正在尝试以SOCKET模式连接到: {self.address}")
            self.instrument = self.rm.open_resource(self.address)

            # 设置通信参数
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self.instrument.timeout = 20000  # 20秒判定超时

            # 给予设备响应时间
            time.sleep(0.2)

            # 直接发送验证指令。query方法内部的逻辑已被修正。
            print("连接已建立，正在验证设备身份...")
            idn = self.query('*IDN?')
            print(f"收到设备ID: {idn}")

            if 'KEITHLEY' in idn.upper() and '2701' in idn:
                # --- 只有在验证成功后，才将状态设为True ---
                self.connected = True
                print("设备验证成功，连接已就绪。")
                return True
            else:
                print("设备验证失败，IDN不匹配或无响应。")
                self.close()
                return False

        except pyvisa.errors.VisaIOError as e:
            print(f"连接失败：发生VISA I/O错误。请检查IP地址、网络和防火墙。")
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
            time.sleep(0.05)
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
            # 1. 设置扫描列表
            self.write(f"ROUT:SCAN {self.scan_list}")

            # 2. 配置全局参数
            self.write(f"SENS:FUNC 'TEMP', {self.scan_list}")
            self.write(f"SENS:TEMP:NPLC {nplc}")  # 设置积分时间

            # 3. 配置通道级参数
            self.write(f"SENS:TEMP:TRAN TC, {self.scan_list}")
            self.write(f"SENS:TC:TYPE {thermocouple_type}, {self.scan_list}")
            self.write(f"SENS:TC:CJON:STATE ON, {self.scan_list}")
            self.write(f"SENS:TEMP:UNIT C, {self.scan_list}")


            # 4. 配置触发和采样
            self.write("TRAC:CLE")
            self.write("TRIG:SOUR IMM")
            self.write("SAMP:COUN 80")
            self.write("ROUT:SCAN:TSO IMM")
            self.write("INIT:CONT OFF")
            self.write("FORM:ELEM READ")
            self.write("TRIG:COUN 1")

            # 添加同步点确保配置完成
            #self.query("*OPC?")

            # 正确查询通道配置（添加空格）
            print("测量函数:", self.query("SENS:FUNC? (@101)"))
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
            self.write("ROUT:SCAN:LSEL INT") #扫描打开
            temperatures = self.instrument.query_ascii_values('READ?', container=list)  # 直接获取列表
            self.write("ROUT:SCAN:LSEL NONE") #扫描关闭
            """
            这里需要说明一下，由于80个通道扫描耗时较高，最终获得的数据间隔可能与设置的不一致
            """
            #print(temperatures)
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
                self.instrument.close()
                print("设备连接已成功关闭。")
            except pyvisa.errors.VisaIOError:
                pass  # 关闭时可能出错，忽略即可
        self.instrument = None
        self.connected = False