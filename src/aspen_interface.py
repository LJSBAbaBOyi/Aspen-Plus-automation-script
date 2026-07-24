"""
Aspen Plus COM 接口封装模块
提供连接 Aspen Plus、读取模块/流股变量、设置参数、运行模拟、获取结果等功能

Copyright (c) 2026 lijunsen & DeepSeek

连接策略（按优先级）:
  1. GetActiveObject → 连接已有进程
  2. Dispatch → 启动新 Aspen 实例
  3. 自动扫描注册表发现所有 Aspen 相关 COM 类
"""

import time
import os
import sys
import struct
import logging
import queue
import threading
import winreg as _wr
import win32com.client
import pythoncom


# ──────────────────────────────────────────────
# 基础路径工具（PyInstaller 打包兼容）
# ──────────────────────────────────────────────

def _get_exe_dir():
    """返回 exe 所在目录（源码运行则返回项目根目录）"""
    if getattr(sys, 'frozen', False):
        # 打包后: sys.executable 指向 exe 路径
        return os.path.dirname(sys.executable)
    else:
        # 源码运行: 项目根目录（src/ 的父目录）
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ──────────────────────────────────────────────
# 日志系统
# ──────────────────────────────────────────────

class AspenLogger:
    """双通道日志：内存缓冲区 + 文件"""

    _instance = None

    @classmethod
    def get(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        self.buffer = []  # 内存缓冲区，供 UI 读取
        self._logger = logging.getLogger("AspenInterface")
        self._logger.setLevel(logging.DEBUG)

        # 控制台 handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        self._logger.addHandler(ch)

        # 文件 handler（写入 exe 所在目录的 logs/ 文件夹）
        base_dir = _get_exe_dir()
        log_dir = os.path.join(base_dir, "logs")
        os.makedirs(log_dir, exist_ok=True)
        fh = logging.FileHandler(
            os.path.join(log_dir, f"aspen_{time.strftime('%Y%m%d_%H%M%S')}.log"),
            encoding="utf-8"
        )
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        self._logger.addHandler(fh)

    def debug(self, msg):
        self._logger.debug(msg)
        self.buffer.append(("DEBUG", msg))

    def info(self, msg):
        self._logger.info(msg)
        self.buffer.append(("INFO", msg))

    def warning(self, msg):
        self._logger.warning(msg)
        self.buffer.append(("WARN", msg))

    def error(self, msg):
        self._logger.error(msg)
        self.buffer.append(("ERROR", msg))

    def get_logs(self, n=200):
        return self.buffer[-n:]


log = AspenLogger.get()


# ──────────────────────────────────────────────
# 注册表扫描工具
# ──────────────────────────────────────────────

def _scan_aspen_progids():
    """全量扫描注册表，找出所有 Aspen 相关的 COM ProgID

    策略：
    - 关键词匹配 apwn / aspen / aspenplus / aspentech / apwshell
    - 排除已知的 ActiveX / 嵌入控件
    - 按 已知有效模式 > 试探 > 控件类 顺序排列
    """
    found = set()
    keywords = ["apwn", "apwshell"]

    # 已知的无效模式（ActiveX 控件 / 嵌入控件）
    bad_patterns = ["activex", "pfsdotnet", "embed", "wpf", "dotnet", "control"]
    # Aspen Plus 的 COM ProgID 固定以 "Apwn." 开头
    # e.g. Apwn.Document.38 (V14), Apwn.Application.38, Apwn.Document.37 (V12)
    good_prefixes = ["apwn."]

    def _is_bad(name):
        nl = name.lower()
        return any(p in nl for p in bad_patterns)

    def _is_good(name):
        nl = name.lower()
        return any(nl.startswith(p) for p in good_prefixes)

    # 搜索 HKEY_CLASSES_ROOT 下所有顶级键
    try:
        key = _wr.OpenKey(_wr.HKEY_CLASSES_ROOT, "")
        i = 0
        while True:
            try:
                name = _wr.EnumKey(key, i)
                name_lower = name.lower()
                if any(kw in name_lower for kw in keywords):
                    if not _is_bad(name):
                        found.add(name)
                i += 1
            except OSError:
                break
        _wr.CloseKey(key)
    except Exception as e:
        log.error(f"注册表扫描失败: {e}")

    # 搜索 CLSID 下的描述
    try:
        clsid_key = _wr.OpenKey(_wr.HKEY_CLASSES_ROOT, "CLSID")
        i = 0
        while True:
            try:
                clsid = _wr.EnumKey(clsid_key, i)
                sub_key = _wr.OpenKey(clsid_key, clsid)
                try:
                    desc = _wr.QueryValueEx(sub_key, "")[0]
                    if any(kw in desc.lower() for kw in keywords):
                        try:
                            progid_key = _wr.OpenKey(sub_key, "ProgID")
                            progid = _wr.QueryValueEx(progid_key, "")[0]
                            if not _is_bad(progid):
                                found.add(progid)
                        except:
                            pass
                except:
                    pass
                _wr.CloseKey(sub_key)
                i += 1
            except OSError:
                break
        _wr.CloseKey(clsid_key)
    except:
        pass

    # 只保留 Document 和 Application 类，排除 OutputFile、Archive 等无关类
    found = {p for p in found if any(k in p.lower() for k in ["document", "application"])}

    # 按有效性排序（好 > 试探 > 坏）
    # Document 类型中：带版本号的 (38.0) > 通用 (Document) > IP 等别名
    # Application 类型优先，但本系统可能不提供 Application 接口
    def _sort_key(x):
        if _is_bad(x):
            return 99  # 排最后
        if _is_good(x):
            xl = x.lower()
            # Application 最高优先级
            if xl.startswith("apwn.application"):
                return 0
            # Document 类型：带明确版本号的优先
            if xl.startswith("apwn.document"):
                # "Apwn.Document.38.0" / "Apwn.Document.38" 标准版本号最高优先级
                if xl in ("apwn.document.38.0", "apwn.document.38", "apwn.document.37", "apwn.document.37.0"):
                    return 1
                # "Apwn.Document.IP.38.0" 等带版本号的别名
                if any(f".{v}" in xl for v in ["38.0", "38", "37"]):
                    return 2
                # "Apwn.Document" 无版本号
                if xl == "apwn.document" or xl == "apwn.document.1":
                    return 3
                # "Apwn.Document.IP" 等别名排在最后
                return 4
            # 其他 apwn. 前缀
            if xl.startswith("apwn."):
                return 5
            if "application" in xl:
                return 6
            if "document" in xl:
                return 7
            return 8
        # 试探类（包含关键词但不在已知列表中）
        return 10

    result = sorted(found, key=_sort_key)
    return result


# ──────────────────────────────────────────────
# 连接工具
# ──────────────────────────────────────────────

def _try_get_active_object(progid):
    """尝试 GetActiveObject 连接到已有进程"""
    try:
        pythoncom.CoInitialize()
    except:
        pass
    try:
        app = win32com.client.GetActiveObject(progid)
        log.info(f"GetActiveObject 成功: {progid} (类=Win32Com)")
    except Exception as e:
        log.debug(f"GetActiveObject({progid}) 失败: {e}")
        return None, None, None

    # 安全获取 Tree（可能 Aspen 未加载文件时 Tree 为空）
    tree = None
    try:
        tree = app.Tree
        log.info(f"app.Tree 获取成功: {type(tree).__name__}")
    except Exception as e:
        log.warning(f"app.Tree 获取失败（可能未加载文件）: {e}")
        # Tree 不可用，但仍返回 app，让上层决定如何处理
        return app, None, "Document"

    return app, tree, "Document"


def _try_get_object_by_file(filepath):
    """尝试用 GetObject 通过文件路径连接 Aspen（可能连接到已运行实例）"""
    try:
        pythoncom.CoInitialize()
    except:
        pass
    try:
        app = win32com.client.GetObject(filepath)
        log.info(f"GetObject 文件成功: {filepath}")
    except Exception as e:
        log.debug(f"GetObject({filepath}) 失败: {e}")
        return None, None, None

    tree = None
    try:
        tree = app.Tree
    except:
        pass
    return app, tree, "Document"


def _try_dispatch(progid):
    """尝试 Dispatch 创建新实例
    优先尝试 Application 层级（适用于 InitFromFile），
    失败后回退到 Document 层级。
    """
    # 先看是否是 Application 层级
    is_app = "application" in progid.lower()

    try:
        app = win32com.client.Dispatch(progid)

        if is_app:
            # Application 对象
            log.info(f"Dispatch Application 成功: {progid}")
            return app, None, "Application"
        else:
            # Document 对象：刚 Dispatch 完处于未初始化状态，
            # 不能设 Visible 或访问 Tree，等 load_file 里 InitFromFile 后再做
            log.info(f"Dispatch Document 成功 (等待 InitFromFile): {progid}")
            return app, None, "Document"
    except Exception as e:
        log.debug(f"Dispatch({progid}) 失败: {e}")
        return None, None, None


# ──────────────────────────────────────────────
# 专用 COM 工作线程
# ──────────────────────────────────────────────
# 所有 Aspen COM 操作必须在此线程上执行，避免跨线程 COM 调用导致崩溃

class _AspenWorker(threading.Thread):
    """专用 COM 工作线程 - 所有 Aspen COM 操作在此线程执行"""

    def __init__(self):
        super().__init__(daemon=True)
        self._req_queue = queue.Queue()   # (func, result_queue) 请求
        self._shutdown = threading.Event()
        # Aspen COM 状态变量（仅在工作线程读写）
        self.app = None
        self.tree = None
        self.is_connected = False
        self.connected_progid = ""
        self.connect_method = ""
        self.dispatch_type = ""
        self.needs_file_load = False
        self.start()

    def run(self):
        """工作线程主循环 - 初始化 COM 并处理请求"""
        try:
            pythoncom.CoInitialize()
        except:
            pass
        log.info("Aspen COM 工作线程已启动")
        while not self._shutdown.is_set():
            try:
                func, res_q = self._req_queue.get(timeout=1)
            except queue.Empty:
                continue
            try:
                result = func(self)
                if res_q is not None:
                    res_q.put(("ok", result))
            except Exception as e:
                log.error(f"COM 工作线程执行异常: {e}")
                if res_q is not None:
                    res_q.put(("error", e))
        # 清理
        self.app = None
        self.tree = None
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        log.info("Aspen COM 工作线程已退出")

    def execute(self, func, timeout=120):
        """在工作线程上执行 func(self)，阻塞等待结果"""
        if self._shutdown.is_set():
            raise RuntimeError("COM 工作线程已关闭")
        res_q = queue.Queue()
        self._req_queue.put((func, res_q))
        status, result = res_q.get(timeout=timeout)
        if status == "error":
            raise result  # 重新抛出异常
        return result

    def execute_async(self, func, callback=None):
        """在工作线程上异步执行 func(self)，不阻塞"""
        if self._shutdown.is_set():
            raise RuntimeError("COM 工作线程已关闭")
        if callback:
            def wrapper(worker):
                try:
                    result = func(worker)
                    callback(True, result)
                except Exception as e:
                    callback(False, e)
                return result
            self._req_queue.put((wrapper, None))
        else:
            self._req_queue.put((func, None))

    def stop(self):
        """关闭工作线程"""
        self._shutdown.set()


# ──────────────────────────────────────────────
# AspenInterface 主类（线程安全封装）
# ──────────────────────────────────────────────

class AspenInterface:
    """Aspen Plus COM 接口封装（线程安全）

    所有 COM 操作在专用工作线程上执行，对外提供阻塞式方法。
    可以在任何线程安全调用（包括临时工作线程），内部自动路由到 COM 工作线程。
    """

    # 以下静态数据可直接访问，无需经过工作线程
    BLOCK_INPUT_PATHS = {
        "Heater": {
            "温度 (Temperature)": "TEMP",
            "压力 (Pressure)": "PRES",
            "气相分率 (Vapor Fraction)": "VFRAC",
            "热负荷 (Duty)": "DUTY",
        },
        "HeatX": {
            "冷侧出口温度 (Cold Outlet Temp)": "TEMP_COLD",
            "热侧出口温度 (Hot Outlet Temp)": "TEMP_HOT",
            "冷侧压降 (Cold Pressure Drop)": "DP_COLD",
            "热侧压降 (Hot Pressure Drop)": "DP_HOT",
        },
        "RadFrac": {
            "回流比 (Reflux Ratio)": "BASIS_RR",
            "塔顶采出率 (Distillate Rate)": "BASIS_DR",
            "塔顶压力 (Condenser Pressure)": "PRES_COL",
            "塔釜压力 (Reboiler Pressure)": "PRES_REB",
            "塔板数 (Number of Stages)": "NSTAGE",
            "进料板位置 (Feed Stage)": "FEED_STAGE",
        },
        "DSTWU": {
            "回流比 (Reflux Ratio)": "RR",
            "理论板数 (Number of Stages)": "NSTAGE",
            "轻关键组分回收率": "LIGHT_KEY_RECOV",
            "重关键组分回收率": "HEAVY_KEY_RECOV",
        },
        "Flash2": {
            "温度 (Temperature)": "TEMP",
            "压力 (Pressure)": "PRES",
            "热负荷 (Duty)": "DUTY",
            "气相分率 (Vapor Fraction)": "VFRAC",
        },
        "Sep": {"出口流股分率 (Split Fraction)": "FRAC"},
        "Mixer": {"压力 (Pressure)": "PRES"},
        "FSplit": {"流股分率 (Split Fraction)": "FRAC"},
        "Compressor": {"出口压力 (Discharge Pressure)": "PRES", "等熵效率 (Isentropic Efficiency)": "EFF"},
        "Pump": {"出口压力 (Discharge Pressure)": "PRES", "效率 (Efficiency)": "EFF"},
        "Valve": {"出口压力 (Outlet Pressure)": "PRES"},
        "Tank": {"温度 (Temperature)": "TEMP", "压力 (Pressure)": "PRES"},
    }

    # ── 流股输入参数预定义列表 ──
    # 来源: 流股input所需参数项.xlsx
    # 格式: (display_name, subpath)
    # 其中 subpath 相对于 \Data\Streams\{stream_name}\Input\
    STREAM_INPUT_SPEC = [
        ("组成基准",              "BASIS\\MIXED"),
        ("计算方法 (电解质)",     "CHEM_METHOD"),
        ("流量基准",              "FLOWBASE\\MIXED"),
        ("自由水计算",            "FREE_WATER\\MIXED"),
        ("闪蒸选项",              "FL_OPTION\\MIXED"),
        ("压力",                  "PRES\\MIXED"),
        ("溶剂组分",              "SOLVENT\\MIXED"),
        ("温度",                  "TEMP\\MIXED"),
        ("总计",                  "TOTAL\\MIXED"),
        ("总流量",                "TOTFLOW\\MIXED"),
        ("汽相分率",              "VFRAC\\MIXED"),
        # 组分流量 FLOW\MIXED\{组分名} 在代码中动态枚举
    ]

    STREAM_INPUT_PATHS = {
        "温度 (Temperature)": "TEMP",
        "压力 (Pressure)": "PRES",
        "总质量流量 (Mass Flow)": "MASSFLOW",
        "总摩尔流量 (Mole Flow)": "MOLEFLOW",
        "气相分率 (Vapor Fraction)": "VFRAC",
    }

    STREAM_OUTPUT_PATHS = {
        "温度 (Temperature)": "TEMP_OUT",
        "压力 (Pressure)": "PRES_OUT",
        "总质量流量 (Mass Flow)": "MASSFLOW_OUT",
        "总摩尔流量 (Mole Flow)": "MOLEFLOW_OUT",
        "气相分率 (Vapor Fraction)": "VFRAC_OUT",
        "摩尔焓 (Molar Enthalpy)": "HMX_OUT",
        "摩尔密度 (Molar Density)": "RHOMX_OUT",
    }

    # ── HAP 属性常量（来自 Aspen Plus 类型库 happ.tlb） ──
    HAP_VALUE = 0
    HAP_UNITROW = 2
    HAP_UNITCOL = 3
    HAP_OPTIONLIST = 5
    HAP_BASIS = 7
    HAP_ENTERABLE = 8
    HAP_UPPERLIMIT = 9
    HAP_LOWERLIMIT = 10
    HAP_VALUEDEFAULT = 11
    HAP_USERENTERED = 12
    HAP_INOUT = 14
    HAP_OUTVAR = 18
    HAP_PROMPT = 19

    # ── 常见 Aspen 单位映射（MET 单位集） ──
    _UNIT_COL_MAP = {
        1: "", 2: "K", 3: "C", 4: "F",
        11: "atm", 12: "bar", 13: "Pa", 14: "kPa",
        21: "kg/hr", 22: "kg/s", 23: "lb/hr",
        31: "kmol/hr", 32: "kmol/s", 33: "mol/s",
        41: "m3/hr", 42: "L/min", 43: "m3/s",
        51: "J/hr", 52: "kJ/hr", 53: "kW", 54: "MW",
        61: "kJ/kmol", 62: "kcal/kmol",
        71: "kg/m3", 72: "kmol/m3",
    }

    def _get_unit_string(self, elem, tree=None):
        """从 COM 节点读取单位字符串

        策略（按优先级）：
        1. Unit Table 查找：HAP_UNITROW(物理量代号)+HAP_UNITCOL(单位代号)→按值匹配单位名
           （最准确，反映用户在 Aspen 中实际选择的单位）
        2. elem.Units 直接属性（可能返回默认/标准单位，不准确）
        3. _UNIT_COL_MAP 硬编码回退：HAP_UNITCOL 代号→已知单位名
        """
        # 策略1: Unit Table 代号查找（最可靠——反映实际单位选择）
        if tree is not None:
            table_unit = self._get_unit_from_table(elem, tree)
            if table_unit:
                return table_unit

        # 策略2: 直接读取 .Units 属性（可能返回标准单位而非实际选择单位）
        try:
            u = elem.Units
            if u is not None:
                unit = str(u).strip()
                if unit:
                    return unit
        except Exception:
            pass

        # 策略3: 硬编码回退（HAP_UNITCOL 代号 → 已知单位名）
        try:
            u_col = elem.AttributeValue(self.HAP_UNITCOL)
            if u_col is not None:
                unit = self._UNIT_COL_MAP.get(int(u_col), "")
                if unit:
                    return unit
        except Exception:
            pass

        return ""

    def _get_data_type(self, elem):
        """尝试获取节点的数据类型字符串
        
        策略：
        1. elem.Type 直接属性
        2. AttributeValue("DataType") 兼容旧方式
        """
        # 策略1: Type 属性
        try:
            dt = str(elem.Type)
            if dt:
                return dt
        except Exception:
            pass
        
        # 策略2: AttributeValue 方式（兼容已有代码）
        try:
            dt = str(elem.AttributeValue("DataType") or "")
            if dt:
                return dt
        except Exception:
            pass
        
        return ""

    def _build_unit_cache(self, tree):
        r"""构建 Unit Table 缓存: {(u_row, u_col): unit_name}

        遍历一次 \Unit Table，构建从 (物理量代号, 单位代号) 到单位名的映射。
        后续查单位直接 O(1) 命中，无需反复遍历。
        """
        cache = {}
        try:
            unit_table = tree.FindNode("\\Unit Table")
            if not unit_table:
                return cache
            for _idx, row_node in self._enum_elements_safe(unit_table.Elements):
                try:
                    row_code = int(row_node.Value)
                except Exception:
                    continue
                for _j, unit_node in self._enum_elements_safe(row_node.Elements):
                    try:
                        col_code = int(unit_node.Value)
                    except Exception:
                        continue
                    name = str(unit_node.Name).strip()
                    if name:
                        cache[(row_code, col_code)] = name
        except Exception:
            pass
        return cache

    def _get_unit_from_table(self, elem, tree):
        r"""通过 Aspen Unit Table 解析单位字符串（带缓存，O(1)）

        Unit Table 路径：\Unit Table
          子节点名 = 物理量名（如 "DENSITY"），子节点值 = 物理量代号
          孙节点名 = 单位名（如 "kg/cum"），孙节点值 = 单位代号
        """
        try:
            u_row = int(elem.AttributeValue(self.HAP_UNITROW))
            u_col = int(elem.AttributeValue(self.HAP_UNITCOL))
        except (TypeError, ValueError, Exception):
            return ""

        # 惰性构建缓存（仅在首次调用时遍历一次 Unit Table）
        if self._unit_cache is None:
            try:
                self._unit_cache = self._build_unit_cache(tree)
            except Exception:
                self._unit_cache = {}

        return self._unit_cache.get((u_row, u_col), "")

    def diagnose_node_attrs(self, path):
        """诊断：读取指定节点的所有 HAP_ 属性原始值，供调试用。
        返回 dict: {"HAP_VALUE": ..., "HAP_UNITROW": ..., "HAP_UNITCOL": ..., ...}
        """
        def _do_get(worker):
            result = {}
            node = worker.tree.FindNode(path)
            if not node:
                return {"error": f"未找到节点: {path}"}
            for attr_name in ("HAP_VALUE", "HAP_UNITROW", "HAP_UNITCOL",
                               "HAP_PROMPT", "HAP_INOUT", "HAP_BASIS"):
                try:
                    val = node.AttributeValue(attr_name)
                    result[attr_name] = val
                except Exception as e:
                    result[attr_name] = f"<ERROR: {e}>"
            # 也读 Units / Type 直接属性
            try:
                result["Units"] = str(node.Units)
            except:
                result["Units"] = "<ERROR>"
            try:
                result["Type"] = str(node.Type)
            except:
                result["Type"] = "<ERROR>"
            # Unit Table 快照
            try:
                ut = worker.tree.FindNode("\\Unit Table")
                if ut:
                    result["UnitTable_children"] = ut.Elements.Count
            except:
                result["UnitTable_children"] = "<ERROR>"
            return result
        try:
            return self._worker.execute(_do_get)
        except Exception as e:
            return {"error": str(e)}

    def _get_node_value(self, node):
        """安全读取节点值"""
        try:
            return node.Value
        except Exception:
            return None

    def __init__(self):
        self._worker = _AspenWorker()
        # 以下是 worker 状态的代理（只读）
        self.is_connected = False
        self.connected_progid = ""
        self.connect_method = ""
        self.dispatch_type = ""
        self.needs_file_load = False
        self._unit_cache = None  # {(u_row, u_col): unit_name} 惰性缓存
        self.loaded_filepath = None  # 当前加载的 .bkp 文件路径（崩溃恢复用）

    def _sync_state(self, worker):
        """同步工作线程状态到主线程可见属性"""
        self.is_connected = worker.is_connected
        self.connected_progid = worker.connected_progid
        self.connect_method = worker.connect_method
        self.dispatch_type = worker.dispatch_type
        self.needs_file_load = worker.needs_file_load

    def _reset_worker(self):
        """重置 COM 工作线程（连接失败 / Aspen崩溃后调用）

        旧的 COM 公寓可能处于损坏状态（如 SGBizLauncher 崩溃后的残留），
        必须先停止旧线程再启动新线程，才能在新 COM 公寓中进行连接。
        """
        log.info("正在重置 COM 工作线程...")
        try:
            self._worker.stop()
            self._worker.join(timeout=3)
        except:
            pass
        self._worker = _AspenWorker()
        self.is_connected = False
        self.connected_progid = ""
        self.connect_method = ""
        self.dispatch_type = ""
        self._unit_cache = None  # 清空 Unit Table 缓存
        self.loaded_filepath = None  # 清空文件路径
        self.needs_file_load = False
        log.info("COM 工作线程已重置")

    # ──────────────────────────────────────────────
    # 连接管理
    # ──────────────────────────────────────────────

    def connect(self):
        """连接到 Aspen Plus。先在主线程尝试 GetActiveObject，避免工作线程 COM 公寓问题。"""
        # ── 主线程上先尝试 GetActiveObject ──
        main_app = None
        main_tree = None
        main_progid = None
        try:
            pythoncom.CoInitialize()
        except:
            pass
        try:
            app = win32com.client.GetActiveObject("Apwn.Document.38.0")
            try:
                tree = app.Tree
            except:
                tree = None
            main_app = app
            main_tree = tree
            main_progid = "Apwn.Document.38.0"
            log.info(f"主线程 GetActiveObject 成功: {main_progid}")
        except Exception as e:
            log.debug(f"主线程 GetActiveObject 失败: {e}")

        def _do_connect(worker):
            # 如果主线程已获取到 app，直接使用
            if main_app is not None and main_tree is not None:
                worker.app = main_app
                worker.tree = main_tree
                worker.is_connected = True
                worker.connected_progid = main_progid
                worker.connect_method = "GetActiveObject"
                worker.dispatch_type = "Document"
                worker.needs_file_load = False
                log.info(f"连接成功 (主线程 GetActiveObject): {main_progid}")
                return True, f"已连接到正在运行的 Aspen Plus ({main_progid})"

            log.info("开始连接 Aspen Plus...")

            # 直接尝试 GetActiveObject → Dispatch，无需注册表扫描
            PROGID = "Apwn.Document.38.0"

            # 1. GetActiveObject（连接已有实例）
            app, tree, obj_type = _try_get_active_object(PROGID)
            if app is not None:
                worker.app = app
                worker.tree = tree
                worker.is_connected = True
                worker.connected_progid = PROGID
                worker.connect_method = "GetActiveObject"
                worker.dispatch_type = "Document"
                worker.needs_file_load = (tree is None)
                if tree is None:
                    log.warning("GetActiveObject 成功但 Tree 为空")
                    return True, "已连接到 Aspen Plus，但未检测到加载的模拟文件"
                log.info(f"连接成功 (GetActiveObject): {PROGID}")
                return True, f"已连接到正在运行的 Aspen Plus"

            # 2. Dispatch（启动新实例）
            app, tree, obj_type = _try_dispatch(PROGID)
            if app is not None:
                worker.app = app
                worker.tree = tree
                worker.is_connected = True
                worker.connected_progid = PROGID
                worker.connect_method = "Dispatch"
                worker.dispatch_type = obj_type
                worker.needs_file_load = True
                log.info(f"连接成功 (Dispatch {obj_type}): {PROGID}")
                return True, "已启动新 Aspen Plus 实例，请选择 .bkp 文件加载"

            # 全部失败
            worker.is_connected = False
            py_bits = 64 if struct.calcsize("P") * 8 == 64 else 32
            msg = (f"连接 Aspen Plus 失败。\nPython: {py_bits} 位\n\n"
                   f"GetActiveObject 和 Dispatch 均失败。\n"
                   f"建议:\n  1. 确认 Aspen Plus V14 已安装\n"
                   f"  2. 以管理员身份运行本工具\n"
                   f"  3. 尝试修复安装 Aspen Plus\n\n"
                   f"详细日志已保存至 logs/ 目录")
            log.error("所有连接方法均失败")
            return False, msg

        try:
            result = self._worker.execute(_do_connect)
            self._sync_state(self._worker)
            return result
        except Exception as e:
            log.error(f"连接异常: {type(e).__name__}: {e}")
            # 连接过程中 COM 可能已损坏，自动重置 worker
            self._reset_worker()
            return False, f"连接失败: {type(e).__name__}: {e}。已重置COM线程，请重试。"

    def reconnect(self):
        """先释放连接 + 重置COM线程，再重新连接（用于从崩溃中恢复）"""
        self.close()
        self._reset_worker()
        return self.connect()

    def close(self):
        """释放 COM 引用（不关闭 Aspen，让用户手动管理 Aspen 生命周期）"""
        def _do_close(worker):
            worker.app = None
            worker.tree = None
            worker.is_connected = False
            worker.connected_progid = ""
            worker.connect_method = ""
            worker.dispatch_type = ""
            worker.needs_file_load = False
            log.info("已断开 Aspen Plus COM 连接（Aspen 进程仍由用户管理）")
            return True, "已断开连接"
        try:
            result = self._worker.execute(_do_close)
            self._sync_state(self._worker)
            return result
        except Exception as e:
            return False, f"断开失败: {e}"

    # ──────────────────────────────────────────────
    # 文件加载
    # ──────────────────────────────────────────────

    def load_file(self, filepath):
        """加载 .bkp 文件到 Aspen Plus"""
        abs_path = os.path.abspath(filepath)
        if not os.path.isfile(abs_path):
            return False, f"文件不存在: {abs_path}"
        self.loaded_filepath = abs_path  # 记住路径，崩溃恢复用

        def _do_load(worker):
            if not worker.app:
                return False, "未连接到 Aspen Plus"
            log.info(f"正在加载文件 (dispatch_type={worker.dispatch_type}): {abs_path}")

            # Dispatch 后 Aspen 进程可能尚未完全初始化，等待片刻
            time.sleep(2)

            if worker.dispatch_type == "Application":
                worker.app.InitFromFile(abs_path)
                time.sleep(1)
                doc = worker.app.ActiveDocument
                worker.tree = doc.Tree
            else:
                # Document 模式：直接用 InitFromFile2（已验证可用）
                try:
                    worker.app.InitFromFile2(abs_path)
                    log.info("文件加载成功: InitFromFile2")
                except Exception as e:
                    err_msg = (f"加载文件失败 ({type(e).__name__}: {e})\n"
                               f"文件: {os.path.basename(abs_path)}\n"
                               f"请确认 .bkp 文件未被其他程序占用且路径正确")
                    log.error(err_msg)
                    return False, err_msg

                # 获取 Tree
                try:
                    worker.tree = worker.app.Tree
                    if worker.tree is None:
                        return False, "文件可能已加载，但无法获取 Tree 对象"
                    log.info("Tree 获取成功")
                    # 显示 Aspen Plus 窗口
                    try:
                        worker.app.Visible = True
                    except:
                        pass
                except Exception as e:
                    log.error(f"Tree 获取失败: {e}")
                    return False, f"文件加载后无法获取 Tree: {e}"

            worker.needs_file_load = False
            log.info(f"文件加载成功: {abs_path}")
            return True, f"已加载: {os.path.basename(abs_path)}"

        import time as _t
        try:
            result = self._worker.execute(_do_load)
            self._sync_state(self._worker)
            return result
        except Exception as e:
            error_detail = f"{type(e).__name__}: {e}"
            log.error(f"文件加载失败: {error_detail}")
            return False, f"加载文件失败: {error_detail}"

    # ──────────────────────────────────────────────
    # 诊断
    # ──────────────────────────────────────────────

    def diagnose(self):
        """诊断 Aspen Plus COM 环境"""
        py_bits = 64 if struct.calcsize("P") * 8 == 64 else 32

        lines = [
            "=" * 58,
            "  Aspen Plus COM 环境诊断报告",
            "=" * 58,
            f"时间: {time.strftime('%Y-%m-%d %H:%M:%S')}",
            f"Python 版本: {sys.version}",
            f"Python 架构: {py_bits} 位",
            f"pywin32 版本: {getattr(win32com.client, '__version__', 'N/A')}",
            "",
            "─" * 58,
            "注册表扫描结果 (HKEY_CLASSES_ROOT 下 Aspen 相关键):",
            "─" * 58,
        ]

        progids = _scan_aspen_progids()
        if progids:
            for p in progids:
                dispatch_ok = "?"
                try:
                    _wr.OpenKey(_wr.HKEY_CLASSES_ROOT, p)
                    dispatch_ok = "✅"
                except:
                    dispatch_ok = "❌"
                lines.append(f"  {dispatch_ok} {p}")
        else:
            lines.append("  ❌ 未找到任何 Aspen 相关 COM 注册项")

        lines.append("")
        lines.append("─" * 58)
        lines.append("系统环境:")
        lines.append("─" * 58)

        for candidate in [
            r"SOFTWARE\AspenTech",
            r"SOFTWARE\WOW6432Node\AspenTech",
        ]:
            try:
                key = _wr.OpenKey(_wr.HKEY_LOCAL_MACHINE, candidate)
                lines.append(f"  ✅ 注册表存在: {candidate}")
                _wr.CloseKey(key)
            except:
                pass

        for var in ["AspenTech_LICENSE", "ASPEN_PLUS_HOME"]:
            val = os.environ.get(var, "")
            if val:
                lines.append(f"  ✅ {var} = {val}")
            else:
                lines.append(f"  ⚠ {var} 未设置")

        lines.append("")
        lines.append("─" * 58)
        lines.append("排查建议:")
        lines.append("─" * 58)
        lines.append("  1. 打开 命令提示符(cmd) → 输入 regedit")
        lines.append("     搜索 HKEY_CLASSES_ROOT\\Apwn 是否存在")
        lines.append("  2. 确认 Aspen Plus 版本与 Python 架构一致")
        lines.append("  3. 尝试以管理员身份运行本工具")
        lines.append("  4. 尝试在 Aspen Plus 中: Help → Diagnostics 修复安装")
        lines.append("")
        lines.append("=" * 58)
        lines.append(f"日志文件: logs/")

        return "\n".join(lines)

    # ──────────────────────────────────────────────
    # 枚举模块与流股
    # ──────────────────────────────────────────────

    def _enum_elements_safe(self, elements):
        """安全枚举 COM Elements 集合，同时尝试 0-based 和 1-based 索引"""
        items = []  # [(index_used, item), ...]
        try:
            count = elements.Count
        except Exception as e:
            log.error(f"获取 Elements.Count 失败: {e}")
            return items
        if count == 0:
            return items
        # 先尝试 0-based
        for i in range(count):
            try:
                item = elements.Item(i)
                if item is not None:
                    items.append((i, item))
            except Exception as e0:
                log.debug(f"  0-based Item({i}) 失败，将尝试 1-based")
                break
        else:
            # 所有 0-based 都成功
            return items
        # 0-based 有失败，改用 1-based
        items.clear()
        for i in range(1, count + 1):
            try:
                item = elements.Item(i)
                if item is not None:
                    items.append((i, item))
            except Exception as e1:
                log.warning(f"  1-based Item({i}) 也失败: {e1}")
                continue
        return items

    def get_blocks(self):
        def _do_get(worker):
            blocks = []
            try:
                if worker.tree is None:
                    log.error("获取模块列表失败: tree 为 None")
                    return blocks
                blocks_node = worker.tree.FindNode(r"\Data\Blocks")
                if not blocks_node:
                    log.warning("路径 \\Data\\Blocks 不存在")
                    return blocks
                elements = blocks_node.Elements
                log.info(f"Blocks Elements.Count = {elements.Count}")
                items = self._enum_elements_safe(elements)
                for idx, child in items:
                    try:
                        blk_name = child.Name
                        blk_type = "Unknown"
                        try:
                            blk_type = child.AttributeValue("ComponentType") or "Unknown"
                        except:
                            pass
                        blocks.append((blk_name, blk_type))
                        log.info(f"  模块[{idx}]: {blk_name} ({blk_type})")
                    except Exception as e:
                        log.warning(f"  模块[{idx}] 读取失败: {e}")
                        continue
                log.info(f"共获取 {len(blocks)} 个模块")
            except Exception as e:
                log.error(f"获取模块列表失败: {e}")
            return blocks
        try:
            return self._worker.execute(_do_get)
        except Exception as e:
            log.error(f"get_blocks 失败: {e}")
            return []

    def get_streams(self):
        def _do_get(worker):
            streams = []
            try:
                if worker.tree is None:
                    log.error("获取流股列表失败: tree 为 None")
                    return streams
                streams_node = worker.tree.FindNode(r"\Data\Streams")
                if not streams_node:
                    log.warning("路径 \\Data\\Streams 不存在")
                    return streams
                elements = streams_node.Elements
                log.info(f"Streams Elements.Count = {elements.Count}")
                items = self._enum_elements_safe(elements)
                for idx, child in items:
                    try:
                        name = child.Name
                        streams.append(name)
                        log.info(f"  流股[{idx}]: {name}")
                    except Exception as e:
                        log.warning(f"  流股[{idx}] 读取失败: {e}")
                        continue
                log.info(f"共获取 {len(streams)} 条流股")
            except Exception as e:
                log.error(f"获取流股列表失败: {e}")
            return streams
        try:
            return self._worker.execute(_do_get)
        except Exception as e:
            log.error(f"get_streams 失败: {e}")
            return []

    # ──────────────────────────────────────────────
    # 获取/设置变量
    # ──────────────────────────────────────────────

    def _get_node_value(self, node_path):
        def _do_get(worker):
            try:
                node = worker.tree.FindNode(node_path)
                if node:
                    return node.Value
                return None
            except:
                return None
        try:
            return self._worker.execute(_do_get)
        except:
            return None

    def _set_node_value(self, node_path, value):
        def _do_set(worker):
            try:
                node = worker.tree.FindNode(node_path)
                if node:
                    node.Value = value
                    return True
                log.warning(f"未找到节点: {node_path}")
                return False
            except Exception as e:
                log.error(f"设置变量失败 [{node_path}]: {e}")
                return False
        try:
            return self._worker.execute(lambda w: _do_set(w))
        except:
            return False

    def get_block_inputs(self, block_name):
        """枚举模块 Input 子节点下的所有可调参数（不依赖硬编码字典）"""
        def _do_get(worker):
            inputs = []
            try:
                input_node = worker.tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Input")
                if not input_node:
                    return inputs
                self._walk_params_on_worker(
                    input_node, f"Blocks\\{block_name}\\Input", inputs, max_depth=2, tree=worker.tree
                )
            except Exception as e:
                log.error(f"获取模块输入变量失败 [{block_name}]: {e}")
            return inputs
        try:
            return self._worker.execute(_do_get)
        except Exception as e:
            log.error(f"get_block_inputs 失败: {e}")
            return []

    def get_block_outputs(self, block_name):
        """枚举模块 Output 子节点下的所有输出参数"""
        def _do_get(worker):
            outputs = []
            try:
                output_node = worker.tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output")
                if not output_node:
                    return outputs
                self._walk_params_on_worker(
                    output_node, f"Blocks\\{block_name}\\Output", outputs, max_depth=2, tree=worker.tree
                )
            except Exception as e:
                log.error(f"获取模块输出变量失败 [{block_name}]: {e}")
            return outputs
        try:
            return self._worker.execute(_do_get)
        except Exception as e:
            log.error(f"get_block_outputs 失败: {e}")
            return []

    def _walk_params_on_worker(self, node, base_path, results, max_depth=3, depth=0, tree=None):
        """递归遍历节点下的所有可读/可写参数

        返回值格式: (display_name, full_path, value, unit, data_type, is_numeric)
        - display_name: 变量名（中文/英文标注）
        - full_path: Aspen Tree 全路径
        - value: 当前值（字符串/数值/None）
        - unit: 单位字符串（如 "C", "atm", "kg/hr"）
        - data_type: 数据类型（"REAL", "TEXT", "INTEGER", "BINARY" 等）
        - is_numeric: 是否为数值型（可作输入变量）
        """
        if depth >= max_depth:
            return
        try:
            elements = node.Elements
        except Exception as e:
            log.debug(f"_walk_params: 获取 Elements 失败 (depth={depth}): {e}")
            return
        # 用索引方式枚举 COM 集合（比 for-in 更可靠）
        try:
            count = elements.Count
            if count == 0:
                return
        except:
            return
        for i in range(count):
            try:
                elem = elements.Item(i)
            except:
                # 索引从 0 开始不成功则尝试从 1 开始
                try:
                    elem = elements.Item(i + 1)
                except:
                    break
            try:
                name = elem.Name
            except:
                continue
            # 跳过 Aspen 内部节点
            if name.startswith("#") or name.startswith("@"):
                continue
            full_path = f"\\Data\\{base_path}\\{name}"
            # 读取值
            val = None
            try:
                val = elem.Value
            except:
                pass
            # 读取数据类型（多策略）
            data_type = self._get_data_type(elem)
            # 读取单位（多策略）
            unit = self._get_unit_string(elem, tree)
            # 判断是否为数值型
            is_numeric = data_type.upper() in ("", "REAL", "INTEGER", "BINARY")
            if is_numeric:
                try:
                    float(val) if val is not None else None
                except (ValueError, TypeError):
                    is_numeric = False
            if val is not None:
                # 有值 → 加入结果
                results.append((name, full_path, val, unit, data_type, is_numeric))
                # 如果也有子节点（如 TEMP 既有值又有 MIXED），也递归展开
                try:
                    if elem.Elements.Count > 0:
                        self._walk_params_on_worker(
                            elem, f"{base_path}\\{name}", results, max_depth, depth + 1, tree
                        )
                except:
                    pass
            else:
                # 无值 → 尝试深入递归（可能是Elements访问失败的容器节点）
                # 递归函数内部会自己处理 Elements 异常，安全
                self._walk_params_on_worker(
                    elem, f"{base_path}\\{name}", results, max_depth, depth + 1, tree
                )
        return

    def get_stream_inputs(self, stream_name):
        """枚举流股 Input 子节点下所有参数（无过滤）"""
        def _do_get(worker):
            inputs = []
            try:
                input_node = worker.tree.FindNode(f"\\Data\\Streams\\{stream_name}\\Input")
                if not input_node:
                    return inputs
                self._walk_params_on_worker(
                    input_node, f"Streams\\{stream_name}\\Input", inputs, max_depth=2, tree=worker.tree
                )
            except Exception as e:
                log.error(f"获取流股输入变量失败 [{stream_name}]: {e}")
            return inputs
        try:
            return self._worker.execute(_do_get)
        except Exception as e:
            log.error(f"get_stream_inputs 失败: {e}")
            return []

    def get_stream_outputs(self, stream_name):
        """枚举流股 Output 子节点下的所有输出参数"""
        def _do_get(worker):
            outputs = []
            try:
                output_node = worker.tree.FindNode(f"\\Data\\Streams\\{stream_name}\\Output")
                if not output_node:
                    return outputs
                self._walk_params_on_worker(
                    output_node, f"Streams\\{stream_name}\\Output", outputs, max_depth=2, tree=worker.tree
                )
            except Exception as e:
                log.error(f"获取流股输出变量失败 [{stream_name}]: {e}")
            return outputs
        try:
            return self._worker.execute(_do_get)
        except Exception as e:
            log.error(f"get_stream_outputs 失败: {e}")
            return []

    def enumerate_params(self, kind, name):
        """通用枚举——kind='block'/'stream', 枚举 Input 和 Output 下所有参数"""
        prefix = "Blocks" if kind == "block" else "Streams"
        base = f"\\Data\\{prefix}\\{name}"
        def _do_get(worker):
            results = []
            for sub in ("Input", "Output"):
                try:
                    node = worker.tree.FindNode(f"{base}\\{sub}")
                    if node:
                        self._walk_params_on_worker(
                            node, f"{prefix}\\{name}\\{sub}", results, max_depth=3, tree=worker.tree
                        )
                except Exception as e:
                    log.error(f"枚举{kind} {sub}参数失败 [{name}]: {e}")
            return results
        try:
            return self._worker.execute(_do_get)
        except:
            return []

    # ──────────────────────────────────────────────
    # 设置变量与运行
    # ──────────────────────────────────────────────

    def set_variable(self, path, value):
        r"""设置 Aspen 变量值。自动处理容器节点（如 \...\TEMP 需穿透到 MIXED）"""
        def _do_set(worker):
            try:
                node = worker.tree.FindNode(path)
                if not node:
                    log.warning(f"set_variable: 未找到节点: {path}")
                    return False
                # 如果节点无值但有子节点，穿透到第一个子节点（如 TEMP→MIXED）
                try:
                    nv = node.Value
                except:
                    nv = None
                if nv is None:
                    try:
                        elements = node.Elements
                        if elements.Count > 0:
                            try:
                                child = elements.Item(0)
                            except:
                                child = elements.Item(1)
                            if child and child.Value is not None:
                                node = child
                    except:
                        pass
                node.Value = float(value)
                log.info(f"set_variable: {path} = {value}")
                return True
            except Exception as e:
                log.error(f"设置变量失败 [{path} = {value}]: {type(e).__name__}: {e}")
                return False
        try:
            return self._worker.execute(_do_set)
        except:
            return False

    def get_variable(self, path):
        """读取指定路径节点的值（精确路径，不兜底避免触发授权崩溃）"""
        def _do_get(worker):
            try:
                node = worker.tree.FindNode(path)
                if node:
                    return node.Value
                return None
            except:
                return None
        try:
            return self._worker.execute(_do_get)
        except:
            return None

    def _read_node_val_unit_dtype(self, worker, node):
        """读取单个节点的值/单位/数据类型，如果是容器（有子节点但无值），自动读第一个子节点"""
        val = None
        try:
            val = node.Value
        except:
            pass
        # 如果有子节点且无值，容器节点 → 读第一个子节点（通常是 MIXED）
        if val is None:
            try:
                elements = node.Elements
                if elements.Count > 0:
                    child = elements.Item(0)
                    try:
                        val = child.Value
                        unit = self._get_unit_string(child, worker.tree)
                        dtype = self._get_data_type(child)
                        return val, unit, dtype
                    except:
                        pass
            except:
                pass
        unit = self._get_unit_string(node, worker.tree)
        dtype = self._get_data_type(node)
        return val, unit, dtype

    def get_node_info(self, path):
        """读取指定路径节点的（值, 单位, 数据类型）"""
        def _do_get(worker):
            try:
                node = worker.tree.FindNode(path)
                if not node:
                    log.warning(f"get_node_info: 未找到节点 [{path}]")
                    return None, "", ""
                return self._read_node_val_unit_dtype(worker, node)
            except Exception as e:
                log.error(f"读取节点信息失败 [{path}]: {e}")
                return None, "", ""
        try:
            return self._worker.execute(_do_get)
        except:
            return None, "", ""

    def reinit(self):
        """重新初始化模拟（不运行），用于失败后重置恢复"""
        def _do_reinit(worker):
            try:
                worker.app.Reinit()
                log.info("模拟已重新初始化 (Reinit)")
                return True
            except Exception as e:
                log.error(f"Reinit 失败: {e}")
                return False
        try:
            return self._worker.execute(_do_reinit)
        except Exception as e:
            log.error(f"Reinit 失败: {e}")
            return False

    def run_simulation(self, timeout=300):
        """运行 Aspen Plus 模拟，返回 (success, message, engine_status)

        Aspen V14 的 Document.Run() 是同步阻塞的，模拟完成或失败后才返回，
        不需要手动轮询 EngineStatus（Document 对象上 EngineStatus 不可用）。
        engine_status: 2=converged(success), 3=simulation_error, 4=server_crash(need reconnect)
        """
        def _do_run(worker):
            if not worker.is_connected:
                return False, "未连接到 Aspen Plus", None
            try:
                # 短暂延迟，让 Aspen 消化完之前的变量设置
                time.sleep(0.5)
                log.info("调用 Run()（同步等待模拟完成）...")
                t0 = time.time()
                worker.app.Run()
                elapsed = time.time() - t0
                log.info(f"Run() 完成，耗时 {elapsed:.1f}s")
                return True, f"模拟完成（{elapsed:.1f}s）", 2
            except Exception as e:
                err_msg = f"{type(e).__name__}: {e}"
                log.error(f"Run() 失败: {err_msg}")
                # RPC_E_SERVERFAULT = 0x80010105 → COM 服务器崩溃，需要重连
                is_server_crash = False
                try:
                    hr = getattr(e, 'hresult', 0)
                    if hr & 0xFFFFFFFF == 0x80010105:
                        is_server_crash = True
                        # 标记 worker 需要重置
                        worker.is_connected = False
                        log.error("检测到 COM 服务器崩溃 (RPC_E_SERVERFAULT)，Aspen 进程可能已终止")
                except:
                    pass
                eng = 4 if is_server_crash else 3
                return False, err_msg, eng
        try:
            return self._worker.execute(_do_run, timeout=timeout + 30)
        except Exception as e:
            return False, f"模拟运行失败: {e}", None

    def reconnect_and_reload(self, filepath=None):
        """COM 服务器崩溃后恢复：重连 + 重新加载 .bkp 文件"""
        fp = filepath or self.loaded_filepath
        if not fp:
            return False, "崩溃恢复失败: 未记录已加载的文件路径"
        log.info(f"开始崩溃恢复流程，文件: {fp}")
        self.close()
        self._reset_worker()
        ok, msg = self.connect()
        if not ok:
            return False, f"重连失败: {msg}"
        ok, msg = self.load_file(fp)
        if not ok:
            return False, f"重连后加载文件失败: {msg}"
        log.info("崩溃恢复完成，连接和文件均已恢复")
        return True, "已恢复连接并重新加载文件"

    @staticmethod
    def input_path_to_output_path(input_path):
        """将输入变量路径转换为输出变量路径。

        Aspen Plus 输出属性名是输入属性名 + _OUT 后缀。
        例: \\Data\\Streams\\S3\\Input\\TEMP\\MIXED
          → \\Data\\Streams\\S3\\Output\\TEMP_OUT\\MIXED
        """
        # 路径格式: \\Data\\Streams\\{stream}\\Input\\{prop}\\{subphase}
        # 转为:      \\Data\\Streams\\{stream}\\Output\\{prop}_OUT\\{subphase}
        parts = input_path.replace("\\", "/").strip("/").split("/")
        if len(parts) >= 6 and parts[3].upper() == "INPUT":
            prop = parts[4]
            parts[3] = "Output"
            parts[4] = prop + "_OUT"
            return "/" + "/".join(parts)
        return input_path  # 格式不匹配则原样返回

    # ──────────────────────────────────────────────
    # 批量工具
    # ──────────────────────────────────────────────

    @staticmethod
    def generate_combinations(var_configs):
        if not var_configs:
            return []
        from itertools import product
        ranges = []
        for path, min_val, max_val, step in var_configs:
            if step <= 0:
                values = [min_val]
            else:
                count = int((max_val - min_val) / step) + 1
                values = [min_val + i * step for i in range(count)]
                if values[-1] < max_val - 1e-10:
                    values.append(max_val)
            ranges.append((path, values))
        paths = [r[0] for r in ranges]
        value_sets = [r[1] for r in ranges]
        return [dict(zip(paths, combo)) for combo in product(*value_sets)]
