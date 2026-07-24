# Aspen Plus 自动调参工具

## 项目概述

通过 Python COM 接口自动化驱动 Aspen Plus 批量模拟，替代人工手动修改参数→运行→记录结果的重复劳动。配置输入变量范围和输出变量，一键运行所有参数组合，结果自动保存为 CSV。

长期目标：构建一套**流程模拟 → 数据采集 → 代理模型 → 多目标优化**的自动化闭环系统。

## 快速开始

### 方式一：运行 EXE（推荐）

从 `dist/` 运行 `Aspen自动化调参工具.exe`，无需安装 Python。

**环境要求**：
- Windows 系统
- Aspen Plus V14 已安装并正常运行

### 方式二：源码运行

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 openpyxl
python src/ui_app.py
```

## 操作流程

```
打开 Aspen → 加载 .bkp → 运行至收敛
          ↓
点「连接 Aspen Plus」（自动关联已有进程）
          ↓
左侧浏览器 → 展开模块/流股 → 双击变量添加
          ↓
配置输入变量范围（最小值/最大值/步长）
选择输出变量
          ↓
点「开始批量模拟」→ 自动遍历参数组合
          ↓
结果保存为 CSV
```

### 关键说明

- **Aspen 由你手动关闭**：程序不管理 Aspen 生命周期，退出时不会自动关闭 Aspen Plus
- **连接方式**：优先 `GetActiveObject` 关联已有进程，失败则 `Dispatch` 启动新实例
- **变量路径**：输入变量路径为 `\Data\Streams\{流股}\Input\{变量}\MIXED`，输出变量自动转为 `\Output\{变量}_OUT`
- **单位解析**：通过 `\Unit Table` 节点的物理量代号和单位代号自动匹配单位

## 项目结构

```
Surrogate-model-modeling/
├── src/                          # 源代码
│   ├── aspen_interface.py        # Aspen COM 接口封装（线程安全）
│   └── ui_app.py                 # 图形界面（tkinter，1280x800）
├── simulation/                   # Aspen Plus 模拟文件
│   ├── 案例2.bkp                 # 测试案例
│   ├── POP-Seperation*           # POP 分离工艺
│   └── aspen_config.log          # 配置导出示例
├── ico/                          # 应用程序图标（多分辨率）
├── dist/                         # 打包输出
│   └── Aspen自动化调参工具.exe
├── docs/
│   ├── 开发日志.md               # 开发过程记录
│   └── AspenWithPython-参考资料/  # Aspen COM 接口参考文档
├── 测试工程/                     # 诊断/测试脚本
├── Aspen自动化调参工具.spec       # PyInstaller 打包配置
├── version_info.txt              # EXE 版本信息（1.0.0）
└── README.md
```

## 技术栈

| 类别         | 工具/库                            |
| ------------ | ---------------------------------- |
| 流程模拟     | Aspen Plus V14                     |
| 自动化接口   | Python `win32com.client` (pywin32) |
| COM 线程模型 | 专用 `_AspenWorker` 守护线程       |
| 图形界面     | tkinter / ttk（蓝白配色）          |
| 结果存储     | CSV                                |
| 打包分发     | PyInstaller 6.17                   |

## 打包

```bash
pyinstaller "Aspen自动化调参工具.spec" --noconfirm
```

输出：`dist/Aspen自动化调参工具.exe`（约 33 MB，含图标 + 版本信息）

## 核心设计

### COM 工作线程隔离

所有 Aspen COM 操作在专用守护线程（`_AspenWorker`）上串行执行，通过消息队列 + 阻塞等待对外暴露线程安全的方法。彻底避免跨线程 COM 调用导致的 Access Violation 崩溃。

### COM 连接策略

1. `GetActiveObject("Apwn.Document.38.0")` — 连接已有 Aspen 进程
2. `Dispatch("Apwn.Document.38.0")` — 启动新 Aspen 实例

### 单位解析

通过参数节点的 `HAP_UNITROW`（物理量代号）和 `HAP_UNITCOL`（单位代号）在 `\Unit Table` 中匹配实际单位名。

### 同步模拟

Aspen V14 的 `Document.Run()` 是同步阻塞方法，调用完成后模拟即结束，无需轮询 EngineStatus。

## 当前进度

- [x] Aspen COM 接口封装（GetActiveObject + Dispatch 双模式）
- [x] 图形界面（变量浏览、参数配置、批量运行、CSV 导出）
- [x] 输入/输出参数区分 + 配置导入导出（JSON 格式 `.log` 文件）
- [x] 单次模拟循环（修改参数 → Run → 读取结果）
- [x] 批量模拟（自动遍历输入参数组合）
- [x] 单位自动解析（基于 Unit Table）
- [x] COM 工作线程隔离（防崩溃）
- [x] EXE 打包（图标 + 版本信息）
- [ ] 数据预处理与探索性分析
- [ ] 代理模型训练（GPR / ANN / SVR）
- [ ] 多目标优化（NSGA-II）

## License

Copyright (c) 2026 lijunsen & DeepSeek
