## App Power Switcher

### 项目简介

App Power Switcher 是一个 Windows 操作系统上的后台应用程序，它能够实时监控当前处于前台的活动窗口，并根据预设的应用程序列表和对应的电源计划映射关系，自动调整系统的电源计划。这有助于在运行诸如游戏、视频编辑软件等需要高性能的应用时自动切换到高性能模式，而在进行日常办公、浏览网页时切换回平衡或节能模式，从而优化性能和电池续航。

应用程序作为一个后台服务运行，并在任务栏通知区域提供一个简单的图标，用户可以通过右键菜单或双击图标来快速退出程序。

### 功能特点

*   **事件驱动的窗口监控:** 使用 Windows API (通过 `ctypes`) 监听窗口焦点变化事件，实现高效、低资源占用的实时前台应用检测。
*   **基于 GUID 的电源计划管理:** 通过调用 Windows `powercfg` 命令行工具来切换电源计划，使用电源计划的 GUID 作为可靠标识符，避免因本地化名称导致的兼容性问题。调用 `powercfg` 时隐藏命令行窗口。
*   **可配置的应用程序映射:** 通过 `.ini` 格式的配置文件灵活定义应用程序进程名与电源计划 GUID 的对应关系，以及默认电源计划和日志级别。
*   **后台运行:** 应用程序的核心逻辑在独立的后台线程中运行，不干扰用户正常操作。
*   **任务栏图标GUI:** 在任务栏通知区域提供图标，支持右键菜单（提供“退出”选项）和双击退出。
*   **详细日志记录:** 应用程序各关键环节（启动、配置加载、事件检测、电源计划切换、错误等）均有详细日志输出，便于问题诊断。

### 系统要求

*   操作系统: Windows 10 或更新版本 (一些 Windows API 在早期版本可能不支持或行为有差异)。
*   Python: Python 3.10 或更新版本。
*   权限: **需要管理员权限才能切换电源计划。** 应用程序启动时应以管理员身份运行（例如通过配置计划任务以最高权限运行）。

### 安装

1.  **克隆或下载项目代码:**
    将项目的代码文件（包含 `main.py`, `src/`, `config/` 等）下载到您的本地目录，例如 `C:\Users\YourName\AppPowerSwitcher`。

2.  **创建 Python 虚拟环境 (推荐):**
    打开命令提示符 (Command Prompt) 或 PowerShell，导航到项目根目录：
    ```bash
    cd C:\Users\YourName\AppPowerSwitcher
    ```
    创建虚拟环境：
    ```bash
    python -m venv .venv
    ```
    激活虚拟环境：
    *   在 Command Prompt 中: `.venv\Scripts\activate`
    *   在 PowerShell 中: `.\.venv\Scripts\Activate.ps1`

3.  **安装依赖:**
    激活虚拟环境后，安装项目所需的依赖库。
    ```bash
    pip install -r requirements.txt
    ```


### 配置

1.  **编辑配置文件:**
    导航到项目根目录下的 `config` 文件夹，打开 `app_config.ini` 文件进行编辑。
    ```bash
    notepad config\app_config.ini
    ```
    根据文件中的注释说明修改 `[General]` 部分的默认电源计划 GUID 和日志级别，以及在 `[ProcessPowerMap]` 部分添加您的应用程序进程名与希望对应的电源计划 GUID 的映射。

2.  **获取电源计划 GUID:**
    要查找您系统上电源计划的 GUID，打开命令提示符或 PowerShell，运行以下命令：
    ```bash
    powercfg /list
    ```
    它会列出所有电源计划及其 GUID。您需要在 `app_config.ini` 中使用这些 GUID。

### 运行应用程序

由于应用程序需要以管理员权限运行才能切换电源计划，并且要实现静默启动，推荐使用以下方式：

1.  **创建静默启动 BAT 脚本:**
    在项目根目录 `AppPowerSwitcher/` 下创建一个名为 `start_silent.bat` 的文件，内容如下（请根据您的实际路径修改）：
    ```batch
    @echo off
    rem IMPORTANT: Replace "C:\Users\YourName\AppPowerSwitcher" with the actual full path to your project root directory.
    set "PROJECT_DIR=C:\Users\YourName\AppPowerSwitcher"

    rem Change directory to the project root. The /d is needed if the project is on a different drive.
    cd /d "%PROJECT_DIR%"

    rem IMPORTANT: Replace "C:\Users\YourName\AppPowerSwitcher\.venv\Scripts\pythonw.exe" with the actual full path to pythonw.exe in your virtual environment or Python installation.
    set "PYTHONW_PATH=%PROJECT_DIR%\.venv\Scripts\pythonw.exe"

    rem Start the main application script using pythonw.exe to run silently (no console window).
    rem Use start "" to ensure the Pythonw process doesn't inherit the BAT window handle, making the BAT return immediately.
    rem The "" is a dummy title.
    start "" "%PYTHONW_PATH%" main.py

    rem Exit the batch script immediately.
    exit /b 0
    ```
    将 `PROJECT_DIR` 和 `PYTHONW_PATH` 变量的值替换为您实际的路径。

2.  **手动测试运行 (需要管理员权限):**
    右键单击 `start_silent.bat` 文件，选择“以管理员身份运行”（Run as administrator）。
    您应该看不到命令行窗口出现。程序会在后台运行，并在任务栏通知区域显示图标。您可以通过Task Manager查看 `pythonw.exe` 进程是否在运行，并通过查看 `logs/app_power_switcher.log` 文件来确认程序是否正常启动并监控窗口切换。

### 开机自动启动 (通过计划任务)

1.  打开“任务计划程序”（Task Scheduler）。
2.  在右侧窗格中，点击“创建任务...”。
3.  在“创建任务”窗口中：
    *   **常规:**
        *   名称: `AppPowerSwitcher Silent Startup`
        *   安全选项: 选择一个用户账户，勾选“**使用最高权限运行**”，并选择“只在用户登录时运行”（或“运行用户帐户”，前者更适合需要GUI交互的应用）。
    *   **触发器:**
        *   点击“新建...”，选择“在用户登录时”。（可选）可以设置一个启动延迟，例如 5 分钟。点击“确定”。
    *   **操作:**
        *   点击“新建...”，操作选择“启动程序”。
        *   程序或脚本: 浏览选择您在项目根目录下的 `start_silent.bat` 文件。
        *   起始于(可选): 输入您项目的**根目录完整路径**（例如 `C:\Users\YourName\AppPowerSwitcher`）。点击“确定”。
    *   **条件:**
        *   根据需要调整电源条件（例如，是否在电池供电时也运行）。
    *   **设置:**
        *   确保“允许按需运行任务”已勾选。
        *   如果任务已经运行，则适用以下规则: 选择“不要启动新实例”。
4.  点击“确定”保存任务。系统可能要求输入密码。

现在，当您选择的用户登录时，该任务将以管理员权限自动静默启动您的应用程序。

### 任务栏图标使用

*   **悬停:** 将鼠标悬停在任务栏图标上，会显示 Tooltip 文本 (`App Power Switcher`)。
*   **左键单击:** 默认无操作 (可以在代码中添加显示主窗口等功能)。
*   **左键双击:** 默认会执行程序退出。
*   **右键单击:** 显示上下文菜单，目前仅包含一个“退出”选项。点击“退出”会关闭应用程序。

### 日志文件

程序的运行日志会输出到项目根目录下的 `logs` 文件夹中的 `app_power_switcher.log` 文件。当需要调试、查看应用程序状态或排查问题时，请检查此文件。您可以在 `config/app_config.ini` 中修改 `log_level` 来调整日志的详细程度。

### 故障排除

*   **程序未能启动/没有任务栏图标:**
    *   检查计划任务配置是否正确，特别是路径、用户账户和“使用最高权限运行”是否勾选。
    *   暂时修改 `start_silent.bat`，移除 `@echo off` 和 `start "" %PYTHONW_PATH%`，改为直接调用 `"%PYTHONW_PATH%" main.py`，然后在命令行窗口中运行 BAT 文件，查看是否有 Python 错误输出。
    *   检查 `logs/app_power_switcher.log` 文件是否有启动错误信息。
    *   确认 `requirements.txt` 中的依赖已经正确安装 (`pip install -r requirements.txt`)。
*   **电源计划切换失败:**
    *   确认应用程序是以**管理员权限**运行的（通过计划任务或手动以管理员身份运行 BAT）。
    *   检查 `config/app_config.ini` 文件中的电源计划 GUID 是否正确（与 `powercfg /list` 输出一致）。
    *   检查 `logs/app_power_switcher.log` 文件是否有 `PowerCfg command failed` 或 `Hint: Power plan switching typically requires Administrator privileges` 等错误信息。
    *   尝试手动在命令提示符中运行 [`powercfg /setactive YOUR_GUID`] 命令，确认该命令本身是否能成功工作。
*   **任务栏右键菜单不显示或不工作:**
    *   确认 `pywin32` 库已正确安装。
    *   检查 `logs/app_power_switcher.log` 文件是否有与 Taskbar Icon 相关的错误信息。
*   **无法静默启动（依然弹出命令行窗口）:**
    *   确认 `start_silent.bat` 文件中使用了 `pythonw.exe` 而不是 `python.exe`。
    *   确认 `start "" "%PYTHONW_PATH%" main.py` 命令中的路径正确，并且使用了 `start ""` 前缀。
    *   确认 `PowerCfgManager` 中调用 `subprocess.run` 时添加了 `creationflags=win32con.CREATE_NO_WINDOW` 参数。

### 项目结构

```
AppPowerSwitcher/
├── config/                  # 配置文件存放目录
│   └── app_config.ini       # 应用程序配置文件
├── docs/                    # 文档目录 (当前为空)
├── logs/                    # 日志文件输出目录 (运行时自动创建)
├── src/                     # 源代码目录 (Python 包根)
│   ├── application/         # 应用层 - 核心业务逻辑
│   │   ├── __init__.py
│   │   └── power_switcher_app.py # 应用核心管理类
│   │
│   ├── infrastructure/      # 基础设施层 - 系统和外部交互
│   │   ├── __init__.py
│   │   │
│   │   ├── configuration/   # 配置管理
│   │   │   ├── __init__.py
│   │   │   └── config_manager.py # 配置文件加载和保存
│   │   │
│   │   ├── power_management/# 电源计划管理
│   │   │   ├── __init__.py
│   │   │   └── power_cfg_manager.py # powercfg 命令交互
│   │   │
│   │   └── windows/         # Windows API 交互 (ctypes)
│   │       ├── __init__.py
│   │       ├── event_listener.py # 窗口事件监听器
│   │       └── process_info.py   # 获取进程信息工具
│   │
│   └── utils/               # 通用工具模块
│       ├── __init__.py
│       └── logging_config.py # 日志配置工具
│
├── tests/                # 测试目录 (当前为空)
├── .venv/                # Python 虚拟环境目录 (如果使用)
├── main.py                 # 主程序入口脚本 (运行此文件)
├── requirements.txt        # 项目依赖列表
├── README.md               # 项目说明文件
└── start_silent.bat        # 静默启动脚本 (手动运行或计划任务使用)
└── app.ico                 # (可选) 任务栏图标文件
