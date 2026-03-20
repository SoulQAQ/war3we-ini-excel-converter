# Warcraft III 物体数据转换工具

[English](README.md) | [中文](README.zh-CN.md)

这是一个使用 Python 开发的工具，用于在 Warcraft III 的物体数据与 `Excel` 表格之间进行互相转换，方便使用者更高效地修改和维护物体数据。

## 功能特性

- 将 Warcraft III 物体数据从 `INI` 转换为 `Excel`
- 将编辑后的 `Excel` 转换回 `INI`
- 便于在表格中整理、筛选和批量修改物体数据
- 当在工具中指定并配置 [`w3x2lni`](https://github.com/sumneko/w3x2lni) 的安装目录后，可使用其相关 `.w3x` 转换功能

## 项目目标

本项目的目标是让 Warcraft III 地图制作者和物体数据编辑者能够更方便地处理数据。

典型工作流如下：

1. 将地图物体数据导出为 `LNI` 或 `INI`
2. 在 `Excel` 中编辑数据
3. 将修改后的数据导回 Warcraft III 可用格式
4. 如果已安装并在工具中配置 [`w3x2lni`](https://github.com/sumneko/w3x2lni)，即可支持相关 `.w3x` 转换流程

## 当前状态

本仓库仍在持续开发中。

- `INI` → `Excel` 转换已可使用
- `Excel` → `INI` 转换计划中 / 开发中
- 图形界面位于 [`script/gui.py`](script/gui.py)
- `w3x2lni` 功能依赖于在工具中配置正确的安装路径

## 仓库结构

- [`script/`](script/)
  - 核心 Python 转换脚本
- [`config/`](config/)
  - 应用配置文件
- [`webui/`](webui/)
  - 桌面图形界面使用的前端文件
- [`build.bat`](build.bat)
  - 使用 PyInstaller 打包 Windows 可执行文件
- [`setup.bat`](setup.bat)
  - 创建 Python 虚拟环境并安装依赖
- [`start.bat`](start.bat)
  - 以源码方式启动图形界面

## 使用方法

### 1. `INI` 转 `Excel`

运行 [`script/ini_to_excel.py`](script/ini_to_excel.py) 中的转换脚本。

### 2. `Excel` 转 `INI`

运行 [`script/excel_to_ini.py`](script/excel_to_ini.py) 中的反向转换脚本。

### 3. 图形界面模式

运行 [`script/gui.py`](script/gui.py) 启动图形界面。

### 4. `w3x2lni` 辅助流程

当你已安装 [`w3x2lni`](https://github.com/sumneko/w3x2lni) 并在工具中指定其安装目录后，项目即可调用其相关能力，支持 `.w3x` 相关转换流程。

## 自行编译 / 构建教程

当前项目主要面向 Windows 环境，并且已经提供了用于环境准备、本地启动和打包的批处理脚本。

### 环境要求

- Windows 10 / 11
- Python `3.12`
- 已安装 Python Launcher for Windows（终端中可使用 `py` 命令）
- 可访问网络，用于安装 [`requirements.txt`](requirements.txt) 中的依赖

[`setup.bat`](setup.bat) 会明确检查 Python `3.12`，创建 `.venv`，并安装以下依赖：

- `pyyaml`
- `openpyxl`
- `pywebview`

### 1. 获取源码

克隆本仓库，或直接下载源码压缩包并解压到本地目录。

### 2. 准备 Python 环境

在项目根目录运行：

```bat
setup.bat
```

该脚本会自动执行以下操作：

1. 检查 `py` 启动器是否存在
2. 检查是否安装了 Python `3.12`
3. 在需要时创建或重建 [`.venv/`](.venv/)
4. 升级 `pip`、`setuptools` 和 `wheel`
5. 安装 [`requirements.txt`](requirements.txt) 中列出的依赖

### 3. 以源码方式启动程序

如果要直接从源码启动桌面界面，运行：

```bat
start.bat
```

[`start.bat`](start.bat) 会在虚拟环境不存在，或者 [`requirements.txt`](requirements.txt) 发生变化时，自动重新调用 [`setup.bat`](setup.bat)。

如果你想手动启动，也可以执行：

```bat
.venv\Scripts\python.exe script\gui.py
```

### 4. 打包生成可执行文件

如果要打包为独立的 Windows 可执行文件，运行：

```bat
build.bat
```

[`build.bat`](build.bat) 会自动完成以下步骤：

1. 调用 [`setup.bat`](setup.bat)
2. 在虚拟环境中安装 `pyinstaller`
3. 清理旧的 `build/` 与 `dist/` 输出
4. 将 [`script/gui.py`](script/gui.py) 打包为单文件窗口程序
5. 把 [`webui/`](webui/) 和 [`config/`](config/) 一并打入最终程序

生成的可执行文件位置为：

- [`dist/INI-Excel转换工具.exe`](dist/INI-Excel转换工具.exe)

### 5. 运行时目录与配置说明

- 主配置文件：[`config/setting.yaml`](config/setting.yaml)
- 桌面界面入口：[`webui/index.html`](webui/index.html)
- 程序默认使用相对路径，如 [`script/gui.py`](script/gui.py) 中定义的 `./rundata/input` 与 `./rundata/output`

如果这些目录还不存在，建议在测试转换流程前手动创建。

## 说明

- 本项目专注于 Warcraft III 物体数据编辑流程。
- `w3x2lni` 相关功能仅在工具中配置了正确的安装路径后可用。
- 生成的文件会保存在 [`rundata/output/`](rundata/output/)。

## 许可证

本项目采用 [GPL-3.0 许可证](https://www.gnu.org/licenses/gpl-3.0.en.html)。
