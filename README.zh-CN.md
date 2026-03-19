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
- [`example/`](example/)
  - 示例 `INI` 数据与样例资源
- [`rundata/`](rundata/)
  - 运行生成的输出文件

## 使用方法

### 1. `INI` 转 `Excel`

运行 [`script/ini_to_excel.py`](script/ini_to_excel.py) 中的转换脚本。

### 2. `Excel` 转 `INI`

运行 [`script/excel_to_ini.py`](script/excel_to_ini.py) 中的反向转换脚本。

### 3. 图形界面模式

运行 [`script/gui.py`](script/gui.py) 启动图形界面。

### 4. `w3x2lni` 辅助流程

当你已安装 [`w3x2lni`](https://github.com/sumneko/w3x2lni) 并在工具中指定其安装目录后，项目即可调用其相关能力，支持 `.w3x` 相关转换流程。

## 说明

- 本项目专注于 Warcraft III 物体数据编辑流程。
- `w3x2lni` 相关功能仅在工具中配置了正确的安装路径后可用。
- 生成的文件会保存在 [`rundata/output/`](rundata/output/)。

## 许可证

本项目采用 [GPL-3.0 许可证](https://www.gnu.org/licenses/gpl-3.0.en.html)。
