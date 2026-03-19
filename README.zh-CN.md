# Warcraft III 物体数据转换工具

[English](README.md) | [中文](README.zh-CN.md)

这是一个使用 Python 开发的工具，用于在 Warcraft III 的 `LNI` 物体数据 / 地图物体数据 与 `Excel` 表格之间进行互相转换，方便使用者更高效地修改和维护物体数据。

## 功能特性

- 将 Warcraft III 物体数据从 `INI` 转换为 `Excel`
- 将编辑后的 `Excel` 转换回 `INI`
- 便于在表格中整理、筛选和批量修改物体数据
- 计划支持在本地安装 [`w3x2lni`](https://github.com/sumneko/w3x2lni) 后，直接调用它实现 `.w3x` 与 `Excel` 之间的转换

## 项目目标

本项目的目标是让 Warcraft III 地图制作者和物体数据编辑者能够更方便地处理数据。

典型工作流如下：

1. 将地图物体数据导出为 `LNI` 或 `INI`
2. 在 `Excel` 中编辑数据
3. 将修改后的数据导回 Warcraft III 可用格式
4. 如需直接处理 `.w3x` 地图，可配合 [`w3x2lni`](https://github.com/sumneko/w3x2lni) 使用，进一步简化流程

## 当前状态

本仓库仍在持续开发中。

- `INI` → `Excel` 转换已可使用
- `Excel` → `INI` 转换计划中 / 开发中
- 图形界面位于 [`script/gui.py`](script/gui.py)

## 仓库结构

- [`script/`](script/)
  - 核心 Python 转换脚本
- [`config/`](config/)
  - 应用配置文件
- [`example/`](example/)
  - 示例 `INI` 数据与样例资源
- [`rundata/`](rundata/)
  - 运行生成的输出文件
- [`w3x2lni/`](w3x2lni/)
  - 内置的上游 [`w3x2lni`](https://github.com/sumneko/w3x2lni) 参考副本

## 使用方法

### 1. `INI` 转 `Excel`

运行 [`script/ini_to_excel.py`](script/ini_to_excel.py) 中的转换脚本。

### 2. `Excel` 转 `INI`

运行 [`script/excel_to_ini.py`](script/excel_to_ini.py) 中的反向转换脚本。

### 3. 图形界面模式

运行 [`script/gui.py`](script/gui.py) 启动图形界面。

## 说明

- 本项目专注于 Warcraft III 物体数据编辑流程。
- 如果你希望使用 `.w3x` 直接转换，请先在本地安装 [`w3x2lni`](https://github.com/sumneko/w3x2lni)，并按项目预期路径放置后再使用计划中的集成功能。
- 生成的文件会保存在 [`rundata/output/`](rundata/output/)。

## 许可证

本项目采用 [GPL-3.0 许可证](https://www.gnu.org/licenses/gpl-3.0.en.html)。
