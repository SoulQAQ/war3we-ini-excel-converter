#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel 转 INI 工具
将 Excel 文件转换为 INI 文件
"""

import configparser
import os
import sys


def excel_to_ini(excel_path, output_path):
    """
    将 Excel 文件转换为 INI 文件
    
    Args:
        excel_path: Excel 文件路径
        output_path: 输出的 INI 文件路径或文件夹路径
    """
    # TODO: 实现 Excel 转 INI 功能
    print(f"Excel 转 INI 功能开发中...")
    print(f"输入：{excel_path}")
    print(f"输出：{output_path}")
    pass


if __name__ == "__main__":
    # 测试调用
    excel_to_ini("./test.xlsx", "./output.ini")
