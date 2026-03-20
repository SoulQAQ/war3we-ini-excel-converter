#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel 转 INI 工具
将 Excel 文件转换为 INI 文件
"""

import os
import sys
from openpyxl import load_workbook


MULTILINE_ELEMENT_SEPARATOR = '----'


def encode_ini_value(value):
    """将 Excel 单元格中的值编码为 INI 中使用的文本格式。"""
    if value is None:
        return ''

    text = str(value).replace('\r\n', '\n').replace('\r', '\n')
    stripped = text.strip()

    if stripped == '':
        return ''

    lines = text.split('\n')
    separator_indexes = [index for index, line in enumerate(lines) if line.strip() == MULTILINE_ELEMENT_SEPARATOR]

    if separator_indexes:
        parts = []
        current_part = []

        for line in lines:
            if line.strip() == MULTILINE_ELEMENT_SEPARATOR:
                parts.append('\n'.join(current_part).strip('\n'))
                current_part = []
            else:
                current_part.append(line)
        parts.append('\n'.join(current_part).strip('\n'))

        normalized_parts = [part for part in parts if part != '']
        if not normalized_parts:
            return '{\n}'

        encoded_parts = [f'[=[\n{part}\n]=]' for part in normalized_parts]
        return '{\n' + ',\n'.join(encoded_parts) + ',\n}'

    if '\n' in text:
        return f'[=[\n{text.strip("\n")}\n]=]'

    return text


def build_ini_lines_from_sheet(ws):
    """将单个工作表转换为 INI 文本行。"""
    lines = []
    max_column = ws.max_column
    max_row = ws.max_row

    property_columns = []
    for col in range(3, max_column + 1):
        prop_name = ws.cell(row=2, column=col).value
        if prop_name is None:
            continue
        prop_name = str(prop_name).strip()
        if not prop_name:
            continue

        comment = ws.cell(row=1, column=col).value
        comment = '' if comment is None else str(comment).strip()
        property_columns.append((col, prop_name, comment))

    for row in range(3, max_row + 1):
        object_id = ws.cell(row=row, column=1).value
        if object_id is None or str(object_id).strip() == '':
            continue

        object_id = str(object_id).strip()
        parent_id = ws.cell(row=row, column=2).value
        parent_id = '' if parent_id is None else str(parent_id).strip()

        lines.append(f'[{object_id}]')
        if parent_id:
            lines.append(f'_parent = "{parent_id}"')

        for col, prop_name, comment in property_columns:
            cell_value = ws.cell(row=row, column=col).value
            if cell_value is None or str(cell_value) == '':
                continue

            encoded_value = encode_ini_value(cell_value)
            if encoded_value == '':
                continue

            if comment:
                lines.append(f'-- {comment}')

            if '\n' in encoded_value:
                encoded_lines = encoded_value.split('\n')
                lines.append(f'{prop_name} = {encoded_lines[0]}')
                lines.extend(encoded_lines[1:])
            else:
                lines.append(f'{prop_name} = {encoded_value}')

        lines.append('')

    return lines


def excel_to_ini(excel_path, output_path):
    """
    将 Excel 文件转换为 INI 文件

    Args:
        excel_path: Excel 文件路径
        output_path: 输出的 INI 文件路径或文件夹路径
    """
    workbook = load_workbook(excel_path)

    if os.path.isdir(output_path):
        output_dir = output_path
        os.makedirs(output_dir, exist_ok=True)

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            lines = build_ini_lines_from_sheet(ws)
            ini_path = os.path.join(output_dir, f'{sheet_name}.ini')
            with open(ini_path, 'w', encoding='utf-8', newline='\n') as file:
                file.write('\n'.join(lines).rstrip() + '\n')
            print(f'INI 文件已创建：{ini_path}')
        return

    parent_dir = os.path.dirname(output_path)
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir)

    ws = workbook[workbook.sheetnames[0]]
    lines = build_ini_lines_from_sheet(ws)
    with open(output_path, 'w', encoding='utf-8', newline='\n') as file:
        file.write('\n'.join(lines).rstrip() + '\n')

    print(f'INI 文件已创建：{output_path}')


if __name__ == "__main__":
    # 测试调用
    excel_to_ini("./test.xlsx", "./output.ini")
