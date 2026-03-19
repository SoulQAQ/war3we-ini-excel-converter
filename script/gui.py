#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
INI 与 Excel 文件互转工具
支持 .ini 配置文件与 .xlsx/.xls Excel 文件之间的相互转换
"""

import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import configparser
import os
import sys

# 导入转换模块
from ini_to_excel import ini_to_excel, get_unique_filename
from excel_to_ini import excel_to_ini


def load_config():
    """加载配置文件"""
    config = configparser.ConfigParser()
    config_path = os.path.join(os.path.dirname(__file__), '..', 'config', 'setting.cfg')
    config.read(config_path, encoding='utf-8')
    return config


def save_config(config):
    """保存配置文件"""
    config_path = os.path.join(os.path.dirname(__file__), '..', 'config', 'setting.cfg')
    with open(config_path, 'w', encoding='utf-8') as f:
        config.write(f)


def load_ini_names(config):
    """加载 INI 文件名称映射"""
    ini_names = {}
    if config.has_section('ini_names'):
        for key, value in config.items('ini_names'):
            ini_names[key.lower()] = value
    return ini_names


def to_absolute_path(path):
    """将相对路径转换为绝对路径"""
    if not path:
        return path
    if os.path.isabs(path):
        return path
    # 相对于工作目录（脚本的父目录）
    base_dir = os.path.join(os.path.dirname(__file__), '..')
    return os.path.normpath(os.path.join(base_dir, path))


def check_and_add_table_folder(folder_path):
    """
    检查文件夹内是否有 table 和 w3x2lni 文件夹，如果有则自动添加 table 层级
    
    Args:
        folder_path: 原始文件夹路径
        
    Returns:
        str: 处理后的路径
    """
    if not folder_path or not os.path.isdir(folder_path):
        return folder_path
    
    table_path = os.path.join(folder_path, 'table')
    w3x2lni_path = os.path.join(folder_path, 'w3x2lni')
    
    # 如果同时存在 table 和 w3x2lni 文件夹，且 table 是文件夹
    if os.path.isdir(table_path) and os.path.isdir(w3x2lni_path):
        return table_path
    
    return folder_path


class ConverterApp:
    """转换工具主应用类"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("INI-Excel 转换工具")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        # 居中窗口
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - 600) // 2
        y = (screen_height - 450) // 2
        root.geometry(f"600x450+{x}+{y}")
        
        # 加载配置
        self.config = load_config()
        self.ini_names = load_ini_names(self.config)
        
        # 变量
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.output_filename = tk.StringVar()
        self.conversion_type = tk.StringVar(value="ini_to_excel")
        
        # 恢复上次设置
        self._restore_settings()
        
        self.create_widgets()
    
    def _restore_settings(self):
        """恢复上次的用户设置"""
        # 从配置文件中读取用户设置
        if self.config.has_section('user_settings'):
            if 'input_path' in self.config['user_settings']:
                self.input_path.set(self.config['user_settings']['input_path'])
            if 'output_path' in self.config['user_settings']:
                self.output_path.set(self.config['user_settings']['output_path'])
            if 'output_filename' in self.config['user_settings']:
                self.output_filename.set(self.config['user_settings']['output_filename'])
            else:
                self.output_filename.set("output")
            if 'conversion_type' in self.config['user_settings']:
                self.conversion_type.set(self.config['user_settings']['conversion_type'])
        else:
            # 默认输出文件名（不带后缀）
            self.output_filename.set("output")
    
    def _save_settings(self):
        """保存当前用户设置到配置文件"""
        if not self.config.has_section('user_settings'):
            self.config.add_section('user_settings')
        
        self.config['user_settings']['input_path'] = self.input_path.get()
        self.config['user_settings']['output_path'] = self.output_path.get()
        self.config['user_settings']['output_filename'] = self.output_filename.get()
        self.config['user_settings']['conversion_type'] = self.conversion_type.get()
        
        save_config(self.config)
    
    def create_widgets(self):
        """创建界面组件"""
        # 主内容区
        main_frame = tk.Frame(self.root, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 转换类型选择
        type_frame = tk.LabelFrame(main_frame, text="转换类型", padx=10, pady=10)
        type_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Radiobutton(
            type_frame,
            text="INI 转 Excel",
            variable=self.conversion_type,
            value="ini_to_excel",
            command=self.on_type_changed
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Radiobutton(
            type_frame,
            text="Excel 转 INI",
            variable=self.conversion_type,
            value="excel_to_ini",
            command=self.on_type_changed
        ).pack(side=tk.LEFT, padx=10)
        
        # 输入路径
        input_frame = tk.LabelFrame(main_frame, text="输入路径", padx=10, pady=10)
        input_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.input_entry = tk.Entry(input_frame, textvariable=self.input_path, width=50)
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.input_btn = tk.Button(
            input_frame,
            text="选择文件夹",
            command=self.select_input_folder,
            width=12
        )
        self.input_btn.pack(side=tk.LEFT, padx=(10, 5))
        
        self.input_file_btn = tk.Button(
            input_frame,
            text="选择文件",
            command=self.select_input_file,
            width=10
        )
        self.input_file_btn.pack(side=tk.LEFT)
        
        # 输出路径
        output_frame = tk.LabelFrame(main_frame, text="输出设置", padx=10, pady=10)
        output_frame.pack(fill=tk.X, pady=(0, 15))
        
        output_path_frame = tk.Frame(output_frame)
        output_path_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.output_entry = tk.Entry(output_path_frame, textvariable=self.output_path, width=50)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        output_browse_btn = tk.Button(
            output_path_frame,
            text="浏览",
            command=self.select_output_folder,
            width=10
        )
        output_browse_btn.pack(side=tk.LEFT, padx=(10, 5))
        
        filename_frame = tk.Frame(output_frame)
        filename_frame.pack(fill=tk.X)
        
        tk.Label(filename_frame, text="文件名:").pack(side=tk.LEFT)
        
        self.filename_entry = tk.Entry(filename_frame, textvariable=self.output_filename, width=40)
        self.filename_entry.pack(side=tk.LEFT, padx=(10, 10))
        
        # 扩展名标签
        self.ext_label = tk.Label(filename_frame, text=".xlsx")
        self.ext_label.pack(side=tk.LEFT)
        
        # 按钮区
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.convert_btn = tk.Button(
            btn_frame,
            text="开始转换",
            command=self.do_convert,
            bg="#27ae60",
            fg="white",
            font=("Microsoft YaHei", 11, "bold"),
            width=20,
            height=2
        )
        self.convert_btn.pack(side=tk.LEFT, padx=10)
        
        tk.Button(
            btn_frame,
            text="退出",
            command=self._on_closing,
            width=12,
            height=2
        ).pack(side=tk.LEFT, padx=10)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_frame = tk.Frame(self.root, bg="#ecf0f1", height=35)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)
        
        status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg="#ecf0f1",
            fg="#7f8c8d",
            font=("Microsoft YaHei", 9)
        )
        status_label.pack(pady=8)
        
        # 窗口关闭时保存设置
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _on_closing(self):
        """窗口关闭时保存设置"""
        self._save_settings()
        self.root.quit()
    
    def on_type_changed(self):
        """转换类型改变时的处理"""
        if self.conversion_type.get() == "ini_to_excel":
            self.ext_label.config(text=".xlsx")
        else:
            self.ext_label.config(text=".ini")
    
    def select_input_folder(self):
        """选择输入文件夹"""
        folder = filedialog.askdirectory(title="选择 INI 文件夹")
        if folder:
            # 自动检测并添加 table 层级
            folder = check_and_add_table_folder(folder)
            self.input_path.set(folder)
    
    def select_input_file(self):
        """选择输入文件"""
        if self.conversion_type.get() == "ini_to_excel":
            filetypes = [("INI 文件", "*.ini"), ("所有文件", "*.*")]
            title = "选择 INI 文件"
        else:
            filetypes = [("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
            title = "选择 Excel 文件"
        
        file = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if file:
            self.input_path.set(file)
    
    def select_output_folder(self):
        """选择输出文件夹"""
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_path.set(folder)
    
    def do_convert(self):
        """执行转换"""
        input_path = self.input_path.get()
        output_path = self.output_path.get()
        filename = self.output_filename.get()
        
        if not input_path:
            messagebox.showwarning("警告", "请选择输入文件/文件夹")
            return
        
        if not os.path.exists(input_path):
            messagebox.showerror("错误", f"输入路径不存在：{input_path}")
            return
        
        # 确保输出目录存在
        if not os.path.exists(output_path):
            os.makedirs(output_path)
        
        # 构建输出文件路径（不带后缀，由 ext_label 显示）
        ext = '.xlsx' if self.conversion_type.get() == "ini_to_excel" else '.ini'
        output_file = os.path.join(output_path, filename + ext)
        
        # 处理文件名冲突
        output_file = get_unique_filename(output_file)
        # 不更新输入框中的文件名
        
        try:
            if self.conversion_type.get() == "ini_to_excel":
                self.status_var.set("正在转换 INI 到 Excel...")
                self.root.update()
                ini_to_excel(input_path, output_file, self.ini_names)
                messagebox.showinfo("完成", f"Excel 文件已创建:\n{output_file}")
            else:
                self.status_var.set("正在转换 Excel 到 INI...")
                self.root.update()
                excel_to_ini(input_path, output_file)
                messagebox.showinfo("完成", f"INI 文件已创建:\n{output_file}")
            
            self.status_var.set("转换完成")
        except Exception as e:
            messagebox.showerror("错误", f"转换失败：{str(e)}")
            self.status_var.set("转换失败")


def show_welcome_window():
    """显示主窗口"""
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()


def main():
    """主函数"""
    # 加载配置
    config = load_config()
    
    # 获取路径配置
    input_path = config.get('paths', 'input_path', fallback='./rundata/input')
    out_path = config.get('paths', 'out_path', fallback='./rundata/output')
    
    print(f"输入路径：{input_path}")
    print(f"输出路径：{out_path}")
    
    # 显示主窗口
    show_welcome_window()


if __name__ == "__main__":
    main()
