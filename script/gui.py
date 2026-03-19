#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
INI 与 Excel 文件互转工具
支持 .ini 配置文件与 .xlsx/.xls Excel 文件之间的相互转换
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import os
import sys
from pathlib import Path

try:
    import yaml
except ImportError as exc:
    raise RuntimeError("缺少 PyYAML 依赖，请先执行: pip install pyyaml") from exc

# 导入转换模块
from ini_to_excel import ini_to_excel, get_unique_filename
from excel_to_ini import excel_to_ini


BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE_DIR / 'config' / 'setting.yaml'


DEFAULT_CONFIG = {
    'ini_names': {
        'ability.ini': '技能',
        'buff.ini': '魔法效果',
        'imp.ini': '导入文件',
        'item.ini': '物品',
        'misc.ini': '平衡常数',
        'unit.ini': '单位',
        'upgrade.ini': '科技',
        'w3i.ini': '地图属性',
    },
    'user_settings': {
        'input_path': './rundata/input',
        'output_path': './rundata/output',
        'output_filename': 'output',
        'conversion_type': 'ini_to_excel',
    },
}


def normalize_relative_path(path_value):
    """将任意路径规范为相对于项目根目录的路径。"""
    if not path_value:
        return ''

    path_obj = Path(path_value)
    if not path_obj.is_absolute():
        path_obj = (BASE_DIR / path_obj).resolve()
    else:
        path_obj = path_obj.resolve()

    try:
        relative = path_obj.relative_to(BASE_DIR)
        return relative.as_posix() or '.'
    except ValueError:
        return os.path.relpath(path_obj, BASE_DIR).replace('\\', '/')


def resolve_config_path(path_value):
    """将配置中的相对路径解析为绝对路径。"""
    if not path_value:
        return ''
    return str((BASE_DIR / path_value).resolve())


def load_config():
    """加载 YAML 配置文件。"""
    if not CONFIG_PATH.exists():
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()

    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f) or {}

    config = {
        'ini_names': dict(DEFAULT_CONFIG['ini_names']),
        'user_settings': dict(DEFAULT_CONFIG['user_settings']),
    }
    config['ini_names'].update(data.get('ini_names', {}) or {})
    config['user_settings'].update(data.get('user_settings', {}) or {})
    return config


def save_config(config):
    """保存 YAML 配置文件。"""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        yaml.safe_dump(config, f, allow_unicode=True, sort_keys=False)


def load_ini_names(config):
    """加载 INI 文件名称映射。"""
    ini_names = {}
    for key, value in (config.get('ini_names') or {}).items():
        ini_names[str(key).lower()] = value
    return ini_names


def check_and_add_table_folder(folder_path):
    """
    检查文件夹内是否有 table 和 w3x2lni 文件夹，如果有则自动添加 table 层级
    """
    if not folder_path or not os.path.isdir(folder_path):
        return folder_path

    table_path = os.path.join(folder_path, 'table')
    w3x2lni_path = os.path.join(folder_path, 'w3x2lni')

    if os.path.isdir(table_path) and os.path.isdir(w3x2lni_path):
        return table_path

    return folder_path


class ConverterApp:
    """转换工具主应用类"""

    def __init__(self, root):
        self.root = root
        self.root.title("INI-Excel 转换工具")
        self.root.geometry("680x460")
        self.root.resizable(False, False)

        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - 680) // 2
        y = (screen_height - 460) // 2
        root.geometry(f"680x460+{x}+{y}")

        self.config = load_config()
        self.ini_names = load_ini_names(self.config)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.output_filename = tk.StringVar()
        self.conversion_type = tk.StringVar(value="ini_to_excel")

        self._restore_settings()
        self.create_widgets()

    def _restore_settings(self):
        """恢复上次的用户设置。"""
        user_settings = self.config.get('user_settings', {})
        self.input_path.set(user_settings.get('input_path', ''))
        self.output_path.set(user_settings.get('output_path', ''))
        self.output_filename.set(user_settings.get('output_filename', 'output'))
        self.conversion_type.set(user_settings.get('conversion_type', 'ini_to_excel'))

    def _save_settings(self):
        """保存当前用户设置到配置文件。"""
        self.config['user_settings'] = {
            'input_path': self.input_path.get(),
            'output_path': self.output_path.get(),
            'output_filename': self.output_filename.get().strip() or 'output',
            'conversion_type': self.conversion_type.get(),
        }
        save_config(self.config)

    def create_widgets(self):
        """创建界面组件。"""
        main_frame = tk.Frame(self.root, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

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

        input_frame = tk.LabelFrame(main_frame, text="输入路径（相对项目根目录）", padx=10, pady=10)
        input_frame.pack(fill=tk.X, pady=(0, 15))

        self.input_entry = tk.Entry(input_frame, textvariable=self.input_path, width=58, state='readonly', readonlybackground='white')
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

        output_frame = tk.LabelFrame(main_frame, text="输出设置", padx=10, pady=10)
        output_frame.pack(fill=tk.X, pady=(0, 15))

        output_path_frame = tk.Frame(output_frame)
        output_path_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(output_path_frame, text="输出目录:").pack(side=tk.LEFT)
        self.output_entry = tk.Entry(output_path_frame, textvariable=self.output_path, width=50, state='readonly', readonlybackground='white')
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        output_browse_btn = tk.Button(
            output_path_frame,
            text="浏览",
            command=self.select_output_folder,
            width=10
        )
        output_browse_btn.pack(side=tk.LEFT, padx=(10, 0))

        filename_frame = tk.Frame(output_frame)
        filename_frame.pack(fill=tk.X)

        tk.Label(filename_frame, text="文件名:").pack(side=tk.LEFT)
        self.filename_entry = tk.Entry(filename_frame, textvariable=self.output_filename, width=40)
        self.filename_entry.pack(side=tk.LEFT, padx=(10, 10))

        self.ext_label = tk.Label(filename_frame, text=".xlsx")
        self.ext_label.pack(side=tk.LEFT)

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

        self.status_var = tk.StringVar(value="就绪")
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

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.on_type_changed()

    def _on_closing(self):
        """窗口关闭时保存设置。"""
        self._save_settings()
        self.root.quit()

    def on_type_changed(self):
        """转换类型改变时的处理。"""
        if self.conversion_type.get() == "ini_to_excel":
            self.ext_label.config(text=".xlsx")
            self.input_btn.config(text="选择文件夹")
        else:
            self.ext_label.config(text=".ini")
            self.input_btn.config(text="选择文件夹")

    def select_input_folder(self):
        """选择输入文件夹。"""
        initial_dir = resolve_config_path(self.input_path.get()) if self.input_path.get() else str(BASE_DIR)
        folder = filedialog.askdirectory(title="选择输入文件夹", initialdir=initial_dir)
        if folder:
            folder = check_and_add_table_folder(folder)
            self.input_path.set(normalize_relative_path(folder))

    def select_input_file(self):
        """选择输入文件。"""
        initial_dir = resolve_config_path(self.input_path.get()) if self.input_path.get() else str(BASE_DIR)
        if self.conversion_type.get() == "ini_to_excel":
            filetypes = [("INI 文件", "*.ini"), ("所有文件", "*.*")]
            title = "选择 INI 文件"
        else:
            filetypes = [("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
            title = "选择 Excel 文件"

        file = filedialog.askopenfilename(title=title, filetypes=filetypes, initialdir=initial_dir)
        if file:
            self.input_path.set(normalize_relative_path(file))

    def select_output_folder(self):
        """选择输出文件夹。"""
        initial_dir = resolve_config_path(self.output_path.get()) if self.output_path.get() else str(BASE_DIR)
        folder = filedialog.askdirectory(title="选择输出文件夹", initialdir=initial_dir)
        if folder:
            self.output_path.set(normalize_relative_path(folder))

    def do_convert(self):
        """执行转换。"""
        input_rel = self.input_path.get().strip()
        output_rel = self.output_path.get().strip()
        filename = self.output_filename.get().strip()

        if not input_rel:
            messagebox.showwarning("警告", "请选择输入文件/文件夹")
            return
        if not output_rel:
            messagebox.showwarning("警告", "请选择输出文件夹")
            return
        if not filename:
            messagebox.showwarning("警告", "请输入输出文件名")
            return

        input_path = resolve_config_path(input_rel)
        output_path = resolve_config_path(output_rel)

        if not os.path.exists(input_path):
            messagebox.showerror("错误", f"输入路径不存在：{input_path}")
            return

        os.makedirs(output_path, exist_ok=True)

        ext = '.xlsx' if self.conversion_type.get() == "ini_to_excel" else '.ini'
        output_file = os.path.join(output_path, filename + ext)
        output_file = get_unique_filename(output_file)

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
    """显示主窗口。"""
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


def main():
    """主函数。"""
    config = load_config()
    input_path = config.get('user_settings', {}).get('input_path', './rundata/input')
    out_path = config.get('user_settings', {}).get('output_path', './rundata/output')

    print(f"输入路径：{input_path}")
    print(f"输出路径：{out_path}")
    show_welcome_window()


if __name__ == "__main__":
    main()
