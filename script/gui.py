#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
INI 与 Excel 文件互转工具 WebView 界面
使用 pywebview 加载 webui/index.html，后续可平滑升级到 Vue。
"""

import os
import sys
import webbrowser
from pathlib import Path
from typing import Any, Dict

try:
    import yaml
except ImportError as exc:
    raise RuntimeError("缺少 PyYAML 依赖，请先执行: pip install pyyaml") from exc

try:
    import webview
except ImportError as exc:
    raise RuntimeError("缺少 pywebview 依赖，请先执行: pip install pywebview") from exc

from ini_to_excel import ini_to_excel, get_unique_filename
from excel_to_ini import excel_to_ini


if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).resolve().parent
    RESOURCE_DIR = Path(getattr(sys, '_MEIPASS', APP_DIR)).resolve()
else:
    APP_DIR = Path(__file__).resolve().parent.parent
    RESOURCE_DIR = APP_DIR

BASE_DIR = APP_DIR
CONFIG_PATH = BASE_DIR / 'config' / 'setting.yaml'
WEBUI_INDEX = RESOURCE_DIR / 'webui' / 'index.html'
HELP_PAGE = RESOURCE_DIR / 'webui' / 'help.html'
GITHUB_URL = 'https://github.com/SoulQAQ/war3we-ini-excel-converter'
W3X2LNI_DOWNLOAD_URL = 'https://github.com/sumneko/w3x2lni'


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
        'w3x2lni_path': '',
    },
    'ui_tips': [
        '建议优先使用已拆解完成的 table 目录做通用规则验证。',
        '若要直接选择 .w3x 地图，请先在设置中配置 w3x2lni 路径。',
        '输出文件名无需手动输入扩展名，程序会自动补全。',
    ],
}


window = None


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
        return {
            'ini_names': dict(DEFAULT_CONFIG['ini_names']),
            'user_settings': dict(DEFAULT_CONFIG['user_settings']),
            'ui_tips': list(DEFAULT_CONFIG['ui_tips']),
        }

    with open(CONFIG_PATH, 'r', encoding='utf-8') as file:
        data = yaml.safe_load(file) or {}

    config = {
        'ini_names': dict(DEFAULT_CONFIG['ini_names']),
        'user_settings': dict(DEFAULT_CONFIG['user_settings']),
        'ui_tips': list(DEFAULT_CONFIG['ui_tips']),
    }
    config['ini_names'].update(data.get('ini_names', {}) or {})
    config['user_settings'].update(data.get('user_settings', {}) or {})

    ui_tips = data.get('ui_tips')
    if isinstance(ui_tips, list) and ui_tips:
        config['ui_tips'] = [str(item) for item in ui_tips if str(item).strip()]

    return config


def save_config(config):
    """保存 YAML 配置文件。"""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_PATH, 'w', encoding='utf-8') as file:
        yaml.safe_dump(config, file, allow_unicode=True, sort_keys=False)


def load_ini_names(config):
    """加载 INI 文件名称映射。"""
    ini_names = {}
    for key, value in (config.get('ini_names') or {}).items():
        ini_names[str(key).lower()] = value
    return ini_names


def check_and_add_table_folder(folder_path):
    """检查文件夹内是否有 table 和 w3x2lni 文件夹，如果有则自动添加 table 层级。"""
    if not folder_path or not os.path.isdir(folder_path):
        return folder_path

    table_path = os.path.join(folder_path, 'table')
    w3x2lni_path = os.path.join(folder_path, 'w3x2lni')

    if os.path.isdir(table_path) and os.path.isdir(w3x2lni_path):
        return table_path

    return folder_path


def find_w2l_path_from_w3x2lni(selected_path: str):
    """根据用户选择的 w3x2lni.exe 推导并验证 w2l.exe 路径。"""
    if not selected_path:
        return None

    selected = Path(selected_path)
    if selected.name.lower() != 'w3x2lni.exe':
        return None

    w2l_path = selected.with_name('w2l.exe')
    if not w2l_path.exists():
        return None

    return str(w2l_path.resolve())


class ConverterApi:
    """暴露给 WebView 前端的桥接接口。"""

    def __init__(self):
        self.config = load_config()
        self.ini_names = load_ini_names(self.config)

    def _refresh_config(self):
        self.config = load_config()
        self.ini_names = load_ini_names(self.config)

    def _save_user_settings(self, input_path: str, output_path: str, output_filename: str, conversion_type: str):
        user_settings = dict(self.config.get('user_settings') or {})
        user_settings.update({
            'input_path': input_path,
            'output_path': output_path,
            'output_filename': output_filename.strip() or 'output',
            'conversion_type': conversion_type,
        })
        self.config['user_settings'] = user_settings
        save_config(self.config)

    def get_initial_state(self, payload: Dict[str, Any] | None = None):
        """返回初始界面状态。"""
        _ = payload
        self._refresh_config()
        user_settings = self.config.get('user_settings', {})
        w3x2lni_path = user_settings.get('w3x2lni_path', '')
        return {
            'input_path': user_settings.get('input_path', ''),
            'output_path': user_settings.get('output_path', ''),
            'output_filename': user_settings.get('output_filename', 'output'),
            'conversion_type': user_settings.get('conversion_type', 'ini_to_excel'),
            'w3x2lni_path': w3x2lni_path,
            'has_w3x2lni': bool(w3x2lni_path),
            'ui_tips': self.config.get('ui_tips', []),
            'help_url': HELP_PAGE.as_uri() if HELP_PAGE.exists() else GITHUB_URL,
            'github_url': GITHUB_URL,
        }

    def pick_input_folder(self, payload: Dict[str, Any] | None = None):
        """选择输入文件夹。"""
        _ = payload
        initial_dir = resolve_config_path(self.config.get('user_settings', {}).get('input_path', '')) or str(BASE_DIR)
        result = window.create_file_dialog(webview.FOLDER_DIALOG, directory=initial_dir)
        if result:
            selected = check_and_add_table_folder(result[0])
            return {'path': normalize_relative_path(selected)}
        return {'path': None}

    def pick_input_file(self, payload: Dict[str, Any] | None = None):
        """选择输入文件。"""
        payload = payload or {}
        conversion_type = payload.get('conversion_type', 'ini_to_excel')
        initial_dir = resolve_config_path(self.config.get('user_settings', {}).get('input_path', '')) or str(BASE_DIR)

        if conversion_type == 'ini_to_excel':
            user_settings = self.config.get('user_settings', {})
            if not user_settings.get('w3x2lni_path'):
                return {
                    'path': None,
                    'success': False,
                    'message': '请先在设置中配置 w3x2lni 路径，之后才能直接选择地图文件。',
                }
            file_types = ('Warcraft III 地图 (*.w3x)', 'All files (*.*)')
        else:
            file_types = ('Excel 文件 (*.xlsx;*.xls)', 'All files (*.*)')

        result = window.create_file_dialog(
            webview.OPEN_DIALOG,
            directory=initial_dir,
            allow_multiple=False,
            file_types=[file_types],
        )
        if result:
            return {'path': normalize_relative_path(result[0]), 'success': True}
        return {'path': None, 'success': True}

    def pick_output_folder(self, payload: Dict[str, Any] | None = None):
        """选择输出文件夹。"""
        _ = payload
        initial_dir = resolve_config_path(self.config.get('user_settings', {}).get('output_path', '')) or str(BASE_DIR)
        result = window.create_file_dialog(webview.FOLDER_DIALOG, directory=initial_dir)
        if result:
            return {'path': normalize_relative_path(result[0])}
        return {'path': None}

    def get_settings(self, payload: Dict[str, Any] | None = None):
        """返回设置面板所需配置。"""
        _ = payload
        self._refresh_config()
        user_settings = self.config.get('user_settings', {})
        return {
            'w3x2lni_path': user_settings.get('w3x2lni_path', ''),
        }

    def pick_w3x2lni_path(self, payload: Dict[str, Any] | None = None):
        """让用户选择 w3x2lni.exe，并返回对应的 w2l.exe 路径。"""
        _ = payload
        configured = self.config.get('user_settings', {}).get('w3x2lni_path', '')
        initial_dir = str(Path(configured).resolve().parent) if configured else str(BASE_DIR)

        result = window.create_file_dialog(
            webview.OPEN_DIALOG,
            directory=initial_dir,
            allow_multiple=False,
            file_types=[('w3x2lni 程序 (w3x2lni.exe)', 'w3x2lni.exe')],
        )
        if not result:
            return {'success': False, 'cancelled': True}

        selected_path = result[0]
        w2l_path = find_w2l_path_from_w3x2lni(selected_path)
        if not w2l_path:
            return {
                'success': False,
                'cancelled': False,
                'message': '未在同目录找到 w2l.exe，当前 w3x2lni 可能已损坏。',
                'download_url': W3X2LNI_DOWNLOAD_URL,
            }

        return {
            'success': True,
            'cancelled': False,
            'selected_path': normalize_relative_path(selected_path),
            'w3x2lni_path': normalize_relative_path(w2l_path),
        }

    def save_settings(self, payload: Dict[str, Any] | None = None):
        """保存设置。"""
        payload = payload or {}
        self._refresh_config()

        raw_path = (payload.get('w3x2lni_path') or '').strip()
        user_settings = dict(self.config.get('user_settings') or {})
        user_settings['w3x2lni_path'] = raw_path
        self.config['user_settings'] = user_settings
        save_config(self.config)

        return {
            'success': True,
            'w3x2lni_path': raw_path,
            'has_w3x2lni': bool(raw_path),
            'message': '设置已保存。',
        }

    def open_external_link(self, payload: Dict[str, Any] | None = None):
        """打开外部链接。"""
        payload = payload or {}
        url = (payload.get('url') or '').strip()
        if not url:
            return {'success': False, 'message': '缺少要打开的链接。'}

        webbrowser.open(url)
        return {'success': True}

    def run_conversion(self, payload: Dict[str, Any] | None = None):
        """执行转换。"""
        payload = payload or {}
        input_rel = (payload.get('input_path') or '').strip()
        output_rel = (payload.get('output_path') or '').strip()
        output_filename = (payload.get('output_filename') or '').strip()
        conversion_type = (payload.get('conversion_type') or 'ini_to_excel').strip()

        if not input_rel:
            return {'success': False, 'message': '请选择输入文件/文件夹'}
        if not output_rel:
            return {'success': False, 'message': '请选择输出文件夹'}
        if not output_filename:
            return {'success': False, 'message': '请输入输出文件名'}

        input_path = resolve_config_path(input_rel)
        output_path = resolve_config_path(output_rel)

        if not os.path.exists(input_path):
            return {'success': False, 'message': f'输入路径不存在：{input_path}'}

        os.makedirs(output_path, exist_ok=True)

        ext = '.xlsx' if conversion_type == 'ini_to_excel' else '.ini'
        output_file = os.path.join(output_path, output_filename + ext)
        output_file = get_unique_filename(output_file)

        try:
            self._refresh_config()
            if conversion_type == 'ini_to_excel':
                ini_to_excel(input_path, output_file, self.ini_names)
                result_message = f'Excel 文件已创建：{normalize_relative_path(output_file)}'
            else:
                excel_to_ini(input_path, output_file)
                result_message = f'INI 文件已创建：{normalize_relative_path(output_file)}'

            self._save_user_settings(input_rel, output_rel, output_filename, conversion_type)
            return {'success': True, 'message': result_message, 'output_file': normalize_relative_path(output_file)}
        except Exception as exc:
            return {'success': False, 'message': f'转换失败：{str(exc)}'}

    def close_window(self, payload: Dict[str, Any] | None = None):
        """关闭窗口。"""
        _ = payload
        if window is not None:
            window.destroy()
        return {'success': True}


def ensure_webui_exists():
    """确保 Web UI 入口文件存在。"""
    if not WEBUI_INDEX.exists():
        raise FileNotFoundError(f'未找到 Web UI 文件：{WEBUI_INDEX}')


def main():
    """主函数。"""
    ensure_webui_exists()

    api = ConverterApi()
    user_settings = api.get_initial_state()

    print(f"运行根目录：{BASE_DIR}")
    print(f"资源目录：{RESOURCE_DIR}")
    print(f"输入路径：{user_settings['input_path']}")
    print(f"输出路径：{user_settings['output_path']}")
    print('启动 WebView 界面...')

    global window
    window = webview.create_window(
        'INI-Excel 转换工具',
        url=WEBUI_INDEX.as_uri(),
        js_api=api,
        width=1120,
        height=820,
        min_size=(980, 700),
        text_select=True,
    )
    webview.start()


if __name__ == '__main__':
    main()
