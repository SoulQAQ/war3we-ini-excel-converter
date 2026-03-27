#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
INI 与 Excel 文件互转工具 WebView 界面
使用 pywebview 加载 webui/index.html，后续可平滑升级为 Vue。
"""

import os
import sys
import subprocess
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

from ini_to_excel import ini_to_excel
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
FAVICON_PATH = RESOURCE_DIR / 'favicon.ico'
LOGO_PATH = RESOURCE_DIR / 'logo.jpg'
HELP_URL = 'http://soul2.cn/read/doc/war3we-ini-excel-converter/help.html'
GITHUB_URL = 'https://github.com/SoulQAQ/war3we-ini-excel-converter'
W3X2LNI_DOWNLOAD_URL = 'https://github.com/sumneko/w3x2lni'


DEFAULT_CONFIG = {
    'ini_names': {
        'ability.ini': '技能',
        'buff.ini': '魔法效果',
        'item.ini': '物品',
        'unit.ini': '单位',
        'upgrade.ini': '科技',
    },
    'user_settings': {
        'input_path': './rundata/input',
        'conversion_type': 'ini_to_excel',
        'w3x2lni_path': '',
        'enable_calc_formula_detection': True,
    },
    'ui_tips': [
        '选择地图文件夹后，Excel 将自动输出到同目录下。',
        '若要直接选择 .w3x 地图，请先在设置中配置 w3x2lni 路径。',
        'xlsx 文件已存在时会询问覆盖或重命名旧文件。',
    ],
}


window = None


def normalize_relative_path(path_value):
    """将任意路径规范为相对于项目根目录的路径。"""
    if not path_value:
        return ''

    path_obj = Path(path_value)
    if not path_obj.is_absolute():
        path_obj = (BASE_DIR / path_value).resolve()
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


def validate_w3x2lni_path(w2l_path: str) -> bool:
    """验证 w3x2lni 路径是否有效（检查 w3x2lni.exe 和 w2l.exe 是否都存在）。"""
    if not w2l_path:
        return False

    w2l = Path(w2l_path)
    if not w2l.exists():
        return False

    w3x2lni_exe = w2l.parent / 'w3x2lni.exe'
    if not w3x2lni_exe.exists():
        return False

    return True


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


def run_w3x2lni_convert(w2l_path: str, w3x_path: str, output_dir: str) -> Dict[str, Any]:
    """调用 w2l.exe 将 w3x 转换为地图文件夹。

    用法: w2l.exe lni <目标绝对路径> <输出位置绝对路径>
    输出位置直接指定为地图名子文件夹路径
    """
    try:
        w3x_abs = str(Path(w3x_path).resolve())
        w3x_name = Path(w3x_path).stem

        # 输出目录直接指定为地图名子文件夹
        output_parent = str(Path(output_dir).resolve())
        # 最终输出路径 - w2l.exe 会在此路径下创建地图文件夹
        output_abs = os.path.join(output_parent, w3x_name)

        # 直接指定输出路径，让 w2l.exe 在该路径下创建文件夹
        cmd = [w2l_path, 'lni', w3x_abs, output_abs]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
            cwd=str(Path(w2l_path).parent)
        )

        if result.returncode != 0:
            return {
                'success': False,
                'message': f'w3x2lni 转换失败：{result.stderr or "未知错误"}',
            }

        # 检查输出文件夹是否存在
        if os.path.isdir(output_abs):
            return {
                'success': True,
                'converted_folder': output_abs,
            }

        # 如果没有创建子文件夹，检查是否直接输出到了 output_parent
        if os.path.isdir(output_parent):
            return {
                'success': True,
                'converted_folder': output_parent,
            }

        return {
            'success': False,
            'message': 'w3x2lni 转换完成但未找到输出文件夹。',
        }

    except subprocess.TimeoutExpired:
        return {
            'success': False,
            'message': 'w3x2lni 转换超时，请检查地图文件是否损坏。',
        }
    except Exception as exc:
        return {
            'success': False,
            'message': f'调用 w3x2lni 失败：{str(exc)}',
        }


def get_output_xlsx_path(input_path: str) -> str:
    """根据输入路径自动确定输出 xlsx 文件路径。

    规则：
    - 如果输入是 table 目录，输出到其父目录下的 <地图名>.xlsx
    - 如果输入是地图文件夹，输出到同目录下的 <文件夹名>.xlsx
    - 如果输入是单个 ini 文件，输出到同目录下的 <文件名>.xlsx
    """
    input_abs = resolve_config_path(input_path)

    if not input_abs or not os.path.exists(input_abs):
        return ''

    input_p = Path(input_abs)

    if input_p.is_file():
        parent_dir = input_p.parent
        base_name = input_p.stem
    else:
        if input_p.name.lower() == 'table':
            parent_dir = input_p.parent
            base_name = input_p.parent.name
        else:
            parent_dir = input_p.parent
            base_name = input_p.name

    return str(parent_dir / f'{base_name}.xlsx')


def rename_existing_file(file_path: str) -> str:
    """重命名已存在的文件，添加 _1, _2, _3 等后缀。"""
    if not os.path.exists(file_path):
        return file_path

    p = Path(file_path)
    parent = p.parent
    stem = p.stem
    suffix = p.suffix

    counter = 1
    while True:
        new_name = f'{stem}_{counter}{suffix}'
        new_path = parent / new_name
        if not new_path.exists():
            return str(new_path)
        counter += 1


class ConverterApi:
    """暴露给 WebView 前端的桥接接口。"""

    def __init__(self):
        self.config = load_config()
        self.ini_names = load_ini_names(self.config)
        self._pending_output_path = None

        # 启动时校验 w3x2lni 路径
        self._validate_and_clean_w3x2lni_path()

    def _validate_and_clean_w3x2lni_path(self):
        """校验 w3x2lni 路径，如果无效则清空配置。"""
        w2l_path = self.config.get('user_settings', {}).get('w3x2lni_path', '')
        if w2l_path and not validate_w3x2lni_path(w2l_path):
            self.config['user_settings']['w3x2lni_path'] = ''
            save_config(self.config)

    def _refresh_config(self):
        self.config = load_config()
        self.ini_names = load_ini_names(self.config)
        self._validate_and_clean_w3x2lni_path()

    def _save_user_settings(self, input_path: str, conversion_type: str):
        user_settings = dict(self.config.get('user_settings') or {})
        user_settings.update({
            'input_path': input_path,
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
        input_rel = user_settings.get('input_path', '')

        # 计算输入路径的绝对路径用于显示
        input_abs = resolve_config_path(input_rel) if input_rel else ''

        # 检查是否存在对应的 xlsx 文件
        xlsx_exists = False
        xlsx_path = ''
        if input_abs and os.path.exists(input_abs):
            xlsx_path = get_output_xlsx_path(input_abs)
            xlsx_exists = os.path.exists(xlsx_path)

        return {
            'input_path': input_rel,
            'input_path_abs': input_abs,
            'conversion_type': user_settings.get('conversion_type', 'ini_to_excel'),
            'w3x2lni_path': w3x2lni_path,
            'has_w3x2lni': bool(w3x2lni_path),
            'enable_calc_formula_detection': bool(user_settings.get('enable_calc_formula_detection', True)),
            'ui_tips': self.config.get('ui_tips', []),
            'help_url': HELP_URL,
            'github_url': GITHUB_URL,
            'xlsx_exists': xlsx_exists,
            'xlsx_path': xlsx_path,
        }

    def pick_input_folder(self, payload: Dict[str, Any] | None = None):
        """选择输入文件夹。"""
        _ = payload
        initial_dir = resolve_config_path(self.config.get('user_settings', {}).get('input_path', '')) or str(BASE_DIR)
        result = window.create_file_dialog(webview.FOLDER_DIALOG, directory=initial_dir)
        if result:
            selected = check_and_add_table_folder(result[0])
            selected_abs = str(Path(selected).resolve())

            # 检查是否存在对应的 xlsx 文件
            xlsx_path = get_output_xlsx_path(selected_abs)
            xlsx_exists = os.path.exists(xlsx_path)

            return {
                'path': normalize_relative_path(selected),
                'path_abs': selected_abs,
                'xlsx_exists': xlsx_exists,
                'xlsx_path': xlsx_path,
            }
        return {'path': None}

    def pick_input_file(self, payload: Dict[str, Any] | None = None):
        """选择输入文件。"""
        payload = payload or {}
        conversion_type = payload.get('conversion_type', 'ini_to_excel')
        initial_dir = resolve_config_path(self.config.get('user_settings', {}).get('input_path', '')) or str(BASE_DIR)

        if conversion_type == 'ini_to_excel':
            user_settings = self.config.get('user_settings', {})
            w2l_path = user_settings.get('w3x2lni_path', '')
            if not w2l_path:
                return {
                    'path': None,
                    'success': False,
                    'message': '请先在设置中配置 w3x2lni 路径，之后才能直接选择地图文件。',
                }
            file_types = 'Warcraft III 地图 (*.w3x)'
        else:
            file_types = 'Excel 文件 (*.xlsx;*.xls)'

        result = window.create_file_dialog(
            webview.OPEN_DIALOG,
            directory=initial_dir,
            allow_multiple=False,
            file_types=[file_types],
        )
        if not result:
            return {'path': None, 'success': True}

        selected_path = result[0]

        if conversion_type == 'ini_to_excel' and selected_path.lower().endswith('.w3x'):
            w2l_path = self.config.get('user_settings', {}).get('w3x2lni_path', '')
            if not w2l_path:
                return {
                    'path': None,
                    'success': False,
                    'message': '请先在设置中配置 w3x2lni 路径。',
                }

            w3x_dir = str(Path(selected_path).parent)
            convert_result = run_w3x2lni_convert(w2l_path, selected_path, w3x_dir)

            if not convert_result.get('success'):
                return {
                    'path': None,
                    'success': False,
                    'message': convert_result.get('message', 'w3x 转换失败。'),
                }

            converted_folder = convert_result.get('converted_folder', '')
            if converted_folder:
                table_path = check_and_add_table_folder(converted_folder)

                # 检查是否存在对应的 xlsx 文件
                xlsx_path = get_output_xlsx_path(table_path)
                xlsx_exists = os.path.exists(xlsx_path)

                return {
                    'path': normalize_relative_path(table_path),
                    'path_abs': str(Path(table_path).resolve()),
                    'success': True,
                    'converted_folder': normalize_relative_path(table_path),
                    'xlsx_exists': xlsx_exists,
                    'xlsx_path': xlsx_path,
                }

        selected_abs = str(Path(selected_path).resolve())

        # 检查是否存在对应的 xlsx 文件
        xlsx_path = get_output_xlsx_path(selected_abs)
        xlsx_exists = os.path.exists(xlsx_path)

        return {
            'path': normalize_relative_path(selected_path),
            'path_abs': selected_abs,
            'success': True,
            'xlsx_exists': xlsx_exists,
            'xlsx_path': xlsx_path,
        }

    def get_settings(self, payload: Dict[str, Any] | None = None):
        """返回设置面板所需配置。"""
        _ = payload
        self._refresh_config()
        user_settings = self.config.get('user_settings', {})
        w3x2lni_path_rel = user_settings.get('w3x2lni_path', '')
        # 返回绝对路径用于显示
        w3x2lni_path_abs = resolve_config_path(w3x2lni_path_rel) if w3x2lni_path_rel else ''
        return {
            'w3x2lni_path': w3x2lni_path_rel,
            'w3x2lni_path_abs': w3x2lni_path_abs,
            'enable_calc_formula_detection': bool(user_settings.get('enable_calc_formula_detection', True)),
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
            file_types=['Executable Files (*.exe)'],
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
            'w3x2lni_path_abs': str(Path(w2l_path).resolve()),
        }

    def save_settings(self, payload: Dict[str, Any] | None = None):
        """保存设置。"""
        payload = payload or {}
        self._refresh_config()

        raw_path = (payload.get('w3x2lni_path') or '').strip()
        enable_calc_formula_detection = bool(payload.get('enable_calc_formula_detection', True))
        user_settings = dict(self.config.get('user_settings') or {})
        user_settings['w3x2lni_path'] = raw_path
        user_settings['enable_calc_formula_detection'] = enable_calc_formula_detection
        self.config['user_settings'] = user_settings
        save_config(self.config)

        return {
            'success': True,
            'w3x2lni_path': raw_path,
            'has_w3x2lni': bool(raw_path),
            'enable_calc_formula_detection': enable_calc_formula_detection,
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

    def open_in_explorer(self, payload: Dict[str, Any] | None = None):
        """在资源管理器中打开文件并选中。"""
        payload = payload or {}
        file_path = (payload.get('path') or '').strip()
        if not file_path:
            return {'success': False, 'message': '缺少文件路径。'}

        file_abs = resolve_config_path(file_path) if not Path(file_path).is_absolute() else file_path

        if not os.path.exists(file_abs):
            return {'success': False, 'message': f'文件不存在：{file_abs}'}

        try:
            subprocess.run(['explorer', '/select,', file_abs], check=False)
            return {'success': True}
        except Exception as exc:
            return {'success': False, 'message': f'打开资源管理器失败：{str(exc)}'}

    def run_conversion(self, payload: Dict[str, Any] | None = None):
        """执行转换。"""
        payload = payload or {}
        input_rel = (payload.get('input_path') or '').strip()
        conversion_type = (payload.get('conversion_type') or 'ini_to_excel').strip()

        if not input_rel:
            return {'success': False, 'message': '请选择输入路径'}

        input_path = resolve_config_path(input_rel)

        if not os.path.exists(input_path):
            return {'success': False, 'message': f'输入路径不存在：{input_path}'}

        output_file = get_output_xlsx_path(input_path)

        if not output_file:
            return {'success': False, 'message': '无法确定输出文件路径'}

        if os.path.exists(output_file):
            self._pending_output_path = output_file
            return {
                'success': True,
                'need_confirm': True,
                'message': f'文件 {os.path.basename(output_file)} 已存在。',
            }

        return self._execute_conversion(input_path, output_file, conversion_type, input_rel)

    def confirm_overwrite(self, payload: Dict[str, Any] | None = None):
        """确认覆盖或重命名旧文件后执行转换。"""
        payload = payload or {}
        input_rel = (payload.get('input_path') or '').strip()
        conversion_type = (payload.get('conversion_type') or 'ini_to_excel').strip()
        overwrite = payload.get('overwrite', False)

        input_path = resolve_config_path(input_rel)
        output_file = get_output_xlsx_path(input_path)

        if not output_file:
            return {'success': False, 'message': '无法确定输出文件路径'}

        if not overwrite and os.path.exists(output_file):
            output_file = rename_existing_file(output_file)

        return self._execute_conversion(input_path, output_file, conversion_type, input_rel)

    def _execute_conversion(self, input_path: str, output_file: str, conversion_type: str, input_rel: str):
        """实际执行转换操作。"""
        try:
            self._refresh_config()
            export_options = {
                'enable_calc_formula_detection': bool(self.config.get('user_settings', {}).get('enable_calc_formula_detection', True))
            }
            if conversion_type == 'ini_to_excel':
                ini_to_excel(input_path, output_file, self.ini_names, export_options)
                result_message = f'Excel 文件已创建：{output_file}'
            else:
                excel_to_ini(input_path, output_file)
                result_message = f'INI 文件已创建：{output_file}'

            self._save_user_settings(input_rel, conversion_type)

            # 检查 xlsx 文件是否存在
            xlsx_exists = os.path.exists(output_file)

            return {
                'success': True,
                'message': result_message,
                'output_file': output_file,
                'xlsx_exists': xlsx_exists,
                'xlsx_path': output_file,
            }
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
    print('启动 WebView 界面...')

    global window
    window = webview.create_window(
        'INI-Excel 转换工具',
        url=WEBUI_INDEX.as_uri(),
        js_api=api,
        width=1120,
        height=660,
        min_size=(980, 650),
        text_select=True,
    )
    if FAVICON_PATH.exists():
        try:
            window.icon = str(FAVICON_PATH)
        except Exception:
            pass
    webview.start(debug=False)


if __name__ == '__main__':
    main()
