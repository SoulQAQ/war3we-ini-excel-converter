#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
INI 与 Excel 文件互转工具 WebView 界面
使用 pywebview 加载 webui/index.html，后续可平滑升级为 Vue。
"""

import os
import re
import subprocess
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
<<<<<<< HEAD
        'conversion_type': 'ini_to_excel',
        'w3x2lni_path': '',
        'enable_calc_formula_detection': True,
        'ini_to_excel': {
            'input_path': './rundata/input',
            'output_path': './rundata/output',
            'output_filename': 'output',
        },
        'excel_to_ini': {
            'input_path': './rundata/output/output.xlsx',
            'output_path': '',
            'output_filename': '',
            'use_original_ini_location': True,
        },
        'w3x': {
            'map_path': '',
            'workspace_path': './rundata/map',
            'packed_map_path': './rundata/output/output.w3x',
        },
    },
    'ui_tips': [
        '建议优先使用已拆解完成的 table 目录做通用规则验证。',
        'Excel 回写会优先写回导出时记录的原始 ini 文件位置。',
        '多等级字段可使用 ---- 分隔，也可以用 calc@10+5 快速展开。',
        '若要直接选择 .w3x 地图，请先在设置中配置 w2l.exe 路径。',
        '输出文件名无需手动输入扩展名，程序会自动补全。',
=======
        'input_path': './rundata/input',
        'conversion_type': 'ini_to_excel',
        'w3x2lni_path': '',
        'enable_calc_formula_detection': True,
        # Excel 转 INI 独立配置
        'excel_to_ini_map_path': '',
        'excel_file_path': '',
    },
    'ui_tips': [
        '选择地图文件夹后，Excel 将自动输出到同目录下。',
        '若要直接选择 .w3x 地图，请先在设置中配置 w3x2lni 路径。',
        'xlsx 文件已存在时会询问覆盖或重命名旧文件。',
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74
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
    config['user_settings']['ini_to_excel'] = dict(DEFAULT_CONFIG['user_settings']['ini_to_excel'])
    config['user_settings']['excel_to_ini'] = dict(DEFAULT_CONFIG['user_settings']['excel_to_ini'])
    config['user_settings']['w3x'] = dict(DEFAULT_CONFIG['user_settings']['w3x'])
    config['ini_names'].update(data.get('ini_names', {}) or {})
    saved_user_settings = data.get('user_settings', {}) or {}
    config['user_settings'].update(saved_user_settings)

    for key in ('ini_to_excel', 'excel_to_ini', 'w3x'):
        nested = dict(DEFAULT_CONFIG['user_settings'][key])
        saved_nested = saved_user_settings.get(key)
        if isinstance(saved_nested, dict):
            nested.update(saved_nested)
        config['user_settings'][key] = nested

    # 兼容旧版单向配置：仅在目标分组缺失时迁移一次，避免每次加载覆盖新分组配置
    has_legacy_fields = any(
        saved_user_settings.get(field) for field in ('input_path', 'output_path', 'output_filename')
    )
    if has_legacy_fields:
        legacy_type = saved_user_settings.get('conversion_type', config['user_settings'].get('conversion_type'))
        legacy_bucket = 'excel_to_ini' if legacy_type == 'excel_to_ini' else 'ini_to_excel'
        saved_bucket = saved_user_settings.get(legacy_bucket)
        if not isinstance(saved_bucket, dict) or not saved_bucket:
            bucket = dict(config['user_settings'].get(legacy_bucket) or {})
            if saved_user_settings.get('input_path'):
                bucket['input_path'] = saved_user_settings.get('input_path')
            if saved_user_settings.get('output_path'):
                bucket['output_path'] = saved_user_settings.get('output_path')
            if saved_user_settings.get('output_filename'):
                bucket['output_filename'] = saved_user_settings.get('output_filename')
            config['user_settings'][legacy_bucket] = bucket

    for legacy_field in ('input_path', 'output_path', 'output_filename'):
        config['user_settings'].pop(legacy_field, None)

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


def make_unique_directory_path(folder_path: str):
    """如果目录已存在，在目录名后追加 _1、_2 等后缀。"""
    if not os.path.exists(folder_path):
        return folder_path

    parent = os.path.dirname(folder_path)
    name = os.path.basename(folder_path)
    match = re.match(r'^(.+?)_(\d+)$', name)
    if match:
        name = match.group(1)

    counter = 1
    while True:
        candidate = os.path.join(parent, f'{name}_{counter}')
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def normalize_w2l_path(selected_path: str):
    """根据用户选择的 exe 推导并验证 w2l.exe 路径。"""
    if not selected_path:
        return None

    selected = Path(selected_path)
    if selected.name.lower() == 'w2l.exe' and selected.exists():
        return str(selected.resolve())

    if selected.name.lower() == 'w3x2lni.exe':
        w2l_path = selected.with_name('w2l.exe')
        if w2l_path.exists():
            return str(w2l_path.resolve())

    return None


def ensure_suffix(path_value: str, suffix: str):
    path = Path(path_value)
    if path.suffix.lower() == suffix.lower():
        return str(path)
    return str(path.with_suffix(suffix))


def is_empty_directory(folder_path: Path):
    return folder_path.is_dir() and not any(folder_path.iterdir())


def get_safe_unpack_directory(folder_path: Path):
    resolved = folder_path.resolve()
    base = BASE_DIR.resolve()
    if resolved == base or resolved == base.parent:
        raise ValueError(f'拆包目录不能是项目根目录或其父目录：{resolved}')
    if resolved.exists() and not is_empty_directory(resolved):
        return Path(make_unique_directory_path(str(resolved))).resolve()
    return resolved


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

<<<<<<< HEAD
    def _save_conversion_settings(
        self,
        input_path: str,
        output_path: str,
        output_filename: str,
        conversion_type: str,
        use_original_ini_location: bool | None = None,
    ):
=======
    def _save_user_settings(self, input_path: str, conversion_type: str):
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74
        user_settings = dict(self.config.get('user_settings') or {})
        for legacy_field in ('input_path', 'output_path', 'output_filename'):
            user_settings.pop(legacy_field, None)
        user_settings['conversion_type'] = conversion_type
        bucket = dict(user_settings.get(conversion_type) or {})
        bucket.update({
            'input_path': input_path,
<<<<<<< HEAD
            'output_path': output_path,
            'output_filename': output_filename.strip(),
=======
            'conversion_type': conversion_type,
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74
        })
        if conversion_type == 'ini_to_excel':
            bucket['output_filename'] = bucket['output_filename'] or 'output'
        elif use_original_ini_location is not None:
            bucket['use_original_ini_location'] = bool(use_original_ini_location)

        user_settings[conversion_type] = bucket
        self.config['user_settings'] = user_settings
        save_config(self.config)

    def _save_w3x_settings(self, payload: Dict[str, Any]):
        user_settings = dict(self.config.get('user_settings') or {})
        w3x_settings = dict(user_settings.get('w3x') or {})
        for key in ('map_path', 'workspace_path', 'packed_map_path'):
            if key in payload:
                w3x_settings[key] = (payload.get(key) or '').strip()
        user_settings['w3x'] = w3x_settings
        self.config['user_settings'] = user_settings
        save_config(self.config)

    def get_initial_state(self, payload: Dict[str, Any] | None = None):
        """返回初始界面状态。"""
        _ = payload
        self._refresh_config()
        user_settings = self.config.get('user_settings', {})
        w3x2lni_path = user_settings.get('w3x2lni_path', '')
<<<<<<< HEAD
        ini_to_excel_settings = user_settings.get('ini_to_excel', {})
        excel_to_ini_settings = user_settings.get('excel_to_ini', {})
        w3x_settings = user_settings.get('w3x', {})
        return {
=======
        input_rel = user_settings.get('input_path', '')

        # 计算输入路径的绝对路径用于显示
        input_abs = resolve_config_path(input_rel) if input_rel else ''

        # 检查是否存在对应的 xlsx 文件
        xlsx_exists = False
        xlsx_path = ''
        if input_abs and os.path.exists(input_abs):
            xlsx_path = get_output_xlsx_path(input_abs)
            xlsx_exists = os.path.exists(xlsx_path)

        # Excel 转 INI 配置
        excel_to_ini_map_rel = user_settings.get('excel_to_ini_map_path', '')
        excel_to_ini_map_abs = resolve_config_path(excel_to_ini_map_rel) if excel_to_ini_map_rel else ''
        excel_file_rel = user_settings.get('excel_file_path', '')
        excel_file_abs = resolve_config_path(excel_file_rel) if excel_file_rel else ''

        return {
            'input_path': input_rel,
            'input_path_abs': input_abs,
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74
            'conversion_type': user_settings.get('conversion_type', 'ini_to_excel'),
            'ini_to_excel': {
                'input_path': ini_to_excel_settings.get('input_path', ''),
                'output_path': ini_to_excel_settings.get('output_path', ''),
                'output_filename': ini_to_excel_settings.get('output_filename', 'output'),
            },
            'excel_to_ini': {
                'input_path': excel_to_ini_settings.get('input_path', ''),
                'output_path': excel_to_ini_settings.get('output_path', ''),
                'output_filename': excel_to_ini_settings.get('output_filename', ''),
                'use_original_ini_location': bool(excel_to_ini_settings.get('use_original_ini_location', True)),
            },
            'w3x': {
                'map_path': w3x_settings.get('map_path', ''),
                'workspace_path': w3x_settings.get('workspace_path', ''),
                'packed_map_path': w3x_settings.get('packed_map_path', ''),
            },
            'w3x2lni_path': w3x2lni_path,
            'w3x2lni_path_abs': resolve_config_path(w3x2lni_path) if w3x2lni_path else '',
            'has_w3x2lni': bool(w3x2lni_path),
            'has_w2l': bool(w3x2lni_path),
            'enable_calc_formula_detection': bool(user_settings.get('enable_calc_formula_detection', True)),
            'ui_tips': self.config.get('ui_tips', []),
            'help_url': HELP_URL,
            'github_url': GITHUB_URL,
            'xlsx_exists': xlsx_exists,
            'xlsx_path': xlsx_path,
            # Excel 转 INI 配置
            'excel_to_ini_map_path': excel_to_ini_map_rel,
            'excel_to_ini_map_path_abs': excel_to_ini_map_abs,
            'excel_file_path': excel_file_rel,
            'excel_file_path_abs': excel_file_abs,
        }

    def pick_input_folder(self, payload: Dict[str, Any] | None = None):
        """选择输入文件夹。"""
        payload = payload or {}
        conversion_type = payload.get('conversion_type', 'ini_to_excel')
        settings = self.config.get('user_settings', {}).get(conversion_type, {})
        initial_dir = resolve_config_path(settings.get('input_path', '')) or str(BASE_DIR)
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
        settings = self.config.get('user_settings', {}).get(conversion_type, {})
        initial_dir = resolve_config_path(settings.get('input_path', '')) or str(BASE_DIR)

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

<<<<<<< HEAD
    def pick_output_folder(self, payload: Dict[str, Any] | None = None):
        """选择输出文件夹。"""
        payload = payload or {}
        conversion_type = payload.get('conversion_type', 'ini_to_excel')
        settings = self.config.get('user_settings', {}).get(conversion_type, {})
        initial_dir = resolve_config_path(settings.get('output_path', '')) or str(BASE_DIR)
        result = window.create_file_dialog(webview.FOLDER_DIALOG, directory=initial_dir)
        if result:
            return {'path': normalize_relative_path(result[0])}
        return {'path': None}
=======
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
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74

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
        """让用户选择 w2l.exe 或 w3x2lni.exe，并返回可调用的 w2l.exe 路径。"""
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
        w2l_path = normalize_w2l_path(selected_path)
        if not w2l_path:
            return {
                'success': False,
                'cancelled': False,
                'message': '请选择 w2l.exe；如果选择 w3x2lni.exe，则同目录必须存在 w2l.exe。',
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
            'has_w2l': bool(raw_path),
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

<<<<<<< HEAD
    def reveal_output_path(self, payload: Dict[str, Any] | None = None):
        """在资源管理器中打开最近生成的文件或目录。"""
        payload = payload or {}
        raw_path = (payload.get('path') or '').strip()
        if not raw_path:
            return {'success': False, 'message': '没有可打开的输出路径。'}

        target = Path(resolve_config_path(raw_path)).resolve()
        if not target.exists():
            return {'success': False, 'message': f'输出路径不存在：{target}'}

        try:
            if sys.platform.startswith('win'):
                creationflags = subprocess.CREATE_NO_WINDOW
                if target.is_file():
                    subprocess.Popen(['explorer', f'/select,{str(target)}'], creationflags=creationflags)
                else:
                    subprocess.Popen(['explorer', str(target)], creationflags=creationflags)
            else:
                open_target = target.parent if target.is_file() else target
                subprocess.Popen(['xdg-open', str(open_target)])
            return {'success': True, 'message': '已打开输出位置。'}
        except Exception as exc:
            return {'success': False, 'message': f'打开输出位置失败：{exc}'}
=======
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
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74

    def run_conversion(self, payload: Dict[str, Any] | None = None):
        """执行转换。"""
        payload = payload or {}
        input_rel = (payload.get('input_path') or '').strip()
        conversion_type = (payload.get('conversion_type') or 'ini_to_excel').strip()
        use_original_ini_location = bool(payload.get('use_original_ini_location', conversion_type == 'excel_to_ini'))

        if not input_rel:
<<<<<<< HEAD
            return {'success': False, 'message': '请选择输入文件/文件夹'}
        if conversion_type == 'ini_to_excel' and not output_rel:
            return {'success': False, 'message': '请选择 Excel 输出文件夹'}
        if conversion_type == 'ini_to_excel' and not output_filename:
            return {'success': False, 'message': '请输入输出文件名'}

        input_path = resolve_config_path(input_rel)
        output_path = resolve_config_path(output_rel) if output_rel else ''
=======
            return {'success': False, 'message': '请选择输入路径'}

        input_path = resolve_config_path(input_rel)
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74

        if not os.path.exists(input_path):
            return {'success': False, 'message': f'输入路径不存在：{input_path}'}

<<<<<<< HEAD
        if output_path:
            os.makedirs(output_path, exist_ok=True)
=======
        output_file = get_output_xlsx_path(input_path)

        if not output_file:
            return {'success': False, 'message': '无法确定输出文件路径'}
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74

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
                output_file = os.path.join(output_path, output_filename + '.xlsx')
                output_file = get_unique_filename(output_file)
                ini_to_excel(input_path, output_file, self.ini_names, export_options)
<<<<<<< HEAD
                result_message = f'Excel 文件已创建：{normalize_relative_path(output_file)}'
                normalized_output = normalize_relative_path(output_file)
            else:
                output_file = ''
                if not use_original_ini_location:
                    if not output_rel:
                        return {'success': False, 'message': '请选择 INI 输出目录，或启用回写原始位置。'}
                    directory_name = output_filename.strip()
                    output_file = os.path.join(output_path, directory_name) if directory_name else output_path
                    output_file = make_unique_directory_path(output_file)

                written_files = excel_to_ini(
                    input_path,
                    output_file or None,
                    self.ini_names,
                    prefer_source_paths=use_original_ini_location,
                )
                file_count = len(written_files or [])
                if use_original_ini_location:
                    first_file = written_files[0] if written_files else input_path
                    result_message = f'INI 已回写到原始位置（{file_count} 个文件）'
                    normalized_output = normalize_relative_path(Path(first_file).parent)
                else:
                    result_message = f'INI 目录已创建：{normalize_relative_path(output_file)}（{file_count} 个文件）'
                    normalized_output = normalize_relative_path(output_file)

            self._save_conversion_settings(input_rel, output_rel, output_filename, conversion_type, use_original_ini_location)
            return {'success': True, 'message': result_message, 'output_file': normalized_output}
        except Exception as exc:
            return {'success': False, 'message': f'转换失败：{str(exc)}'}

    def pick_w3x_map_file(self, payload: Dict[str, Any] | None = None):
        """选择需要拆包的 .w3x 地图文件。"""
        _ = payload
        settings = self.config.get('user_settings', {}).get('w3x', {})
        initial_dir = resolve_config_path(settings.get('map_path', '')) or str(BASE_DIR)
=======
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

    def pick_excel_file(self, payload: Dict[str, Any] | None = None):
        """选择 Excel 文件。"""
        payload = payload or {}
        map_folder = payload.get('map_folder', '')
        initial_dir = map_folder if map_folder and os.path.isdir(map_folder) else str(BASE_DIR)
        
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74
        result = window.create_file_dialog(
            webview.OPEN_DIALOG,
            directory=initial_dir,
            allow_multiple=False,
<<<<<<< HEAD
            file_types=['Warcraft III 地图 (*.w3x)'],
        )
        if result:
            return {'path': normalize_relative_path(result[0]), 'success': True}
        return {'path': None, 'success': True}

    def pick_w3x_workspace_folder(self, payload: Dict[str, Any] | None = None):
        """选择地图拆包目录。"""
        _ = payload
        settings = self.config.get('user_settings', {}).get('w3x', {})
        initial_dir = resolve_config_path(settings.get('workspace_path', '')) or str(BASE_DIR)
        result = window.create_file_dialog(webview.FOLDER_DIALOG, directory=initial_dir)
        if result:
            return {'path': normalize_relative_path(result[0]), 'success': True}
        return {'path': None, 'success': True}

    def pick_packed_w3x_file(self, payload: Dict[str, Any] | None = None):
        """选择装包输出 .w3x 文件路径的父目录，并生成输出文件名。"""
        payload = payload or {}
        settings = self.config.get('user_settings', {}).get('w3x', {})
        initial_dir = resolve_config_path(settings.get('packed_map_path', '')) or str(BASE_DIR)
        result = window.create_file_dialog(webview.FOLDER_DIALOG, directory=initial_dir)
        if not result:
            return {'path': None, 'success': True}

        filename = (payload.get('filename') or '').strip()
        if not filename:
            filename = 'output.w3x'
        path = Path(result[0]) / filename
        return {'path': normalize_relative_path(ensure_suffix(str(path), '.w3x')), 'success': True}

    def _get_w2l_path(self):
        user_settings = self.config.get('user_settings', {})
        configured = user_settings.get('w3x2lni_path', '')
        if not configured:
            raise ValueError('请先在设置中配置 w2l.exe 路径。')

        w2l_path = Path(resolve_config_path(configured)).resolve()
        if not w2l_path.exists() or w2l_path.name.lower() != 'w2l.exe':
            raise ValueError(f'w2l.exe 路径无效：{w2l_path}')
        return w2l_path

    def _run_w2l(self, args):
        w2l_path = self._get_w2l_path()
        completed = subprocess.run(
            [str(w2l_path), *[str(arg) for arg in args]],
            cwd=str(w2l_path.parent),
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform.startswith('win') else 0,
        )
        output = '\n'.join(part for part in [completed.stdout.strip(), completed.stderr.strip()] if part)
        if completed.returncode != 0:
            raise RuntimeError(output or f'w2l.exe 执行失败，退出码：{completed.returncode}')
        return output

    def unpack_w3x(self, payload: Dict[str, Any] | None = None):
        """调用 w2l.exe unpack 完成 w3x 拆包。"""
        payload = payload or {}
        map_rel = (payload.get('map_path') or '').strip()
        workspace_rel = (payload.get('workspace_path') or '').strip()
        if not map_rel:
            return {'success': False, 'message': '请选择 .w3x 地图文件。'}
        if not workspace_rel:
            return {'success': False, 'message': '请选择拆包输出目录。'}

        map_path = Path(resolve_config_path(map_rel)).resolve()
        workspace_path = get_safe_unpack_directory(Path(resolve_config_path(workspace_rel)))
        if not map_path.exists():
            return {'success': False, 'message': f'地图文件不存在：{map_path}'}

        try:
            workspace_path.mkdir(parents=True, exist_ok=True)
            self._run_w2l(['unpack', map_path, workspace_path])
            table_path = workspace_path / 'table'
            normalized_workspace = normalize_relative_path(workspace_path)
            self._save_w3x_settings({'map_path': map_rel, 'workspace_path': normalized_workspace})
            return {
                'success': True,
                'message': f'地图已拆包：{normalized_workspace}',
                'workspace_path': normalized_workspace,
                'table_path': normalize_relative_path(table_path) if table_path.exists() else '',
            }
        except Exception as exc:
            return {'success': False, 'message': f'拆包失败：{exc}'}

    def pack_w3x(self, payload: Dict[str, Any] | None = None):
        """调用 w2l.exe pack 完成地图目录装包。"""
        payload = payload or {}
        workspace_rel = (payload.get('workspace_path') or '').strip()
        packed_rel = (payload.get('packed_map_path') or '').strip()
        if not workspace_rel:
            return {'success': False, 'message': '请选择已拆包的地图目录。'}
        if not packed_rel:
            return {'success': False, 'message': '请选择装包输出 .w3x 路径。'}

        workspace_path = Path(resolve_config_path(workspace_rel)).resolve()
        packed_path = Path(ensure_suffix(resolve_config_path(packed_rel), '.w3x')).resolve()
        if not workspace_path.exists():
            return {'success': False, 'message': f'拆包目录不存在：{workspace_path}'}

        try:
            packed_path.parent.mkdir(parents=True, exist_ok=True)
            self._run_w2l(['pack', workspace_path, packed_path])
            normalized_packed = normalize_relative_path(packed_path)
            self._save_w3x_settings({'workspace_path': workspace_rel, 'packed_map_path': normalized_packed})
            return {
                'success': True,
                'message': f'地图已装包：{normalized_packed}',
                'packed_map_path': normalized_packed,
                'output_file': normalized_packed,
            }
        except Exception as exc:
            return {'success': False, 'message': f'装包失败：{exc}'}
=======
            file_types=['Excel 文件 (*.xlsx;*.xls)'],
        )
        if result:
            selected = result[0]
            return {
                'path': normalize_relative_path(selected),
                'path_abs': str(Path(selected).resolve()),
            }
        return {'path': None}

    def auto_detect_excel_file(self, payload: Dict[str, Any] | None = None):
        """自动检测地图文件夹内的同名 Excel 文件。"""
        payload = payload or {}
        map_folder = payload.get('map_folder', '')
        if not map_folder or not os.path.isdir(map_folder):
            return {'excel_path': None}
        
        # 获取文件夹名称
        folder_name = Path(map_folder).name
        # 如果是 table 目录，使用父目录名称
        if folder_name.lower() == 'table':
            folder_name = Path(map_folder).parent.name
        
        # 检查同目录下的同名 xlsx 文件
        parent_dir = Path(map_folder).parent if Path(map_folder).name.lower() == 'table' else Path(map_folder)
        xlsx_path = parent_dir / f'{folder_name}.xlsx'
        
        if xlsx_path.exists():
            return {
                'excel_path': str(xlsx_path),
                'excel_path_rel': normalize_relative_path(str(xlsx_path)),
            }
        
        return {'excel_path': None}

    def run_excel_to_ini_conversion(self, payload: Dict[str, Any] | None = None):
        """执行 Excel 转 INI 转换。"""
        payload = payload or {}
        map_folder = (payload.get('map_folder') or '').strip()
        excel_file = (payload.get('excel_file') or '').strip()
        
        if not map_folder:
            return {'success': False, 'message': '请选择地图文件夹'}
        if not excel_file:
            return {'success': False, 'message': '请选择 Excel 文件'}
        if not os.path.isdir(map_folder):
            return {'success': False, 'message': f'地图文件夹不存在：{map_folder}'}
        if not os.path.isfile(excel_file):
            return {'success': False, 'message': f'Excel 文件不存在：{excel_file}'}
        
        try:
            self._refresh_config()
            excel_to_ini(map_folder, excel_file)
            
            # 保存配置
            user_settings = dict(self.config.get('user_settings') or {})
            user_settings['excel_to_ini_map_path'] = normalize_relative_path(map_folder)
            user_settings['excel_file_path'] = normalize_relative_path(excel_file)
            self.config['user_settings'] = user_settings
            save_config(self.config)
            
            return {
                'success': True,
                'message': f'INI 文件已写入：{map_folder}',
            }
        except Exception as exc:
            return {'success': False, 'message': f'转换失败：{str(exc)}'}

    def save_excel_to_ini_paths(self, payload: Dict[str, Any] | None = None):
        """保存 Excel 转 INI 的路径配置。"""
        payload = payload or {}
        map_path = (payload.get('map_path') or '').strip()
        excel_path = (payload.get('excel_path') or '').strip()
        
        self._refresh_config()
        user_settings = dict(self.config.get('user_settings') or {})
        user_settings['excel_to_ini_map_path'] = map_path
        user_settings['excel_file_path'] = excel_path
        self.config['user_settings'] = user_settings
        save_config(self.config)
        
        return {'success': True}
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74

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
<<<<<<< HEAD
    print(f"INI 输入路径：{user_settings['ini_to_excel']['input_path']}")
    print(f"Excel 输入路径：{user_settings['excel_to_ini']['input_path']}")
=======
    print(f"输入路径：{user_settings['input_path']}")
>>>>>>> c7000fd00ff82f4af44147a5d8ffb0a672e2ed74
    print('启动 WebView 界面...')

    global window
    window = webview.create_window(
        'INI-Excel 转换工具',
        url=WEBUI_INDEX.as_uri(),
        js_api=api,
        width=1120,
        height=680,
        min_size=(980, 670),
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
