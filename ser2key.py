#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
Serial to keyboard:
Author: Akihiko Fujita
Version: 1.6

Copyright 2025 Akihiko Fujita

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

import ctypes
from ctypes import wintypes
import serial
from serial.tools import list_ports
import threading
import time
import configparser
import pystray
from pystray import MenuItem as item
import sys
import os
import logging
from logging.handlers import RotatingFileHandler
from contextlib import contextmanager
import codecs
import re
from datetime import datetime
import subprocess
import struct

if os.name != 'nt':
    raise OSError('ser2key は Windows 専用のアプリケーションです。')


# 定数定義
DEFAULT_ICON_SIZE = (32, 32)
ICON_FILES = ['f.ico']
APP_NAME = 'ser2key'
BAUDRATE_OPTIONS = [1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200]
BYTESIZE_OPTIONS = [5, 6, 7, 8]
PARITY_OPTIONS = ['N', 'E', 'O', 'M', 'S']
STOPBITS_OPTIONS = [1, 1.5, 2]
TIMEOUT_OPTIONS = [0.1, 0.5, 1, 2, 5]
ENCODING_OPTIONS = ['shift_jis', 'ascii', 'utf-8']
ENCODING_LABELS = {
    'shift_jis': 'Shift-JIS',
    'ascii': 'ASCII',
    'utf-8': 'UTF-8',
}


# Nuitka のワンファイル実行時に設定されるパス情報を取得
def _get_nuitka_onefile_dirs():
    dirs = []
    # 解凍先ディレクトリ（同梱データはここに展開される）
    temp_dir = os.environ.get('NUITKA_ONEFILE_TEMP')
    if temp_dir:
        dirs.append(os.path.abspath(temp_dir))

    # 元の実行ファイルが存在するディレクトリ（config.ini をここに置く想定）
    parent_dir = os.environ.get('NUITKA_ONEFILE_PARENT')
    if parent_dir:
        dirs.append(os.path.abspath(parent_dir))

    return dirs


# Nuitka Onefile の場合は、元の実行ファイルの配置場所を最優先で返す
def _get_executable_directory():
    nuitka_dirs = _get_nuitka_onefile_dirs()
    if nuitka_dirs:
        return nuitka_dirs[-1]

    # 起動時のコマンドラインに含まれるパス（実行ファイルが置かれているディレクトリ）
    argv_dir = os.path.dirname(os.path.abspath(sys.argv[0])) if sys.argv else ''

    # frozen 実行時は sys.executable が展開先を指すことがあるため argv ベースを優先
    if getattr(sys, 'frozen', False):
        if argv_dir:
            return argv_dir
        return os.path.dirname(os.path.abspath(sys.executable))

    return argv_dir or os.path.dirname(os.path.abspath(__file__))


def _get_storage_directory():
    base_dir = os.getenv('APPDATA')
    if not base_dir:
        base_dir = os.path.expanduser('~')
    return os.path.join(base_dir, APP_NAME)


APP_DIR = _get_executable_directory()
STORAGE_DIR = _get_storage_directory()
CONFIG_CANDIDATE_PATHS = [
    os.path.join(APP_DIR, 'config.ini'),
    os.path.join(STORAGE_DIR, 'config.ini'),
]
LOG_FILENAME = 'ser2key.log'


def _get_preferred_storage_dir():
    if os.path.exists(CONFIG_CANDIDATE_PATHS[0]):
        return APP_DIR
    return STORAGE_DIR


def ensure_storage_directory(storage_dir=None):
    target_dir = storage_dir or STORAGE_DIR
    try:
        os.makedirs(target_dir, exist_ok=True)
    except OSError as exc:
        raise RuntimeError(
            f'設定保存用ディレクトリ {target_dir} を作成できません: {exc}'
        ) from exc


DATETIME_TOKEN_PATTERN = re.compile(r'\{(DATE|TIME|DATETIME)(?::([^}]*))?\}')
DEFAULT_DATETIME_FORMATS = {
    'DATE': '%Y-%m-%d',
    'TIME': '%H:%M:%S',
    'DATETIME': '%Y-%m-%d %H:%M:%S',
}


RECONNECT_DELAY = 5     # 再接続までの待機時間（秒）
ACTIVITY_TIMEOUT = 300  # アクティビティタイムアウト（秒）
MONITOR_INTERVAL = 60   # モニタリング間隔（秒）
MAX_ERRORS = 10         # 最大エラー回数
CLIPBOARD_TIMEOUT = 5   # クリップボード操作タイムアウト（秒）
CLIPBRD_E_CANT_OPEN = 0x800401D0
CLIPBRD_E_EMPTY = 0x800401D1

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002

VK_CONTROL = 0x11
VK_V = 0x56
VK_RETURN = 0x0D
KEYEVENTF_KEYUP = 0x0002
IMAGE_ICON = 1
LR_LOADFROMFILE = 0x00000010
DIB_RGB_COLORS = 0
DI_NORMAL = 0x0003

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
gdi32 = ctypes.windll.gdi32
ole32 = ctypes.OleDLL('ole32')

ole32.OleInitialize.restype = ctypes.c_long
ole32.OleInitialize.argtypes = [ctypes.c_void_p]
ole32.OleGetClipboard.restype = ctypes.c_long
ole32.OleGetClipboard.argtypes = [ctypes.POINTER(ctypes.c_void_p)]
ole32.OleSetClipboard.restype = ctypes.c_long
ole32.OleSetClipboard.argtypes = [ctypes.c_void_p]
ole32.OleFlushClipboard.restype = ctypes.c_long
ole32.OleFlushClipboard.argtypes = []

S_OK = 0
S_FALSE = 1


class _BitmapInfoHeader(ctypes.Structure):
    _fields_ = [
        ('biSize', wintypes.DWORD),
        ('biWidth', wintypes.LONG),
        ('biHeight', wintypes.LONG),
        ('biPlanes', wintypes.WORD),
        ('biBitCount', wintypes.WORD),
        ('biCompression', wintypes.DWORD),
        ('biSizeImage', wintypes.DWORD),
        ('biXPelsPerMeter', wintypes.LONG),
        ('biYPelsPerMeter', wintypes.LONG),
        ('biClrUsed', wintypes.DWORD),
        ('biClrImportant', wintypes.DWORD),
    ]


class _BitmapInfo(ctypes.Structure):
    _fields_ = [
        ('bmiHeader', _BitmapInfoHeader),
        ('bmiColors', wintypes.DWORD * 3),
    ]



class _IUnknownVTable(ctypes.Structure):
    _fields_ = [
        ('QueryInterface', ctypes.c_void_p),
        ('AddRef', ctypes.c_void_p),
        ('Release', ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)),
    ]


class _IUnknown(ctypes.Structure):
    _fields_ = [('lpVtbl', ctypes.POINTER(_IUnknownVTable))]


class ClipboardBackupData:
    """保持しているクリップボードデータのバックアップ"""

    def __init__(self, data_object=None, is_empty=False):
        self.data_object = data_object
        self.is_empty = is_empty

    # COM オブジェクトの参照カウントを解放
    def release(self):
        if self.data_object:
            obj_ptr = ctypes.cast(ctypes.c_void_p(self.data_object), ctypes.POINTER(_IUnknown))
            release = obj_ptr.contents.lpVtbl.contents.Release
            release(obj_ptr)
            self.data_object = None


_ole_thread_state = threading.local()


# 現在のスレッドで OLE を初期化
def ensure_ole_initialized():
    if getattr(_ole_thread_state, 'initialized', False):
        return

    hr = ole32.OleInitialize(None)
    if hr not in (S_OK, S_FALSE):
        raise ClipboardError(f"OLE の初期化に失敗しました (HRESULT=0x{hr:08X})")

    _ole_thread_state.initialized = True


# クリップボードが空であるか確認
def is_clipboard_empty():
    try:
        with open_clipboard():
            kernel32.SetLastError(0)
            first_format = user32.EnumClipboardFormats(0)
            if first_format == 0 and kernel32.GetLastError() == 0:
                return True
    except ClipboardError:
        pass
    return False


# クリップボードの全データを退避
def backup_clipboard():
    ensure_ole_initialized()

    empty = is_clipboard_empty()

    deadline = time.time() + CLIPBOARD_TIMEOUT
    last_error = None

    while True:
        data_object = ctypes.c_void_p()
        hr = ole32.OleGetClipboard(ctypes.byref(data_object))
        if hr >= 0:
            return ClipboardBackupData(data_object.value, empty)

        last_error = hr & 0xFFFFFFFF
        if last_error == CLIPBRD_E_EMPTY:
            return ClipboardBackupData(None, True)

        if last_error == CLIPBRD_E_CANT_OPEN and time.time() < deadline:
            time.sleep(0.05)
            continue

        raise ClipboardError(
            f"クリップボードの退避に失敗しました (HRESULT=0x{last_error:08X})"
        )


# 退避したクリップボード内容を復元
def restore_clipboard(backup):
    if not backup:
        return

    try:
        ensure_ole_initialized()
        if backup.is_empty:
            clear_clipboard()
        elif backup.data_object:
            hr = ole32.OleSetClipboard(ctypes.c_void_p(backup.data_object))
            if hr < 0:
                raise ClipboardError(f"クリップボードの復元に失敗しました (HRESULT=0x{hr & 0xFFFFFFFF:08X})")
        else:
            clear_clipboard()
    finally:
        backup.release()

class ClipboardError(Exception):
    """クリップボード操作に失敗した場合の例外"""


@contextmanager
# クリップボードを安全に開くためのコンテキストマネージャ
def open_clipboard():
    deadline = time.time() + CLIPBOARD_TIMEOUT
    opened = False
    while time.time() < deadline:
        if user32.OpenClipboard(None):
            opened = True
            break
        time.sleep(0.05)

    if not opened:
        raise ClipboardError('クリップボードを開けませんでした')
    try:
        yield
    finally:
        if opened:
            user32.CloseClipboard()


# 現在のクリップボード文字列を取得
def get_clipboard_text():
    with open_clipboard():
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return None
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            raise ClipboardError('クリップボードデータのロックに失敗しました')
        try:
            data = ctypes.wstring_at(ptr)
        finally:
            kernel32.GlobalUnlock(handle)
        return data


# クリップボードに文字列を設定
def set_clipboard_text(text):
    if text is None:
        return clear_clipboard()

    buffer = ctypes.create_unicode_buffer(text)
    size = ctypes.sizeof(buffer)
    handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, size)
    if not handle:
        raise ClipboardError('メモリ確保に失敗しました')

    ptr = kernel32.GlobalLock(handle)
    if not ptr:
        kernel32.GlobalFree(handle)
        raise ClipboardError('メモリロックに失敗しました')

    try:
        ctypes.memmove(ptr, ctypes.addressof(buffer), size)
    finally:
        kernel32.GlobalUnlock(handle)

    with open_clipboard():
        user32.EmptyClipboard()
        if not user32.SetClipboardData(CF_UNICODETEXT, handle):
            kernel32.GlobalFree(handle)
            raise ClipboardError('クリップボードへの設定に失敗しました')


# クリップボードを空にする
def clear_clipboard():
    with open_clipboard():
        user32.EmptyClipboard()


# キーイベントを送信
def _key_event(vk_code, key_up=False):
    scan_code = user32.MapVirtualKeyW(vk_code, 0)
    flags = KEYEVENTF_KEYUP if key_up else 0
    user32.keybd_event(vk_code, scan_code, flags, 0)


# Ctrl+V と必要に応じて Enter キーを送信
def send_ctrl_v(add_enter=False):
    _key_event(VK_CONTROL)
    _key_event(VK_V)
    _key_event(VK_V, key_up=True)
    _key_event(VK_CONTROL, key_up=True)

    if add_enter:
        time.sleep(0.05)
        _key_event(VK_RETURN)
        _key_event(VK_RETURN, key_up=True)


# ログ設定を初期化
def setup_logging():
    storage_dir = _get_preferred_storage_dir()
    try:
        if storage_dir == STORAGE_DIR:
            ensure_storage_directory(storage_dir=storage_dir)
    except RuntimeError as exc:
        raise RuntimeError(str(exc)) from exc

    logger = logging.getLogger('ser2key')
    logger.setLevel(logging.INFO)

    log_file = os.path.join(storage_dir, LOG_FILENAME)
    log_path = os.path.abspath(log_file)
    handler_exists = any(
        isinstance(handler, RotatingFileHandler) and
        getattr(handler, 'baseFilename', None) == log_path
        for handler in logger.handlers
    )

    if not handler_exists:
        handler = RotatingFileHandler(
            log_file,
            maxBytes=512*1024,
            backupCount=3,
            encoding='utf-8'
        )

        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger


class SerialKeyboardEmulator:
    """シリアル通信とキーボードエミュレーションを管理するクラス"""

    def __init__(self):
        self.serial_config = None
        self.settings_config = None
        self.tray_icon = None
        self.is_running = True
        self._lock = threading.Lock()
        self.logger = setup_logging()
        self.last_activity = time.time()
        self.error_count = 0
        self.available_ports = []
        self._config_parser = None
        self._reconnect_event = threading.Event()
        self.config_path = None
        self.output_config = {
            'header_template': '',
            'footer_template': '',
        }
        self._cleanup_event = threading.Event()

    def _resolve_config_path(self):
        for path in CONFIG_CANDIDATE_PATHS:
            if os.path.exists(path):
                if path.endswith('config.ini') and path != CONFIG_CANDIDATE_PATHS[0]:
                    self.logger.info(f"設定ファイルを互換パスから読み込みます: {path}")
                return path
        return CONFIG_CANDIDATE_PATHS[0]

    def _persist_config(self):
        if not self._config_parser or not self.config_path:
            return

        directory = os.path.dirname(self.config_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)

        with open(self.config_path, 'w', encoding='utf-8') as f:
            self._config_parser.write(f)

    def _decode_template(self, text, field_name):
        if not text:
            return ''
        try:
            return codecs.decode(text.encode('utf-8'), 'unicode_escape')
        except Exception as exc:
            self.logger.warning(
                f"[output] セクションの {field_name} のエスケープシーケンス解釈に失敗しました: {exc}"
            )
            return text

    def _render_template(self, template, now):
        if not template:
            return ''

        def replace(match):
            token = match.group(1)
            fmt = match.group(2) or DEFAULT_DATETIME_FORMATS[token]
            try:
                return now.strftime(fmt)
            except ValueError as exc:
                self.logger.warning(
                    f"[output] {token} の書式 '{fmt}' が不正なため置換できませんでした: {exc}"
                )
                return match.group(0)

        return DATETIME_TOKEN_PATTERN.sub(replace, template)

    def _apply_output_templates(self, payload):
        if payload is None:
            return ''

        now = datetime.now()
        header_template = self.output_config.get('header_template', '') if self.output_config else ''
        footer_template = self.output_config.get('footer_template', '') if self.output_config else ''
        header = self._render_template(header_template, now)
        footer = self._render_template(footer_template, now)
        return f"{header}{payload}{footer}"

    def _wait_clipboard_update(self, expected_text):
        deadline = time.time() + CLIPBOARD_TIMEOUT
        while time.time() < deadline:
            try:
                if get_clipboard_text() == expected_text:
                    return True
            except ClipboardError:
                time.sleep(0.05)
                continue
            time.sleep(0.05)
        return False

    def _restore_clipboard_safely(self, clipboard_backup):
        try:
            restore_clipboard(clipboard_backup)
        except ClipboardError:
            self.logger.warning("クリップボードの復元に失敗しました")

    def _update_error_state(self, success):
        with self._lock:
            if success:
                self.error_count = 0
            else:
                self.error_count += 1

    def _apply_clipboard_and_paste(self, formatted_payload, add_enter):
        clipboard_backup = None
        try:
            clipboard_backup = backup_clipboard()
            set_clipboard_text(formatted_payload)

            if not self._wait_clipboard_update(formatted_payload):
                raise ClipboardError('クリップボードの内容が更新されませんでした')

            send_ctrl_v(add_enter)
        finally:
            self._restore_clipboard_safely(clipboard_backup)


    # 終了要求または再接続要求を監視しながら待機
    def _wait_for_stop_or_reconnect(self, timeout):
        end_time = time.time() + timeout

        while self.is_running and time.time() < end_time:
            remaining = end_time - time.time()
            wait_time = min(0.1, remaining)
            if wait_time <= 0:
                break
            if self._reconnect_event.wait(wait_time):
                break


    # リソースの解放処理
    def cleanup(self):
        self.is_running = False
        self._reconnect_event.set()

        if self._cleanup_event.is_set():
            return

        self._cleanup_event.set()
        self.logger.info("アプリケーションのクリーンアップを実行")
        icon = self.tray_icon
        if icon:
            try:
                icon.visible = False
            except Exception:
                pass
            try:
                icon.stop()
            except Exception as exc:
                self.logger.debug(f"トレイアイコン停止時の例外: {exc}")


    # デフォルト設定ファイルの作成
    def create_default_config(self):
        config = configparser.ConfigParser()

        config['serial'] = {
            'port': 'COM8',
            'baudrate': '9600',
            'bytesize': '8',
            'parity': 'N',
            'stopbits': '1',
            'timeout': '1'
        }
        
        config['settings'] = {
            'add_enter': 'true',
            'encoding': 'shift_jis',
            'buffer_msec': '0'
        }

        config['output'] = {
            'header': '',
            'footer': ''
        }

        self._config_parser = config

        errors = []
        for candidate in CONFIG_CANDIDATE_PATHS:
            self.config_path = candidate
            try:
                self._persist_config()
            except OSError as exc:
                errors.append((candidate, exc))
                self.logger.warning(
                    f"設定ファイル {candidate} を作成できませんでした: {exc}"
                )
                continue

            self.logger.info(f"デフォルト設定ファイルを作成しました ({candidate})")
            self.output_config = {
                'header_template': '',
                'footer_template': '',
            }
            return

        message = "; ".join(
            f"{path}: {exc}" for path, exc in errors
        ) or '未知の理由'
        raise RuntimeError(f"設定ファイルを作成できませんでした ({message})")


    # シリアル通信の設定値を検証
    def validate_serial_config(self, config):
        baudrate = config['baudrate']
        if baudrate <= 0:
            raise ValueError(f"不正なボーレート: {baudrate}")

        self.refresh_available_ports(update_menu=False)
        self.logger.info(f"利用可能なポート: {self.available_ports}")

        if config['port'] not in self.available_ports:
            self.logger.warning(f"設定されたポート {config['port']} は現在利用可能なポート一覧にありません")
            # ここでもエラーは発生させず、警告のみ


    # 利用可能なシリアルポート一覧を更新
    def refresh_available_ports(self, update_menu=True):
        ports = [port.device for port in list_ports.comports()]
        ports.sort()
        self.available_ports = ports
        self.logger.info(f"ポート一覧を更新: {self.available_ports}")

        if update_menu and self.tray_icon:
            self.update_tray_menu()


    # タスクトレイアイコンを関連付け
    def attach_tray_icon(self, icon):
        self.tray_icon = icon
        self.update_tray_menu()


    # 現在選択されているポートかどうかを判定
    def _is_selected_port(self, port):
        with self._lock:
            current = self.serial_config.get('port') if self.serial_config else None
        return current == port


    # タスクトレイメニューを更新
    def update_tray_menu(self):
        if not self.tray_icon:
            return

        def refresh_ports_action(icon, menu_item):
            self.refresh_available_ports()

        def ensure_option_in_list(options, current):
            if current is None or current in options:
                return options
            return [current] + list(options)

        def format_value(value, unit=None):
            if isinstance(value, float) and value.is_integer():
                formatted = str(int(value))
            else:
                formatted = str(value)
            if unit:
                return f"{formatted} {unit}"
            return formatted

        def build_option_menu(label, key, options, unit=None, config_getter=None, update_handler=None, formatter=None):
            if config_getter is None or update_handler is None:
                raise ValueError("config_getter と update_handler は必須です")

            def get_current():
                return config_getter(key)

            current = get_current()
            option_values = ensure_option_in_list(options, current)
            menu_items = []
            for option in option_values:
                def make_action(value):
                    def select_action(icon, menu_item):
                        update_handler(key, value)

                    return select_action

                def make_checked(value):
                    def is_checked(menu_item):
                        return get_current() == value

                    return is_checked

                label_text = formatter(option) if formatter else format_value(option, unit=unit)
                menu_items.append(
                    item(
                        label_text,
                        make_action(option),
                        checked=make_checked(option)
                    )
                )
            if not menu_items:
                menu_items.append(item('選択肢なし', None, enabled=False))
            return item(label, pystray.Menu(*menu_items))

        def get_serial_current(key):
            with self._lock:
                return self.serial_config.get(key) if self.serial_config else None

        def get_settings_current(key):
            with self._lock:
                return self.settings_config.get(key) if self.settings_config else None

        def format_encoding(value):
            return ENCODING_LABELS.get(value, value)

        menu_definition = [
            item('ポート一覧を更新', refresh_ports_action),
            pystray.Menu.SEPARATOR
        ]

        if self.available_ports:
            for port in self.available_ports:
                def make_select_action(target_port):
                    def select_action(icon, menu_item):
                        self.update_serial_port(target_port)

                    return select_action

                def make_checked(target_port):
                    def is_checked(menu_item):
                        return self._is_selected_port(target_port)

                    return is_checked

                menu_definition.append(
                    item(
                        f'接続: {port}',
                        make_select_action(port),
                        checked=make_checked(port)
                    )
                )
        else:
            menu_definition.append(item('ポートが見つかりません', None, enabled=False))

        menu_definition.extend([
            pystray.Menu.SEPARATOR,
            item(
                'シリアル設定',
                pystray.Menu(
                    build_option_menu(
                        'ボーレート',
                        'baudrate',
                        BAUDRATE_OPTIONS,
                        config_getter=get_serial_current,
                        update_handler=self.update_serial_setting
                    ),
                    build_option_menu(
                        'データ長',
                        'bytesize',
                        BYTESIZE_OPTIONS,
                        config_getter=get_serial_current,
                        update_handler=self.update_serial_setting
                    ),
                    build_option_menu(
                        'パリティ',
                        'parity',
                        PARITY_OPTIONS,
                        config_getter=get_serial_current,
                        update_handler=self.update_serial_setting
                    ),
                    build_option_menu(
                        'ストップビット',
                        'stopbits',
                        STOPBITS_OPTIONS,
                        config_getter=get_serial_current,
                        update_handler=self.update_serial_setting
                    ),
                    build_option_menu(
                        'タイムアウト',
                        'timeout',
                        TIMEOUT_OPTIONS,
                        unit='秒',
                        config_getter=get_serial_current,
                        update_handler=self.update_serial_setting
                    ),
                )
            ),
            item(
                'デコード設定',
                pystray.Menu(
                    build_option_menu(
                        '文字コード',
                        'encoding',
                        ENCODING_OPTIONS,
                        config_getter=get_settings_current,
                        update_handler=self.update_settings_setting,
                        formatter=format_encoding
                    ),
                )
            ),
            pystray.Menu.SEPARATOR,
            item('終了', self.stop_application)
        ])

        try:
            self.tray_icon.menu = pystray.Menu(*menu_definition)
        except Exception as e:
            self.logger.error(f"トレイメニュー更新エラー: {e}")


    # 利用するシリアルポートを更新し再接続を要求
    def update_serial_port(self, port):
        if not self.serial_config:
            self.logger.error("シリアル設定が初期化されていません")
            return

        current_port = self.serial_config.get('port')
        if port == current_port:
            self.logger.info(f"ポート {port} は既に選択されています")
            return
        if port not in self.available_ports:
            self.logger.warning(f"ポート {port} は現在の一覧に存在しませんが接続を試行します")

        with self._lock:
            self.serial_config['port'] = port
            if self._config_parser and 'serial' in self._config_parser and self.config_path:
                self._config_parser['serial']['port'] = port
                try:
                    self._persist_config()
                except OSError as exc:
                    self.logger.error(
                        f"設定ファイル {self.config_path} への書き込みに失敗しました: {exc}"
                    )

        self.logger.info(f"シリアルポートを {port} に更新しました。再接続を実行します")
        self._reconnect_event.set()
        self.update_tray_menu()


    # シリアル設定値を更新し再接続を要求
    def update_serial_setting(self, key, value):
        if not self.serial_config:
            self.logger.error("シリアル設定が初期化されていません")
            return

        with self._lock:
            current_value = self.serial_config.get(key)

        if current_value == value:
            self.logger.info(f"{key} は既に {value} に設定されています")
            return

        if key == 'parity' and isinstance(value, str):
            value = value.upper()

        with self._lock:
            self.serial_config[key] = value
            if self._config_parser and 'serial' in self._config_parser and self.config_path:
                self._config_parser['serial'][key] = str(value)
                try:
                    self._persist_config()
                except OSError as exc:
                    self.logger.error(
                        f"設定ファイル {self.config_path} への書き込みに失敗しました: {exc}"
                    )

        self.logger.info(f"{key} を {value} に更新しました。再接続を実行します")
        self._reconnect_event.set()
        self.update_tray_menu()


    # 一般設定値を更新
    def update_settings_setting(self, key, value):
        if not self.settings_config:
            self.logger.error("一般設定が初期化されていません")
            return

        if key == 'encoding' and isinstance(value, str):
            try:
                value = codecs.lookup(value).name
            except LookupError:
                self.logger.error(f"サポートされていないエンコーディング: {value}")
                return

        with self._lock:
            current_value = self.settings_config.get(key)

        if current_value == value:
            self.logger.info(f"{key} は既に {value} に設定されています")
            return

        with self._lock:
            self.settings_config[key] = value
            if self._config_parser and 'settings' in self._config_parser and self.config_path:
                self._config_parser['settings'][key] = str(value)
                try:
                    self._persist_config()
                except OSError as exc:
                    self.logger.error(
                        f"設定ファイル {self.config_path} への書き込みに失敗しました: {exc}"
                    )

        self.logger.info(f"{key} を {value} に更新しました")
        self.update_tray_menu()


    # アプリケーション終了処理
    def stop_application(self, icon, menu_item=None):
        self.logger.info("ユーザー操作によりアプリケーションを終了します")
        self.cleanup()


    # 一般設定の検証
    def validate_settings_config(self, config):
        encoding_name = config['encoding']
        try:
            normalized = codecs.lookup(encoding_name).name
        except (LookupError, TypeError):
            raise ValueError(f"サポートされていないエンコーディング: {encoding_name}")

        config['encoding'] = normalized


    # 設定ファイルを読み込む
    def read_config(self):
        config_path = self._resolve_config_path()

        if not os.path.exists(config_path):
            self.logger.warning("設定ファイルが見つかりません。デフォルト設定を作成します。")
            self.config_path = config_path
            self.create_default_config()
            config_path = self.config_path
        else:
            self.config_path = config_path

        config = configparser.ConfigParser()
        try:
            read_files = config.read(config_path, encoding='utf-8')
        except OSError as exc:
            raise RuntimeError(f"設定ファイル {config_path} を読み込めませんでした: {exc}") from exc

        if not read_files:
            raise RuntimeError(f"設定ファイル {config_path} の読み込みに失敗しました")

        self._config_parser = config
        self.logger.info(f"設定ファイルを読み込みました: {config_path}")

        if not config.has_section('output'):
            config.add_section('output')

        try:
            self.serial_config = {
                'port': config['serial']['port'],
                'baudrate': int(config['serial']['baudrate']),
                'bytesize': int(config['serial']['bytesize']),
                'parity': config['serial']['parity'],
                'stopbits': float(config['serial']['stopbits']),
                'timeout': float(config['serial']['timeout'])
            }
            self.settings_config = {
                'add_enter': config.getboolean('settings', 'add_enter', fallback=True),
                'encoding': config.get('settings', 'encoding', fallback='sjis'),
                'buffer_msec': config.getint('settings', 'buffer_msec', fallback=0)
            }

            raw_header = config.get('output', 'header', fallback='')
            raw_footer = config.get('output', 'footer', fallback='')
            self.output_config = {
                'header_template': self._decode_template(raw_header, 'header'),
                'footer_template': self._decode_template(raw_footer, 'footer'),
            }

            self.validate_serial_config(self.serial_config)
            self.validate_settings_config(self.settings_config)

            if self._config_parser and self.config_path:
                try:
                    self._config_parser['settings']['encoding'] = self.settings_config['encoding']
                except KeyError:
                    pass

        except KeyError as e:
            raise KeyError(f"設定キーが見つかりません: {e}")
        except ValueError as e:
            raise ValueError(f"不正な設定値です: {e}")

    # シリアルデータを処理してキーボード入力をシミュレート
    def process_serial_data(self, data):
        if not data:
            return

        payload = str(data)
        formatted_payload = self._apply_output_templates(payload)

        self.logger.info(f"受信: {payload}")
        if formatted_payload != payload:
            self.logger.info(f"整形後: {formatted_payload}")
        self.last_activity = time.time()
        with self._lock:
            add_enter = self.settings_config.get('add_enter', False) if self.settings_config else False

        try:
            self._apply_clipboard_and_paste(formatted_payload, add_enter)
        except ClipboardError as e:
            self.logger.error(f"クリップボード操作エラー: {e}")
            self._update_error_state(False)
        except Exception as e:
            self.logger.error(f"入力操作エラー: {e}")
            self._update_error_state(False)
        else:
            self._update_error_state(True)

    # シリアルポートからデータを読み取り、キーボード入力に変換
    def read_serial_and_type(self):
        while self.is_running:
            try:
                with self._lock:
                    serial_kwargs = dict(self.serial_config) if self.serial_config else None
                    encoding = self.settings_config.get('encoding') if self.settings_config else 'utf-8'

                if not serial_kwargs:
                    self.logger.error("シリアル設定が初期化されていません。再試行します")
                    time.sleep(RECONNECT_DELAY)
                    continue

                port = serial_kwargs.get('port', '未設定')
                self.logger.info(f"シリアルポート {port} への接続を試行します")

                with serial.Serial(**serial_kwargs) as ser:
                    if self.tray_icon:
                        self.tray_icon.title = f"ser2key - 接続中 ({port})"
                    self.logger.info("シリアルポートに接続しました")
                    with self._lock:
                        self.error_count = 0
                    self._reconnect_event.clear()

                    buffer_data = []
                    buffer_start_time = None

                    while self.is_running and not self._reconnect_event.is_set():
                        with self._lock:
                            buffer_duration_ms = self.settings_config.get('buffer_msec', 0) if self.settings_config else 0

                        if buffer_duration_ms <= 0 and buffer_data:
                            combined = '\n'.join(buffer_data)
                            if combined:
                                self.process_serial_data(combined)
                            buffer_data = []
                            buffer_start_time = None

                        if ser.in_waiting > 0:
                            try:
                                data = ser.readline().decode(encoding).rstrip('\r\n')
                                if not data:
                                    pass
                                elif buffer_duration_ms <= 0:
                                    self.process_serial_data(data)
                                else:
                                    if buffer_start_time is None:
                                        buffer_start_time = time.monotonic()
                                        buffer_data = []
                                    buffer_data.append(data)
                            except UnicodeDecodeError as e:
                                self.logger.error(f"デコードエラー: {e}")
                                with self._lock:
                                    self.error_count += 1
                            except Exception as e:
                                self.logger.error(f"予期せぬエラー: {e}")
                                with self._lock:
                                    self.error_count += 1

                        if buffer_duration_ms > 0 and buffer_start_time is not None:
                            elapsed = (time.monotonic() - buffer_start_time) * 1000
                            if elapsed >= buffer_duration_ms:
                                combined = '\n'.join(buffer_data)
                                if combined:
                                    self.process_serial_data(combined)
                                buffer_data = []
                                buffer_start_time = None

                        if self._reconnect_event.is_set():
                            self.logger.info("再接続要求を受け取りました")
                            if buffer_data:
                                combined = '\n'.join(buffer_data)
                                if combined:
                                    self.process_serial_data(combined)
                            buffer_data = []
                            buffer_start_time = None
                            break

                        time.sleep(0.05)

                    if buffer_data:
                        combined = '\n'.join(buffer_data)
                        if combined:
                            self.process_serial_data(combined)

            except serial.SerialException as e:
                if self.tray_icon:
                    self.tray_icon.title = f"ser2key - 接続失敗 ({e})"
                self.logger.error(f"シリアル通信エラー: {e}")
                with self._lock:
                    self.error_count += 1
                self._wait_for_stop_or_reconnect(RECONNECT_DELAY)
            except Exception as e:
                if self.tray_icon:
                    self.tray_icon.title = f"ser2key - エラー ({e})"
                self.logger.error(f"予期せぬエラー: {e}")
                with self._lock:
                    self.error_count += 1
                self._wait_for_stop_or_reconnect(RECONNECT_DELAY)
            finally:
                if self.tray_icon and self.is_running:
                    self.tray_icon.title = "ser2key - 切断"

            if self._reconnect_event.is_set():
                self._reconnect_event.clear()
                if self.is_running:
                    time.sleep(1)


class ApplicationMonitor:
    """アプリケーションの状態を監視するクラス"""

    def __init__(self, emulator):
        self.emulator = emulator
        self.logger = logging.getLogger('ser2key.monitor')

    # 定期的な状態確認
    def monitor(self):
        while self.emulator.is_running:
            try:
                # アクティビティタイムアウトチェック
                if time.time() - self.emulator.last_activity > ACTIVITY_TIMEOUT:
                    self.logger.warning("長時間データ受信がありません")

                # エラー回数チェック
                with self.emulator._lock:
                    current_errors = self.emulator.error_count
                if current_errors >= MAX_ERRORS:
                    self.logger.error("エラー回数が上限を超えました。アプリケーションを停止します")
                    self.emulator.cleanup()
                    break

                time.sleep(MONITOR_INTERVAL)

            except Exception as e:
                self.logger.error(f"モニタリングエラー: {e}")


class SimpleIconImage:
    """Pillow 非依存で pystray が扱えるシンプルな画像表現"""

    def __init__(self, width, height, data, mode='RGBA'):
        self.width = width
        self.height = height
        self.size = (width, height)
        self.mode = mode
        self._data = bytes(data)

    def copy(self):
        return SimpleIconImage(self.width, self.height, self._data, mode=self.mode)

    def resize(self, size, resample=None):
        target_w, target_h = size
        source_w, source_h = self.size
        if (target_w, target_h) == (source_w, source_h):
            return self.copy()

        # 最近傍補間で縮小/拡大
        resized = bytearray(target_w * target_h * 4)
        data_rgba = self._to_rgba()
        for y in range(target_h):
            src_y = int(y * source_h / target_h)
            for x in range(target_w):
                src_x = int(x * source_w / target_w)
                src_index = (src_y * source_w + src_x) * 4
                dst_index = (y * target_w + x) * 4
                resized[dst_index:dst_index + 4] = data_rgba[src_index:src_index + 4]
        return SimpleIconImage(target_w, target_h, bytes(resized), mode='RGBA')

    def _to_rgba(self):
        if self.mode == 'RGBA':
            return self._data
        if self.mode == 'BGRA':
            converted = bytearray()
            for i in range(0, len(self._data), 4):
                b, g, r, a = self._data[i:i + 4]
                converted.extend((r, g, b, a))
            return bytes(converted)
        if self.mode == 'RGB':
            converted = bytearray()
            for i in range(0, len(self._data), 3):
                r, g, b = self._data[i:i + 3]
                converted.extend((r, g, b, 255))
            return bytes(converted)
        raise ValueError('サポートされていないモードです')

    def _to_bgra(self):
        if self.mode == 'BGRA':
            return self._data
        if self.mode == 'RGBA':
            converted = bytearray()
            for i in range(0, len(self._data), 4):
                r, g, b, a = self._data[i:i + 4]
                converted.extend((b, g, r, a))
            return bytes(converted)
        if self.mode == 'RGB':
            converted = bytearray()
            for i in range(0, len(self._data), 3):
                r, g, b = self._data[i:i + 3]
                converted.extend((b, g, r, 255))
            return bytes(converted)
        raise ValueError('サポートされていないモードです')

    def convert(self, mode):
        if mode == self.mode:
            return self
        if mode == 'RGBA':
            return SimpleIconImage(self.width, self.height, self._to_rgba(), mode='RGBA')
        if mode == 'BGRA':
            return SimpleIconImage(self.width, self.height, self._to_bgra(), mode='BGRA')
        if mode == 'RGB':
            rgba = self._to_rgba()
            rgb = bytearray()
            for i in range(0, len(rgba), 4):
                r, g, b = rgba[i:i + 3]
                rgb.extend((r, g, b))
            return SimpleIconImage(self.width, self.height, bytes(rgb), mode='RGB')
        raise ValueError('サポートされていないモードです')

    def tobytes(self, *args):
        if not args:
            return self._to_rgba()
        if args[0] == 'raw':
            if len(args) > 1 and args[1] == 'BGRA':
                return self._to_bgra()
            if len(args) > 1 and args[1] == 'RGBA':
                return self._to_rgba()
            if len(args) > 1 and args[1] == 'RGB':
                if self.mode == 'RGB':
                    return self._data
                rgba = self._to_rgba()
                rgb = bytearray()
                for i in range(0, len(rgba), 4):
                    r, g, b = rgba[i:i + 3]
                    rgb.extend((r, g, b))
                return bytes(rgb)
        return self._to_rgba()

    def save(self, fp, format=None):
        if format and format.upper() != 'ICO':
            raise ValueError('サポートされていないフォーマットです')

        width, height = self.size
        bgra = self._to_bgra()

        row_bytes = width * 4
        pixel_data = bytearray()
        for y in range(height - 1, -1, -1):
            start = y * row_bytes
            pixel_data.extend(bgra[start:start + row_bytes])

        mask_row_bytes = ((width + 31) // 32) * 4
        mask_data = bytearray(mask_row_bytes * height)

        info_header = struct.pack(
            '<IIIHHIIIIII',
            40,
            width,
            height * 2,
            1,
            32,
            0,
            len(pixel_data) + len(mask_data),
            0,
            0,
            0,
            0,
        )

        image_data = info_header + pixel_data + mask_data

        ico_header = struct.pack('<HHH', 0, 1, 1)
        dir_entry = struct.pack(
            '<BBBBHHII',
            width if width < 256 else 0,
            height if height < 256 else 0,
            0,
            0,
            1,
            32,
            len(image_data),
            6 + 16,
        )
        fp.write(ico_header)
        fp.write(dir_entry)
        fp.write(image_data)

class TrayIconManager:
    """タスクトレイアイコンを管理するクラス"""

    _BitmapInfoHeader = _BitmapInfoHeader
    _BitmapInfo = _BitmapInfo

    @staticmethod
    # デフォルトアイコンを作成
    def create_default_icon():
        width, height = DEFAULT_ICON_SIZE
        data = bytearray()
        for y in range(height):
            for x in range(width):
                if x < width // 2:
                    data.extend((0, 120, 215, 255))  # Windows ライクな青
                else:
                    data.extend((255, 255, 255, 255))
        return SimpleIconImage(width, height, data)

    @staticmethod
    def _load_with_pillow(image_path):
        try:
            from PIL import Image
        except Exception:
            return None

        try:
            with Image.open(image_path) as img:
                return img.convert('RGBA')
        except Exception as exc:
            logging.getLogger('ser2key').warning(f"Pillow によるアイコン読み込みに失敗: {exc}")
            return None

    @staticmethod
    def _load_with_win32_icon(image_path, target_size):
        if not image_path.lower().endswith('.ico'):
            return None

        width, height = target_size
        hicon = user32.LoadImageW(
            None,
            image_path,
            IMAGE_ICON,
            width,
            height,
            LR_LOADFROMFILE,
        )
        if not hicon:
            return None

        hdc = gdi32.CreateCompatibleDC(None)
        if not hdc:
            user32.DestroyIcon(hicon)
            return None

        bmi = TrayIconManager._BitmapInfo()
        bmi.bmiHeader.biSize = ctypes.sizeof(TrayIconManager._BitmapInfoHeader)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0
        bmi.bmiHeader.biSizeImage = width * height * 4

        bits = ctypes.c_void_p()
        hbitmap = gdi32.CreateDIBSection(
            hdc,
            ctypes.byref(bmi),
            DIB_RGB_COLORS,
            ctypes.byref(bits),
            None,
            0,
        )
        if not hbitmap:
            gdi32.DeleteDC(hdc)
            user32.DestroyIcon(hicon)
            return None

        old_obj = gdi32.SelectObject(hdc, hbitmap)
        try:
            user32.DrawIconEx(hdc, 0, 0, hicon, width, height, 0, None, DI_NORMAL)
            if not bits:
                return None
            data = ctypes.string_at(bits, width * height * 4)
            return SimpleIconImage(width, height, data, mode='BGRA')
        finally:
            if old_obj:
                gdi32.SelectObject(hdc, old_obj)
            gdi32.DeleteObject(hbitmap)
            gdi32.DeleteDC(hdc)
            user32.DestroyIcon(hicon)

    @staticmethod
    # アイコン画像を作成または読み込む
    def create_icon_image():
        candidate_paths = []

        # PyInstaller 互換の _MEIPASS があれば最優先で使用
        if hasattr(sys, '_MEIPASS'):
            candidate_paths.extend(
                os.path.join(sys._MEIPASS, icon_file) for icon_file in ICON_FILES
            )

        # Nuitka Onefile で解凍されたデータパスを優先的に参照
        for onefile_dir in _get_nuitka_onefile_dirs():
            candidate_paths.extend(
                os.path.join(onefile_dir, icon_file) for icon_file in ICON_FILES
            )

        # 実行ファイルと同じ場所に展開されたデータファイルを優先的に参照
        candidate_paths.extend(
            os.path.join(APP_DIR, icon_file) for icon_file in ICON_FILES
        )

        # カレントディレクトリにもあれば参照（デバッグ実行などを考慮）
        candidate_paths.extend(
            os.path.abspath(icon_file) for icon_file in ICON_FILES
        )

        for image_path in candidate_paths:
            if os.path.exists(image_path):
                win32_image = TrayIconManager._load_with_win32_icon(
                    image_path,
                    DEFAULT_ICON_SIZE,
                )
                if win32_image is not None:
                    return win32_image
                pillow_image = TrayIconManager._load_with_pillow(image_path)
                if pillow_image is not None:
                    return pillow_image

        return TrayIconManager.create_default_icon()


# pystray が SimpleIconImage を扱えるようにパッチ適用
def ensure_pystray_image_support():
    try:
        from pystray import _base
    except Exception:
        return

    try:
        from pystray import _win32  # type: ignore
    except Exception:
        _win32 = None

    def patch_icon_class(icon_cls):
        if not hasattr(icon_cls, '_ser2key_image_patch') and hasattr(icon_cls, '_assert_image'):
            original = icon_cls._assert_image

            def _patched(self, image):
                if isinstance(image, SimpleIconImage):
                    return image
                return original(self, image)

            icon_cls._assert_image = _patched
            icon_cls._ser2key_image_patch = True

    patch_icon_class(_base.Icon)
    if _win32 is not None:
        patch_icon_class(_win32.Icon)


# エラーメッセージをポップアップで表示
def show_error_message(message):
    user32.MessageBoxW(None, str(message), "エラー", 0x00000010)


# 管理者権限時は通常権限プロセスでトレイアイコンを表示
def launch_tray_proxy():
    try:
        # 標準ユーザーで同じEXEを起動（UAC昇格せず）
        exe_path = os.path.abspath(sys.argv[0])
        subprocess.Popen(
            ['cmd', '/c', 'start', '', exe_path, '--tray-proxy'],
            shell=False,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
    except Exception as e:
        logging.getLogger('ser2key').warning(f"トレイプロキシ起動失敗: {e}")


# メイン処理
def main():
    logger = logging.getLogger('ser2key')

    # --- 多重起動防止処理 ---
    mutex_name = "Global\\ser2key_mutex"
    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, False, mutex_name)
    last_error = kernel32.GetLastError()
    ERROR_ALREADY_EXISTS = 183

    if last_error == ERROR_ALREADY_EXISTS:
        show_error_message("すでに ser2key が起動中です。")
        sys.exit(0)
    # ---------------------------------

    try:
        emulator = SerialKeyboardEmulator()
        emulator.read_config()

        # --- COM初期化済みトレイスレッド関数 ---
        def tray_thread():
            try:
                ole32.OleInitialize(None)  # COM初期化（管理者権限でも安定）
                ensure_pystray_image_support()
                icon = pystray.Icon("ser2key")
                icon.icon = TrayIconManager.create_icon_image()
                icon.title = "ser2key - 初期化中"

                emulator.refresh_available_ports(update_menu=False)
                emulator.attach_tray_icon(icon)

                icon.run()
            except Exception as e:
                logger.error(f"トレイアイコンスレッドエラー: {e}")
            finally:
                try:
                    ole32.OleUninitialize()
                except Exception:
                    pass

        # --- トレイスレッド起動 ---
        tray_thread_obj = threading.Thread(target=tray_thread, daemon=True)
        tray_thread_obj.start()

        # モニタリングスレッド
        monitor = ApplicationMonitor(emulator)
        threading.Thread(target=monitor.monitor, daemon=True).start()

        # シリアル通信スレッド
        threading.Thread(target=emulator.read_serial_and_type, daemon=True).start()

        # メインループ（待機）
        while emulator.is_running:
            time.sleep(1)

    except Exception as e:
        logger.error(f"アプリケーションエラー: {e}")
        show_error_message(str(e))
        sys.exit(1)
    finally:
        if 'emulator' in locals():
            emulator.cleanup()

        # --- ミューテックス解放 ---
        if mutex:
            kernel32.ReleaseMutex(mutex)
            kernel32.CloseHandle(mutex)

if __name__ == "__main__":

    main()


