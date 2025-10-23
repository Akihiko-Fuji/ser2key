#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
Serial to keyboard:
Author: Akihiko Fujita
Version: 1.4

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


if os.name != 'nt':
    raise OSError('ser2key は Windows 専用のアプリケーションです。')


# 定数定義
DEFAULT_ICON_SIZE = (32, 32)
ICON_FILE = 'f.png'
APP_NAME = 'ser2key'


def _get_storage_directory():
    base_dir = os.getenv('APPDATA')
    if not base_dir:
        base_dir = os.path.expanduser('~')
    return os.path.join(base_dir, APP_NAME)


STORAGE_DIR = _get_storage_directory()
CONFIG_FILE = os.path.join(STORAGE_DIR, 'config.ini')
LOG_FILE = os.path.join(STORAGE_DIR, 'ser2key.log')


def ensure_storage_directory():
    try:
        os.makedirs(STORAGE_DIR, exist_ok=True)
    except OSError as exc:
        raise RuntimeError(
            f'設定保存用ディレクトリ {STORAGE_DIR} を作成できません: {exc}'
        ) from exc


RECONNECT_DELAY = 5     # 再接続までの待機時間（秒）
ACTIVITY_TIMEOUT = 300  # アクティビティタイムアウト（秒）
MONITOR_INTERVAL = 60   # モニタリング間隔（秒）
MAX_ERRORS = 10         # 最大エラー回数
CLIPBOARD_TIMEOUT = 5   # クリップボード操作タイムアウト（秒）
CLIPBRD_E_CANT_OPEN = 0x800401D0

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002

VK_CONTROL = 0x11
VK_V = 0x56
VK_RETURN = 0x0D
KEYEVENTF_KEYUP = 0x0002

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
ole32 = ctypes.OleDLL('ole32')

ole32.OleInitialize.restype = wintypes.HRESULT
ole32.OleInitialize.argtypes = [ctypes.c_void_p]
ole32.OleGetClipboard.restype = wintypes.HRESULT
ole32.OleGetClipboard.argtypes = [ctypes.POINTER(ctypes.c_void_p)]
ole32.OleSetClipboard.restype = wintypes.HRESULT
ole32.OleSetClipboard.argtypes = [ctypes.c_void_p]
ole32.OleFlushClipboard.restype = wintypes.HRESULT
ole32.OleFlushClipboard.argtypes = []

S_OK = 0
S_FALSE = 1


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

    def release(self):
        """COM オブジェクトの参照カウントを解放"""
        if self.data_object:
            obj_ptr = ctypes.cast(ctypes.c_void_p(self.data_object), ctypes.POINTER(_IUnknown))
            release = obj_ptr.contents.lpVtbl.contents.Release
            release(obj_ptr)
            self.data_object = None


_ole_thread_state = threading.local()


def ensure_ole_initialized():
    """現在のスレッドで OLE を初期化"""
    if getattr(_ole_thread_state, 'initialized', False):
        return

    hr = ole32.OleInitialize(None)
    if hr not in (S_OK, S_FALSE):
        raise ClipboardError(f"OLE の初期化に失敗しました (HRESULT=0x{hr:08X})")

    _ole_thread_state.initialized = True


def is_clipboard_empty():
    """クリップボードが空であるか確認"""
    try:
        with open_clipboard():
            kernel32.SetLastError(0)
            first_format = user32.EnumClipboardFormats(0)
            if first_format == 0 and kernel32.GetLastError() == 0:
                return True
    except ClipboardError:
        pass
    return False


def backup_clipboard():
    """クリップボードの全データを退避"""
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
        if last_error == CLIPBRD_E_CANT_OPEN and time.time() < deadline:
            time.sleep(0.05)
            continue

        raise ClipboardError(
            f"クリップボードの退避に失敗しました (HRESULT=0x{last_error:08X})"
        )


def restore_clipboard(backup):
    """退避したクリップボード内容を復元"""
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
def open_clipboard():
    """クリップボードを安全に開くためのコンテキストマネージャ"""
    if not user32.OpenClipboard(None):
        raise ClipboardError('クリップボードを開けませんでした')
    try:
        yield
    finally:
        user32.CloseClipboard()


def get_clipboard_text():
    """現在のクリップボード文字列を取得"""
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


def set_clipboard_text(text):
    """クリップボードに文字列を設定"""
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


def clear_clipboard():
    """クリップボードを空にする"""
    with open_clipboard():
        user32.EmptyClipboard()


def _key_event(vk_code, key_up=False):
    """キーイベントを送信"""
    scan_code = user32.MapVirtualKeyW(vk_code, 0)
    flags = KEYEVENTF_KEYUP if key_up else 0
    user32.keybd_event(vk_code, scan_code, flags, 0)


def send_ctrl_v(add_enter=False):
    """Ctrl+V と必要に応じて Enter キーを送信"""
    _key_event(VK_CONTROL)
    _key_event(VK_V)
    _key_event(VK_V, key_up=True)
    _key_event(VK_CONTROL, key_up=True)

    if add_enter:
        time.sleep(0.05)
        _key_event(VK_RETURN)
        _key_event(VK_RETURN, key_up=True)


def setup_logging():
    """ログ設定を初期化"""
    try:
        ensure_storage_directory()
    except RuntimeError as exc:
        raise RuntimeError(str(exc)) from exc

    logger = logging.getLogger('ser2key')
    logger.setLevel(logging.INFO)

    log_path = os.path.abspath(LOG_FILE)
    handler_exists = any(
        isinstance(handler, RotatingFileHandler) and
        getattr(handler, 'baseFilename', None) == log_path
        for handler in logger.handlers
    )

    if not handler_exists:
        handler = RotatingFileHandler(
            LOG_FILE,
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

    def cleanup(self):
        """リソースの解放処理"""
        self.is_running = False
        self.logger.info("アプリケーションのクリーンアップを実行")
        self._reconnect_event.set()

    def create_default_config(self):
        """デフォルト設定ファイルの作成"""
        try:
            ensure_storage_directory()
        except RuntimeError as exc:
            self.logger.error(str(exc))
            raise

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
            'encoding': 'sjis',
            'buffer_msec': '0'
        }

        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                config.write(f)
        except OSError as exc:
            message = f"設定ファイル {CONFIG_FILE} を作成できません: {exc}"
            self.logger.error(message)
            raise RuntimeError(message) from exc
        self.logger.info("デフォルト設定ファイルを作成しました")
        self._config_parser = config

    def validate_serial_config(self, config):
        """シリアル通信の設定値を検証"""
        valid_bauds = [300, 1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200]
        if config['baudrate'] not in valid_bauds:
            raise ValueError(f"不正なボーレート: {config['baudrate']}")

        self.refresh_available_ports(update_menu=False)
        self.logger.info(f"利用可能なポート: {self.available_ports}")

        if config['port'] not in self.available_ports:
            self.logger.warning(f"設定されたポート {config['port']} は現在利用可能なポート一覧にありません")
            # ここでもエラーは発生させず、警告のみ

    def refresh_available_ports(self, update_menu=True):
        """利用可能なシリアルポート一覧を更新"""
        ports = [port.device for port in list_ports.comports()]
        ports.sort()
        self.available_ports = ports
        self.logger.info(f"ポート一覧を更新: {self.available_ports}")

        if update_menu and self.tray_icon:
            self.update_tray_menu()

    def attach_tray_icon(self, icon):
        """タスクトレイアイコンを関連付け"""
        self.tray_icon = icon
        self.update_tray_menu()

    def _is_selected_port(self, port):
        """現在選択されているポートかどうかを判定"""
        with self._lock:
            current = self.serial_config.get('port') if self.serial_config else None
        return current == port

    def update_tray_menu(self):
        """タスクトレイメニューを更新"""
        if not self.tray_icon:
            return

        def refresh_ports_action(icon, menu_item):
            self.refresh_available_ports()

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
            item('終了', self.stop_application)
        ])

        try:
            self.tray_icon.menu = pystray.Menu(*menu_definition)
        except Exception as e:
            self.logger.error(f"トレイメニュー更新エラー: {e}")

    def update_serial_port(self, port):
        """利用するシリアルポートを更新し再接続を要求"""
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
            if self._config_parser and 'serial' in self._config_parser:
                self._config_parser['serial']['port'] = port
                try:
                    ensure_storage_directory()
                except RuntimeError as exc:
                    self.logger.error(str(exc))
                else:
                    try:
                        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                            self._config_parser.write(f)
                    except OSError as exc:
                        self.logger.error(
                            f"設定ファイル {CONFIG_FILE} への書き込みに失敗しました: {exc}"
                        )

        self.logger.info(f"シリアルポートを {port} に更新しました。再接続を実行します")
        self._reconnect_event.set()
        self.update_tray_menu()

    def stop_application(self, icon, menu_item=None):
        """アプリケーション終了処理"""
        self.logger.info("ユーザー操作によりアプリケーションを終了します")
        self.cleanup()
        if icon:
            try:
                icon.visible = False
            except Exception:
                pass
            icon.stop()

    def validate_settings_config(self, config):
        """一般設定の検証"""
        valid_encodings = ['utf-8', 'sjis', 'ascii']
        if config['encoding'].lower() not in valid_encodings:
            raise ValueError(f"サポートされていないエンコーディング: {config['encoding']}")

    def read_config(self):
        """設定ファイルを読み込む"""
        try:
            ensure_storage_directory()
        except RuntimeError as exc:
            self.logger.error(str(exc))
            raise

        config = configparser.ConfigParser()

        if not os.path.exists(CONFIG_FILE):
            self.logger.warning("設定ファイルが見つかりません。デフォルト設定を作成します。")
            self.create_default_config()

        config.read(CONFIG_FILE, encoding='utf-8')
        self._config_parser = config

        try:
            self.serial_config = {
                'port': config['serial']['port'],
                'baudrate': int(config['serial']['baudrate']),
                'bytesize': int(config['serial']['bytesize']),
                'parity': config['serial']['parity'],
                'stopbits': int(config['serial']['stopbits']),
                'timeout': float(config['serial']['timeout'])
            }
            self.settings_config = {
                'add_enter': config.getboolean('settings', 'add_enter', fallback=True),
                'encoding': config.get('settings', 'encoding', fallback='sjis'),
                'buffer_msec': config.getint('settings', 'buffer_msec', fallback=0)
            }
            
            self.validate_serial_config(self.serial_config)
            self.validate_settings_config(self.settings_config)
            
        except KeyError as e:
            raise KeyError(f"設定キーが見つかりません: {e}")
        except ValueError as e:
            raise ValueError(f"不正な設定値です: {e}")

    def process_serial_data(self, data):
        """シリアルデータを処理してキーボード入力をシミュレート"""
        if not data:
            return

        self.logger.info(f"受信: {data}")
        self.last_activity = time.time()

        with self._lock:
            clipboard_backup = None
            clipboard_ready = False
            try:
                clipboard_backup = backup_clipboard()
                set_clipboard_text(data)

                timeout = time.time() + CLIPBOARD_TIMEOUT
                while time.time() < timeout:
                    try:
                        if get_clipboard_text() == data:
                            clipboard_ready = True
                            break
                    except ClipboardError:
                        time.sleep(0.05)
                        continue
                    time.sleep(0.05)

                if not clipboard_ready:
                    raise ClipboardError('クリップボードの内容が更新されませんでした')

                send_ctrl_v(self.settings_config.get('add_enter', False))
            except ClipboardError as e:
                self.logger.error(f"クリップボード操作エラー: {e}")
                self.error_count += 1
            except Exception as e:
                self.logger.error(f"入力操作エラー: {e}")
                self.error_count += 1
            else:
                self.error_count = 0
            finally:
                try:
                    restore_clipboard(clipboard_backup)
                except ClipboardError:
                    self.logger.warning("クリップボードの復元に失敗しました")

    def read_serial_and_type(self):
        """シリアルポートからデータを読み取り、キーボード入力に変換"""
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
                time.sleep(RECONNECT_DELAY)
            except Exception as e:
                if self.tray_icon:
                    self.tray_icon.title = f"ser2key - エラー ({e})"
                self.logger.error(f"予期せぬエラー: {e}")
                with self._lock:
                    self.error_count += 1
                time.sleep(RECONNECT_DELAY)
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

    def monitor(self):
        """定期的な状態確認"""
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


class TrayIconManager:
    """タスクトレイアイコンを管理するクラス"""

    @staticmethod
    def create_default_icon():
        """デフォルトアイコンを作成"""
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
    def create_icon_image():
        """アイコン画像を作成または読み込む"""
        image_path = os.path.join(sys._MEIPASS, ICON_FILE) if hasattr(sys, '_MEIPASS') else ICON_FILE

        if os.path.exists(image_path):
            pillow_image = TrayIconManager._load_with_pillow(image_path)
            if pillow_image is not None:
                return pillow_image

        return TrayIconManager.create_default_icon()


def ensure_pystray_image_support():
    """pystray が SimpleIconImage を扱えるようにパッチ適用"""
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

def show_error_message(message):
    """エラーメッセージをポップアップで表示"""
    user32.MessageBoxW(None, str(message), "エラー", 0x00000010)

def main():
    """メイン処理"""
    logger = logging.getLogger('ser2key')
    try:
        emulator = SerialKeyboardEmulator()
        emulator.read_config()

        # タスクトレイアイコンの設定
        ensure_pystray_image_support()
        icon = pystray.Icon("ser2key")
        icon.icon = TrayIconManager.create_icon_image()
        icon.title = "ser2key - 初期化中"

        emulator.refresh_available_ports(update_menu=False)
        emulator.attach_tray_icon(icon)

        # モニタリングスレッドの開始
        monitor = ApplicationMonitor(emulator)
        monitor_thread = threading.Thread(target=monitor.monitor, daemon=True)
        monitor_thread.start()

        # シリアル通信スレッドの開始
        serial_thread = threading.Thread(target=emulator.read_serial_and_type, daemon=True)
        serial_thread.start()

        icon.run()

    except Exception as e:
        logger.error(f"アプリケーションエラー: {e}")
        show_error_message(str(e))
        sys.exit(1)
    finally:
        if 'emulator' in locals():
            emulator.cleanup()

if __name__ == "__main__":

    main()







