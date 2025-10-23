#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
Serial to keyboard:
Author: Akihiko Fujita
Version: 1.3

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

import serial␊
from serial.tools import list_ports␊
import pyautogui␊
import pyperclip␊
import threading␊
import time␊
import configparser␊
import tkinter as tk␊
from tkinter import messagebox␊
from PIL import Image, ImageDraw␊
import pystray␊
from pystray import MenuItem as item␊
import sys␊
import os␊
import logging␊
from logging.handlers import RotatingFileHandler␊
import psutil␊


# 定数定義
DEFAULT_ICON_SIZE = (32, 32) 
DEFAULT_ICON_COLOR = 'white'
CONFIG_FILE = 'config.ini'
ICON_FILE = 'f.png'
LOG_FILE = 'ser2key.log'
RECONNECT_DELAY = 5     # 再接続までの待機時間（秒）
ACTIVITY_TIMEOUT = 300  # アクティビティタイムアウト（秒）
MONITOR_INTERVAL = 60   # モニタリング間隔（秒）
MAX_ERRORS = 10         # 最大エラー回数
CLIPBOARD_TIMEOUT = 5   # クリップボード操作タイムアウト（秒）
RESOURCE_WARNING_THRESHOLD = {
    'memory': 90,       # メモリ使用率警告閾値（%）
    'cpu': 80           # CPU使用率警告閾値（%）
}

def setup_logging():
    """ログ設定を初期化"""
    logger = logging.getLogger('ser2key')
    logger.setLevel(logging.INFO)
    
    handler = RotatingFileHandler(
        LOG_FILE,
        maxBytes=512*1024,
        backupCount=3
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
            'encoding': 'sjis'
        }

        with open(CONFIG_FILE, 'w') as f:
            config.write(f)
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

        menu_definition = [
            item('ポート一覧を更新', lambda icon, menu_item: self.refresh_available_ports()),
            pystray.Menu.SEPARATOR
        ]

        if self.available_ports:
            for port in self.available_ports:
                menu_definition.append(
                    item(
                        f'接続: {port}',
                        lambda icon, menu_item, port=port: self.update_serial_port(port),
                        checked=lambda menu_item, port=port: self._is_selected_port(port)
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
                with open(CONFIG_FILE, 'w') as f:
                    self._config_parser.write(f)

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
        config = configparser.ConfigParser()

        if not os.path.exists(CONFIG_FILE):
            self.logger.warning("設定ファイルが見つかりません。デフォルト設定を作成します。")
            self.create_default_config()

        config.read(CONFIG_FILE)
        self._config_parser = config

        try:
            self.serial_config = {
                'port': config['serial']['port'],
                'baudrate': int(config['serial']['baudrate']),
                'bytesize': int(config['serial']['bytesize']),
                'parity': config['serial']['parity'],
                'stopbits': int(config['serial']['stopbits']),
                'timeout': int(config['serial']['timeout'])
            }
            self.settings_config = {
                'add_enter': config.getboolean('settings', 'add_enter', fallback=True),
                'encoding': config.get('settings', 'encoding', fallback='sjis')
            }
            
            self.validate_serial_config(self.serial_config)
            self.validate_settings_config(self.settings_config)
            
        except KeyError as e:
            raise KeyError(f"設定キーが見つかりません: {e}")
        except ValueError as e:
            raise ValueError(f"不正な設定値です: {e}")

    def check_system_resources(self):
        """システムリソースの状態確認"""
        try:
            # メモリ使用率の確認
            memory_percent = psutil.virtual_memory().percent
            if memory_percent > RESOURCE_WARNING_THRESHOLD['memory']:
                self.logger.warning(f"メモリ使用率が高くなっています: {memory_percent}%")
            
            # CPU使用率の確認
            cpu_percent = psutil.cpu_percent(interval=1)
            if cpu_percent > RESOURCE_WARNING_THRESHOLD['cpu']:
                self.logger.warning(f"CPU使用率が高くなっています: {cpu_percent}%")
            
        except Exception as e:
            self.logger.error(f"リソース監視エラー: {e}")

    def process_serial_data(self, data):
        """シリアルデータを処理してキーボード入力をシミュレート"""
        if not data:
            return

        self.logger.info(f"受信: {data}")
        self.last_activity = time.time()

        with self._lock:  # クリップボード操作をスレッドセーフに
            original_clipboard = None
            try:
                original_clipboard = pyperclip.paste()
                pyperclip.copy(data)

                # タイムアウト付きの待機を実装
                timeout = time.time() + CLIPBOARD_TIMEOUT
                while not pyperclip.paste() == data and time.time() < timeout:
                    time.sleep(0.1)
                
                pyautogui.hotkey('ctrl', 'v')
                
                if self.settings_config['add_enter']:
                    pyautogui.press('enter')
            except Exception as e:
                self.logger.error(f"クリップボード操作エラー: {e}")
                self.error_count += 1
            finally:
                if original_clipboard is not None:
                    try:
                        pyperclip.copy(original_clipboard)
                    except Exception:
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
                    self._reconnect_event.clear()

                    while self.is_running and not self._reconnect_event.is_set():
                        if ser.in_waiting > 0:
                            try:
                                data = ser.readline().decode(encoding).strip()
                                self.process_serial_data(data)
                            except UnicodeDecodeError as e:
                                self.logger.error(f"デコードエラー: {e}")
                                self.error_count += 1
                            except Exception as e:
                                self.logger.error(f"予期せぬエラー: {e}")
                                self.error_count += 1

                        # システムリソースの定期チェック
                        if time.time() - self.last_activity > MONITOR_INTERVAL:
                            self.check_system_resources()

                        if self._reconnect_event.is_set():
                            self.logger.info("再接続要求を受け取りました")
                            break

                        time.sleep(0.05)

            except serial.SerialException as e:
                if self.tray_icon:
                    self.tray_icon.title = f"ser2key - 接続失敗 ({e})"
                self.logger.error(f"シリアル通信エラー: {e}")
                self.error_count += 1
                time.sleep(RECONNECT_DELAY)
            except Exception as e:
                if self.tray_icon:
                    self.tray_icon.title = f"ser2key - エラー ({e})"
                self.logger.error(f"予期せぬエラー: {e}")
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
                if self.emulator.error_count >= MAX_ERRORS:
                    self.logger.error("エラー回数が上限を超えました。アプリケーションを再起動します")
                    self.restart_application()
                
                time.sleep(MONITOR_INTERVAL)
                
            except Exception as e:
                self.logger.error(f"モニタリングエラー: {e}")
    
    def restart_application(self):
        """アプリケーションの再起動処理"""
        self.emulator.cleanup()
        os.execv(sys.executable, ['python'] + sys.argv)

class TrayIconManager:
    """タスクトレイアイコンを管理するクラス"""
    
    @staticmethod
    def create_default_icon():
        """デフォルトアイコンを作成"""
        image = Image.new('RGB', DEFAULT_ICON_SIZE, color=DEFAULT_ICON_COLOR)
        draw = ImageDraw.Draw(image)
        draw.rectangle([0, 0, DEFAULT_ICON_SIZE[0]-1, DEFAULT_ICON_SIZE[1]-1], outline='black')
        draw.text((20, 25), 'S2K', fill='black')
        return image

    @staticmethod
    def create_icon_image():
        """アイコン画像を作成または読み込む"""
        try:
            image_path = os.path.join(sys._MEIPASS, ICON_FILE) if hasattr(sys, '_MEIPASS') else ICON_FILE
            return Image.open(image_path) if os.path.exists(image_path) else TrayIconManager.create_default_icon()
        except Exception as e:
            logging.getLogger('ser2key').error(f"アイコン読み込みエラー: {e}")
            return TrayIconManager.create_default_icon()

def show_error_message(message):
    """エラーメッセージをポップアップで表示"""
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("エラー", message)
    root.destroy()

def main():
    """メイン処理"""
    logger = logging.getLogger('ser2key')
    try:
        emulator = SerialKeyboardEmulator()
        emulator.read_config()

        # タスクトレイアイコンの設定
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
