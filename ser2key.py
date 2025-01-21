#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
Serial to keyboard:
Author: Akihiko Fujita
Version: 1.0

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

import serial
import pyautogui
import pyperclip
import threading
import time
import configparser
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import pystray
from pystray import MenuItem as item
import sys
import os

# エラーメッセージをポップアップで表示
def show_error_message(message):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", message)
    root.destroy()

# config.ini を読み込む関数
def read_config():
    config = configparser.ConfigParser()
    config_path = 'config.ini'
    
    if not os.path.exists(config_path):
        show_error_message(f"{config_path} not found. Please make sure the file exists.")
        sys.exit(1)
    
    config.read(config_path)
    try:
        serial_config = {
            'port': config['serial']['port'],
            'baudrate': int(config['serial']['baudrate']),
            'bytesize': int(config['serial']['bytesize']),
            'parity': config['serial']['parity'],
            'stopbits': int(config['serial']['stopbits']),
            'timeout': int(config['serial']['timeout'])
        }
        settings_config = {
            'add_enter': config.getboolean('settings', 'add_enter', fallback=True),
            'encoding': config.get('settings', 'encoding', fallback='sjis')
        }
    except KeyError as e:
        show_error_message(f"Missing configuration key: {e}")
        sys.exit(1)
    except ValueError as e:
        show_error_message(f"Invalid configuration value: {e}")
        sys.exit(1)
    
    return serial_config, settings_config

# シリアルデータの読み取りとキーボード入力のシミュレーション
def read_serial_and_type(serial_config, settings_config, tray_icon):
    while True:
        try:
            with serial.Serial(**serial_config) as ser:
                tray_icon.title = "ser2key - Connected"
                while True:
                    if ser.in_waiting > 0:
                        try:
                            data = ser.readline().decode(settings_config['encoding']).strip()  # 設定されたエンコードでデコード
                            if data:
                                print(f"Received: {data}")

                                # 現在のクリップボードの内容を退避
                                original_clipboard = pyperclip.paste()

                                # 新しいデータをクリップボードにコピー
                                pyperclip.copy(data)

                                # クリップボードのデータを貼り付け
                                pyautogui.hotkey('ctrl', 'v')

                                # 必要に応じて Enter キーを押す
                                if settings_config['add_enter']:
                                    pyautogui.press('enter')

                                # クリップボードの内容を元に戻す
                                pyperclip.copy(original_clipboard)
                        except UnicodeDecodeError as e:
                            print(f"デコードエラー: {e}")
                        except Exception as e:
                            print(f"予期せぬエラー: {e}")
        except serial.SerialException as e:
            tray_icon.title = f"ser2kley - {e}"
            time.sleep(5)  # 5秒待ってから再接続を試みる

# スレッドを使用してシリアル読み取りを実行
def main(tray_icon):
    serial_config, settings_config = read_config()
    serial_thread = threading.Thread(target=read_serial_and_type, args=(serial_config, settings_config, tray_icon))
    serial_thread.daemon = True
    serial_thread.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("終了します")

# タスクトレイアイコンの設定
def create_image():
    # 実行ファイルに埋め込まれた f.png を適切に読み込む
    if hasattr(sys, '_MEIPASS'):
        image_path = os.path.join(sys._MEIPASS, 'f.png')
    else:
        image_path = 'f.png'

    image = Image.open(image_path)
    return image

def on_exit(icon, item):
    icon.stop()
    sys.exit()

def setup_tray_icon():
    icon = pystray.Icon("ser2key")
    icon.icon = create_image()
    icon.menu = pystray.Menu(item('Exit', on_exit))
    threading.Thread(target=main, args=(icon,), daemon=True).start()
    icon.run()

if __name__ == "__main__":
    setup_tray_icon()
