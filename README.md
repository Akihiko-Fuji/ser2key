# ser2key - Serial to Keyboard  

## 🚀 Overview  
**ser2key** is a lightweight Python utility that converts serial input into keyboard events. It listens for incoming serial data and simulates keystrokes. Designed for efficiency, it runs in the **system tray** for seamless background operation.  

## ✨ Features
- 🔧 **Configurable Serial Communication** – Adjust settings via `config.ini` (baud rate, parity, timeout, etc.).
- ⌨️ **Clipboard-based Keystroke Simulation** – Pastes received data as keyboard input.
- 🚀 **Designed for High-Speed Data Input"** – Outperforms existing tools, especially with **Japanese, Chinese, and Korean characters**.
- 🖥️ **System Tray Integration** – Quick access and status monitoring.

---

## ⚙️ Configuration Tips

### Serial port settings
Edit the values under the `[serial]` and `[settings]` sections of `config.ini` to control the default port, baud rate, data length, parity, stop bits, encoding, and automatic Enter key behaviour. All of these can also be adjusted at runtime from the system tray menu, and any change is written back to the configuration file for the next launch.

### Formatting output with headers and footers
Use the `[output]` section of `config.ini` to wrap received serial data before it is pasted. The `header` value is inserted before the serial payload and `footer` after it, making it easy to add things like prefixes, suffixes, or line breaks without changing the device firmware.

Both fields support:

- **Escape sequences** such as `\n` (newline), `\r` (carriage return), `\t` (tab), and Unicode escapes like `\u3001`.
- **Date/Time tokens** wrapped in braces. Available tokens are `{DATE}`, `{TIME}`, and `{DATETIME}`. You can optionally specify a [Python `strftime` format](https://docs.python.org/3/library/datetime.html#strftime-and-strptime-format-codes) after a colon, e.g. `{DATE:%Y/%m/%d}` or `{TIME:%H時%M分%S秒}`. The default formats are:
  - `{DATE}` → `YYYY-MM-DD`
  - `{TIME}` → `HH:MM:SS`
  - `{DATETIME}` → `YYYY-MM-DD HH:MM:SS`

Example:

```ini
[output]
header={DATE:%Y%m%d}\t
footer=\r\n--\n
```

This configuration pastes the current date followed by a tab before the serial data, and appends a Windows-style line ending plus a separator line afterwards.

---

## 🔥 Performance Advantage  
Designed with **barcode & QR code readers** in mind, `ser2key` delivers superior speed when handling multilingual text.  
⚡ **Several times faster** than [Keyence AutoID Keyboard Wedge](https://www.keyence.co.jp/support/codereader/blsrus/soft/#d12) when processing **Japanese, Chinese, or Korean** characters from QR codes.  

---

## 💻 Supported Platforms  
✅ **Windows** (tested) – Runs in the system tray.  
⚠️ **Other OS** – Requires modification due to system tray dependencies, and In addition, the serial communication process validation is for Windows OS and needs to be corrected.<BR>
🔌 **Compatible with all serial devices** – Recognized as a COM port by Windows, including RS-232C, USB serial adapters, and Bluetooth SPP mode. 

---

## Update history
1.3 Fixed to allow selection of the COM port to connect to from the system tray icon.
    By limiting the execution environment to Windows and modifying the libraries used, we reduced the size of the executable file.
1.2 Revised entire code, added validation checks for anomalous values in configuration files, etc.<BR>
1.1 Enhanced Error Handling.<BR>
1.0 Release Version.<BR>

---

## 📥 Download  
Includes Windows x64 executable and configuration files.
📌 [ser2key.zip](https://github.com/Akihiko-Fuji/ser2key/raw/refs/heads/main/ser2key.zip)  

