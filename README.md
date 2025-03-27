# ser2key - Serial to Keyboard  

## 🚀 Overview  
**ser2key** is a lightweight Python utility that converts serial input into keyboard events. It listens for incoming serial data and simulates keystrokes. Designed for efficiency, it runs in the **system tray** for seamless background operation.  

## ✨ Features  
- 🔧 **Configurable Serial Communication** – Adjust settings via `config.ini` (baud rate, parity, timeout, etc.).  
- ⌨️ **Clipboard-based Keystroke Simulation** – Pastes received data as keyboard input.  
- 🚀 **Designed for High-Speed Data Input"** – Outperforms existing tools, especially with **Japanese, Chinese, and Korean characters**.  
- 🖥️ **System Tray Integration** – Quick access and status monitoring.  


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
1.2 Revised entire code, added validation checks for anomalous values in configuration files, etc.<BR>
1.1 Enhanced Error Handling.<BR>
1.0 Release Version.<BR>

---

## 📥 Download  
Includes Windows x64 executable and configuration files.
📌 [ser2key.zip](https://github.com/Akihiko-Fuji/ser2key/raw/refs/heads/main/ser2key.zip)  

