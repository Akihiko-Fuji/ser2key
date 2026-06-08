import logging
import os
import threading
import unittest
from unittest.mock import Mock, patch

from ser2key_core import create_config_parser

if os.name == 'nt':
    import ser2key


@unittest.skipUnless(os.name == 'nt', 'Windows 専用アプリケーションのテスト')
class ConfigurationUpdateTests(unittest.TestCase):
    def setUp(self):
        self.emulator = ser2key.SerialKeyboardEmulator.__new__(
            ser2key.SerialKeyboardEmulator
        )
        self.emulator.logger = Mock()
        self.emulator._lock = threading.Lock()
        self.emulator._reconnect_event = Mock()
        self.emulator.update_tray_menu = Mock()
        self.emulator.available_ports = ['COM7', 'COM8']
        self.emulator.serial_config = {
            'port': 'COM7',
            'baudrate': 9600,
            'bytesize': 8,
            'parity': 'N',
            'stopbits': 1.0,
            'timeout': 1.0,
        }
        self.emulator.settings_config = {
            'add_enter': True,
            'encoding': 'shift_jis',
            'buffer_msec': 0,
        }
        self.emulator._config_parser = create_config_parser()
        self.emulator._config_parser.read_dict({
            'serial': self.emulator.serial_config,
            'settings': self.emulator.settings_config,
        })
        self.emulator.config_path = 'config.ini'
        self.emulator._persist_config = Mock(side_effect=OSError('write failed'))

    def test_serial_port_rollback_refreshes_menu_without_reconnect(self):
        self.emulator.update_serial_port('COM8')

        self.assertEqual(self.emulator.serial_config['port'], 'COM7')
        self.assertEqual(self.emulator._config_parser['serial']['port'], 'COM7')
        self.emulator.update_tray_menu.assert_called_once_with()
        self.emulator._reconnect_event.set.assert_not_called()

    def test_serial_setting_rollback_refreshes_menu_without_reconnect(self):
        self.emulator.update_serial_setting('baudrate', 19200)

        self.assertEqual(self.emulator.serial_config['baudrate'], 9600)
        self.assertEqual(self.emulator._config_parser['serial']['baudrate'], '9600')
        self.emulator.update_tray_menu.assert_called_once_with()
        self.emulator._reconnect_event.set.assert_not_called()

    def test_settings_rollback_refreshes_menu(self):
        self.emulator.update_settings_setting('buffer_msec', 250)

        self.assertEqual(self.emulator.settings_config['buffer_msec'], 0)
        self.assertEqual(self.emulator._config_parser['settings']['buffer_msec'], '0')
        self.emulator.update_tray_menu.assert_called_once_with()

    def test_every_settings_key_is_validated_before_update(self):
        invalid_values = {
            'encoding': 'not-a-codec',
            'buffer_msec': -1,
            'add_enter': 'yes',
        }

        for key, value in invalid_values.items():
            with self.subTest(key=key):
                self.emulator.update_settings_setting(key, value)

        self.emulator._persist_config.assert_not_called()
        self.emulator.update_tray_menu.assert_not_called()
        self.assertEqual(self.emulator.logger.error.call_count, len(invalid_values))


@unittest.skipUnless(os.name == 'nt', 'Windows 専用アプリケーションのテスト')
class StartupErrorTests(unittest.TestCase):
    def test_logging_setup_failure_uses_fallback_logger(self):
        fallback_logger = Mock(spec=logging.Logger)
        error = RuntimeError('log directory unavailable')

        with (
            patch.object(ser2key.logging, 'getLogger', return_value=fallback_logger),
            patch.object(ser2key, 'setup_logging', side_effect=error),
            patch.object(ser2key, 'show_error_message') as show_error_message,
            self.assertRaises(SystemExit) as exit_context,
        ):
            ser2key.main()

        self.assertEqual(exit_context.exception.code, 1)
        fallback_logger.error.assert_called_once_with(
            'アプリケーションエラー: log directory unavailable'
        )
        show_error_message.assert_called_once_with('log directory unavailable')


if __name__ == '__main__':
    unittest.main()
