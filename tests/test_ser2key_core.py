import math
import unittest
from datetime import datetime

from ser2key_core import (
    create_config_parser,
    decode_output_template,
    render_output_template,
    validate_serial_config,
    validate_settings_config,
)


class OutputTemplateTests(unittest.TestCase):
    def test_decodes_escapes_without_corrupting_japanese(self):
        self.assertEqual(
            decode_output_template(r'受付\tデータ\r\n'),
            '受付\tデータ\r\n',
        )

    def test_preserves_unknown_escape_sequences(self):
        self.assertEqual(decode_output_template(r'C:\queue\item'), r'C:\queue\item')

    def test_preserves_incomplete_escape_sequences(self):
        self.assertEqual(decode_output_template(r'末尾\u12'), r'末尾\u12')

    def test_config_parser_preserves_strftime_percent_signs(self):
        parser = create_config_parser()
        parser.read_string('[output]\nheader={DATE:%Y%m%d}\n')
        self.assertEqual(parser.get('output', 'header'), '{DATE:%Y%m%d}')

    def test_renders_default_and_custom_datetime_tokens(self):
        now = datetime(2026, 6, 8, 14, 5, 9)
        self.assertEqual(
            render_output_template(
                '{DATE} {TIME} {DATETIME:%Y%m%d-%H%M%S}', now
            ),
            '2026-06-08 14:05:09 20260608-140509',
        )

    def test_leaves_unknown_tokens_unchanged(self):
        now = datetime(2026, 6, 8, 14, 5, 9)
        self.assertEqual(render_output_template('{USER} {DATE}', now), '{USER} 2026-06-08')



class ConfigurationValidationTests(unittest.TestCase):
    def setUp(self):
        self.serial_config = {
            'port': 'COM7',
            'baudrate': 9600,
            'bytesize': 8,
            'parity': 'N',
            'stopbits': 1.0,
            'timeout': 1.0,
        }
        self.settings_config = {
            'add_enter': True,
            'encoding': 'shift_jis',
            'buffer_msec': 0,
        }

    def test_accepts_default_configuration(self):
        validate_serial_config(self.serial_config)
        validate_settings_config(self.settings_config)
        self.assertEqual(self.settings_config['encoding'], 'shift_jis')

    def test_rejects_invalid_serial_values(self):
        invalid_values = {
            'port': '',
            'baudrate': 0,
            'bytesize': 9,
            'parity': 'X',
            'stopbits': 3,
            'timeout': -1,
        }
        for key, value in invalid_values.items():
            with self.subTest(key=key):
                config = dict(self.serial_config)
                config[key] = value
                with self.assertRaises(ValueError):
                    validate_serial_config(config)

    def test_rejects_boolean_and_non_finite_serial_values(self):
        for key, value in (
            ('bytesize', 8.0),
            ('stopbits', True),
            ('timeout', math.inf),
            ('timeout', math.nan),
        ):
            with self.subTest(key=key, value=value):
                config = dict(self.serial_config)
                config[key] = value
                with self.assertRaises(ValueError):
                    validate_serial_config(config)

    def test_rejects_invalid_application_values(self):
        for key, value in (
            ('encoding', 'not-a-codec'),
            ('buffer_msec', -1),
            ('buffer_msec', 60_001),
            ('add_enter', 'yes'),
        ):
            with self.subTest(key=key, value=value):
                config = dict(self.settings_config)
                config[key] = value
                with self.assertRaises(ValueError):
                    validate_settings_config(config)


if __name__ == '__main__':
    unittest.main()
