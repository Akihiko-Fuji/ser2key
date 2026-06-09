import math
import unittest
from datetime import datetime

from ser2key_core import (
    ENCODING_OPTIONS,
    create_config_parser,
    decode_output_template,
    render_output_template,
    validate_serial_config,
    validate_settings_config,
)
from ser2key_i18n import (
    LANGUAGE_OPTIONS,
    language_from_locale,
    language_from_windows_lang_id,
    parity_label,
    translate,
)


class LocalizationTests(unittest.TestCase):
    def test_maps_supported_windows_locales(self):
        expected = {
            'ja-JP': 'ja',
            'en-US': 'en',
            'ko-KR': 'ko',
            'zh-CN': 'zh',
            'zh-TW': 'zh',
        }
        for locale_name, language in expected.items():
            with self.subTest(locale_name=locale_name):
                self.assertEqual(language_from_locale(locale_name), language)

    def test_unsupported_or_missing_locale_falls_back_to_english(self):
        self.assertEqual(language_from_locale('fr-FR'), 'en')
        self.assertEqual(language_from_locale(None), 'en')

    def test_maps_windows_ui_language_ids(self):
        expected = {
            0x0411: 'ja',
            0x0409: 'en',
            0x0412: 'ko',
            0x0804: 'zh',
            0x040C: 'en',
        }
        for lang_id, language in expected.items():
            with self.subTest(lang_id=lang_id):
                self.assertEqual(language_from_windows_lang_id(lang_id), language)

    def test_every_language_has_translated_parity_names(self):
        expected_japanese = {
            'N': 'なし',
            'E': '偶数',
            'O': '奇数',
            'M': 'マーク',
            'S': 'スペース',
        }
        for parity, label in expected_japanese.items():
            self.assertEqual(parity_label('ja', parity), label)
        for language in LANGUAGE_OPTIONS:
            for parity in ('N', 'E', 'O', 'M', 'S'):
                with self.subTest(language=language, parity=parity):
                    self.assertNotEqual(parity_label(language, parity), parity)

    def test_translates_dynamic_menu_labels(self):
        self.assertEqual(translate('en', 'connect', port='COM7'), 'Connect: COM7')
        self.assertEqual(translate('ko', 'exit'), '종료')
        self.assertEqual(translate('zh', 'language'), '语言')


class OutputTemplateTests(unittest.TestCase):
    def test_decodes_escapes_without_corrupting_japanese(self):
        self.assertEqual(
            decode_output_template(r'受付\tデータ\r\n'),
            '受付\tデータ\r\n',
        )

    def test_preserves_unknown_escape_sequences(self):
        self.assertEqual(decode_output_template(r'C:\queue\item'), r'C:\queue\item')

    def test_decodes_all_documented_escape_sequences(self):
        escaped = r'\\\a\b\f\n\r\t\v\x41\u3001\U0001F600'
        expected = '\\' + '\a\b\f\n\r\t\v' + 'A、😀'
        self.assertEqual(decode_output_template(escaped), expected)

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


class SerialDataDecodingTests(unittest.TestCase):
    def test_includes_korean_and_chinese_legacy_encodings(self):
        self.assertIn('euc_kr', ENCODING_OPTIONS)
        self.assertIn('gb18030', ENCODING_OPTIONS)

    def test_decodes_korean_euc_kr_data(self):
        encoded = b'\xc7\xd1\xb1\xb9\xbe\xee'
        self.assertEqual(encoded.decode('euc_kr'), '한국어')

    def test_decodes_chinese_gb18030_data(self):
        encoded = b'\xd6\xd0\xce\xc4'
        self.assertEqual(encoded.decode('gb18030'), '中文')


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
            'language': 'ja',
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
            ('language', 'fr'),
        ):
            with self.subTest(key=key, value=value):
                config = dict(self.settings_config)
                config[key] = value
                with self.assertRaises(ValueError):
                    validate_settings_config(config)


if __name__ == '__main__':
    unittest.main()
