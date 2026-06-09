"""Localization helpers and user-interface messages for ser2key."""

from __future__ import annotations

from typing import Final

LANGUAGE_OPTIONS: Final[tuple[str, ...]] = ('ja', 'en', 'ko', 'zh')
DEFAULT_LANGUAGE: Final = 'en'

LANGUAGE_LABELS: Final[dict[str, str]] = {
    'ja': '日本語',
    'en': 'English',
    'ko': '한국어',
    'zh': '中文',
}

TRANSLATIONS: Final[dict[str, dict[str, str]]] = {
    'ja': {
        'refresh_ports': 'ポート一覧を更新',
        'connect': '接続: {port}',
        'no_ports': 'ポートが見つかりません',
        'serial_settings': 'シリアル設定',
        'baudrate': 'ボーレート',
        'bytesize': 'データ長',
        'parity': 'パリティ',
        'stopbits': 'ストップビット',
        'timeout': 'タイムアウト',
        'seconds': '秒',
        'decode_settings': 'デコード設定',
        'encoding': '文字コード',
        'language': '言語',
        'exit': '終了',
        'no_options': '選択肢なし',
        'initializing': '初期化中',
        'connected': '接続中 ({port})',
        'connection_failed': '接続失敗 ({error})',
        'error': 'エラー ({error})',
        'disconnected': '切断',
    },
    'en': {
        'refresh_ports': 'Refresh ports',
        'connect': 'Connect: {port}',
        'no_ports': 'No ports found',
        'serial_settings': 'Serial settings',
        'baudrate': 'Baud rate',
        'bytesize': 'Data bits',
        'parity': 'Parity',
        'stopbits': 'Stop bits',
        'timeout': 'Timeout',
        'seconds': 'sec',
        'decode_settings': 'Decode settings',
        'encoding': 'Character encoding',
        'language': 'Language',
        'exit': 'Exit',
        'no_options': 'No options',
        'initializing': 'Initializing',
        'connected': 'Connected ({port})',
        'connection_failed': 'Connection failed ({error})',
        'error': 'Error ({error})',
        'disconnected': 'Disconnected',
    },
    'ko': {
        'refresh_ports': '포트 목록 새로 고침',
        'connect': '연결: {port}',
        'no_ports': '포트를 찾을 수 없습니다',
        'serial_settings': '직렬 통신 설정',
        'baudrate': '보드 레이트',
        'bytesize': '데이터 비트',
        'parity': '패리티',
        'stopbits': '정지 비트',
        'timeout': '시간 제한',
        'seconds': '초',
        'decode_settings': '디코딩 설정',
        'encoding': '문자 인코딩',
        'language': '언어',
        'exit': '종료',
        'no_options': '선택 항목 없음',
        'initializing': '초기화 중',
        'connected': '연결됨 ({port})',
        'connection_failed': '연결 실패 ({error})',
        'error': '오류 ({error})',
        'disconnected': '연결 끊김',
    },
    'zh': {
        'refresh_ports': '刷新端口列表',
        'connect': '连接：{port}',
        'no_ports': '未找到端口',
        'serial_settings': '串口设置',
        'baudrate': '波特率',
        'bytesize': '数据位',
        'parity': '校验位',
        'stopbits': '停止位',
        'timeout': '超时',
        'seconds': '秒',
        'decode_settings': '解码设置',
        'encoding': '字符编码',
        'language': '语言',
        'exit': '退出',
        'no_options': '无可用选项',
        'initializing': '正在初始化',
        'connected': '已连接 ({port})',
        'connection_failed': '连接失败 ({error})',
        'error': '错误 ({error})',
        'disconnected': '已断开',
    },
}

PARITY_LABELS: Final[dict[str, dict[str, str]]] = {
    'ja': {'N': 'なし', 'E': '偶数', 'O': '奇数', 'M': 'マーク', 'S': 'スペース'},
    'en': {'N': 'None', 'E': 'Even', 'O': 'Odd', 'M': 'Mark', 'S': 'Space'},
    'ko': {'N': '없음', 'E': '짝수', 'O': '홀수', 'M': '마크', 'S': '스페이스'},
    'zh': {'N': '无', 'E': '偶数', 'O': '奇数', 'M': '标记', 'S': '空格'},
}


def language_from_windows_lang_id(lang_id: int) -> str:
    """Map a Windows LANGID to a supported language using its primary ID."""
    primary_language = int(lang_id) & 0x03FF
    return {
        0x04: 'zh',
        0x09: 'en',
        0x11: 'ja',
        0x12: 'ko',
    }.get(primary_language, DEFAULT_LANGUAGE)


def language_from_locale(locale_name: str | None) -> str:
    """Map a Windows locale name such as ``ja-JP`` to a supported language."""
    normalized = (locale_name or '').strip().lower().replace('_', '-')
    primary = normalized.split('-', 1)[0]
    aliases = {'ja': 'ja', 'en': 'en', 'ko': 'ko', 'zh': 'zh'}
    return aliases.get(primary, DEFAULT_LANGUAGE)


def normalize_language(language: str | None) -> str:
    """Validate and normalize a configured language code."""
    normalized = (language or '').strip().lower()
    if normalized not in LANGUAGE_OPTIONS:
        raise ValueError(f'Unsupported language: {language}')
    return normalized


def translate(language: str, key: str, **values: object) -> str:
    """Return a localized UI message, falling back to English."""
    messages = TRANSLATIONS.get(language, TRANSLATIONS[DEFAULT_LANGUAGE])
    template = messages.get(key, TRANSLATIONS[DEFAULT_LANGUAGE].get(key, key))
    return template.format(**values)


def parity_label(language: str, parity: str) -> str:
    """Return the user-facing name of a pyserial parity code."""
    labels = PARITY_LABELS.get(language, PARITY_LABELS[DEFAULT_LANGUAGE])
    return labels.get(parity, parity)
