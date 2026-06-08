"""Platform-independent configuration and output helpers for ser2key."""

from __future__ import annotations

import codecs
from collections.abc import Callable, Mapping, MutableMapping
import configparser
import math
from datetime import datetime
import re
from typing import Final, Optional

BAUDRATE_OPTIONS: Final[tuple[int, ...]] = (
    1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200
)
BYTESIZE_OPTIONS: Final[tuple[int, ...]] = (5, 6, 7, 8)
PARITY_OPTIONS: Final[tuple[str, ...]] = ('N', 'E', 'O', 'M', 'S')
STOPBITS_OPTIONS: Final[tuple[float, ...]] = (1, 1.5, 2)
TIMEOUT_OPTIONS: Final[tuple[float, ...]] = (0.1, 0.5, 1, 2, 5)
ENCODING_OPTIONS: Final[tuple[str, ...]] = ('shift_jis', 'ascii', 'utf-8')
MAX_BUFFER_MSEC: Final = 60_000

DATETIME_TOKEN_PATTERN = re.compile(r'\{(DATE|TIME|DATETIME)(?::([^}]*))?\}')
ESCAPE_SEQUENCE_PATTERN = re.compile(
    r"\\(?:[\\abfnrtv]|x[0-9a-fA-F]{2}|u[0-9a-fA-F]{4}|U[0-9a-fA-F]{8})"
)
DEFAULT_DATETIME_FORMATS = {
    'DATE': '%Y-%m-%d',
    'TIME': '%H:%M:%S',
    'DATETIME': '%Y-%m-%d %H:%M:%S',
}


def decode_output_template(text: str) -> str:
    """Decode supported backslash escapes without corrupting literal Unicode."""
    if not text:
        return ''

    def decode_match(match: re.Match[str]) -> str:
        return codecs.decode(match.group(0), 'unicode_escape')

    return ESCAPE_SEQUENCE_PATTERN.sub(decode_match, text)


def create_config_parser() -> configparser.ConfigParser:
    """Create a parser that preserves percent signs used by strftime."""
    return configparser.ConfigParser(interpolation=None)


def render_output_template(
    template: str,
    now: datetime,
    on_error: Optional[Callable[[str, str, ValueError], None]] = None,
) -> str:
    """Expand date/time tokens in an output template."""
    if not template:
        return ''

    def replace(match: re.Match[str]) -> str:
        token = match.group(1)
        fmt = match.group(2) or DEFAULT_DATETIME_FORMATS[token]
        try:
            return now.strftime(fmt)
        except ValueError as exc:
            if on_error is not None:
                on_error(token, fmt, exc)
            return match.group(0)

    return DATETIME_TOKEN_PATTERN.sub(replace, template)


def validate_serial_config(config: Mapping[str, object]) -> None:
    """Validate values accepted by pyserial and the tray configuration UI."""
    port = config.get('port')
    if not isinstance(port, str) or not port.strip():
        raise ValueError('シリアルポートが設定されていません')

    baudrate = config.get('baudrate')
    if not isinstance(baudrate, int) or isinstance(baudrate, bool) or baudrate <= 0:
        raise ValueError(f"不正なボーレート: {baudrate}")

    bytesize = config.get('bytesize')
    if (
        not isinstance(bytesize, int)
        or isinstance(bytesize, bool)
        or bytesize not in BYTESIZE_OPTIONS
    ):
        raise ValueError(f"不正なデータ長: {bytesize}")

    parity = config.get('parity')
    if parity not in PARITY_OPTIONS:
        raise ValueError(f"不正なパリティ: {parity}")

    stopbits = config.get('stopbits')
    if (
        not isinstance(stopbits, (int, float))
        or isinstance(stopbits, bool)
        or stopbits not in STOPBITS_OPTIONS
    ):
        raise ValueError(f"不正なストップビット: {stopbits}")

    timeout = config.get('timeout')
    if (
        not isinstance(timeout, (int, float))
        or isinstance(timeout, bool)
        or not math.isfinite(timeout)
        or timeout < 0
    ):
        raise ValueError(f"不正なタイムアウト: {timeout}")


def validate_settings_config(config: MutableMapping[str, object]) -> None:
    """Validate application settings and normalize the codec name in place."""
    encoding_name = config.get('encoding')
    try:
        normalized = codecs.lookup(str(encoding_name)).name
    except (LookupError, TypeError):
        raise ValueError(
            f"サポートされていないエンコーディング: {encoding_name}"
        ) from None

    buffer_msec = config.get('buffer_msec')
    if (
        not isinstance(buffer_msec, int)
        or isinstance(buffer_msec, bool)
        or not 0 <= buffer_msec <= MAX_BUFFER_MSEC
    ):
        raise ValueError(
            f"buffer_msec は 0 から {MAX_BUFFER_MSEC} の整数で指定してください: "
            f"{buffer_msec}"
        )

    add_enter = config.get('add_enter')
    if not isinstance(add_enter, bool):
        raise ValueError(f"add_enter は真偽値で指定してください: {add_enter}")

    config['encoding'] = normalized
