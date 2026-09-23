"""Конфигурация из .env. Никаких секретов в коде — только чтение переменных окружения."""
from __future__ import annotations

import os
from pathlib import Path
from zoneinfo import ZoneInfo

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


def _int(name: str, default: int) -> int:
    return int(os.environ.get(name, default))


API_ID = _int("API_ID", 0)
API_HASH = os.environ.get("API_HASH", "")
PHONE_NUMBER = os.environ.get("PHONE_NUMBER", "")
SESSION_NAME = os.environ.get("SESSION_NAME", "nas_connect_poster")
SESSION_PATH = str(BASE_DIR / SESSION_NAME)

DAILY_POST_LIMIT = _int("DAILY_POST_LIMIT", 25)
WARMUP_DAYS = _int("WARMUP_DAYS", 14)
WARMUP_START_LIMIT = _int("WARMUP_START_LIMIT", 5)

POST_DELAY_MIN = _int("POST_DELAY_MIN", 420)
POST_DELAY_MAX = _int("POST_DELAY_MAX", 1500)

BERLIN_TZ = ZoneInfo("Europe/Berlin")
NIGHT_START_HOUR = _int("NIGHT_START_HOUR", 23)
NIGHT_END_HOUR = _int("NIGHT_END_HOUR", 7)

JOIN_DAILY_LIMIT = _int("JOIN_DAILY_LIMIT", 6)
JOIN_DELAY_MIN_MINUTES = _int("JOIN_DELAY_MIN_MINUTES", 15)
JOIN_DELAY_MAX_MINUTES = _int("JOIN_DELAY_MAX_MINUTES", 40)

DB_PATH = str(BASE_DIR / os.environ.get("DB_PATH", "state.db"))

ALERT_PEER = os.environ.get("ALERT_PEER", "me")

GROUPS_PATH = BASE_DIR / "groups.json"
TEMPLATES_DIR = BASE_DIR / "templates"

# Абсолютный жёсткий потолок постов/сутки независимо от прогрева — правило
# антиспама из ТЗ, не должно превышаться никакой конфигурацией.
HARD_DAILY_POST_CEILING = 25
