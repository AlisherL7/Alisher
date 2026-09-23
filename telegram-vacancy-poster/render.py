"""Шаблонизатор текста объявления.

Юридически обязательные формулировки (AGG §11, §107 GewO, §266a StGB /
§404 SGB III — см. комментарии в OFFER_BLOCK/REQUIREMENTS_BLOCK) — это
неизменная часть текста. Варьируются только приветствие, порядок блоков,
эмодзи и заголовок — ради того, чтобы хэш сообщения не совпадал между
отправками, а не ради изменения смысла.
"""
from __future__ import annotations

import hashlib
import json
import random

import config

_VARIANT_FILES = sorted(config.TEMPLATES_DIR.glob("variant_*.txt"))
_FRAGMENTS = json.loads((config.TEMPLATES_DIR / "fragments.json").read_text(encoding="utf-8"))


def _offer_block(bullet: str) -> str:
    return (
        "Предлагаем:\n"
        f"{bullet} 16-18 €/час в зависимости от квалификации\n"
        f"{bullet} Unterkunft wird gestellt\n"  # НЕ "жильё бесплатно" — §107 Abs.2 GewO
        f"{bullet} 30 дней оплачиваемого отпуска\n"
        f"{bullet} Рабочая одежда\n"
        f"{bullet} Официальное трудоустройство с Sozialversicherung\n"
        f"{bullet} Долгосрочная занятость"
    )


def _requirements_block(bullet: str) -> str:
    return (
        "Требования:\n"
        f"{bullet} Führerschein Klasse B\n"
        f"{bullet} Arbeitserlaubnis für Deutschland\n"  # НЕ "Германия и/или Польша" — §404 SGB III
        f"{bullet} Готовность работать на улице\n"
        f"{bullet} Опыт не обязателен"
    )


def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def render_once(rng: random.Random) -> str:
    if not _VARIANT_FILES:
        raise RuntimeError(f"Нет файлов templates/variant_*.txt в {config.TEMPLATES_DIR}")

    variant_path = rng.choice(_VARIANT_FILES)
    template = variant_path.read_text(encoding="utf-8")

    bullet = rng.choice(_FRAGMENTS["emoji_bullet"])
    order = rng.choice(_FRAGMENTS["block_orders"])
    blocks_map = {"offer": _offer_block(bullet), "requirements": _requirements_block(bullet)}
    blocks = "\n\n".join(blocks_map[name] for name in order)

    text = template.format(
        greeting=rng.choice(_FRAGMENTS["greetings"]),
        emoji_header=rng.choice(_FRAGMENTS["emoji_header"]),
        closing=rng.choice(_FRAGMENTS["closing"]),
        blocks=blocks,
    )
    # Схлопнуть возможные пустые строки от пустых плейсхолдеров (пустой greeting/emoji).
    lines = [line.rstrip() for line in text.splitlines()]
    cleaned = "\n".join(lines)
    while "\n\n\n" in cleaned:
        cleaned = cleaned.replace("\n\n\n", "\n\n")
    return cleaned.strip() + "\n"


def render_unique(rng: random.Random, recent_hashes: set[str], max_attempts: int = 30) -> tuple[str, str]:
    """Генерирует текст, чей хэш не встречается в recent_hashes."""
    for _ in range(max_attempts):
        text = render_once(rng)
        h = _sha256(text)
        if h not in recent_hashes:
            return text, h
    # Не удалось найти уникальный за max_attempts попыток (маловероятно
    # при 6 шаблонах * фрагментах) — возвращаем последний вариант как есть.
    return text, h
