"""Постепенное вступление в группы из groups.json.

Массовый join — сильный сигнал антиспама, поэтому: максимум
JOIN_DAILY_LIMIT (по умолчанию 6) в сутки, паузы
JOIN_DELAY_MIN_MINUTES..JOIN_DELAY_MAX_MINUTES между вступлениями.

Запуск отдельно от poster.py: python joiner.py (или через
systemd/telegram-vacancy-joiner.service). Использует ту же сессию Telethon.
"""
from __future__ import annotations

import asyncio
import logging
import random
from datetime import datetime, timedelta, timezone

from telethon import TelegramClient, errors
from telethon.tl.functions.channels import JoinChannelRequest

import config
import state
from poster import alert_saved_messages, is_night, seconds_until_night_end

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("joiner")


def groups_to_join(groups: list[dict]) -> list[dict]:
    return [g for g in groups if not g.get("dead") and not state.has_joined(g["id"])]


async def join_one(client: TelegramClient, group: dict) -> str:
    """Возвращает 'ok' / 'flood_stop' / 'skip'."""
    group_id = group["id"]
    try:
        await client(JoinChannelRequest(group["username"]))
        state.record_join(group_id, "joined")
        log.info("Вступили в %s", group_id)
        return "ok"
    except errors.UserAlreadyParticipantError:
        state.record_join(group_id, "joined", "already participant")
        log.info("Уже состоим в %s", group_id)
        return "ok"
    except errors.PeerFloodError as e:
        state.record_join(group_id, "error", str(e))
        state.set_full_stop(24, "PeerFloodError (join)")
        log.error("PeerFloodError при join — полная остановка на 24ч")
        await alert_saved_messages(
            client,
            f"🚨 PeerFloodError при вступлении в {group_id}. Остановка на 24 часа.",
        )
        return "flood_stop"
    except errors.FloodWaitError as e:
        state.record_join(group_id, "error", str(e))
        log.warning("FloodWaitError при join, ждём %sс", e.seconds)
        await asyncio.sleep(e.seconds + 60)
        return "skip"
    except (errors.ChannelPrivateError, errors.UserBannedInChannelError, errors.ChatWriteForbiddenError) as e:
        state.record_join(group_id, "error", str(e))
        state.mark_group_dead(group_id, type(e).__name__)
        log.error("Группа %s помечена dead: %s", group_id, e)
        return "skip"
    except Exception as e:
        state.record_join(group_id, "error", str(e))
        log.exception("Неожиданная ошибка при вступлении в %s", group_id)
        return "skip"


async def main() -> None:
    state.init_db()
    rng = random.Random()
    client = TelegramClient(config.SESSION_PATH, config.API_ID, config.API_HASH)
    await client.start(phone=config.PHONE_NUMBER or None)
    log.info("Клиент авторизован (joiner)")

    async with client:
        while True:
            full_stop = state.get_full_stop()
            if full_stop:
                until = datetime.fromisoformat(full_stop["until"])
                sleep_s = max(60.0, (until - datetime.now(timezone.utc)).total_seconds())
                log.info("Полная остановка до %s, спим", until.isoformat())
                await asyncio.sleep(min(sleep_s, 3600))
                continue

            now_berlin = datetime.now(config.BERLIN_TZ)
            if is_night(now_berlin):
                sleep_s = seconds_until_night_end(now_berlin)
                log.info("Ночь, спим %.0fс", sleep_s)
                await asyncio.sleep(min(sleep_s, 3600))
                continue

            if state.joins_done_today() >= config.JOIN_DAILY_LIMIT:
                tomorrow = (now_berlin + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
                sleep_s = (tomorrow - now_berlin).total_seconds()
                log.info("Дневной лимит join исчерпан, спим до полуночи")
                await asyncio.sleep(min(sleep_s, 3600))
                continue

            groups = state.load_groups()
            candidates = groups_to_join(groups)
            if not candidates:
                log.info("Нечего вступать — все группы уже joined или dead, спим час")
                await asyncio.sleep(3600)
                continue

            group = rng.choice(candidates)
            result = await join_one(client, group)
            if result == "flood_stop":
                continue

            pause_minutes = rng.uniform(config.JOIN_DELAY_MIN_MINUTES, config.JOIN_DELAY_MAX_MINUTES)
            log.info("Пауза %.1f мин перед следующим join", pause_minutes)
            await asyncio.sleep(pause_minutes * 60)


if __name__ == "__main__":
    asyncio.run(main())
