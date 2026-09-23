"""Основной цикл рассылки объявления по группам.

Запуск: python poster.py (или через systemd/telegram-vacancy-poster.service).
Первый запуск интерактивен — Telethon попросит код подтверждения для
PHONE_NUMBER и создаст файл сессии SESSION_NAME.session.
"""
from __future__ import annotations

import argparse
import asyncio
import logging
import random
from datetime import datetime, timedelta, timezone

from telethon import TelegramClient, errors

import config
import render
import state

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("poster")


def warmup_daily_limit() -> int:
    if config.WARMUP_DAYS <= 0:
        return min(config.DAILY_POST_LIMIT, config.HARD_DAILY_POST_CEILING)
    first_run = state.get_first_run()
    day_index = (datetime.now(timezone.utc) - first_run).days
    if day_index >= config.WARMUP_DAYS:
        limit = config.DAILY_POST_LIMIT
    else:
        span = config.DAILY_POST_LIMIT - config.WARMUP_START_LIMIT
        limit = round(config.WARMUP_START_LIMIT + span * day_index / config.WARMUP_DAYS)
    return max(1, min(limit, config.HARD_DAILY_POST_CEILING))


def is_night(now_berlin: datetime) -> bool:
    h = now_berlin.hour
    if config.NIGHT_START_HOUR > config.NIGHT_END_HOUR:
        return h >= config.NIGHT_START_HOUR or h < config.NIGHT_END_HOUR
    return config.NIGHT_START_HOUR <= h < config.NIGHT_END_HOUR


def seconds_until_night_end(now_berlin: datetime) -> float:
    end = now_berlin.replace(hour=config.NIGHT_END_HOUR, minute=0, second=0, microsecond=0)
    if end <= now_berlin:
        end += timedelta(days=1)
    return (end - now_berlin).total_seconds()


def seconds_until_berlin_midnight(now_berlin: datetime) -> float:
    tomorrow = (now_berlin + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    return (tomorrow - now_berlin).total_seconds()


def eligible_groups(groups: list[dict]) -> list[dict]:
    now = datetime.now(timezone.utc)
    result = []
    for g in groups:
        if g.get("dead"):
            continue
        if not g.get("requires_approval"):
            continue
        last = state.last_post_time(g["id"])
        if last is not None:
            min_interval = timedelta(hours=g.get("min_interval_hours", 24))
            if now - last < min_interval:
                continue
        result.append(g)
    return result


async def alert_saved_messages(client: TelegramClient, text: str) -> None:
    try:
        await client.send_message(config.ALERT_PEER, text)
    except Exception:
        log.exception("Не удалось отправить алерт в Saved Messages")


async def send_to_group(
    client: TelegramClient, group: dict, rng: random.Random, dry_run: bool = False
) -> str:
    """Возвращает 'ok' / 'flood_stop' (нужна полная остановка цикла) / 'skip'."""
    text, text_hash = render.render_unique(rng, state.recent_hashes())
    group_id = group["id"]
    if dry_run:
        log.info("[DRY-RUN] в %s (%s) ушёл бы текст:\n%s", group_id, group["username"], text)
        return "ok"
    try:
        await client.send_message(group["username"], text)
        state.record_post(group_id, text_hash, "sent")
        log.info("Отправлено в %s", group_id)
        return "ok"
    except errors.PeerFloodError as e:
        state.record_post(group_id, text_hash, "error", str(e))
        state.set_full_stop(24, "PeerFloodError")
        log.error("PeerFloodError — полная остановка на 24ч")
        await alert_saved_messages(
            client,
            f"🚨 PeerFloodError на группе {group_id}. Остановка рассылки на 24 часа. "
            f"Это метка антиспама Telegram — ретраить нельзя, дождитесь окончания паузы.",
        )
        return "flood_stop"
    except errors.SlowModeWaitError as e:
        state.record_post(group_id, text_hash, "error", str(e))
        log.warning("SlowModeWaitError в %s, ждём %sс + пропускаем группу", group_id, e.seconds)
        await asyncio.sleep(e.seconds + 10)
        return "skip"
    except errors.FloodWaitError as e:
        state.record_post(group_id, text_hash, "error", str(e))
        log.warning("FloodWaitError, ждём %sс", e.seconds)
        await asyncio.sleep(e.seconds + 60)
        return "skip"
    except (errors.ChatWriteForbiddenError, errors.UserBannedInChannelError, errors.ChannelPrivateError) as e:
        state.record_post(group_id, text_hash, "error", str(e))
        state.mark_group_dead(group_id, type(e).__name__)
        log.error("Группа %s помечена dead: %s", group_id, e)
        return "skip"
    except Exception as e:
        state.record_post(group_id, text_hash, "error", str(e))
        log.exception("Неожиданная ошибка при отправке в %s", group_id)
        return "skip"


async def run_cycle(
    client: TelegramClient, rng: random.Random, dry_run: bool = False, once: bool = False
) -> bool:
    """Возвращает False, если работу пора завершать (--once / --dry-run отработали)."""
    groups = state.load_groups()
    order = eligible_groups(groups)
    rng.shuffle(order)

    for group in order:
        full_stop = state.get_full_stop()
        if full_stop:
            log.info("Полная остановка активна (%s), пропускаем цикл", full_stop["reason"])
            return True

        now_berlin = datetime.now(config.BERLIN_TZ)
        if is_night(now_berlin) and not dry_run:
            log.info("Ночная пауза, откладываем до утра")
            return True

        if state.posts_sent_today() >= warmup_daily_limit():
            log.info("Дневной лимит достигнут (%d)", warmup_daily_limit())
            return True

        result = await send_to_group(client, group, rng, dry_run)
        if result == "flood_stop":
            return True
        if once and result == "ok":
            log.info("--once: один пост сделан, выходим")
            return False

        pause = 2.0 if dry_run else rng.uniform(config.POST_DELAY_MIN, config.POST_DELAY_MAX)
        log.info("Пауза %.0fс перед следующей группой", pause)
        await asyncio.sleep(pause)

    return not dry_run


async def main(dry_run: bool = False, once: bool = False) -> None:
    state.init_db()
    rng = random.Random()
    client = TelegramClient(config.SESSION_PATH, config.API_ID, config.API_HASH)
    await client.start(phone=config.PHONE_NUMBER or None)
    log.info("Клиент авторизован, дневной лимит сегодня: %d", warmup_daily_limit())
    if dry_run:
        log.info("DRY-RUN: ничего отправлено не будет, только показ текстов")

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
            if is_night(now_berlin) and not dry_run:
                sleep_s = seconds_until_night_end(now_berlin)
                log.info("Ночь, спим %.0fс до конца ночной паузы", sleep_s)
                await asyncio.sleep(min(sleep_s, 3600))
                continue

            if state.posts_sent_today() >= warmup_daily_limit():
                sleep_s = seconds_until_berlin_midnight(now_berlin)
                log.info("Лимит на сегодня исчерпан, спим %.0fс до полуночи по Берлину", sleep_s)
                await asyncio.sleep(min(sleep_s, 3600))
                continue

            groups = state.load_groups()
            if not eligible_groups(groups):
                log.info("Нет подходящих групп сейчас (min_interval/approval), спим 30 мин")
                if dry_run:
                    return
                await asyncio.sleep(1800)
                continue

            if not await run_cycle(client, rng, dry_run, once):
                return


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Рассылка вакансии NAS-Connect по Telegram-группам")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="ничего не отправлять: показать, какой текст в какую группу ушёл бы, и выйти",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="отправить ровно один пост и выйти (для первой боевой проверки)",
    )
    args = parser.parse_args()
    asyncio.run(main(dry_run=args.dry_run, once=args.once))
