"""SQLite-состояние: лог отправок, лог join'ов, полная остановка (PeerFlood),
первый запуск (для расчёта прогрева). Плюс helpers для groups.json — единственный
источник правды по dead-флагу группы (мутируется на месте, чтобы человек видел
причину прямо в файле, который сам же редактирует руками)."""
from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from typing import Iterator, Optional

import config

_SCHEMA = """
CREATE TABLE IF NOT EXISTS posts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    group_id TEXT NOT NULL,
    template_hash TEXT NOT NULL,
    status TEXT NOT NULL,          -- 'sent' | 'error'
    error TEXT,
    sent_at TEXT NOT NULL          -- ISO8601 UTC
);
CREATE INDEX IF NOT EXISTS idx_posts_group ON posts(group_id);
CREATE INDEX IF NOT EXISTS idx_posts_sent_at ON posts(sent_at);

CREATE TABLE IF NOT EXISTS joins (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    group_id TEXT NOT NULL,
    status TEXT NOT NULL,          -- 'joined' | 'error'
    error TEXT,
    joined_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_joins_joined_at ON joins(joined_at);

CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"""


@contextmanager
def _conn() -> Iterator[sqlite3.Connection]:
    conn = sqlite3.connect(config.DB_PATH)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with _conn() as conn:
        conn.executescript(_SCHEMA)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _today_start_utc() -> str:
    now_berlin = datetime.now(config.BERLIN_TZ)
    midnight_berlin = now_berlin.replace(hour=0, minute=0, second=0, microsecond=0)
    return midnight_berlin.astimezone(timezone.utc).isoformat()


# --- posts ---

def record_post(group_id: str, template_hash: str, status: str, error: Optional[str] = None) -> None:
    with _conn() as conn:
        conn.execute(
            "INSERT INTO posts (group_id, template_hash, status, error, sent_at) VALUES (?, ?, ?, ?, ?)",
            (group_id, template_hash, status, error, _now_iso()),
        )


def posts_sent_today() -> int:
    with _conn() as conn:
        row = conn.execute(
            "SELECT COUNT(*) FROM posts WHERE status = 'sent' AND sent_at >= ?",
            (_today_start_utc(),),
        ).fetchone()
        return row[0]


def last_post_time(group_id: str) -> Optional[datetime]:
    with _conn() as conn:
        row = conn.execute(
            "SELECT sent_at FROM posts WHERE group_id = ? AND status = 'sent' ORDER BY sent_at DESC LIMIT 1",
            (group_id,),
        ).fetchone()
        return datetime.fromisoformat(row[0]) if row else None


def recent_hashes(limit: int = 20) -> set[str]:
    with _conn() as conn:
        rows = conn.execute(
            "SELECT template_hash FROM posts WHERE status = 'sent' ORDER BY sent_at DESC LIMIT ?",
            (limit,),
        ).fetchall()
        return {r[0] for r in rows}


# --- joins ---

def record_join(group_id: str, status: str, error: Optional[str] = None) -> None:
    with _conn() as conn:
        conn.execute(
            "INSERT INTO joins (group_id, status, error, joined_at) VALUES (?, ?, ?, ?)",
            (group_id, status, error, _now_iso()),
        )


def joins_done_today() -> int:
    with _conn() as conn:
        row = conn.execute(
            "SELECT COUNT(*) FROM joins WHERE status = 'joined' AND joined_at >= ?",
            (_today_start_utc(),),
        ).fetchone()
        return row[0]


def has_joined(group_id: str) -> bool:
    with _conn() as conn:
        row = conn.execute(
            "SELECT 1 FROM joins WHERE group_id = ? AND status = 'joined' LIMIT 1",
            (group_id,),
        ).fetchone()
        return row is not None


# --- meta: first run (для прогрева) ---

def get_first_run() -> datetime:
    with _conn() as conn:
        row = conn.execute("SELECT value FROM meta WHERE key = 'first_run'").fetchone()
        if row:
            return datetime.fromisoformat(row[0])
        now = datetime.now(timezone.utc)
        conn.execute(
            "INSERT INTO meta (key, value) VALUES ('first_run', ?)", (now.isoformat(),)
        )
        return now


# --- meta: полная остановка после PeerFloodError ---

def set_full_stop(hours: float, reason: str) -> None:
    until = datetime.now(timezone.utc) + timedelta(hours=hours)
    payload = json.dumps({"until": until.isoformat(), "reason": reason})
    with _conn() as conn:
        conn.execute(
            "INSERT INTO meta (key, value) VALUES ('full_stop', ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (payload,),
        )


def get_full_stop() -> Optional[dict]:
    with _conn() as conn:
        row = conn.execute("SELECT value FROM meta WHERE key = 'full_stop'").fetchone()
        if not row:
            return None
        payload = json.loads(row[0])
        until = datetime.fromisoformat(payload["until"])
        if datetime.now(timezone.utc) >= until:
            return None
        return payload


def clear_full_stop() -> None:
    with _conn() as conn:
        conn.execute("DELETE FROM meta WHERE key = 'full_stop'")


# --- groups.json helpers ---

def load_groups() -> list[dict]:
    return json.loads(config.GROUPS_PATH.read_text(encoding="utf-8"))


def save_groups(groups: list[dict]) -> None:
    config.GROUPS_PATH.write_text(
        json.dumps(groups, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def mark_group_dead(group_id: str, reason: str) -> None:
    groups = load_groups()
    changed = False
    for g in groups:
        if g["id"] == group_id and not g.get("dead"):
            g["dead"] = True
            g["notes"] = (g.get("notes") or "").strip()
            g["notes"] = (g["notes"] + f" | dead: {reason} ({_now_iso()})").strip(" |")
            changed = True
    if changed:
        save_groups(groups)
