"""Отчёт: куда ушло, где забанили (dead), где были flood-ошибки.

Запуск: python stats.py — печатает отчёт в stdout и сохраняет копию в
reports/YYYY-MM-DD-stats.txt.
"""
from __future__ import annotations

import sqlite3
from datetime import datetime

import config
import state

REPORTS_DIR = config.BASE_DIR / "reports"


def generate_report() -> str:
    state.init_db()
    groups = state.load_groups()
    lines: list[str] = []
    now = datetime.now(config.BERLIN_TZ)
    lines.append(f"Отчёт по рассылке — {now.strftime('%Y-%m-%d %H:%M %Z')}")
    lines.append("=" * 60)

    with sqlite3.connect(config.DB_PATH) as conn:
        total_sent = conn.execute("SELECT COUNT(*) FROM posts WHERE status='sent'").fetchone()[0]
        total_errors = conn.execute("SELECT COUNT(*) FROM posts WHERE status='error'").fetchone()[0]
        total_joined = conn.execute("SELECT COUNT(*) FROM joins WHERE status='joined'").fetchone()[0]
        lines.append(f"Всего отправлено успешно: {total_sent}")
        lines.append(f"Всего ошибок при отправке: {total_errors}")
        lines.append(f"Всего вступлений в группы: {total_joined}")
        lines.append("")

        lines.append("По группам:")
        for g in groups:
            gid = g["id"]
            sent = conn.execute(
                "SELECT COUNT(*) FROM posts WHERE group_id=? AND status='sent'", (gid,)
            ).fetchone()[0]
            errors_row = conn.execute(
                "SELECT COUNT(*), MAX(error) FROM posts WHERE group_id=? AND status='error'", (gid,)
            ).fetchone()
            err_count, last_error = errors_row
            joined = state.has_joined(gid)
            dead_marker = " [DEAD]" if g.get("dead") else ""
            approval_marker = "" if g.get("requires_approval") else " [БЕЗ РАЗРЕШЕНИЯ — пропускается]"
            lines.append(
                f"  {gid}{dead_marker}{approval_marker}: joined={joined}, sent={sent}, errors={err_count}"
                + (f", last_error={last_error}" if last_error else "")
            )

        lines.append("")
        full_stop = state.get_full_stop()
        if full_stop:
            lines.append(f"⚠️ Активна полная остановка: {full_stop}")
        else:
            lines.append("Полная остановка не активна.")

    return "\n".join(lines)


def main() -> None:
    report = generate_report()
    print(report)
    REPORTS_DIR.mkdir(exist_ok=True)
    out_path = REPORTS_DIR / f"{datetime.now().strftime('%Y-%m-%d')}-stats.txt"
    out_path.write_text(report + "\n", encoding="utf-8")
    print(f"\nСохранено в {out_path}")


if __name__ == "__main__":
    main()
