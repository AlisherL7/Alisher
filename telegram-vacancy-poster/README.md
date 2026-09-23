# telegram-vacancy-poster

Постинг вакансии NAS-Connect GmbH (Glasfaser-Tiefbau, Niedersachsen) в
русско- и узбекоязычные Telegram-группы «работа в Германии» от личного
аккаунта через MTProto (Telethon).

## Риск аккаунта — прочитать до первого запуска

Инструмент рассчитан на **один** аккаунт, без обхода лимитов и без смены IP.
Главный фактор риска — не количество постов само по себе, а возраст и
«история» аккаунта: Telegram оценивает доверие к аккаунту по времени жизни
номера, наличию обычной переписки/звонков и по жалобам на спам. Новый
аккаунт, который сразу начинает постить в десятки групп, попадает под
spam-restriction (`PeerFloodError`) быстрее всего, а снять её можно только
через @SpamBot и не всегда.

Практика:

- **Если есть старый аккаунт с обычной историей — используйте его.** Код и
  конфиг не меняются, меняются только `API_ID`/`API_HASH`/`PHONE_NUMBER` в
  `.env`.
- Для нового аккаунта в `config.py` заложен **прогрев**: лимит постов
  растёт линейно от `WARMUP_START_LIMIT` (по умолчанию 5) до
  `DAILY_POST_LIMIT` за `WARMUP_DAYS` (по умолчанию 14) дней с первого
  запуска. Для старого аккаунта поставьте `WARMUP_DAYS=0`.
- Жёсткий потолок `HARD_DAILY_POST_CEILING = 25` в `config.py` не
  обходится никакой конфигурацией.
- `PeerFloodError` = полная остановка на 24 часа + алерт в Saved Messages.
  Ретраить нельзя: это уже пометка антиспама, повторные попытки ухудшают
  положение.

## Юридические ограничения, зашитые в текст

Текст объявления в `render.py` (блоки `_offer_block` / `_requirements_block`)
не варьируется в части формулировок, обязательных по праву:

- заголовок с `(m/w/d)` — § 11 AGG, отсутствие маркера полов —
  повод для иска по § 15 AGG;
- **`Arbeitserlaubnis für Deutschland`**, а не «Германия и/или Польша» —
  польский work permit не даёт права работать в Германии; наём без
  разрешения — Schwarzarbeit (§ 404 SGB III, § 266a StGB);
- **`Unterkunft wird gestellt`**, а не «жильё бесплатно» — Unterkunft это
  Sachbezug и по § 107 Abs. 2 GewO не может замещать часть Mindestlohn;
- никаких требований к национальности, возрасту, полу.

Варьируются только приветствие, порядок блоков «Предлагаем»/«Требования»,
эмодзи и заголовок — чтобы хэш текста не совпадал между отправками.

Не делает и делать не следует:

- рассылка в личные сообщения — § 7 UWG, abmahnfähig;
- парсинг участников групп;
- несколько аккаунтов параллельно, смена IP между отправками.

Кроме того, `poster.py` **пропускает группы с `requires_approval: false`** —
постить только туда, где владелец группы дал разрешение на объявление.

## Установка на VPS

```bash
# 1. Пользователь и каталог
sudo adduser --system --group --home /opt/telegram-vacancy-poster nasconnect
sudo -u nasconnect git clone <repo> /opt/telegram-vacancy-poster
cd /opt/telegram-vacancy-poster/telegram-vacancy-poster

# 2. Виртуальное окружение
python3.11 -m venv .venv
.venv/bin/pip install -r requirements.txt

# 3. Конфиг
cp .env.example .env
# получить API_ID/API_HASH на https://my.telegram.org/apps, вписать в .env
```

### Разовая авторизация сессии (интерактивно)

Telethon при первом запуске запросит код подтверждения из Telegram — это
нужно сделать руками на VPS, автоматизировать нельзя:

```bash
.venv/bin/python poster.py
# ввести номер (если не задан в .env), затем код из Telegram,
# при включённой 2FA — пароль
```

Создастся файл `<SESSION_NAME>.session`. Он равносилен доступу к аккаунту —
не коммитить, права `chmod 600`.

### systemd

```bash
sudo cp systemd/telegram-vacancy-*.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now telegram-vacancy-joiner    # сначала вступаем
# после того как вступили в нужные группы:
sudo systemctl enable --now telegram-vacancy-poster
journalctl -u telegram-vacancy-poster -f
```

Оба процесса — долгоживущие asyncio-циклы, которые сами спят между
действиями и учитывают ночное окно, поэтому `.timer` не используется:
timer + внешний lock для процесса с многочасовыми паузами только создаёт
риск параллельного запуска.

## Файлы

| Файл | Назначение |
|---|---|
| `config.py` | Конфиг из `.env`, потолки, таймзона Берлина |
| `groups.json` | Список групп: `username`, `min_interval_hours`, `requires_approval`, `dead`, правила группы |
| `templates/variant_*.txt` | 6 скелетов объявления |
| `templates/fragments.json` | Варианты приветствий/эмодзи/порядка блоков |
| `render.py` | Сборка текста + проверка, что хэш не повторяет последние отправки |
| `state.py` | SQLite: лог постов, лог join'ов, полная остановка, dead-флаги в `groups.json` |
| `poster.py` | Основной цикл рассылки, вся обработка ошибок Telethon |
| `joiner.py` | Постепенное вступление в группы |
| `stats.py` | Отчёт в stdout + `reports/YYYY-MM-DD-stats.txt` |

## Эксплуатация

```bash
.venv/bin/python stats.py     # куда ушло, где dead, где ошибки
```

Добавление группы: дописать объект в `groups.json` с
`requires_approval: true` и заполненным `approved_by` (кто и когда дал
разрешение — на случай претензий), затем `systemctl restart
telegram-vacancy-joiner`. Группы читаются из файла на каждом цикле,
рестарт poster'а не обязателен.

Снятие `dead`: поправить флаг в `groups.json` руками. Автоматически
`dead` не снимается — если группа забанила аккаунт, повторная попытка
это очередная жалоба.

## Матрица обработки ошибок

| Ошибка | Реакция |
|---|---|
| `FloodWaitError` | `sleep(e.seconds + 60)`, лог, следующая группа |
| `SlowModeWaitError` | `sleep(e.seconds + 10)`, группа пропускается |
| `PeerFloodError` | полная остановка 24 ч + алерт в Saved Messages, без ретраев |
| `ChatWriteForbiddenError` / `UserBannedInChannelError` / `ChannelPrivateError` | `dead=true` в `groups.json`, группа больше не трогается |
| любая другая | лог, следующая группа |
