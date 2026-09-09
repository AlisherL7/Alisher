# -*- coding: utf-8 -*-
"""
Генератор Excel-журнала операций из выгрузки Notion.

    python3 scripts/journal_excel.py --data operations.json --out Zhurnal.xlsx

Формат operations.json — список объектов:
    {"Дата": "2026-09-01", "Название": "...", "Сумма": -8.88, "Поток": "Расход",
     "Раздел": "...", "Категория": "...", "Нужность": "...", "Статус": "...",
     "Способ": "...", "Кому / от кого": "...", "Заметки": "..."}

Файл с данными в репозиторий НЕ коммитится — репозиторий публичный.
Совместимость: Excel, LibreOffice, Apple Numbers (без именованных диапазонов,
без функций, которых нет в Numbers).
"""
import argparse, json, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule

C_DARK, C_HEAD, C_SUB = "1F3864", "2E5FA3", "D6E4F7"
C_IN, C_CALC, C_RES, C_WARN = "FFF7D6", "EDEDED", "DFF3DF", "FFD6D6"
F_TITLE = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
F_H1 = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
F_H2 = Font(name="Calibri", size=11, bold=True, color=C_DARK)
F_B  = Font(name="Calibri", size=10, bold=True)
F_N  = Font(name="Calibri", size=10)
F_S  = Font(name="Calibri", size=9, color="606060")
F_RES = Font(name="Calibri", size=12, bold=True, color="006100")
F_RED = Font(name="Calibri", size=10, bold=True, color="9C0006")
FILL_TITLE = PatternFill("solid", fgColor=C_DARK)
FILL_HEAD = PatternFill("solid", fgColor=C_HEAD)
FILL_SUB  = PatternFill("solid", fgColor=C_SUB)
FILL_IN   = PatternFill("solid", fgColor=C_IN)
FILL_CALC = PatternFill("solid", fgColor=C_CALC)
FILL_RES  = PatternFill("solid", fgColor=C_RES)
FILL_WARN = PatternFill("solid", fgColor=C_WARN)
thin = Side(style="thin", color="B0B0B0")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
EUR = '#,##0.00\\ "€"'
DATE = 'DD.MM.YYYY'
PCT = '0.0%'

POTOK  = ["Расход", "Доход"]
RAZDEL = ["Займы", "Семья", "Постоянные", "Повседневные", "Транспорт", "Покупки", "Вопросы", "Доходы"]
KATEG  = ["Займы друзьям", "Семья", "Переводы за границу", "Онлайн-покупки", "Связь и подписки",
          "Наличные", "Продукты", "Еда вне дома", "Самокаты", "Киоск", "Транспорт и топливо",
          "Страховки", "Банковские сборы", "Прочее / не опознано", "Доход"]
NUZHN  = ["Обязательное", "Необязательное", "Возвратное", "Неизвестно", "Доход"]
STATUS = ["Оплачен", "Открыт", "Возврат ожидается", "Возвращено"]
SPOSOB = ["Карта", "Наличные", "Перевод", "Списание"]

COLS = [("Дата", 12), ("Название", 32), ("Сумма", 12), ("Поток", 10), ("Раздел", 14),
        ("Категория", 22), ("Нужность", 16), ("Статус", 18), ("Способ", 12),
        ("Кому / от кого", 22), ("Заметки", 40), ("Месяц", 10), ("Реальный расход", 15)]

HR = 4            # строка заголовка
R0 = 5            # первая строка данных
NROWS = 500       # запас строк для новых операций
R1 = R0 + NROWS - 1


def title_row(ws, row, text, span):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = F_TITLE; c.fill = FILL_TITLE
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 26


def section(ws, row, text, span, col=1):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=span)
    c = ws.cell(row=row, column=col, value=text)
    c.font = F_H1; c.fill = FILL_HEAD
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 19


def build(ops, out):
    wb = Workbook()

    # ---------------- ЖУРНАЛ ----------------
    jr = wb.active
    jr.title = "Журнал"
    jr.sheet_view.showGridLines = False
    title_row(jr, 1, "ЖУРНАЛ ОПЕРАЦИЙ — сюда вносятся все доходы и расходы", len(COLS))
    jr["A2"] = ("Расход пиши со знаком минус, доход — плюс. Жёлтые колонки заполняешь ты, серые считаются сами. "
                "Пустые строки внизу — для новых операций, формулы там уже стоят. "
                "Фото чеков живут в Notion — в Excel их положить нельзя.")
    jr["A2"].font = F_S
    jr.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))

    for i, (name, w) in enumerate(COLS, start=1):
        c = jr.cell(row=HR, column=i, value=name)
        c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        jr.column_dimensions[get_column_letter(i)].width = w
    jr.row_dimensions[HR].height = 28

    for i in range(NROWS):
        r = R0 + i
        op = ops[i] if i < len(ops) else None
        if op:
            d = dt.date.fromisoformat(op["Дата"])
            jr.cell(row=r, column=1, value=d)
            for col, key in [(2, "Название"), (4, "Поток"), (5, "Раздел"), (6, "Категория"),
                             (7, "Нужность"), (8, "Статус"), (9, "Способ"),
                             (10, "Кому / от кого"), (11, "Заметки")]:
                v = op.get(key)
                if v:
                    jr.cell(row=r, column=col, value=v)
            jr.cell(row=r, column=3, value=op["Сумма"])
        jr.cell(row=r, column=1).number_format = DATE
        jr.cell(row=r, column=3).number_format = EUR
        jr.cell(row=r, column=12,
                value=f'=IF($A{r}="","",YEAR($A{r})&"-"&IF(MONTH($A{r})<10,"0","")&MONTH($A{r}))')
        jr.cell(row=r, column=13,
                value=f'=IF($C{r}="","",IF(OR($D{r}="Доход",$G{r}="Возвратное"),0,$C{r}))')
        jr.cell(row=r, column=13).number_format = EUR
        for c in range(1, len(COLS) + 1):
            cell = jr.cell(row=r, column=c)
            cell.border = BOX; cell.font = F_N
            cell.fill = FILL_CALC if c in (12, 13) else FILL_IN
            if c in (1, 3, 4, 12, 13):
                cell.alignment = Alignment(horizontal="center")

    for lst, col in [(POTOK, "D"), (RAZDEL, "E"), (KATEG, "F"),
                     (NUZHN, "G"), (STATUS, "H"), (SPOSOB, "I")]:
        dv = DataValidation(type="list", formula1='"' + ",".join(lst) + '"', allow_blank=True)
        jr.add_data_validation(dv)
        dv.add(f"{col}{R0}:{col}{R1}")

    jr.conditional_formatting.add(f"C{R0}:C{R1}",
        CellIsRule(operator="lessThan", formula=["0"], font=Font(color="9C0006")))
    jr.conditional_formatting.add(f"C{R0}:C{R1}",
        CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="006100", bold=True)))
    jr.freeze_panes = "C5"

    J = f"Журнал!"
    A, C, D, E, F, G, H, L = (f"{J}$A${R0}:$A${R1}", f"{J}$C${R0}:$C${R1}", f"{J}$D${R0}:$D${R1}",
                              f"{J}$E${R0}:$E${R1}", f"{J}$F${R0}:$F${R1}", f"{J}$G${R0}:$G${R1}",
                              f"{J}$H${R0}:$H${R1}", f"{J}$L${R0}:$L${R1}")
    M = f"{J}$M${R0}:$M${R1}"
    JN = f"{J}$B${R0}:$B${R1}"
    JK = f"{J}$J${R0}:$J${R1}"

    # ---------------- ОБЗОР ----------------
    ov = wb.create_sheet("Обзор")
    ov.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", [2, 34, 16, 12, 46]):
        ov.column_dimensions[col].width = w
    title_row(ov, 1, "ОБЗОР — всё считается автоматически из листа «Журнал»", 5)
    ov["B2"] = "Ничего вручную вписывать не нужно. Добавил операцию в журнал — цифры здесь обновились."
    ov["B2"].font = F_S

    r = 4
    section(ov, r, "ИТОГИ ЗА ВСЁ ВРЕМЯ", 5, col=2); r += 1
    lines = [
        ("Пришло", f'=SUMIFS({C},{D},"Доход")', EUR, "все поступления"),
        ("Ушло", f'=SUMIFS({C},{D},"Расход")', EUR, "все списания"),
        ("Разница", f"=C{r}-0+C{r+1}", EUR, "приход минус расход"),
        ("Реальные траты (без займов)", f"=SUM({M})", EUR, "то, что действительно потрачено"),
        ("Выдано в долг", f'=SUMIFS({C},{G},"Возвратное")', EUR, "деньги у друзей"),
        ("Ждёт возврата", f'=SUMIFS({C},{H},"Возврат ожидается")', EUR, "ещё не вернули"),
        ("Операций всего", f'=COUNTIF({A},">0")', '0', ""),
    ]
    first = r
    for i, (label, formula, fmt, note) in enumerate(lines):
        rr = first + i
        if label == "Разница":
            formula = f"=C{first}+C{first+1}"
        ov.cell(row=rr, column=2, value=label).font = F_B
        c = ov.cell(row=rr, column=3, value=formula)
        c.number_format = fmt; c.border = BOX; c.fill = FILL_RES if i < 4 else FILL_CALC
        c.font = F_RES if i == 3 else F_B
        c.alignment = Alignment(horizontal="center")
        ov.cell(row=rr, column=5, value=note).font = F_S
    r = first + len(lines) + 1

    section(ov, r, "ПО РАЗДЕЛАМ", 5, col=2); r += 1
    hdr = r
    for i, h in enumerate(["Раздел", "Сумма", "Операций", "Доля от расходов"], start=2):
        c = ov.cell(row=hdr, column=i, value=h); c.font = F_B; c.fill = FILL_SUB; c.border = BOX
        c.alignment = Alignment(horizontal="center")
    r += 1
    rz_first = r
    for rz in RAZDEL:
        ov.cell(row=r, column=2, value=rz).font = F_N
        ov.cell(row=r, column=3, value=f'=SUMIFS({C},{E},$B{r},{D},"Расход")').number_format = EUR
        ov.cell(row=r, column=4, value=f'=COUNTIFS({E},$B{r})').number_format = '0'
        ov.cell(row=r, column=5, value=f'=IF($C${first+1}=0,0,$C{r}/$C${first+1})').number_format = PCT
        for c in range(2, 6):
            ov.cell(row=r, column=c).border = BOX
            ov.cell(row=r, column=c).fill = FILL_CALC
            ov.cell(row=r, column=c).font = F_N
        r += 1
    ov.cell(row=r, column=2, value="ИТОГО").font = F_B
    ov.cell(row=r, column=3, value=f"=SUM(C{rz_first}:C{r-1})").number_format = EUR
    for c in range(2, 6):
        ov.cell(row=r, column=c).border = BOX; ov.cell(row=r, column=c).fill = FILL_SUB
        ov.cell(row=r, column=c).font = F_B
    r += 2

    section(ov, r, "ПО КАТЕГОРИЯМ", 5, col=2); r += 1
    for i, h in enumerate(["Категория", "Сумма", "Операций", "Доля"], start=2):
        c = ov.cell(row=r, column=i, value=h); c.font = F_B; c.fill = FILL_SUB; c.border = BOX
        c.alignment = Alignment(horizontal="center")
    r += 1
    for kt in KATEG:
        ov.cell(row=r, column=2, value=kt).font = F_N
        ov.cell(row=r, column=3, value=f'=SUMIFS({C},{F},$B{r},{D},"Расход")').number_format = EUR
        ov.cell(row=r, column=4, value=f'=COUNTIFS({F},$B{r})').number_format = '0'
        ov.cell(row=r, column=5, value=f'=IF($C${first+1}=0,0,$C{r}/$C${first+1})').number_format = PCT
        for c in range(2, 6):
            ov.cell(row=r, column=c).border = BOX; ov.cell(row=r, column=c).fill = FILL_CALC
            ov.cell(row=r, column=c).font = F_N
        r += 1
    r += 1

    section(ov, r, "ПО МЕСЯЦАМ", 5, col=2); r += 1
    for i, h in enumerate(["Месяц", "Приход", "Расход", "Реальные траты"], start=2):
        c = ov.cell(row=r, column=i, value=h); c.font = F_B; c.fill = FILL_SUB; c.border = BOX
        c.alignment = Alignment(horizontal="center")
    r += 1
    y, m = 2026, 8
    for _ in range(18):
        key = f"{y}-{m:02d}"
        ov.cell(row=r, column=2, value=key).font = F_N
        ov.cell(row=r, column=3, value=f'=SUMIFS({C},{L},$B{r},{D},"Доход")').number_format = EUR
        ov.cell(row=r, column=4, value=f'=SUMIFS({C},{L},$B{r},{D},"Расход")').number_format = EUR
        ov.cell(row=r, column=5, value=f'=SUMIFS({M},{L},$B{r})').number_format = EUR
        for c in range(2, 6):
            ov.cell(row=r, column=c).border = BOX; ov.cell(row=r, column=c).fill = FILL_CALC
            ov.cell(row=r, column=c).font = F_N
        m += 1
        if m == 13:
            m = 1; y += 1
        r += 1

    # ---------------- ЗАЙМЫ ----------------
    zm = wb.create_sheet("Займы")
    zm.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", [2, 30, 16, 14, 40]):
        zm.column_dimensions[col].width = w
    title_row(zm, 1, "КТО МНЕ ДОЛЖЕН", 5)
    zm["B2"] = "Считается по колонке «Кому / от кого» в журнале. Вернули — поставь в журнале статус «Возвращено»."
    zm["B2"].font = F_S
    people = sorted({o.get("Кому / от кого", "") for o in ops
                     if o.get("Нужность") == "Возвратное" and o.get("Кому / от кого")})
    for i, h in enumerate(["Человек", "Ждёт возврата", "Операций", "Уже вернул"], start=2):
        c = zm.cell(row=4, column=i, value=h); c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
        c.alignment = Alignment(horizontal="center")
    r = 5
    for p in people:
        zm.cell(row=r, column=2, value=p).font = F_B
        zm.cell(row=r, column=3, value=f'=SUMIFS({C},{JK},$B{r},{H},"Возврат ожидается")').number_format = EUR
        zm.cell(row=r, column=4, value=f'=COUNTIFS({JK},$B{r},{H},"Возврат ожидается")').number_format = '0'
        zm.cell(row=r, column=5, value=f'=SUMIFS({C},{JK},$B{r},{H},"Возвращено")').number_format = EUR
        for c in range(2, 6):
            zm.cell(row=r, column=c).border = BOX; zm.cell(row=r, column=c).fill = FILL_CALC
            zm.cell(row=r, column=c).font = F_N
            zm.cell(row=r, column=c).alignment = Alignment(horizontal="center" if c > 2 else "left")
        r += 1
    zm.cell(row=r, column=2, value="ИТОГО").font = F_RES
    zm.cell(row=r, column=3, value=f"=SUM(C5:C{r-1})").number_format = EUR
    zm.cell(row=r, column=5, value=f"=SUM(E5:E{r-1})").number_format = EUR
    for c in range(2, 6):
        zm.cell(row=r, column=c).border = BOX; zm.cell(row=r, column=c).fill = FILL_RES
        zm.cell(row=r, column=c).font = F_RES
    zm.cell(row=r + 2, column=2,
            value="Это не расход, а дебиторка: пока не вернули, деньги твои только на бумаге.").font = F_S

    # ---------------- ПОСТОЯННЫЕ ----------------
    ps = wb.create_sheet("Постоянные")
    ps.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", [2, 34, 16, 12, 44]):
        ps.column_dimensions[col].width = w
    title_row(ps, 1, "ПОСТОЯННЫЕ ПЛАТЕЖИ — база, от которой считается всё остальное", 5)
    ps["B2"] = "Списываются сами каждый месяц. Это первое, что стоит пересмотреть, если нужно снизить расходы."
    ps["B2"].font = F_S
    fixed = sorted({o["Название"] for o in ops if o.get("Раздел") == "Постоянные"})
    for i, h in enumerate(["Платёж", "Всего списано", "Раз", "Заметка"], start=2):
        c = ps.cell(row=4, column=i, value=h); c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
        c.alignment = Alignment(horizontal="center")
    r = 5
    notes = {o["Название"]: o.get("Заметки", "") for o in ops if o.get("Раздел") == "Постоянные"}
    for nm in fixed:
        ps.cell(row=r, column=2, value=nm).font = F_B
        ps.cell(row=r, column=3, value=f'=SUMIFS({C},{JN},$B{r},{E},"Постоянные")').number_format = EUR
        ps.cell(row=r, column=4, value=f'=COUNTIFS({JN},$B{r},{E},"Постоянные")').number_format = '0'
        ps.cell(row=r, column=5, value=notes.get(nm, "")).font = F_S
        for c in range(2, 6):
            ps.cell(row=r, column=c).border = BOX; ps.cell(row=r, column=c).fill = FILL_CALC
            if c != 5:
                ps.cell(row=r, column=c).font = F_N if c == 2 else F_B
        r += 1
    ps.cell(row=r, column=2, value="ИТОГО").font = F_RES
    ps.cell(row=r, column=3, value=f"=SUM(C5:C{r-1})").number_format = EUR
    for c in range(2, 6):
        ps.cell(row=r, column=c).border = BOX; ps.cell(row=r, column=c).fill = FILL_RES
        ps.cell(row=r, column=c).font = F_RES

    # ---------------- ИНСТРУКЦИЯ ----------------
    ins = wb.create_sheet("Как это работает")
    ins.sheet_view.showGridLines = False
    for col, w in zip("ABC", [3, 34, 92]):
        ins.column_dimensions[col].width = w
    title_row(ins, 1, "КАК ПОЛЬЗОВАТЬСЯ ЖУРНАЛОМ", 3)
    blocks = [
        ("SEC", "ГДЕ ЧТО ЛЕЖИТ"),
        ("N", "Notion — «Журнал операций»", "Основное место. Заносишь траты с телефона, прикрепляешь фото чека. Есть виды: «Чеки (фото)», «Кто мне должен», «По разделам», «Разобраться»."),
        ("N", "Excel — этот файл", "Для расчётов и отчётов. Фото сюда положить нельзя — такой возможности в xlsx нет. Всё остальное считается формулами."),
        ("SP",),
        ("SEC", "КАК ОНИ СВЯЗАНЫ"),
        ("N", "Направление одно: Notion → Excel", "Notion — источник правды. Этот файл пересобирается из него. Так сделано специально: двусторонняя синхронизация между файлом и базой всегда рано или поздно даёт конфликты и потерю данных."),
        ("N", "Как обновить Excel", "Напиши Claude: «обнови журнал». Он выгрузит текущее состояние из Notion и пришлёт новый файл. Занимает меньше минуты."),
        ("N", "Если внёс что-то только в Excel", "Скажи об этом при обновлении — иначе новая выгрузка из Notion затрёт эти строки."),
        ("SP",),
        ("SEC", "КАК ЗАНОСИТЬ ОПЕРАЦИИ"),
        ("N", "Сумма", "Расход — со знаком минус (−8,88). Доход — плюс (2000). Знак определяет всё остальное."),
        ("N", "Поток", "Расход или Доход. По нему считаются итоги."),
        ("N", "Раздел", "Крупная группа: Займы, Семья, Постоянные, Повседневные, Транспорт, Покупки, Вопросы, Доходы."),
        ("N", "Категория", "Детально: продукты, самокаты, подписки и так далее."),
        ("N", "Нужность", "Обязательное / Необязательное / Возвратное / Неизвестно. «Возвратное» — это займы: они не попадают в реальные траты."),
        ("N", "Статус", "Оплачен / Открыт (ещё vorgemerkt) / Возврат ожидается / Возвращено."),
        ("N", "Кому / от кого", "Обязательно для займов — по этой колонке считается лист «Займы»."),
        ("SP",),
        ("SEC", "ПОЧЕМУ ЗАЙМЫ СЧИТАЮТСЯ ОТДЕЛЬНО"),
        ("N", "Займ — не расход", "Это дебиторка: деньги ушли, но остаются твоими. Колонка «Реальный расход» их обнуляет, поэтому в обзоре видно настоящее потребление, а не раздутую цифру."),
        ("N", "Когда вернули", "Меняешь статус на «Возвращено» и заводишь отдельную строку с плюсом (Поток = Доход). Тогда и долг закроется, и приход будет виден."),
        ("SP",),
        ("SEC", "ФОРМАТ"),
        ("N", "Открывается везде", "Excel, Google Sheets, LibreOffice, Numbers. Именованных диапазонов нет, экзотических функций нет — Numbers ничего не сломает."),
        ("N", "Что Numbers всё же уберёт", "Подсветку по условию и выпадающие списки. Формулы и цифры останутся."),
        ("N", "Свободные строки", "В журнале 500 строк, формулы в них уже стоят. Просто пиши в следующую пустую."),
    ]
    r = 3
    for b in blocks:
        if b[0] == "SEC":
            section(ins, r, b[1], 3); r += 1
        elif b[0] == "SP":
            r += 1
        else:
            ins.cell(row=r, column=2, value=b[1]).font = F_B
            c = ins.cell(row=r, column=3, value=b[2]); c.font = F_N
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ins.row_dimensions[r].height = max(15, 13 * (1 + len(b[2]) // 95))
            r += 1

    wb.active = 0
    wb.save(out)
    return len(ops)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    with open(a.data, encoding="utf-8") as f:
        ops = json.load(f)
    ops.sort(key=lambda o: (o["Дата"], o["Название"]))
    n = build(ops, a.out)
    print(f"OK: {n} операций -> {a.out}")
