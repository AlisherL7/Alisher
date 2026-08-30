# -*- coding: utf-8 -*-
"""
Gehaltsrechner / Расчёт зарплаты — Constantia Intensivpflege GmbH.
Совместим с Excel, LibreOffice и Apple Numbers:
  · без именованных диапазонов (Numbers их не поддерживает)
  · время смен — обычные числа-часы, не значения времени
  · ограниченные диапазоны вместо ссылок на весь столбец
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule

OUT = "/home/user/Alisher/Gehaltsrechner_Karimov.xlsx"

C_DARK, C_HEAD, C_SUB = "1F3864", "2E5FA3", "D6E4F7"
C_IN, C_CALC, C_RES, C_WARN = "FFF7D6", "EDEDED", "DFF3DF", "FFD6D6"

F_TITLE = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
F_H1    = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
F_H2    = Font(name="Calibri", size=11, bold=True, color=C_DARK)
F_B     = Font(name="Calibri", size=10, bold=True)
F_N     = Font(name="Calibri", size=10)
F_S     = Font(name="Calibri", size=9, color="606060")
F_RES   = Font(name="Calibri", size=12, bold=True, color="006100")
F_RED   = Font(name="Calibri", size=10, bold=True, color="9C0006")

FILL_TITLE = PatternFill("solid", fgColor=C_DARK)
FILL_HEAD  = PatternFill("solid", fgColor=C_HEAD)
FILL_SUB   = PatternFill("solid", fgColor=C_SUB)
FILL_IN    = PatternFill("solid", fgColor=C_IN)
FILL_CALC  = PatternFill("solid", fgColor=C_CALC)
FILL_RES   = PatternFill("solid", fgColor=C_RES)
FILL_WARN  = PatternFill("solid", fgColor=C_WARN)

thin = Side(style="thin", color="B0B0B0")
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)

FMT_EUR = '#,##0.00\\ "€"'
FMT_H   = '#,##0.00\\ "h"'
FMT_PCT = '0.00%'
FMT_D   = 'DD.MM.YYYY'
FMT_UHR = '0.00\\ "час"'

wb = Workbook()

def title_row(ws, row, text, span=6):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = F_TITLE; c.fill = FILL_TITLE
    c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
    ws.row_dimensions[row].height = 26

def section(ws, row, text, span=6):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = F_H1; c.fill = FILL_HEAD
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 19

# =====================================================================
# ЛИСТ 1 — ИНСТРУКЦИЯ
# =====================================================================
ws = wb.active
ws.title = "Инструкция"
ws.sheet_view.showGridLines = False
for col, w in zip("ABCDEF", [3, 40, 48, 30, 14, 14]):
    ws.column_dimensions[col].width = w

title_row(ws, 1, "РАСЧЁТ ЗАРПЛАТЫ · Alisher Karimov · Constantia Intensivpflege GmbH", 6)
ws["B3"] = "Источники: Arbeitsvertrag от 01.08.2026 (§§ 1–19), EzB от 13.07.2026, ArbZG, EStG (§ 3b, § 32a, § 39b), SvEV, BUrlG, EFZG, PflegeArbbV"
ws["B3"].font = F_S

rows = [
 ("SEC", "ЛИСТЫ ФАЙЛА"),
 ("N", "Настройки", "Все параметры: ставка, надбавки, время смен, лимит переработки, налоговый класс. Жёлтое — можно менять."),
 ("N", "Dienstplan", "График месяца: ND / TD / FREI, статус (Arbeit / Urlaub / Krank), «x» = Eingesprungen. Август 2026 заполнен."),
 ("N", "Расчёт", "Брутто, надбавки, счёт часов, соцвзносы, налоги, нетто."),
 ("N", "Lohnsteuer", "Подоходный налог по методу официального Programmablaufplan."),
 ("N", "Контроль ArbZG", "Проверка по закону о рабочем времени."),
 ("N", "Права и риски", "Что в договоре есть, чего нет, и что из этого следует."),
 ("N", "Feiertage", "Праздники Нижней Саксонии 2026/2027."),
 ("N", "Мои данные", "Пустые поля для SV-Nummer, Steuer-ID и т.д. Заполни у себя — файл лежит в ПУБЛИЧНОМ репозитории GitHub."),
 ("N", "Год", "Сводка по 12 месяцам и счёт переносимых часов."),
 ("SP",),
 ("SEC", "ВАЖНО ПРО ФОРМАТ"),
 ("N", "Время смен — это ЧИСЛА", "20 = 20:00, 8 = 08:00, 6,5 = 06:30. Так сделано специально: Numbers на iPhone/Mac ломает формулы с настоящими значениями времени."),
 ("N", "Numbers убирает подсветку", "Цветные предупреждения Numbers удаляет. Поэтому все предупреждения продублированы ТЕКСТОМ в последней колонке Dienstplan и на листе «Контроль ArbZG»."),
 ("N", "Если что-то показывает ошибку", "Открой в Excel, Google Sheets или LibreOffice — там работает всё. В Numbers формулы тоже работают, но подсветка и выпадающие списки пропадают."),
 ("SP",),
 ("SEC", "ЧТО ГОВОРИТ ДОГОВОР (§ 3 Vergütung)"),
 ("N", "Часовая ставка", "27,00 € брутто (§ 3 Abs. 1)"),
 ("N", "Оклад в месяц", "4.676,40 € брутто → норма 4.676,40 / 27,00 = 173,20 ч/мес"),
 ("N", "Рабочее время", "40 ч в неделю, пятидневка, БЕЗ учёта перерывов (§ 2 Abs. 1)"),
 ("N", "Ночь 20:00–06:00", "+25 % (§ 3 Abs. 2 a)"),
 ("N", "Воскресенье 00:00–24:00", "+50 % (§ 3 Abs. 2 b)"),
 ("N", "Праздник с отгулом (mit FZA)", "+35 % (§ 3 Abs. 2 c)"),
 ("N", "Праздник без отгула (ohne FZA)", "+135 % (§ 3 Abs. 2 d)"),
 ("N", "Совпадение надбавок", "Платится ТОЛЬКО САМАЯ ВЫСОКАЯ. Ночь в воскресенье = 50 %, не 75 %."),
 ("N", "Твои смены", "ND 20:00–08:00 = 12 ч, из них 10 ч ночных (20:00–06:00) и 2 ч обычных (06:00–08:00). TD 08:00–20:00 = 12 ч, ночных 0."),
 ("N", "Переработка", "§ 2 Abs. 3: первые 8 ч/мес включены в оклад. Работодатель на практике оплачивает максимум 60 ч/мес, остальное переносит в следующий месяц отгулами — в договоре этого правила НЕТ, см. лист «Права и риски»."),
 ("N", "Выплата", "15-го числа следующего месяца (§ 3 Abs. 3)"),
 ("N", "Срок претензий", "3 месяца — § 14 Ausschlussfrist. Пропустил — право пропало."),
 ("SP",),
 ("SEC", "НАЛОГИ И ВЗНОСЫ"),
 ("N", "§ 3b EStG", "Надбавки за ночь/воскресенье/праздник свободны от подоходного налога: ночь до 25 %, воскресенье до 50 %, праздник до 125 % (01.05, 25./26.12 и 24.12 с 14:00 — до 150 %). База — максимум 50 €/ч."),
 ("N", "§ 1 Abs. 1 SvEV", "От соцвзносов надбавки свободны только с базы 25,00 €/ч. Ставка 27,00 € → часть надбавки, посчитанная с 2,00 €/ч, взносами облагается. Файл делит автоматически."),
 ("N", "Отпуск и больничный", "Надбавки, выплаченные за дни отпуска/болезни, налогом ОБЛАГАЮТСЯ — § 3b требует фактической работы. Ставь статус Urlaub/Krank."),
 ("N", "Налоговый класс 3, 1 ребёнок", "Считается по Splittingtarif. Ребёнок с 2025 г. → доплата бездетных в Pflegeversicherung (0,6 %) больше НЕ удерживается — проверь это в Lohnabrechnung."),
 ("SP",),
 ("SEC", "ТОЧНОСТЬ"),
 ("N", "Часы, надбавки, брутто", "Точно — по формулам договора."),
 ("N", "Соцвзносы", "Точно — ставки и предельные базы 2026."),
 ("N", "Подоходный налог", "Оценка ±несколько евро. Классы 5 и 6 не поддерживаются."),
 ("N", "Нетто", "Оценка. Официальный документ — Lohnabrechnung. Этот файл нужен, чтобы её проверить."),
]
r = 5
for item in rows:
    if item[0] == "SEC":
        section(ws, r, item[1], 6); r += 1
    elif item[0] == "SP":
        r += 1
    else:
        ws.cell(row=r, column=2, value=item[1]).font = F_B
        c = ws.cell(row=r, column=3, value=item[2]); c.font = F_N
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        ws.row_dimensions[r].height = max(15, 13 * (1 + len(item[2]) // 100))
        r += 1
ws.cell(row=r + 1, column=2, value="Файл — вспомогательный инструмент, не юридическая консультация. При споре: профсоюз ver.di или Fachanwalt für Arbeitsrecht.").font = F_S

# =====================================================================
# ЛИСТ 2 — НАСТРОЙКИ  (прямые ссылки, без именованных диапазонов)
# =====================================================================
st = wb.create_sheet("Настройки")
st.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [2, 44, 36, 16, 66]):
    st.column_dimensions[col].width = w
title_row(st, 1, "НАСТРОЙКИ / EINSTELLUNGEN — жёлтые поля можно менять", 5)

P = {}
row_ptr = [3]

def sec(text):
    r = row_ptr[0]
    st.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    c = st.cell(row=r, column=2, value=text); c.font = F_H1; c.fill = FILL_HEAD
    c.alignment = Alignment(vertical="center", indent=1)
    st.row_dimensions[r].height = 19
    row_ptr[0] += 1

def par(ru, de, value, name=None, fmt=None, note="", inp=True):
    r = row_ptr[0]
    st.cell(row=r, column=2, value=ru).font = F_B
    st.cell(row=r, column=3, value=de).font = F_N
    c = st.cell(row=r, column=4, value=value)
    c.font = Font(name="Calibri", size=10, bold=True, color="7F3F00" if inp else "202020")
    c.fill = FILL_IN if inp else FILL_CALC
    c.border = BOX
    c.alignment = Alignment(horizontal="center")
    if fmt: c.number_format = fmt
    n = st.cell(row=r, column=5, value=note); n.font = F_S
    n.alignment = Alignment(wrap_text=True, vertical="center")
    if name: P[name] = f"Настройки!$D${r}"
    row_ptr[0] += 1
    return r

def spacer():
    row_ptr[0] += 1

sec("1. РАСЧЁТНЫЙ МЕСЯЦ")
par("Год", "Jahr", 2026, "Jahr", "0")
par("Месяц (1–12)", "Monat", 8, "Monat", "0")
r_tage = par("Дней в месяце", "Tage im Monat", "", "Tage", "0", inp=False)
par("Федеральная земля", "Bundesland", "Niedersachsen", "Land", None, "Праздники — с листа «Feiertage»")
spacer()

sec("2. ОПЛАТА — § 3 Arbeitsvertrag")
r_lohn = par("Часовая ставка", "Stundenlohn (Grundlohn)", 27.00, "Lohn", FMT_EUR, "§ 3 Abs. 1")
r_geh  = par("Оклад брутто в месяц", "Monatsvergütung brutto", 4676.40, "Gehalt", FMT_EUR, "§ 3 Abs. 1")
par("Плановые часы в месяц", "Sollstunden / Monat", "", "SollStd", FMT_H, "оклад / ставка = 173,20 ч", inp=False)
par("Часов в неделю", "Wochenarbeitszeit", 40, "WStd", '0.0\\ "h"', "§ 2 Abs. 1")
par("Переработка включена в оклад (ч)", "pauschal abgegoltene Überstunden", 8, "UeFrei", '0\\ "h"', "§ 2 Abs. 3 — до 8 ч/мес уже в окладе")
par("Максимум оплачиваемой переработки (ч)", "max. auszahlbare Überstunden", 60, "UeMax", '0\\ "h"',
    "Практика работодателя: не более 60 ч/мес. В ДОГОВОРЕ такого правила нет — см. лист «Права и риски»")
par("Оплачивать переработку?", "Überstunden auszahlen (ja/nein)", "ja", "UeZahlen", None, "«nein» = всё идёт в счёт часов отгулами")
par("Надбавка за переработку", "Überstundenzuschlag", 0.00, "UeZuschlag", FMT_PCT, "Договором не предусмотрена")
par("Перенос часов из прошлого месяца", "Übertrag aus Vormonat", 0.00, "UeVortrag", '0.00\\ "h"', "Остаток счёта часов (Arbeitszeitkonto) на начало месяца")
par("Премия за Eingesprungen (за смену)", "Einspringprämie je Dienst", 0.00, "Praemie", FMT_EUR,
    "В договоре НЕТ. Впиши сумму, о которой договорились. Требуй подтверждение по e-mail — см. «Права и риски»")
par("Фактор неполного месяца", "Faktor (Ein-/Austritt)", 1.00, "Faktor", '0.0000', "1,00 = полный месяц")
spacer()

sec("3. НАДБАВКИ — § 3 Abs. 2 Arbeitsvertrag")
par("Ночная надбавка", "Nachtzuschlag", 0.25, "ZNacht", FMT_PCT, "§ 3 Abs. 2 a)")
par("Ночь: с (час)", "Nachtarbeit von", 20, "NachtVon", "0", "20 = 20:00")
par("Ночь: до (час)", "Nachtarbeit bis", 6, "NachtBis", "0", "6 = 06:00")
par("Воскресная надбавка", "Sonntagszuschlag", 0.50, "ZSonn", FMT_PCT, "§ 3 Abs. 2 b)")
par("Праздник С отгулом", "Feiertag mit FZA", 0.35, "ZFmit", FMT_PCT, "§ 3 Abs. 2 c)")
par("Праздник БЕЗ отгула", "Feiertag ohne FZA", 1.35, "ZFohne", FMT_PCT, "§ 3 Abs. 2 d)")
par("Праздники по умолчанию", "Feiertagsmodus", "ohne FZA", "FZAStd", None, "«mit FZA» / «ohne FZA»; в графике можно задать по дню (колонка FZA: j / n)")
spacer()

sec("4. ЛЬГОТЫ ПО НАДБАВКАМ — § 3b EStG / § 1 SvEV")
par("Макс. необлагаемая ночная", "Höchstsatz Nacht steuerfrei", 0.25, "HNacht", FMT_PCT, "§ 3b Abs. 1 Nr. 1 EStG")
par("Макс. необлагаемая воскресная", "Höchstsatz Sonntag steuerfrei", 0.50, "HSonn", FMT_PCT, "§ 3b Abs. 1 Nr. 2 EStG")
par("Макс. необлагаемая праздничная", "Höchstsatz Feiertag steuerfrei", 1.25, "HFeier", FMT_PCT, "§ 3b Abs. 1 Nr. 3 EStG; для 01.05 и 25./26.12 — 150 % (в колонке D листа «Feiertage»)")
par("Предел базы: налог", "Grundlohngrenze steuerfrei", 50.00, "GLSteuer", FMT_EUR, "§ 3b Abs. 2 EStG")
par("Предел базы: соцвзносы", "Grundlohngrenze SV-frei", 25.00, "GLSV", FMT_EUR, "§ 1 Abs. 1 Nr. 1 SvEV — твои 27 € выше, часть надбавок облагается взносами")
spacer()

sec("5. ВРЕМЯ СМЕН — ЧИСЛА-ЧАСЫ (20 = 20:00, 8 = 08:00, 6,5 = 06:30)")
par("TD (Tagdienst) начало", "TD Beginn", 8, "TDvon", FMT_UHR, "08:00")
par("TD (Tagdienst) конец", "TD Ende", 20, "TDbis", FMT_UHR, "20:00 → 12 часов")
par("TD перерыв (часы)", "TD Pause", 0.00, "TDpause", '0.00\\ "h"', "§ 4 ArbZG при смене >9 ч требует 45 мин. Впиши 0,75 если работодатель вычитает перерыв")
par("ND (Nachtdienst) начало", "ND Beginn", 20, "NDvon", FMT_UHR, "20:00")
par("ND (Nachtdienst) конец", "ND Ende", 8, "NDbis", FMT_UHR, "08:00 следующего дня → 12 часов, из них 10 ночных")
par("ND перерыв (часы)", "ND Pause", 0.00, "NDpause", '0.00\\ "h"', "§ 4 ArbZG при смене >9 ч требует 45 мин")
spacer()

sec("6. НАЛОГИ И СОЦВЗНОСЫ 2026")
par("Налоговый класс", "Steuerklasse", 3, "StKl", "0", "Класс 3 → Splittingtarif. Классы 5 и 6 не поддерживаются")
par("Церковный налог?", "Kirchensteuerpflicht (ja/nein)", "nein", "KiSt", None, "")
par("Ставка церковного налога", "Kirchensteuersatz", 0.09, "KiStSatz", FMT_PCT, "Нижняя Саксония 9 %")
par("Детские вычеты (Zähler)", "Kinderfreibeträge", 1.0, "Kinder", "0.0", "Класс 3 → 1,0 на ребёнка. Влияет только на Soli и церковный налог")
par("Бездетный (доплата в PV)?", "kinderlos ab 23 (ja/nein)", "nein", "Kinderlos", None,
    "У тебя ребёнок 2025 г. → «nein». Доплата 0,6 % удерживаться НЕ должна — проверь в Lohnabrechnung")
par("Взнос KV общий", "KV allgemeiner Beitragssatz", 0.146, "KVSatz", FMT_PCT, "делится пополам")
par("Доп. взнос кассы", "kassenindiv. Zusatzbeitrag", 0.029, "KVZusatz", FMT_PCT, "средний 2026 — 2,9 %. Поставь ставку СВОЕЙ кассы")
par("Взнос PV", "Pflegeversicherung", 0.036, "PVSatz", FMT_PCT, "делится пополам")
par("Доплата бездетным PV", "Zuschlag Kinderlose", 0.006, "PVZuschlag", FMT_PCT, "применяется только если поле выше = «ja»")
par("Взнос RV", "Rentenversicherung", 0.186, "RVSatz", FMT_PCT, "→ 9,3 % с работника")
par("Взнос AV", "Arbeitslosenversicherung", 0.026, "AVSatz", FMT_PCT, "→ 1,3 % с работника")
par("Предел базы KV/PV в месяц", "BBG KV/PV monatlich", 5812.50, "BBGKV", FMT_EUR, "2026: 69.750 €/год")
par("Предел базы RV/AV в месяц", "BBG RV/AV monatlich", 8450.00, "BBGRV", FMT_EUR, "2026: 101.400 €/год")
spacer()

sec("7. ПАРАМЕТРЫ ПОДОХОДНОГО НАЛОГА 2026")
par("Вычет на работника", "Arbeitnehmer-Pauschbetrag", 1230.00, "ANPausch", FMT_EUR, "§ 9a EStG, в год")
par("Вычет особых расходов", "Sonderausgaben-Pauschbetrag", 36.00, "SAPausch", FMT_EUR, "§ 10c EStG, в год")
par("Вычет одинокому родителю", "Entlastungsbetrag Alleinerziehende", 4260.00, "EntlastAE", FMT_EUR, "только класс 2")
par("Учитывать AV в Vorsorgepauschale?", "AV-Teilbetrag ansetzen (ja/nein)", "nein", "AVinVP", None,
    "Новое с 2026 (§ 39b Abs. 2 S. 5 Nr. 3 Bst. e EStG), но только в пределах 1.900 € вместе с KV/PV — у тебя они уже выше")
par("Порог Soli (налог за год)", "Soli-Freigrenze", 20350.00, "SoliFrei", FMT_EUR, "одинокий; для класса 3 порог удваивается — при твоём доходе Soli всё равно 0")
par("Базовый вычет", "Grundfreibetrag 2026", 12348.00, "GFB", FMT_EUR, "§ 32a EStG")
par("Граница 1-й зоны", "Ende Zone 1", 17799.00, "Z1", FMT_EUR, "§ 32a EStG 2026")
par("Граница 2-й зоны", "Ende Zone 2", 69878.00, "Z2", FMT_EUR, "§ 32a EStG 2026")
par("Граница 3-й зоны", "Ende Zone 3", 277825.00, "Z3", FMT_EUR, "§ 32a EStG 2026")

# формулы, зависящие от других ячеек этого же листа
st.cell(row=r_tage, column=4, value=f"=DAY(EOMONTH(DATE({P['Jahr']},{P['Monat']},1),0))")
st.cell(row=int(P['SollStd'].rsplit('$', 1)[1]), column=4, value=f"={P['Gehalt']}/{P['Lohn']}")

dv1 = DataValidation(type="list", formula1='"ja,nein"', allow_blank=True)
st.add_data_validation(dv1)
for nm in ["UeZahlen", "KiSt", "Kinderlos", "AVinVP"]:
    dv1.add(st[P[nm].split("!")[1].replace("$", "")])
dv2 = DataValidation(type="list", formula1='"mit FZA,ohne FZA"', allow_blank=True)
st.add_data_validation(dv2)
dv2.add(st[P["FZAStd"].split("!")[1].replace("$", "")])

# =====================================================================
# ЛИСТ 3 — FEIERTAGE   (строки 4..43 — данные)
# =====================================================================
ft = wb.create_sheet("Feiertage")
ft.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [14, 34, 22, 20, 50]):
    ft.column_dimensions[col].width = w
title_row(ft, 1, "ПРАЗДНИКИ / GESETZLICHE FEIERTAGE — Niedersachsen", 5)
ft["A2"] = "Дописывать можно до строки 43. Колонка D — максимальная НЕОБЛАГАЕМАЯ надбавка по § 3b EStG."
ft["A2"].font = F_S
for i, h in enumerate(["Дата", "Название (Feiertag)", "Земля", "§3b Höchstsatz", "Примечание"], start=1):
    c = ft.cell(row=3, column=i, value=h); c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

FEIERTAGE = [
 (dt.date(2026,1,1),  "Neujahr",                   "bundesweit",    1.25, ""),
 (dt.date(2026,4,3),  "Karfreitag",                "bundesweit",    1.25, ""),
 (dt.date(2026,4,6),  "Ostermontag",               "bundesweit",    1.25, ""),
 (dt.date(2026,5,1),  "Tag der Arbeit",            "bundesweit",    1.50, "§ 3b Abs. 3 Nr. 3 — до 150 %"),
 (dt.date(2026,5,14), "Christi Himmelfahrt",       "bundesweit",    1.25, ""),
 (dt.date(2026,5,25), "Pfingstmontag",             "bundesweit",    1.25, ""),
 (dt.date(2026,10,3), "Tag der Deutschen Einheit", "bundesweit",    1.25, ""),
 (dt.date(2026,10,31),"Reformationstag",           "Niedersachsen", 1.25, ""),
 (dt.date(2026,12,25),"1. Weihnachtstag",          "bundesweit",    1.50, "§ 3b Abs. 3 Nr. 3 — до 150 %"),
 (dt.date(2026,12,26),"2. Weihnachtstag",          "bundesweit",    1.50, "§ 3b Abs. 3 Nr. 3 — до 150 %"),
 (dt.date(2027,1,1),  "Neujahr",                   "bundesweit",    1.25, ""),
 (dt.date(2027,3,26), "Karfreitag",                "bundesweit",    1.25, ""),
 (dt.date(2027,3,29), "Ostermontag",               "bundesweit",    1.25, ""),
 (dt.date(2027,5,1),  "Tag der Arbeit",            "bundesweit",    1.50, "суббота"),
 (dt.date(2027,5,6),  "Christi Himmelfahrt",       "bundesweit",    1.25, ""),
 (dt.date(2027,5,17), "Pfingstmontag",             "bundesweit",    1.25, ""),
 (dt.date(2027,10,3), "Tag der Deutschen Einheit", "bundesweit",    1.25, "воскресенье"),
 (dt.date(2027,10,31),"Reformationstag",           "Niedersachsen", 1.25, "воскресенье"),
 (dt.date(2027,12,25),"1. Weihnachtstag",          "bundesweit",    1.50, "суббота"),
 (dt.date(2027,12,26),"2. Weihnachtstag",          "bundesweit",    1.50, "воскресенье"),
]
for i, (d, name, land, hs, note) in enumerate(FEIERTAGE):
    r = 4 + i
    ft.cell(row=r, column=1, value=d).number_format = FMT_D
    ft.cell(row=r, column=2, value=name)
    ft.cell(row=r, column=3, value=land)
    ft.cell(row=r, column=4, value=hs).number_format = FMT_PCT
    ft.cell(row=r, column=5, value=note)
for r in range(4, 44):
    for cc in range(1, 6):
        ft.cell(row=r, column=cc).border = BOX
        ft.cell(row=r, column=cc).font = F_S if cc == 5 else F_N
        if r >= 4 + len(FEIERTAGE):
            ft.cell(row=r, column=cc).fill = FILL_IN

FT_D = "Feiertage!$A$4:$A$43"
FT_N = "Feiertage!$B$4:$B$43"
FT_H = "Feiertage!$D$4:$D$43"

r = 46
ft.cell(row=r, column=1, value="НЕ праздники, но льгота по § 3b EStG:").font = F_H2
r += 1
for txt in ["· 24.12 с 14:00 — до 150 % необлагаемо",
            "· 31.12 с 14:00 — до 125 % необлагаемо",
            "· ночь 00:00–04:00, если смена начата до полуночи — до 40 % необлагаемо (§ 3b Abs. 3 Nr. 1)",
            "· договор Constantia платит за эти часы обычные ставки — льгота касается только налога, не суммы выплаты"]:
    ft.cell(row=r, column=1, value=txt).font = F_S
    ft.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1

# =====================================================================
# ЛИСТ 4 — DIENSTPLAN
# =====================================================================
dp = wb.create_sheet("Dienstplan")
dp.sheet_view.showGridLines = False
title_row(dp, 1, "ГРАФИК И РАСЧЁТ ЧАСОВ / DIENSTPLAN", 21)
dp["A2"] = ("ND = Nachtdienst 20:00–08:00 · TD = Tagdienst 08:00–20:00 · FREI = выходной · «x» = eingesprungen · "
            "статус: Arbeit / Urlaub / Krank / Fortbildung · FZA: j = праздник с отгулом, n = без отгула. "
            "Время — числа-часы (20 = 20:00). Смена относится к дню, в который НАЧАЛАСЬ.")
dp["A2"].font = F_S
dp.merge_cells("A2:U2")

HEAD = [
 ("Дата\nDatum", 11), ("День", 6), ("Праздник\nFeiertag", 20),
 ("Смена\nND/TD/FREI", 11), ("Статус\nStatus", 11), ("Sprn", 6), ("FZA", 6),
 ("Начало\nчас", 8), ("Конец\nчас", 8), ("Пауза\nч", 7),
 ("Длит.\nч", 7), ("Часы\nоплач.", 8),
 ("Ночь\nч", 7), ("Воскр.\nч", 7), ("Праздн.\nч", 8), ("Обычн.\nч", 8),
 ("Надбавка\n€", 11), ("из них\nбез налога", 11), ("из них\nбез взносов", 12),
 ("Отдых\nдо смены ч", 11), ("Предупреждение (ArbZG)", 36),
]
for i, (h, w) in enumerate(HEAD, start=1):
    c = dp.cell(row=4, column=i, value=h)
    c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    dp.column_dimensions[get_column_letter(i)].width = w
dp.row_dimensions[4].height = 34

R0, R1 = 5, 35
TR = 36

PLAN = {1:"FREI",2:"ND",3:"ND",4:"ND",5:"ND",6:"FREI",7:"FREI",8:"FREI",9:"ND",10:"ND",
        11:"ND",12:"ND",13:"FREI",14:"FREI",15:"ND",16:"ND",17:"FREI",18:"FREI",19:"TD",20:"TD",
        21:"FREI",22:"TD",23:"TD",24:"ND",25:"ND",26:"ND",27:"FREI",28:"ND",29:"ND",30:"ND",31:"ND"}
SPRN = {19, 20, 28}

for i, r in enumerate(range(R0, R1 + 1)):
    n = i + 1
    dienst = PLAN.get(n, "")
    dp.cell(row=r, column=1, value=f'=IF({n}>{P["Tage"]},"",DATE({P["Jahr"]},{P["Monat"]},{n}))').number_format = FMT_D
    dp.cell(row=r, column=2, value=f'=IF($A{r}="","",CHOOSE(WEEKDAY($A{r},2),"Пн","Вт","Ср","Чт","Пт","Сб","Вс"))')
    dp.cell(row=r, column=3, value=f'=IF($A{r}="","",IFERROR(INDEX({FT_N},MATCH($A{r},{FT_D},0)),""))')
    dp.cell(row=r, column=4, value=dienst)
    dp.cell(row=r, column=5, value=("Arbeit" if dienst in ("ND", "TD") else ""))
    dp.cell(row=r, column=6, value=("x" if n in SPRN else ""))
    dp.cell(row=r, column=7, value="")
    dp.cell(row=r, column=8, value=f'=IF($D{r}="ND",{P["NDvon"]},IF($D{r}="TD",{P["TDvon"]},""))').number_format = '0.00'
    dp.cell(row=r, column=9, value=f'=IF($D{r}="ND",{P["NDbis"]},IF($D{r}="TD",{P["TDbis"]},""))').number_format = '0.00'
    dp.cell(row=r, column=10, value=f'=IF($D{r}="ND",{P["NDpause"]},IF($D{r}="TD",{P["TDpause"]},""))').number_format = '0.00'
    dp.cell(row=r, column=11, value=f'=IF($H{r}="","",$I{r}-$H{r}+IF($I{r}<=$H{r},24,0))').number_format = '0.00'
    dp.cell(row=r, column=12, value=f'=IF($K{r}="","",MAX(0,$K{r}-$J{r}))').number_format = '0.00'
    # служебные колонки W..AO
    dp.cell(row=r, column=23, value=f'=IF($H{r}="","",$H{r})')                                             # W начало
    dp.cell(row=r, column=24, value=f'=IF($H{r}="","",$I{r}+IF($I{r}<=$H{r},24,0))')                       # X конец
    dp.cell(row=r, column=25, value=f'=IF(OR($K{r}="",$K{r}=0),0,$L{r}/$K{r})')                            # Y пауза-фактор
    dp.cell(row=r, column=26, value=f'=IF($A{r}="",0,IF(WEEKDAY($A{r},2)=7,1,0))')                         # Z вс сегодня
    dp.cell(row=r, column=27, value=f'=IF($A{r}="",0,IF(WEEKDAY($A{r}+1,2)=7,1,0))')                       # AA вс завтра
    dp.cell(row=r, column=28, value=f'=IF($A{r}="",0,MIN(1,COUNTIF({FT_D},$A{r})))')                       # AB праздник сегодня
    dp.cell(row=r, column=29, value=f'=IF($A{r}="",0,MIN(1,COUNTIF({FT_D},$A{r}+1)))')                     # AC праздник завтра
    dp.cell(row=r, column=30, value=f'=IFERROR(INDEX({FT_H},MATCH($A{r},{FT_D},0)),{P["HFeier"]})')        # AD
    dp.cell(row=r, column=31, value=f'=IFERROR(INDEX({FT_H},MATCH($A{r}+1,{FT_D},0)),{P["HFeier"]})')      # AE
    dp.cell(row=r, column=32, value=(f'=IF($W{r}="",0,MAX(0,MIN($X{r},{P["NachtBis"]})-MAX($W{r},0))'
                                     f'+MAX(0,MIN($X{r},24)-MAX($W{r},{P["NachtVon"]})))'))               # AF ночь сегодня
    dp.cell(row=r, column=33, value=f'=IF($W{r}="",0,MAX(0,MIN($X{r},24+{P["NachtBis"]})-MAX($W{r},24)))') # AG ночь завтра
    dp.cell(row=r, column=34, value=f'=IF($W{r}="",0,MAX(0,MIN($X{r},24)-MAX($W{r},0)))')                  # AH часы сегодня
    dp.cell(row=r, column=35, value=f'=IF($W{r}="",0,MAX(0,MIN($X{r},48)-MAX($W{r},24)))')                 # AI часы завтра
    dp.cell(row=r, column=36, value=f'=($Z{r}*$AB{r}*$AH{r}+$AA{r}*$AC{r}*$AI{r})*$Y{r}')                  # AJ вс И праздник
    dp.cell(row=r, column=37, value=f'=$O{r}-$AJ{r}')                                                      # AK только праздник
    dp.cell(row=r, column=38, value=f'=$N{r}-$AJ{r}')                                                      # AL только вс
    dp.cell(row=r, column=39, value=f'=$M{r}-(MAX($Z{r},$AB{r})*$AF{r}+MAX($AA{r},$AC{r})*$AG{r})*$Y{r}')  # AM только ночь
    dp.cell(row=r, column=40, value=(f'=IF($G{r}="j",{P["ZFmit"]},IF($G{r}="n",{P["ZFohne"]},'
                                     f'IF({P["FZAStd"]}="mit FZA",{P["ZFmit"]},{P["ZFohne"]})))'))         # AN ставка праздника
    dp.cell(row=r, column=41, value=(f'=IF($O{r}=0,{P["HFeier"]},'
                                     f'($AB{r}*$AH{r}*$AD{r}+$AC{r}*$AI{r}*$AE{r})*$Y{r}/$O{r})'))        # AO предел праздника
    # видимые результаты
    dp.cell(row=r, column=13, value=f'=($AF{r}+$AG{r})*$Y{r}').number_format = '0.00'
    dp.cell(row=r, column=14, value=f'=($Z{r}*$AH{r}+$AA{r}*$AI{r})*$Y{r}').number_format = '0.00'
    dp.cell(row=r, column=15, value=f'=($AB{r}*$AH{r}+$AC{r}*$AI{r})*$Y{r}').number_format = '0.00'
    dp.cell(row=r, column=16, value=f'=IF($L{r}="","",MAX(0,$L{r}-$O{r}-$AL{r}-$AM{r}))').number_format = '0.00'
    dp.cell(row=r, column=17, value=(f'=ROUND({P["Lohn"]}*($AK{r}*$AN{r}+$AJ{r}*MAX($AN{r},{P["ZSonn"]})'
                                     f'+$AL{r}*{P["ZSonn"]}+$AM{r}*{P["ZNacht"]}),2)')).number_format = FMT_EUR
    for col, base in ((18, P["GLSteuer"]), (19, P["GLSV"])):
        dp.cell(row=r, column=col, value=(
            f'=IF($E{r}<>"Arbeit",0,ROUND(MIN({P["Lohn"]},{base})*'
            f'($AK{r}*MIN($AN{r},$AO{r})+$AJ{r}*MIN(MAX($AN{r},{P["ZSonn"]}),$AO{r})'
            f'+$AL{r}*MIN({P["ZSonn"]},{P["HSonn"]})+$AM{r}*MIN({P["ZNacht"]},{P["HNacht"]})),2))'
        )).number_format = FMT_EUR
    if r == R0:
        dp.cell(row=r, column=20, value='')
    else:
        dp.cell(row=r, column=20, value=f'=IF(OR($W{r}="",$W{r-1}=""),"",24+$W{r}-$X{r-1})').number_format = '0.0'
    dp.cell(row=r, column=21, value=(
        f'=IF($L{r}="","",TRIM(IF($L{r}>10,"смена >10 ч (§ 3 / § 6 ArbZG)  ","")'
        f'&IF(AND($T{r}<>"",$T{r}<11),"отдых <11 ч (§ 5 ArbZG)  ","")'
        f'&IF(AND($L{r}>9,$J{r}<0.75),"перерыв <45 мин (§ 4 ArbZG)  ","")'
        f'&IF(AND($L{r}>6,$L{r}<=9,$J{r}<0.5),"перерыв <30 мин (§ 4 ArbZG)","")))'))

for r in range(R0, R1 + 1):
    for c in range(1, 22):
        cell = dp.cell(row=r, column=c)
        cell.border = BOX; cell.font = F_N
        if c in (4, 5, 6, 7, 8, 9, 10):
            cell.fill = FILL_IN; cell.alignment = Alignment(horizontal="center")
        elif c in (17, 18, 19):
            cell.fill = FILL_RES; cell.alignment = Alignment(horizontal="center")
        elif c == 21:
            cell.font = F_RED
            cell.alignment = Alignment(horizontal="left")
        else:
            cell.fill = FILL_CALC; cell.alignment = Alignment(horizontal="center")

dp.cell(row=TR, column=1, value="ИТОГО ЗА МЕСЯЦ")
dp.merge_cells(start_row=TR, start_column=1, end_row=TR, end_column=5)
dp.cell(row=TR, column=6, value=f'=COUNTIF(F{R0}:F{R1},"x")').number_format = '0\\ "x"'
for c, fmt in [(11,'0.00'),(12,'0.00'),(13,'0.00'),(14,'0.00'),(15,'0.00'),(16,'0.00'),
               (17,FMT_EUR),(18,FMT_EUR),(19,FMT_EUR)]:
    cl = get_column_letter(c)
    dp.cell(row=TR, column=c, value=f'=SUM({cl}{R0}:{cl}{R1})').number_format = fmt
for c in range(1, 22):
    cc = dp.cell(row=TR, column=c)
    if c <= 5:
        cc.font = F_H1; cc.fill = FILL_HEAD
        cc.alignment = Alignment(indent=1, vertical="center")
    else:
        cc.font = Font(name="Calibri", size=11, bold=True, color=C_DARK)
        cc.fill = FILL_SUB
        cc.alignment = Alignment(horizontal="center", vertical="center")
    cc.border = BOX
dp.row_dimensions[TR].height = 22

for dvspec, col in [('"ND,TD,FREI"', "D"), ('"Arbeit,Urlaub,Krank,Fortbildung"', "E"),
                    ('"x"', "F"), ('"j,n"', "G")]:
    dv = DataValidation(type="list", formula1=dvspec, allow_blank=True)
    dp.add_data_validation(dv)
    dv.add(f"{col}{R0}:{col}{R1}")

dp.conditional_formatting.add(f"U{R0}:U{R1}",
    FormulaRule(formula=[f'LEN($U{R0})>0'], fill=FILL_WARN, font=F_RED))
dp.conditional_formatting.add(f"A{R0}:C{R1}",
    FormulaRule(formula=[f'WEEKDAY($A{R0},2)>5'], fill=PatternFill("solid", fgColor="FCE4D6")))

for c in range(22, 42):
    dp.column_dimensions[get_column_letter(c)].hidden = True
dp.freeze_panes = "D5"

# =====================================================================
# ЛИСТ 5 — РАСЧЁТ
# =====================================================================
ab = wb.create_sheet("Расчёт")
ab.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [2, 48, 40, 16, 62]):
    ab.column_dimensions[col].width = w
title_row(ab, 1, "РАСЧЁТ ЗАРПЛАТЫ ЗА МЕСЯЦ / GEHALTSABRECHNUNG", 5)
MONNAMES = '"Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"'
ab["B2"] = f'=CONCATENATE("Месяц: ",CHOOSE({P["Monat"]},{MONNAMES})," ",{P["Jahr"]},"   ·   Alisher Karimov   ·   Constantia Intensivpflege GmbH")'
ab["B2"].font = F_H2
ab.merge_cells("B2:E2")

AR = {}
ptr = [4]

def asec(text):
    r = ptr[0]
    ab.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    c = ab.cell(row=r, column=2, value=text); c.font = F_H1; c.fill = FILL_HEAD
    c.alignment = Alignment(vertical="center", indent=1)
    ab.row_dimensions[r].height = 19
    ptr[0] += 1

def aline(ru, de, formula, key=None, fmt=FMT_EUR, note="", big=False, warn=False):
    r = ptr[0]
    ab.cell(row=r, column=2, value=ru).font = F_RES if big else F_B
    ab.cell(row=r, column=3, value=de).font = F_N
    c3 = ab.cell(row=r, column=4, value=formula)
    if fmt: c3.number_format = fmt
    c3.font = F_RES if big else (F_RED if warn else Font(name="Calibri", size=10, bold=True))
    c3.fill = FILL_RES if big else (FILL_WARN if warn else FILL_CALC)
    c3.border = BOX
    c3.alignment = Alignment(horizontal="center")
    c4 = ab.cell(row=r, column=5, value=note); c4.font = F_S
    c4.alignment = Alignment(wrap_text=True, vertical="center")
    if big: ab.row_dimensions[r].height = 22
    if key: AR[key] = r
    ptr[0] += 1
    return r

def aspace():
    ptr[0] += 1

D = lambda k: f"D{AR[k]}"

asec("A. ЧАСЫ / ARBEITSZEIT")
aline("Отработано часов (факт)", "Ist-Stunden", f"=Dienstplan!L{TR}", "ist", '0.00', "Сумма с листа Dienstplan, включая дни отпуска и болезни")
aline("Плановая норма месяца", "Sollstunden", f"=ROUND({P['SollStd']}*{P['Faktor']},2)", "soll", '0.00', "оклад / ставка = 173,20 ч")
aline("Разница за месяц", "Differenz", f"={D('ist')}-{D('soll')}", "diff", '0.00', "«+» переработка, «−» недоработка (оклад платится полностью)")
aline("Ночных часов", "Nachtstunden", f"=Dienstplan!M{TR}", "nh", '0.00', "20:00–06:00")
aline("Воскресных часов", "Sonntagsstunden", f"=Dienstplan!N{TR}", "sh", '0.00', "")
aline("Праздничных часов", "Feiertagsstunden", f"=Dienstplan!O{TR}", "fh", '0.00', "")
aline("Смен «Eingesprungen»", "Einspringdienste", f"=Dienstplan!F{TR}", "sprn", '0', "")
aspace()

asec("B. СЧЁТ ЧАСОВ / ARBEITSZEITKONTO")
aline("Перенос из прошлого месяца", "Übertrag aus Vormonat", f"={P['UeVortrag']}", "vor", '0.00', "Впиши на листе «Настройки»")
aline("Счёт часов до выплаты", "Stundenkonto vor Auszahlung", f"={D('vor')}+{D('diff')}", "konto", '0.00', "")
aline("Пропадает: включено в оклад", "pauschal abgegolten (§ 2 Abs. 3)", f"=MIN(MAX(0,{D('diff')}),{P['UeFrei']})", "abg", '0.00', "8 ч/мес — единственное, что договор разрешает не оплачивать")
aline("Доступно к оплате", "auszahlbar", f"=MAX(0,{D('konto')}-{D('abg')})", "verf", '0.00', "Отрицательный счёт к оплате не даёт — оклад всё равно выплачивается полностью")
aline("Оплачено в этом месяце", "ausgezahlte Überstunden", f'=IF({P["UeZahlen"]}="ja",MAX(0,MIN({D("verf")},{P["UeMax"]})),0)', "ueh", '0.00', "ограничение 60 ч/мес — практика работодателя")
aline("ПЕРЕНОС НА СЛЕДУЮЩИЙ МЕСЯЦ", "Übertrag in den Folgemonat", f"={D('konto')}-{D('abg')}-{D('ueh')}", "konto2", '0.00',
      "Эту цифру впиши в «Настройки» следующего месяца. Минус = отработано меньше нормы, часы уходят в долг. И требуй письменного подтверждения остатка!", warn=True)
aline("Стоимость непогашенных часов", "Wert des Stundenkontos", f"=ROUND({D('konto2')}*{P['Lohn']},2)", "kwert", FMT_EUR,
      "Деньги, которые работодатель пока не выплатил. Не потеряй их из-за § 14 Ausschlussfrist", warn=True)
aline("Оплачено бы БЕЗ лимита 60 ч", "ohne 60-Stunden-Deckel", f"=ROUND({D('verf')}*{P['Lohn']},2)", "ohnedeckel", FMT_EUR,
      "Столько работодатель должен был бы выплатить в этом месяце по § 3 Abs. 1, если бы лимита не было")
aspace()

asec("C. БРУТТО / BRUTTOENTGELT")
aline("Оклад", "Grundvergütung § 3 Abs. 1", f"=ROUND({P['Gehalt']}*{P['Faktor']},2)", "grund", FMT_EUR, "4.676,40 € при полном месяце")
aline("Оплата переработки", "Überstundenvergütung", f"=ROUND({D('ueh')}*{P['Lohn']}*(1+{P['UeZuschlag']}),2)", "uev", FMT_EUR, "часы × 27,00 €")
aline("Надбавки всего", "Zeitzuschläge gesamt", f"=Dienstplan!Q{TR}", "zus", FMT_EUR, "начисляются на ВСЕ отработанные часы, включая переработку")
aline("   из них без подоходного налога", "   steuerfrei § 3b EStG", f"=Dienstplan!R{TR}", "zusstf", FMT_EUR, "")
aline("   из них без соцвзносов", "   SV-frei § 1 SvEV", f"=Dienstplan!S{TR}", "zussvf", FMT_EUR, "меньше, т.к. база ограничена 25,00 €/ч")
aline("   облагается налогом", "   steuerpflichtiger Anteil", f"={D('zus')}-{D('zusstf')}", "zusstp", FMT_EUR, "")
aline("   облагается взносами", "   SV-pflichtiger Anteil", f"={D('zus')}-{D('zussvf')}", "zussvp", FMT_EUR, "")
aline("Премии за Eingesprungen", "Einspringprämien", f"=ROUND({D('sprn')}*{P['Praemie']},2)", "prae", FMT_EUR, "облагается полностью; сумму задай в «Настройках»")
aline("БРУТТО ВСЕГО", "GESAMTBRUTTO", f"={D('grund')}+{D('uev')}+{D('zus')}+{D('prae')}", "brutto", FMT_EUR, "", big=True)
aline("Налогооблагаемое брутто", "steuerpflichtiges Brutto", f"={D('grund')}+{D('uev')}+{D('prae')}+{D('zusstp')}", "stbrutto", FMT_EUR, "база подоходного налога")
aline("База для соцвзносов", "SV-Brutto", f"={D('grund')}+{D('uev')}+{D('prae')}+{D('zussvp')}", "svbrutto", FMT_EUR, "база KV/PV/RV/AV")
aspace()

asec("D. СОЦВЗНОСЫ (доля работника)")
aline("Медицинское страхование", "Krankenversicherung", f"=ROUND(MIN({D('svbrutto')},{P['BBGKV']})*({P['KVSatz']}/2+{P['KVZusatz']}/2),2)", "kv", FMT_EUR, "7,3 % + половина доп. взноса")
aline("Страхование по уходу", "Pflegeversicherung", f'=ROUND(MIN({D("svbrutto")},{P["BBGKV"]})*({P["PVSatz"]}/2+IF({P["Kinderlos"]}="ja",{P["PVZuschlag"]},0)),2)', "pv", FMT_EUR, "1,8 %; доплата бездетных не применяется — есть ребёнок")
aline("Пенсионное страхование", "Rentenversicherung", f"=ROUND(MIN({D('svbrutto')},{P['BBGRV']})*{P['RVSatz']}/2,2)", "rv", FMT_EUR, "9,3 %")
aline("Страхование по безработице", "Arbeitslosenversicherung", f"=ROUND(MIN({D('svbrutto')},{P['BBGRV']})*{P['AVSatz']}/2,2)", "av", FMT_EUR, "1,3 %")
aline("СОЦВЗНОСЫ ВСЕГО", "SV gesamt", f"={D('kv')}+{D('pv')}+{D('rv')}+{D('av')}", "svges", FMT_EUR, "")
aspace()

asec("E. НАЛОГИ")
aline("Подоходный налог", "Lohnsteuer", "=0", "lst", FMT_EUR, "налоговый класс 3 → Splittingtarif")
aline("Надбавка солидарности", "Solidaritätszuschlag", "=0", "soli", FMT_EUR, "при твоём доходе = 0")
aline("Церковный налог", "Kirchensteuer", "=0", "kist", FMT_EUR, "")
aline("НАЛОГИ ВСЕГО", "Steuern gesamt", f"={D('lst')}+{D('soli')}+{D('kist')}", "stges", FMT_EUR, "")
aspace()

asec("F. НА РУКИ")
aline("НЕТТО (на счёт)", "NETTOVERDIENST", f"={D('brutto')}-{D('svges')}-{D('stges')}", "netto", FMT_EUR,
      "Выплата до 15-го числа следующего месяца (§ 3 Abs. 3)", big=True)
aline("Нетто за час фактической работы", "Nettostundenlohn", f'=IF({D("ist")}=0,0,ROUND({D("netto")}/{D("ist")},2))', "nsl", FMT_EUR, "")
aline("Доля отчислений", "Abzugsquote", f"=IF({D('brutto')}=0,0,({D('svges')}+{D('stges')})/{D('brutto')})", "qu", FMT_PCT, "")
aspace()

asec("G. СПРАВОЧНО: расходы работодателя")
aline("KV работодатель", "AG-Anteil KV", f"=ROUND(MIN({D('svbrutto')},{P['BBGKV']})*({P['KVSatz']}/2+{P['KVZusatz']}/2),2)", "agkv", FMT_EUR, "")
aline("PV работодатель", "AG-Anteil PV", f"=ROUND(MIN({D('svbrutto')},{P['BBGKV']})*{P['PVSatz']}/2,2)", "agpv", FMT_EUR, "")
aline("RV работодатель", "AG-Anteil RV", f"=ROUND(MIN({D('svbrutto')},{P['BBGRV']})*{P['RVSatz']}/2,2)", "agrv", FMT_EUR, "")
aline("AV работодатель", "AG-Anteil AV", f"=ROUND(MIN({D('svbrutto')},{P['BBGRV']})*{P['AVSatz']}/2,2)", "agav", FMT_EUR, "")
aline("Всего расходы работодателя", "Gesamtkosten Arbeitgeber", f"={D('brutto')}+{D('agkv')}+{D('agpv')}+{D('agrv')}+{D('agav')}", "agges", FMT_EUR, "без U1/U2 и страхования от несчастных случаев")
aspace()

asec("H. ПРОВЕРКА ПО ЗАКОНУ")
aline("Ср. часов в неделю", "Ø Wochenarbeitszeit", f"=ROUND({D('ist')}/({P['Tage']}/7),2)", "wk", '0.00',
      "§ 3 ArbZG: в среднем не более 48 ч/нед за 6 месяцев")
aline("Смен длиннее 10 ч", "Dienste > 10 h", f'=COUNTIF(Dienstplan!L{R0}:L{R1},">10")', "l10", '0',
      "§ 3 ArbZG: максимум 10 ч. Для ночных § 6 Abs. 2 строже — 8 ч, до 10 ч с компенсацией за месяц")
aline("Ночных смен длиннее 8 ч", "Nachtdienste > 8 h", f'=COUNTIFS(Dienstplan!D{R0}:D{R1},"ND",Dienstplan!L{R0}:L{R1},">8")', "nd8", '0',
      "§ 6 Abs. 2 ArbZG — самое серьёзное нарушение в твоём графике")
aline("Нарушений отдыха 11 ч", "Ruhezeit < 11 h", f'=COUNTIFS(Dienstplan!T{R0}:T{R1},"<11",Dienstplan!T{R0}:T{R1},">0")', "ru", '0',
      "§ 5 ArbZG; в уходе допустимо 10 ч с компенсацией в течение месяца")
aline("Смен, задевающих воскресенье", "Sonntagsarbeit", f'=COUNTIF(Dienstplan!N{R0}:N{R1},">0")', "so", '0',
      "§ 11 ArbZG: минимум 15 воскресений в году свободны; Ersatzruhetag в течение 2 недель")
aline("Ставка ≥ Pflegemindestlohn?", "Pflegemindestlohn 21,03 €", f'=IF({P["Lohn"]}>=21.03,"да / ja","НЕТ — проверить!")', "pml", None,
      "7. PflegeArbbV, Pflegefachkraft с 01.07.2026")

for r in range(4, ptr[0]):
    ab.cell(row=r, column=2).alignment = Alignment(indent=1, vertical="center", wrap_text=True)

# =====================================================================
# ЛИСТ 6 — LOHNSTEUER
# =====================================================================
ls = wb.create_sheet("Lohnsteuer")
ls.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGH", [2, 44, 40, 16, 56, 3, 14, 18]):
    ls.column_dimensions[col].width = w
title_row(ls, 1, "ПОДОХОДНЫЙ НАЛОГ / LOHNSTEUER 2026 — налоговый класс 3 (оценка)", 5)
ls["B2"] = "Метод официального Programmablaufplan: годовой прогноз → вычеты → тариф § 32a EStG → делим на 12."
ls["B2"].font = F_S
ls.merge_cells("B2:E2")

LR = {}
lptr = [4]

def lsec(text):
    r = lptr[0]
    ls.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    c = ls.cell(row=r, column=2, value=text); c.font = F_H1; c.fill = FILL_HEAD
    c.alignment = Alignment(vertical="center", indent=1)
    lptr[0] += 1

def lline(ru, de, formula, key=None, fmt=FMT_EUR, note="", big=False):
    r = lptr[0]
    ls.cell(row=r, column=2, value=ru).font = F_RES if big else F_B
    ls.cell(row=r, column=3, value=de).font = F_N
    c = ls.cell(row=r, column=4, value=formula)
    if fmt: c.number_format = fmt
    c.border = BOX
    c.font = F_RES if big else Font(name="Calibri", size=10, bold=True)
    c.fill = FILL_RES if big else FILL_CALC
    c.alignment = Alignment(horizontal="center")
    n = ls.cell(row=r, column=5, value=note); n.font = F_S
    n.alignment = Alignment(wrap_text=True, vertical="center")
    if key: LR[key] = r
    lptr[0] += 1
    return r

ls["G3"] = "Коэффициенты § 32a EStG 2026"; ls["G3"].font = F_H2
coef = [("G4", 914.51), ("G5", 1400.0), ("G6", 173.10), ("G7", 2397.0),
        ("G8", 1034.87), ("G9", 0.42), ("G10", 11135.63), ("G11", 0.45), ("G12", 19470.38)]
labels = ["зона 1 · a", "зона 1 · b", "зона 2 · a", "зона 2 · b", "зона 2 · c",
          "зона 3 · %", "зона 3 · вычет", "зона 4 · %", "зона 4 · вычет"]
for (cell, val), lab in zip(coef, labels):
    ls[cell] = val
    ls[cell].number_format = '0.00'
    ls[cell].fill = FILL_IN; ls[cell].border = BOX; ls[cell].font = F_N
    ls[cell].alignment = Alignment(horizontal="center")
    ls.cell(row=int(cell[1:]), column=8, value=lab).font = F_S

GFB, Z1, Z2, Z3 = P["GFB"], P["Z1"], P["Z2"], P["Z3"]
def tarif(x):
    return (f'IF({x}<={GFB},0,'
            f'IF({x}<={Z1},ROUNDDOWN(($G$4*(({x}-{GFB})/10000)+$G$5)*(({x}-{GFB})/10000),0),'
            f'IF({x}<={Z2},ROUNDDOWN(($G$6*(({x}-{Z1})/10000)+$G$7)*(({x}-{Z1})/10000)+$G$8,0),'
            f'IF({x}<={Z3},ROUNDDOWN($G$9*{x}-$G$10,0),'
            f'ROUNDDOWN($G$11*{x}-$G$12,0)))))')

lsec("1. ИСХОДНЫЕ ДАННЫЕ")
lline("Налогооблагаемое брутто за месяц", "steuerpflichtiger Monatslohn", f"=Расчёт!D{AR['stbrutto']}", "mon", FMT_EUR,
      "оклад + переработка + премии + облагаемая часть надбавок")
lline("Прогноз на год", "hochgerechneter Jahresarbeitslohn", f"=ROUND(D{LR['mon']}*12,2)", "jahr", FMT_EUR,
      "как в ELStAM: месяц × 12. В месяц с большой переработкой налог удерживают завышенно — часть вернётся по Steuererklärung")
lline("Налоговый класс", "Steuerklasse", f"={P['StKl']}", "stkl", '0', "")
lptr[0] += 1

lsec("2. ВЫЧЕТЫ (в год)")
lline("Вычет на работника", "Arbeitnehmer-Pauschbetrag", f"=IF({P['StKl']}<6,{P['ANPausch']},0)", "anp", FMT_EUR, "§ 9a Nr. 1a EStG")
lline("Особые расходы", "Sonderausgaben-Pauschbetrag", f"=IF({P['StKl']}<6,{P['SAPausch']},0)", "sap", FMT_EUR, "§ 10c EStG")
lline("Одинокому родителю", "Entlastungsbetrag Alleinerziehende", f"=IF({P['StKl']}=2,{P['EntlastAE']},0)", "ent", FMT_EUR, "только класс 2")
lptr[0] += 1

lsec("3. VORSORGEPAUSCHALE (§ 39b Abs. 2 S. 5 Nr. 3 EStG)")
lline("Пенсионное страхование", "Teilbetrag Rentenversicherung", f"=ROUND(MIN(D{LR['jahr']},{P['BBGRV']}*12)*{P['RVSatz']}/2,2)", "vrv", FMT_EUR, "100 % доли работника")
lline("Медицинское страхование", "Teilbetrag Krankenversicherung", f"=ROUND(MIN(D{LR['jahr']},{P['BBGKV']}*12)*({P['KVSatz']}/2+{P['KVZusatz']}/2),2)", "vkv", FMT_EUR, "")
lline("Страхование по уходу", "Teilbetrag Pflegeversicherung", f'=ROUND(MIN(D{LR["jahr"]},{P["BBGKV"]}*12)*({P["PVSatz"]}/2+IF({P["Kinderlos"]}="ja",{P["PVZuschlag"]},0)),2)', "vpv", FMT_EUR, "")
lline("Страхование по безработице", "Teilbetrag Arbeitslosenversicherung",
      f'=IF({P["AVinVP"]}="ja",ROUND(MIN(MIN(D{LR["jahr"]},{P["BBGRV"]}*12)*{P["AVSatz"]}/2,MAX(0,1900-D{LR["vkv"]}-D{LR["vpv"]})),2),0)', "vav", FMT_EUR,
      "новое с 2026, засчитывается лишь в пределах 1.900 € вместе с KV/PV")
lline("Vorsorgepauschale всего", "Vorsorgepauschale gesamt", f"=D{LR['vrv']}+D{LR['vkv']}+D{LR['vpv']}+D{LR['vav']}", "vges", FMT_EUR, "")
lptr[0] += 1

lsec("4. ТАРИФ")
lline("Налогооблагаемый доход", "zu versteuerndes Einkommen (zvE)",
      f"=MAX(0,ROUNDDOWN(D{LR['jahr']}-D{LR['anp']}-D{LR['sap']}-D{LR['ent']}-D{LR['vges']},0))", "zve", FMT_EUR,
      "округляется вниз до полного евро")
lline("Основной тариф", "Grundtarif § 32a Abs. 1", f"={tarif('D' + str(LR['zve']))}", "gt", FMT_EUR, "классы 1, 2, 4")
lline("Тариф со сплиттингом", "Splittingtarif § 32a Abs. 5", f"=2*({tarif('ROUNDDOWN(D' + str(LR['zve']) + '/2,0)')})", "sp", FMT_EUR, "класс 3 — твой случай")
lline("Подоходный налог за год", "Jahreslohnsteuer", f"=IF({P['StKl']}=3,D{LR['sp']},D{LR['gt']})", "jlst", FMT_EUR, "")
lline("ПОДОХОДНЫЙ НАЛОГ ЗА МЕСЯЦ", "Lohnsteuer monatlich", f"=ROUND(D{LR['jlst']}/12,2)", "mlst", FMT_EUR, "", big=True)
lptr[0] += 1

lsec("5. SOLI И ЦЕРКОВНЫЙ НАЛОГ")
lline("Порог Soli для твоего класса", "Soli-Freigrenze", f"=IF({P['StKl']}=3,{P['SoliFrei']}*2,{P['SoliFrei']})", "sfg", FMT_EUR, "для класса 3 порог удваивается")
lline("Soli за год", "Solidaritätszuschlag Jahr",
      f"=IF(D{LR['jlst']}<=D{LR['sfg']},0,ROUND(MIN(0.055*D{LR['jlst']},0.119*(D{LR['jlst']}-D{LR['sfg']})),2))", "soliJ", FMT_EUR,
      "зона смягчения 11,9 %")
lline("Soli за месяц", "Solidaritätszuschlag monatlich", f"=ROUND(D{LR['soliJ']}/12,2)", "soliM", FMT_EUR, "")
lline("Церковный налог за месяц", "Kirchensteuer monatlich", f'=IF({P["KiSt"]}="ja",ROUND(D{LR["mlst"]}*{P["KiStSatz"]},2),0)', "kistM", FMT_EUR, "")
lptr[0] += 1

wr = lptr[0]
ls.merge_cells(start_row=wr, start_column=2, end_row=wr, end_column=5)
wc = ls.cell(row=wr, column=2, value=(f'=IF(OR({P["StKl"]}=5,{P["StKl"]}=6),'
    '"ВНИМАНИЕ: для классов 5 и 6 действует другой алгоритм (§ 39b Abs. 2 S. 7 EStG) — расчёт НЕ подходит.",'
    '"Оценка. Официальная сумма — в Lohnabrechnung. Сверить: bmf-steuerrechner.de (там выбери Steuerklasse 3, Kinderfreibetrag 1,0)")'))
wc.font = F_B; wc.fill = FILL_WARN
wc.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ls.row_dimensions[wr].height = 30

ab.cell(row=AR['lst'],  column=4, value=f"=Lohnsteuer!D{LR['mlst']}")
ab.cell(row=AR['soli'], column=4, value=f"=Lohnsteuer!D{LR['soliM']}")
ab.cell(row=AR['kist'], column=4, value=f"=Lohnsteuer!D{LR['kistM']}")

# =====================================================================
# ЛИСТ 7 — КОНТРОЛЬ ArbZG
# =====================================================================
kz = wb.create_sheet("Контроль ArbZG")
kz.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGH", [2, 16, 16, 12, 10, 14, 46, 20]):
    kz.column_dimensions[col].width = w
title_row(kz, 1, "КОНТРОЛЬ РАБОЧЕГО ВРЕМЕНИ / ARBEITSZEITGESETZ", 8)
kz["B2"] = "Считаются смены, НАЧИНАЮЩИЕСЯ в этом месяце. Стыки месяцев проверь вручную."
kz["B2"].font = F_S
for i, h in enumerate(["Неделя с", "по", "Часов", "Смен", "Ночных", "Статус"], start=2):
    c = kz.cell(row=4, column=i, value=h); c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center")
for i in range(6):
    r = 5 + i
    kz.cell(row=r, column=2, value=f"=DATE({P['Jahr']},{P['Monat']},1)-WEEKDAY(DATE({P['Jahr']},{P['Monat']},1),2)+1+7*{i}").number_format = FMT_D
    kz.cell(row=r, column=3, value=f"=B{r}+6").number_format = FMT_D
    kz.cell(row=r, column=4, value=(f'=SUMIFS(Dienstplan!$L${R0}:$L${R1},Dienstplan!$A${R0}:$A${R1},">="&B{r},'
                                    f'Dienstplan!$A${R0}:$A${R1},"<="&C{r})')).number_format = '0.00'
    kz.cell(row=r, column=5, value=(f'=COUNTIFS(Dienstplan!$A${R0}:$A${R1},">="&B{r},Dienstplan!$A${R0}:$A${R1},"<="&C{r},'
                                    f'Dienstplan!$L${R0}:$L${R1},">0")')).number_format = '0'
    kz.cell(row=r, column=6, value=(f'=COUNTIFS(Dienstplan!$A${R0}:$A${R1},">="&B{r},Dienstplan!$A${R0}:$A${R1},"<="&C{r},'
                                    f'Dienstplan!$D${R0}:$D${R1},"ND")')).number_format = '0'
    kz.cell(row=r, column=7, value=f'=IF(D{r}>48,"ПРЕВЫШЕНИЕ: больше 48 ч (§ 3 ArbZG)",IF(D{r}>40,"выше нормы 40 ч","в норме"))')
    for c in range(2, 8):
        cc = kz.cell(row=r, column=c)
        cc.border = BOX; cc.font = F_N; cc.fill = FILL_CALC
        cc.alignment = Alignment(horizontal="center")
kz.conditional_formatting.add("D5:D10", CellIsRule(operator="greaterThan", formula=["48"], fill=FILL_WARN, font=F_RED))

r = 12
section(kz, r, "ЧТО ГОВОРИТ ЗАКОН", 8); r += 1
GESETZ = [
 ("§ 3 ArbZG", "Рабочий день — максимум 8 ч. До 10 ч можно, только если за 6 месяцев (или 24 недели) средняя не выше 8 ч в день, то есть около 48 ч в неделю."),
 ("§ 4 ArbZG", "Перерыв: при работе более 6 ч — 30 мин, более 9 ч — 45 мин. Можно дробить по 15 мин. Более 6 ч подряд без перерыва запрещено."),
 ("§ 5 ArbZG", "Отдых между сменами 11 ч. В уходе допускается 10 ч, если в течение месяца компенсируется другой сменой отдыха в 12 ч."),
 ("§ 6 Abs. 2 ArbZG", "ГЛАВНОЕ ДЛЯ ТЕБЯ: у ночного работника смена максимум 8 ч, до 10 ч — только если за календарный месяц (или 4 недели) средняя не выше 8 ч. Твои 12-часовые ND этому не соответствуют."),
 ("§ 7 ArbZG", "Смены длиннее возможны лишь при существенной доле Bereitschaftsdienst И на основании Tarifvertrag или Betriebsvereinbarung. В EzB (п. 42) прямо указано, что работодатель НЕ связан тарифом — значит такой основы нет."),
 ("§ 6 Abs. 3 ArbZG", "Право на бесплатное медобследование каждые 3 года (после 50 лет — ежегодно), за счёт работодателя."),
 ("§ 6 Abs. 5 ArbZG", "Право на соразмерную надбавку за ночную работу ИЛИ на оплачиваемые отгулы. Ориентир BAG — 25 %; при постоянной ночной работе суды присуждали до 30 %."),
 ("§ 11 ArbZG", "Минимум 15 воскресений в году свободны. За работу в воскресенье — Ersatzruhetag в течение 2 недель, за праздник — в течение 8 недель."),
 ("§ 16 ArbZG", "Работодатель обязан фиксировать всё время сверх 8 ч и хранить записи 2 года. Веди свой учёт — он решает исход спора."),
 ("§ 22 ArbZG", "Нарушения — штраф до 15.000 € для работодателя. Жалоба: Gewerbeaufsichtsamt Hannover (Staatliches Gewerbeaufsichtsamt), анонимно тоже принимают."),
 ("§ 3b EStG", "Надбавки за ночь/воскресенье/праздник не облагаются подоходным налогом: ночь 25 %, воскресенье 50 %, праздник 125/150 %, база до 50 €/ч."),
 ("§ 1 SvEV", "От соцвзносов надбавки свободны только с базы до 25 €/ч. При ставке 27 € часть надбавок облагается взносами."),
 ("§ 11 BUrlG", "Отпускные — по среднему заработку за последние 13 недель, ВКЛЮЧАЯ надбавки (сверхурочные не входят). Проверяй, что надбавки учли."),
 ("§ 4 EFZG", "На больничном платят по принципу Lohnausfall — то, что заработал бы, включая надбавки за смены по графику."),
]
for para, txt in GESETZ:
    kz.cell(row=r, column=2, value=para).font = F_B
    c = kz.cell(row=r, column=3, value=txt); c.font = F_N
    c.alignment = Alignment(wrap_text=True, vertical="top")
    kz.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
    kz.row_dimensions[r].height = max(15, 13 * (1 + len(txt) // 105))
    r += 1

# =====================================================================
# ЛИСТ 8 — ПРАВА И РИСКИ
# =====================================================================
pr = wb.create_sheet("Права и риски")
pr.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [2, 26, 40, 44, 44]):
    pr.column_dimensions[col].width = w
title_row(pr, 1, "ПРАВА И РИСКИ — что в договоре есть, чего нет", 5)
pr["B2"] = "Проверено по всем 19 параграфам договора от 01.08.2026 и по EzB от 13.07.2026."
pr["B2"].font = F_S
for i, h in enumerate(["Тема", "Что в договоре", "Что по закону / на практике", "Что делать"], start=2):
    c = pr.cell(row=4, column=i, value=h); c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center")

RISKS = [
 ("Einspringprämie",
  "НЕТ НИ СЛОВА. § 4 Abs. 1: любые Sonderleistungen сверх договора — добровольные, права на них нет.",
  "НО § 4 Abs. 1 Satz 3: оговорка о добровольности НЕ действует, если выплата основана на индивидуальной договорённости (individuelle Vertragsabrede). Устная договорённость тоже считается, но её надо доказать.",
  "Напиши руководителю e-mail: «Bitte bestätigen Sie mir die vereinbarte Einspringprämie von … € je kurzfristig übernommenem Dienst.» Ответ по e-mail = доказательство. Без этого премию могут прекратить в любой момент."),
 ("Лимит 60 ч переработки",
  "НЕТ ТАКОГО ПРАВИЛА. § 2 Abs. 3 разрешает не оплачивать только 8 ч в месяц. Всё остальное — обычное рабочее время по § 3 Abs. 1 (27 €/ч).",
  "Замена оплаты на отгулы (Freizeitausgleich) требует соглашения сторон или Betriebsvereinbarung. Работодатель не может ввести это в одностороннем порядке. Часы не пропадают — они на счёте (Arbeitszeitkonto).",
  "Само по себе накопление часов не страшно. Опасен § 14: требуй КАЖДЫЙ месяц письменное подтверждение остатка счёта часов. Это и фиксирует требование, и не даёт сроку истечь."),
 ("Ausschlussfrist 3 месяца",
  "§ 14: любые требования пропадают, если не заявлены письменно в течение 3 месяцев с момента, когда они стали к оплате.",
  "Зарплата за август платится 15 сентября → срок по августу истекает 15 декабря. Для перенесённых часов момент возникновения требования спорный — не рискуй.",
  "Достаточно e-mail с формулировкой: «Hiermit mache ich die Vergütung der im Monat … geleisteten … Überstunden geltend.» Храни копию. Это занимает 2 минуты и сохраняет тысячи евро."),
 ("12-часовые ночные смены",
  "§ 2 Abs. 2: обязан работать в ночь, в выходные и праздники по распоряжению — «soweit dies gesetzlich zulässig ist».",
  "§ 6 Abs. 2 ArbZG: у ночного работника смена максимум 8 ч, до 10 ч только с компенсацией в течение месяца. 12 ч допустимы лишь при большой доле Bereitschaftsdienst и только на базе тарифа или Betriebsvereinbarung (§ 7 ArbZG). По EzB п. 42 работодатель НЕ tarifgebunden.",
  "Оговорка «soweit gesetzlich zulässig» работает в твою пользу: сверх законного предела ты не обязан. Если давят — Gewerbeaufsichtsamt Hannover. Обращение может быть анонимным."),
 ("Ночная надбавка 25 %",
  "§ 3 Abs. 2 a): 25 % за 20:00–06:00.",
  "§ 6 Abs. 5 ArbZG даёт право на «angemessener Zuschlag». BAG: 25 % — норма при обычной ночной работе; при Dauernachtarbeit (как у тебя — почти только ночи) суды присуждали 30 %.",
  "Реальный аргумент на переговорах о повышении. 5 % от 170 ночных часов — это около 230 € в месяц."),
 ("Надбавки на переработку",
  "§ 3 Abs. 2: надбавки начисляются «je Stunde auf die (Grund-)Vergütung» — на каждый час, без исключений для переработки.",
  "Значит за переработанные ночные часы надбавка 25 % тоже положена. Частая ошибка расчётчиков — начислить надбавки только на часы в пределах нормы.",
  "Сверь в Lohnabrechnung: количество часов с надбавкой должно совпасть с колонками «Ночь/Воскр./Праздн.» этого файла, а не быть меньше."),
 ("Надбавки в отпуске и на больничном",
  "§ 5 и § 6 отсылают к закону.",
  "§ 11 BUrlG: отпускные — среднее за 13 недель ВКЛЮЧАЯ надбавки. § 4 EFZG: больничный — по принципу Lohnausfall, тоже с надбавками за смены по графику.",
  "Проверь: если в месяц с отпуском надбавки резко упали — это ошибка расчёта. Такие надбавки, в отличие от обычных, облагаются налогом (§ 3b требует фактической работы) — файл это учитывает."),
 ("Праздники 135 %",
  "§ 3 Abs. 2 d): 135 % за праздник без отгула, 35 % — если дают отгул.",
  "Разница огромная: 12-часовая смена в праздник — 437,40 € надбавки без отгула против 113,40 € с отгулом. Свободны от налога только 125 %, остальные 10 % облагаются.",
  "Заранее уточняй по каждому празднику: с отгулом или без. В файле есть колонка FZA — ставь j (с отгулом) или n (без)."),
 ("Служебная машина",
  "§ 12 Abs. 1: машина предоставляется, ЛЮБОЕ личное использование запрещено — включая дорогу из дома до места работы.",
  "Раз личного использования нет, нет и geldwerter Vorteil. Правило 1 % применяться НЕ должно.",
  "Проверь Lohnabrechnung: если там есть строка «geldwerter Vorteil Pkw» или «1 %-Regelung» — это ошибка, ты платишь лишний налог и взносы."),
 ("Дорога до работы",
  "§ 2 Abs. 5: дорога дом–работа рабочим временем не считается.",
  "Это законно. НО поездки МЕЖДУ пациентами в течение смены — рабочее время и должны оплачиваться.",
  "Если ездишь между объектами — фиксируй это время и требуй оплаты."),
 ("Отпуск",
  "§ 5 Abs. 1: «по закону». В EzB стояло 15 дней — но это было для 21 ч/нед.",
  "§ 3 BUrlG при 5-дневной неделе — 20 рабочих дней. Плюс дополнительный оплачиваемый отпуск по PflegeArbbV. При работе по сменам отпуск пересчитывается по фактическим рабочим дням.",
  "Требуй письменного подтверждения количества дней на полной ставке. 15 дней для Vollzeit было бы нарушением закона."),
 ("EzB и договор расходятся",
  "EzB от 13.07.2026: 21 ч/нед, 2.184,00 €, старт 15.07.26. Договор от 01.08.2026: 40 ч/нед, 4.676,40 €.",
  "Действует договор — § 18 Abs. 2 прямо заменяет все предыдущие соглашения. Но у Ausländerbehörde в деле лежит версия с 21 часом.",
  "Расхождение между разрешением на работу и фактической занятостью — риск для твоего вида на жительство. Проси работодателя подать новую EzB на полную ставку и оставь себе копию."),
 ("Штрафы по договору",
  "§ 9: до одного месячного оклада за нарушение конфиденциальности или невыход на работу. § 17: запрет переманивать сотрудников 24 месяца, тоже месячный оклад.",
  "Такие оговорки часто признают недействительными, но спорить дорого.",
  "Просто знай: увольнение без соблюдения срока = штраф до 4.676,40 €."),
 ("Выплата и проверка",
  "§ 3 Abs. 3: 15-го числа следующего месяца, безналично.",
  "Работодатель обязан выдать понятную расчётку (§ 108 GewO) с разбивкой по видам надбавок.",
  "Каждый месяц сверяй расчётку с этим файлом. Расхождение больше пары евро — пиши сразу, не жди."),
]
r = 5
for tema, vertrag, gesetz, tun in RISKS:
    pr.cell(row=r, column=2, value=tema).font = F_H2
    for col, txt in [(3, vertrag), (4, gesetz), (5, tun)]:
        c = pr.cell(row=r, column=col, value=txt)
        c.font = F_N if col != 5 else F_B
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = BOX
        c.fill = FILL_CALC if col != 5 else FILL_RES
    pr.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    pr.cell(row=r, column=2).border = BOX
    pr.cell(row=r, column=2).fill = FILL_SUB
    longest = max(len(vertrag), len(gesetz), len(tun))
    pr.row_dimensions[r].height = max(46, 11.5 * (1 + longest // 42))
    r += 1

# =====================================================================
# ЛИСТ 9 — МОИ ДАННЫЕ  (пустые поля: репозиторий публичный)
# =====================================================================
md = wb.create_sheet("Мои данные")
md.sheet_view.showGridLines = False
for col, w in zip("ABCD", [2, 34, 34, 60]):
    md.column_dimensions[col].width = w
title_row(md, 1, "МОИ ДАННЫЕ / PERSÖNLICHE DATEN", 4)
md["B2"] = ("ПОЛЯ ОСТАВЛЕНЫ ПУСТЫМИ НАМЕРЕННО: этот файл лежит в публичном репозитории GitHub. "
            "Заполни их у себя на телефоне/компьютере и НЕ загружай файл с данными обратно в репозиторий. "
            "Steuer-ID и SV-Nummer — это идентификаторы, по которым возможно мошенничество.")
md["B2"].font = F_RED
md.merge_cells("B2:D2")
md.row_dimensions[2].height = 34
md["B2"].alignment = Alignment(wrap_text=True, vertical="center")

FIELDS = [
 ("SEC", "ЛИЧНОЕ"),
 ("F", "Фамилия, имя", "Name, Vorname", "Karimov, Alisher"),
 ("F", "Дата рождения", "Geburtsdatum", "07.02.1998"),
 ("F", "Адрес", "Anschrift", ""),
 ("F", "Номер соцстрахования", "Sozialversicherungsnummer", ""),
 ("F", "Налоговый номер", "Steuer-Identifikationsnummer", ""),
 ("F", "Больничная касса", "Krankenkasse", ""),
 ("F", "Доп. взнос кассы, %", "Zusatzbeitrag", ""),
 ("SEC", "НАЛОГИ"),
 ("F", "Налоговый класс", "Steuerklasse", "3"),
 ("F", "Детские вычеты", "Kinderfreibeträge", "1,0"),
 ("F", "Ребёнок, дата рождения", "Kind, geboren am", ""),
 ("F", "Церковный налог", "Kirchensteuer", "nein"),
 ("SEC", "РАБОТА"),
 ("F", "Работодатель", "Arbeitgeber", "Constantia Intensivpflege GmbH"),
 ("F", "Адрес работодателя", "Anschrift AG", "Theaterstraße 1, 30159 Hannover"),
 ("F", "Betriebsnummer", "Betriebsnummer", "70027532"),
 ("F", "Должность", "Tätigkeit", "Examinierte Pflegefachkraft, außerklinische Intensivpflege 1:1"),
 ("F", "Начало работы по договору", "Beginn", "01.08.2026"),
 ("F", "Испытательный срок", "Probezeit", "в договоре не предусмотрен"),
 ("F", "Контактное лицо", "Kontaktperson", "Christopher Kruhl · 0511 43834924 · info@constantia-intensiv.de"),
]
r = 4
for item in FIELDS:
    if item[0] == "SEC":
        md.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = md.cell(row=r, column=2, value=item[1]); c.font = F_H1; c.fill = FILL_HEAD
        c.alignment = Alignment(vertical="center", indent=1)
        r += 1
        continue
    md.cell(row=r, column=2, value=item[1]).font = F_B
    c = md.cell(row=r, column=3, value=item[2]); c.font = F_N
    v = md.cell(row=r, column=4, value=item[3])
    v.fill = FILL_IN if item[3] == "" else FILL_CALC
    v.border = BOX; v.font = F_N
    v.alignment = Alignment(wrap_text=True, vertical="center")
    r += 1

# =====================================================================
# ЛИСТ 10 — ГОД
# =====================================================================
jr = wb.create_sheet("Год")
jr.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHI", [2, 14, 11, 13, 13, 13, 13, 13, 16]):
    jr.column_dimensions[col].width = w
title_row(jr, 1, "СВОДКА ЗА ГОД / JAHRESÜBERSICHT", 9)
jr["B2"] = "Каждый месяц: копия файла → расчёт → перенеси итоги сюда. Колонка «Счёт часов» — остаток переносимых часов на конец месяца."
jr["B2"].font = F_S
for i, h in enumerate(["Месяц", "Часы", "Брутто", "Надбавки", "Соцвзносы", "Налоги", "Нетто", "Счёт часов"], start=2):
    c = jr.cell(row=4, column=i, value=h); c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
MON = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]
for i, m in enumerate(MON):
    r = 5 + i
    jr.cell(row=r, column=2, value=m).font = F_B
    jr.cell(row=r, column=2).border = BOX
    for c in range(3, 10):
        cc = jr.cell(row=r, column=c)
        cc.border = BOX; cc.fill = FILL_IN; cc.font = F_N
        cc.number_format = '0.00' if c in (3, 9) else FMT_EUR
        cc.alignment = Alignment(horizontal="center")
r = 17
jr.cell(row=r, column=2, value="ИТОГО").font = F_H1
jr.cell(row=r, column=2).fill = FILL_HEAD
jr.cell(row=r, column=2).border = BOX
for c in range(3, 9):
    cl = get_column_letter(c)
    cc = jr.cell(row=r, column=c, value=f"=SUM({cl}5:{cl}16)")
    cc.number_format = '0.00' if c == 3 else FMT_EUR
    cc.font = Font(name="Calibri", size=11, bold=True, color=C_DARK)
    cc.fill = FILL_SUB; cc.border = BOX
    cc.alignment = Alignment(horizontal="center")

jr.cell(row=19, column=2, value="Зачем это нужно:").font = F_H2
for i, t in enumerate([
  "· Не пропустить срок § 14 (3 месяца) ни по одному месяцу.",
  "· Видеть, как растёт счёт неоплаченных часов — при лимите 60 ч/мес он может копиться бесконечно.",
  "· Понять, выгодна ли Steuererklärung: при неровном доходе почти всегда возвращают деньги.",
  "· Контроль среднего рабочего времени за 6 месяцев (§ 3 ArbZG — не более 48 ч/нед).",
  "· База для отпускных (§ 11 BUrlG — среднее за 13 недель, включая надбавки).",
]):
    jr.cell(row=20 + i, column=2, value=t).font = F_N
    jr.merge_cells(start_row=20 + i, start_column=2, end_row=20 + i, end_column=9)

wb.active = 0
wb.save(OUT)
print("OK ->", OUT)
