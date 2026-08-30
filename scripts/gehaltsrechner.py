# -*- coding: utf-8 -*-
"""
Generator: Gehaltsrechner / Расчёт зарплаты
Constantia Intensivpflege GmbH - Alisher Karimov
Grundlagen: Arbeitsvertrag v. 01.08.2026 (Vollzeit), EzB 02/2024,
            ArbZG, EStG (§ 3b, § 32a, § 39b), SvEV, BUrlG, EFZG.
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName

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
FMT_T   = 'HH:MM'
FMT_D   = 'DD.MM.YYYY'

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
# 1. ИНСТРУКЦИЯ
# =====================================================================
ws = wb.active
ws.title = "Инструкция"
ws.sheet_view.showGridLines = False
for col, w in zip("ABCDEF", [3, 42, 46, 30, 14, 14]):
    ws.column_dimensions[col].width = w

title_row(ws, 1, "РАСЧЁТ ЗАРПЛАТЫ · Alisher Karimov · Constantia Intensivpflege GmbH", 6)
ws["B3"] = "Основано на: Arbeitsvertrag от 01.08.2026, EzB (Erklärung zum Beschäftigungsverhältnis) 13.07.2026, ArbZG, EStG, SvEV, BUrlG, EFZG"
ws["B3"].font = F_S

rows = [
 ("SEC", "КАК ПОЛЬЗОВАТЬСЯ ФАЙЛОМ"),
 ("N", "1. Лист «Настройки»", "Все параметры: ставка, проценты надбавок, время смен, налоговый класс. Жёлтые поля можно менять."),
 ("N", "2. Лист «Dienstplan»", "График на месяц: ND / TD / FREI, статус (Arbeit / Urlaub / Krank), «x» если Eingesprungen. Август 2026 уже заполнен по твоему скриншоту."),
 ("N", "3. Лист «Расчёт»", "Автоматически: брутто, надбавки, налоги, соцвзносы, нетто."),
 ("N", "4. Лист «Lohnsteuer»", "Детальный расчёт подоходного налога (метод официального Programmablaufplan ELStAM)."),
 ("N", "5. Лист «Контроль ArbZG»", "Проверка по закону о рабочем времени: >10 ч, отдых <11 ч, перерывы, воскресенья."),
 ("N", "6. Лист «Feiertage»", "Праздники Нижней Саксонии 2026/2027. Строки можно дописывать."),
 ("N", "7. Лист «Год»", "Сводка по 12 месяцам — итоги каждого месяца вносишь вручную."),
 ("SP",),
 ("SEC", "ЧТО ГОВОРИТ ТВОЙ ДОГОВОР (§ 3 Arbeitsvertrag)"),
 ("N", "Часовая ставка (Stundenlohn)", "27,00 € брутто"),
 ("N", "Оклад в месяц (Monatsvergütung)", "4.676,40 € брутто  →  4.676,40 / 27,00 = 173,20 плановых часов в месяц"),
 ("N", "Рабочее время (§ 2 Abs. 1)", "40 часов в неделю, пятидневка, без учёта перерывов"),
 ("N", "Ночные 20:00–06:00 (Nachtarbeit)", "+25 %"),
 ("N", "Воскресенье 00:00–24:00 (Sonntagsarbeit)", "+50 %"),
 ("N", "Праздник с отгулом (Feiertag mit FZA)", "+35 %"),
 ("N", "Праздник без отгула (Feiertag ohne FZA)", "+135 %"),
 ("N", "Совпадение надбавок", "Платится ТОЛЬКО САМАЯ ВЫСОКАЯ. Ночь в воскресенье = 50 %, а не 75 %. Файл считает именно так."),
 ("N", "Переработки (§ 2 Abs. 3)", "Первые 8 часов сверх нормы в месяц уже включены в оклад. Всё, что свыше, должно оплачиваться отдельно."),
 ("N", "Переработки — главное", "Переработка засчитывается, только если её ПРИКАЗАЛИ или заранее одобрили. Утверждённый Dienstplan и есть приказ. Храни скриншоты графика: при 226 часах против нормы 173,2 речь идёт примерно о 1.200 € в месяц."),
 ("N", "Выплата (§ 3 Abs. 3)", "15-го числа следующего месяца, безналично"),
 ("N", "Срок претензий (§ 14 Ausschlussfrist)", "ВАЖНО: 3 месяца с даты, когда деньги должны были прийти. Позже право теряется. Требование в текстовой форме — достаточно e-mail."),
 ("SP",),
 ("SEC", "НАЛОГИ: ПОЧЕМУ НАДБАВКИ ВЫГОДНЫ"),
 ("N", "§ 3b EStG", "Надбавки за ночь/воскресенье/праздник НЕ облагаются подоходным налогом в пределах: ночь 25 %, воскресенье 50 %, праздник 125 % (24.12 с 14:00, 25./26.12, 01.05 — 150 %). База — не более 50 €/час."),
 ("N", "§ 1 Abs. 1 SvEV", "От СОЦВЗНОСОВ надбавки свободны только с базы до 25,00 €/час. Твоя ставка 27,00 € → часть надбавки, посчитанная с 2,00 €/час, взносами ОБЛАГАЕТСЯ. Файл делит это автоматически."),
 ("N", "Праздник без отгула 135 %", "Свободны от налога только 125 %, оставшиеся 10 % — налогооблагаемые. Учтено."),
 ("N", "Отпуск и больничный", "Надбавки, которые продолжают платить во время отпуска/болезни, налогом ОБЛАГАЮТСЯ (§ 3b требует фактической работы). Ставь статус «Urlaub»/«Krank» — файл посчитает правильно."),
 ("SP",),
 ("SEC", "ЧТО ПРОВЕРИТЬ / О ЧЁМ СПРОСИТЬ РАБОТОДАТЕЛЯ"),
 ("N", "Точное время смен", "По умолчанию TD 06:00–20:00 и ND 20:00–06:00. ПОСТАВЬ СВОИ реальные часы на листе «Настройки» — от этого зависит весь расчёт."),
 ("N", "Перерывы (Pause)", "По умолчанию 0. § 4 ArbZG: при работе >6 ч нужен перерыв 30 мин, >9 ч — 45 мин. Если работодатель вычитает перерыв из оплаты, впиши его."),
 ("N", "Einspringprämie", "Премия за «Eingesprungen» в договоре НЕ прописана. Узнай сумму (типично 50–150 € за смену) и впиши в «Настройки». Это Sonderleistung — облагается налогом и взносами полностью."),
 ("N", "Ночная надбавка по закону", "§ 6 Abs. 5 ArbZG даёт право на «angemessener Zuschlag». Ориентир BAG — 25 %, при постоянной ночной работе (Dauernachtarbeit) суды присуждали до 30 %. Ты работаешь почти только ночами → есть аргумент просить 30 %."),
 ("N", "Два разных документа", "EzB (для ведомства): 21 ч/нед, 2.184,00 €, 15 дней отпуска. Договор от 01.08.2026: 40 ч/нед, 4.676,40 €. Действует ДОГОВОР (§ 18 Abs. 2 — заменяет предыдущие соглашения). Убедись, что работодатель подал новую EzB на полную ставку."),
 ("N", "Отпуск", "Договор § 5: «по закону» = 20 рабочих дней при 5-дневной неделе (§ 3 BUrlG), плюс дополнительный отпуск по PflegeArbbV. 15 дней из EzB были для неполной ставки."),
 ("N", "Минимальная зарплата в уходе", "Pflegemindestlohn для Pflegefachkraft с 01.07.2026 — 21,03 €/ч. Твои 27,00 € выше. ОК."),
 ("SP",),
 ("SEC", "ТОЧНОСТЬ РАСЧЁТА"),
 ("N", "Брутто, надбавки, часы", "Точно — по формулам договора."),
 ("N", "Соцвзносы", "Точно — ставки и предельные базы 2026."),
 ("N", "Подоходный налог", "ОЦЕНКА, ±несколько евро. В официальном алгоритме есть дополнительные округления. Для налоговых классов 5 и 6 расчёт НЕ подходит."),
 ("N", "Нетто", "Оценка. Официальный документ — только Lohnabrechnung работодателя. Этот файл нужен, чтобы её ПРОВЕРИТЬ."),
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
        ws.row_dimensions[r].height = max(15, 13 * (1 + len(item[2]) // 95))
        r += 1

ws.cell(row=r + 1, column=2, value="Файл — вспомогательный инструмент, не юридическая консультация. При споре: профсоюз ver.di или Fachanwalt für Arbeitsrecht.").font = F_S

# =====================================================================
# 2. НАСТРОЙКИ
# =====================================================================
st = wb.create_sheet("Настройки")
st.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [2, 44, 34, 16, 62]):
    st.column_dimensions[col].width = w
title_row(st, 1, "НАСТРОЙКИ / EINSTELLUNGEN — жёлтые поля можно менять", 5)

NAMES = {}
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
    if name: NAMES[name] = f"Настройки!$D${r}"
    row_ptr[0] += 1
    return r

def spacer():
    row_ptr[0] += 1

sec("1. РАСЧЁТНЫЙ МЕСЯЦ")
par("Год", "Jahr", 2026, "Jahr", "0")
par("Месяц (1–12)", "Monat", 8, "Monat", "0")
par("Дней в месяце", "Tage im Monat", "=DAY(EOMONTH(DATE(Jahr,Monat,1),0))", "Tage", "0", inp=False)
par("Федеральная земля", "Bundesland", "Niedersachsen", "Land", None, "Праздники берутся с листа «Feiertage»")
spacer()

sec("2. ОПЛАТА — § 3 Arbeitsvertrag")
par("Часовая ставка", "Stundenlohn (Grundlohn)", 27.00, "Lohn", FMT_EUR, "§ 3 Abs. 1 Arbeitsvertrag")
par("Оклад брутто в месяц", "Monatsvergütung brutto", 4676.40, "Gehalt", FMT_EUR, "§ 3 Abs. 1 Arbeitsvertrag")
par("Плановые часы в месяц", "Sollstunden / Monat", "=Gehalt/Lohn", "SollStd", FMT_H, "оклад / ставка = 173,20 ч", inp=False)
par("Часов в неделю", "Wochenarbeitszeit", 40, "WStd", '0.0\\ "h"', "§ 2 Abs. 1 Arbeitsvertrag")
par("Переработки включены в оклад (ч/мес)", "pauschal abgegoltene Überstunden", 8, "UeFrei", '0\\ "h"', "§ 2 Abs. 3: до 8 ч/мес уже оплачены окладом")
par("Оплачивать переработки свыше?", "Überstunden auszahlen (ja/nein)", "ja", "UeZahlen", None, "«ja» = часы сверх (норма + 8 ч) оплачиваются по ставке")
par("Надбавка за переработку", "Überstundenzuschlag", 0.00, "UeZuschlag", FMT_PCT, "Договором не предусмотрена. Если договоришься — впиши, напр. 25 %")
par("Премия за Eingesprungen (за смену)", "Einspringprämie je Dienst", 0.00, "Praemie", FMT_EUR, "В договоре НЕТ. Узнай сумму у работодателя. Облагается налогом и взносами полностью")
par("Фактор неполного месяца", "Faktor (Ein-/Austritt)", 1.00, "Faktor", '0.0000', "1,00 = полный месяц. При неполном: отработанные рабочие дни / все рабочие дни месяца")
spacer()

sec("3. НАДБАВКИ — § 3 Abs. 2 Arbeitsvertrag")
par("Ночная надбавка", "Nachtzuschlag", 0.25, "ZNacht", FMT_PCT, "§ 3 Abs. 2 a)")
par("Ночь: с (час)", "Nachtarbeit von", 20, "NachtVon", "0", "20:00")
par("Ночь: до (час)", "Nachtarbeit bis", 6, "NachtBis", "0", "06:00")
par("Воскресная надбавка", "Sonntagszuschlag", 0.50, "ZSonn", FMT_PCT, "§ 3 Abs. 2 b), 00:00–24:00")
par("Праздник С отгулом", "Feiertag mit FZA", 0.35, "ZFmit", FMT_PCT, "§ 3 Abs. 2 c)")
par("Праздник БЕЗ отгула", "Feiertag ohne FZA", 1.35, "ZFohne", FMT_PCT, "§ 3 Abs. 2 d)")
par("Праздники по умолчанию", "Feiertagsmodus", "ohne FZA", "FZAStd", None, "«mit FZA» / «ohne FZA». В графике можно переопределить по дню (колонка FZA)")
par("Правило совпадения", "nur höchster Zuschlag", "ja", "NurHoechster", None, "§ 3 Abs. 2: при совпадении платится только самая высокая надбавка")
spacer()

sec("4. ЛЬГОТЫ ПО НАДБАВКАМ — § 3b EStG / § 1 SvEV")
par("Макс. необлагаемая ночная", "Höchstsatz Nacht steuerfrei", 0.25, "HNacht", FMT_PCT, "§ 3b Abs. 1 Nr. 1 EStG (для 00:00–04:00 при начале до полуночи — 40 %)")
par("Макс. необлагаемая воскресная", "Höchstsatz Sonntag steuerfrei", 0.50, "HSonn", FMT_PCT, "§ 3b Abs. 1 Nr. 2 EStG")
par("Макс. необлагаемая праздничная", "Höchstsatz Feiertag steuerfrei", 1.25, "HFeier", FMT_PCT, "§ 3b Abs. 1 Nr. 3 EStG; 24.12 с 14:00, 25./26.12, 01.05 — 150 %")
par("Предел базы: налог", "Grundlohngrenze steuerfrei", 50.00, "GLSteuer", FMT_EUR, "§ 3b Abs. 2 EStG — база не более 50 €/ч")
par("Предел базы: соцвзносы", "Grundlohngrenze SV-frei", 25.00, "GLSV", FMT_EUR, "§ 1 Abs. 1 Nr. 1 SvEV — база не более 25 €/ч. Твои 27 € выше → часть надбавок облагается взносами")
spacer()

sec("5. ВРЕМЯ СМЕН — проверь и поставь свои!")
par("TD (Tagdienst) начало", "TD Beginn", dt.time(6, 0), "TDvon", FMT_T, "ПРЕДПОЛОЖЕНИЕ — уточни фактическое время")
par("TD (Tagdienst) конец", "TD Ende", dt.time(20, 0), "TDbis", FMT_T, "ПРЕДПОЛОЖЕНИЕ — уточни фактическое время")
par("TD перерыв (часы)", "TD Pause", 0.00, "TDpause", FMT_H, "§ 4 ArbZG: >6 ч → 30 мин, >9 ч → 45 мин")
par("ND (Nachtdienst) начало", "ND Beginn", dt.time(20, 0), "NDvon", FMT_T, "ПРЕДПОЛОЖЕНИЕ — уточни фактическое время")
par("ND (Nachtdienst) конец", "ND Ende", dt.time(6, 0), "NDbis", FMT_T, "Переход через полночь учитывается автоматически")
par("ND перерыв (часы)", "ND Pause", 0.00, "NDpause", FMT_H, "§ 4 ArbZG: >6 ч → 30 мин")
spacer()

sec("6. НАЛОГИ И СОЦВЗНОСЫ 2026")
par("Налоговый класс", "Steuerklasse", 1, "StKl", "0", "1 / 2 / 3 / 4. Для классов 5 и 6 расчёт налога НЕ подходит")
par("Церковный налог?", "Kirchensteuerpflicht (ja/nein)", "nein", "KiSt", None, "")
par("Ставка церковного налога", "Kirchensteuersatz", 0.09, "KiStSatz", FMT_PCT, "Нижняя Саксония: 9 % от подоходного налога")
par("Детские вычеты (Zähler)", "Kinderfreibeträge", 0.0, "Kinder", "0.0", "Влияет только на Soli и церковный налог")
par("Бездетный (доплата в PV)?", "kinderlos ab 23 (ja/nein)", "ja", "Kinderlos", None, "§ 55 Abs. 3 SGB XI: +0,6 % платит только работник")
par("Взнос KV общий", "KV allgemeiner Beitragssatz", 0.146, "KVSatz", FMT_PCT, "делится пополам")
par("Доп. взнос кассы", "kassenindiv. Zusatzbeitrag", 0.029, "KVZusatz", FMT_PCT, "средний по 2026 — 2,9 %. Поставь ставку СВОЕЙ кассы. Делится пополам")
par("Взнос PV", "Pflegeversicherung", 0.036, "PVSatz", FMT_PCT, "делится пополам")
par("Доплата бездетным PV", "Zuschlag Kinderlose", 0.006, "PVZuschlag", FMT_PCT, "полностью на работнике")
par("Взнос RV", "Rentenversicherung", 0.186, "RVSatz", FMT_PCT, "делится пополам → 9,3 %")
par("Взнос AV", "Arbeitslosenversicherung", 0.026, "AVSatz", FMT_PCT, "делится пополам → 1,3 %")
par("Предел базы KV/PV в месяц", "BBG KV/PV monatlich", 5812.50, "BBGKV", FMT_EUR, "2026: 69.750 €/год")
par("Предел базы RV/AV в месяц", "BBG RV/AV monatlich", 8450.00, "BBGRV", FMT_EUR, "2026: 101.400 €/год")
spacer()

sec("7. ПАРАМЕТРЫ ПОДОХОДНОГО НАЛОГА 2026")
par("Вычет на работника", "Arbeitnehmer-Pauschbetrag", 1230.00, "ANPausch", FMT_EUR, "§ 9a EStG, в год")
par("Вычет особых расходов", "Sonderausgaben-Pauschbetrag", 36.00, "SAPausch", FMT_EUR, "§ 10c EStG, в год")
par("Вычет одинокому родителю", "Entlastungsbetrag Alleinerziehende", 4260.00, "EntlastAE", FMT_EUR, "только налоговый класс 2")
par("Учитывать AV в Vorsorgepauschale?", "AV-Teilbetrag ansetzen (ja/nein)", "nein", "AVinVP", None, "С 2026 есть новый Teilbetrag AV (§ 39b Abs. 2 S. 5 Nr. 3 Bst. e EStG), но только в пределах 1.900 € вместе с KV/PV. У тебя KV+PV уже выше → «nein»")
par("Порог Soli (налог за год)", "Soli-Freigrenze", 20350.00, "SoliFrei", FMT_EUR, "2026, одинокий. Ниже порога Soli = 0")
par("Базовый вычет", "Grundfreibetrag 2026", 12348.00, "GFB", FMT_EUR, "§ 32a EStG")
par("Граница 1-й зоны", "Ende Zone 1", 17799.00, "Z1", FMT_EUR, "§ 32a EStG 2026")
par("Граница 2-й зоны", "Ende Zone 2", 69878.00, "Z2", FMT_EUR, "§ 32a EStG 2026")
par("Граница 3-й зоны", "Ende Zone 3", 277825.00, "Z3", FMT_EUR, "§ 32a EStG 2026")

for n, ref in NAMES.items():
    wb.defined_names.add(DefinedName(n, attr_text=ref))

dv1 = DataValidation(type="list", formula1='"ja,nein"', allow_blank=True)
st.add_data_validation(dv1)
for nm in ["UeZahlen", "KiSt", "Kinderlos", "AVinVP", "NurHoechster"]:
    dv1.add(st[NAMES[nm].split("!")[1].replace("$", "")])
dv2 = DataValidation(type="list", formula1='"mit FZA,ohne FZA"', allow_blank=True)
st.add_data_validation(dv2)
dv2.add(st[NAMES["FZAStd"].split("!")[1].replace("$", "")])

# =====================================================================
# 3. FEIERTAGE
# =====================================================================
ft = wb.create_sheet("Feiertage")
ft.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [14, 34, 22, 20, 50]):
    ft.column_dimensions[col].width = w
title_row(ft, 1, "ПРАЗДНИКИ / GESETZLICHE FEIERTAGE — Niedersachsen", 5)
ft["A2"] = "Строки можно дописывать. Колонка D = максимальная НЕОБЛАГАЕМАЯ надбавка по § 3b EStG."
ft["A2"].font = F_S
for i, h in enumerate(["Дата", "Название (Feiertag)", "Земля", "§3b Höchstsatz", "Примечание"], start=1):
    c = ft.cell(row=3, column=i, value=h); c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

FEIERTAGE = [
 (dt.date(2026,1,1),  "Neujahr",                   "bundesweit",    1.25, ""),
 (dt.date(2026,4,3),  "Karfreitag",                "bundesweit",    1.25, ""),
 (dt.date(2026,4,6),  "Ostermontag",               "bundesweit",    1.25, ""),
 (dt.date(2026,5,1),  "Tag der Arbeit",            "bundesweit",    1.50, "§ 3b Abs. 3 Nr. 3: до 150 %"),
 (dt.date(2026,5,14), "Christi Himmelfahrt",       "bundesweit",    1.25, ""),
 (dt.date(2026,5,25), "Pfingstmontag",             "bundesweit",    1.25, ""),
 (dt.date(2026,10,3), "Tag der Deutschen Einheit", "bundesweit",    1.25, ""),
 (dt.date(2026,10,31),"Reformationstag",           "Niedersachsen", 1.25, ""),
 (dt.date(2026,12,25),"1. Weihnachtstag",          "bundesweit",    1.50, "§ 3b Abs. 3 Nr. 3: до 150 %"),
 (dt.date(2026,12,26),"2. Weihnachtstag",          "bundesweit",    1.50, "§ 3b Abs. 3 Nr. 3: до 150 %"),
 (dt.date(2027,1,1),  "Neujahr",                   "bundesweit",    1.25, ""),
 (dt.date(2027,3,26), "Karfreitag",                "bundesweit",    1.25, ""),
 (dt.date(2027,3,29), "Ostermontag",               "bundesweit",    1.25, ""),
 (dt.date(2027,5,1),  "Tag der Arbeit",            "bundesweit",    1.50, ""),
 (dt.date(2027,5,6),  "Christi Himmelfahrt",       "bundesweit",    1.25, ""),
 (dt.date(2027,5,17), "Pfingstmontag",             "bundesweit",    1.25, ""),
 (dt.date(2027,10,3), "Tag der Deutschen Einheit", "bundesweit",    1.25, "воскресенье"),
 (dt.date(2027,10,31),"Reformationstag",           "Niedersachsen", 1.25, "воскресенье"),
 (dt.date(2027,12,25),"1. Weihnachtstag",          "bundesweit",    1.50, ""),
 (dt.date(2027,12,26),"2. Weihnachtstag",          "bundesweit",    1.50, "воскресенье"),
]
r = 4
for d, name, land, hs, note in FEIERTAGE:
    ft.cell(row=r, column=1, value=d).number_format = FMT_D
    ft.cell(row=r, column=2, value=name)
    ft.cell(row=r, column=3, value=land)
    ft.cell(row=r, column=4, value=hs).number_format = FMT_PCT
    ft.cell(row=r, column=5, value=note)
    for cc in range(1, 6):
        ft.cell(row=r, column=cc).border = BOX
        ft.cell(row=r, column=cc).font = F_S if cc == 5 else F_N
    r += 1

r += 1
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
# 4. DIENSTPLAN
# =====================================================================
dp = wb.create_sheet("Dienstplan")
dp.sheet_view.showGridLines = False
title_row(dp, 1, "ГРАФИК И РАСЧЁТ ЧАСОВ / DIENSTPLAN", 21)
dp["A2"] = ("ND = Nachtdienst · TD = Tagdienst · FREI = выходной · «x» в колонке Sprn = eingesprungen. "
            "Статус: Arbeit / Urlaub / Krank / Fortbildung. Жёлтые колонки — ввод, серые — расчёт. "
            "Смена, начатая в этот день, целиком относится к этому дню, даже если заканчивается утром следующего.")
dp["A2"].font = F_S
dp.merge_cells("A2:U2")

HEAD = [
 ("Дата\nDatum", 11), ("День", 6), ("Праздник\nFeiertag", 20),
 ("Смена\nDienst", 9), ("Статус\nStatus", 11), ("Sprn", 6), ("FZA", 6),
 ("Начало\nBeginn", 8), ("Конец\nEnde", 8), ("Пауза\nh", 7),
 ("Длит.\nh", 7), ("Часы\nоплач.", 8),
 ("Ночь\nh", 7), ("Воскр.\nh", 7), ("Праздн.\nh", 8), ("Обычн.\nh", 8),
 ("Надбавка\n€", 11), ("из них\nбез налога", 11), ("из них\nбез взносов", 12),
 ("Отдых\nдо смены h", 11), ("Предупреждение (ArbZG)", 34),
]
for i, (h, w) in enumerate(HEAD, start=1):
    c = dp.cell(row=4, column=i, value=h)
    c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    dp.column_dimensions[get_column_letter(i)].width = w
dp.row_dimensions[4].height = 34

R0, R1 = 5, 35

PLAN = {1:"FREI",2:"ND",3:"ND",4:"ND",5:"ND",6:"FREI",7:"FREI",8:"FREI",9:"ND",10:"ND",
        11:"ND",12:"ND",13:"FREI",14:"FREI",15:"ND",16:"ND",17:"FREI",18:"FREI",19:"TD",20:"TD",
        21:"FREI",22:"TD",23:"TD",24:"ND",25:"ND",26:"ND",27:"FREI",28:"ND",29:"ND",30:"ND",31:"ND"}
SPRN = {19, 20, 28}

for i, r in enumerate(range(R0, R1 + 1)):
    n = i + 1
    dienst = PLAN.get(n, "")
    dp.cell(row=r, column=1, value=f'=IF({n}>Tage,"",DATE(Jahr,Monat,{n}))').number_format = FMT_D
    dp.cell(row=r, column=2, value=f'=IF($A{r}="","",CHOOSE(WEEKDAY($A{r},2),"Пн","Вт","Ср","Чт","Пт","Сб","Вс"))')
    dp.cell(row=r, column=3, value=f'=IF($A{r}="","",IFERROR(INDEX(Feiertage!$B:$B,MATCH($A{r},Feiertage!$A:$A,0)),""))')
    dp.cell(row=r, column=4, value=dienst)
    dp.cell(row=r, column=5, value=("Arbeit" if dienst in ("ND", "TD") else ""))
    dp.cell(row=r, column=6, value=("x" if n in SPRN else ""))
    dp.cell(row=r, column=7, value="")
    dp.cell(row=r, column=8, value=f'=IF($D{r}="ND",NDvon,IF($D{r}="TD",TDvon,""))').number_format = FMT_T
    dp.cell(row=r, column=9, value=f'=IF($D{r}="ND",NDbis,IF($D{r}="TD",TDbis,""))').number_format = FMT_T
    dp.cell(row=r, column=10, value=f'=IF($D{r}="ND",NDpause,IF($D{r}="TD",TDpause,""))').number_format = '0.00'
    dp.cell(row=r, column=11, value=f'=IF($H{r}="","",($I{r}-$H{r}+IF($I{r}<=$H{r},1,0))*24)').number_format = '0.00'
    dp.cell(row=r, column=12, value=f'=IF($K{r}="","",MAX(0,$K{r}-$J{r}))').number_format = '0.00'
    dp.cell(row=r, column=23, value=f'=IF($H{r}="","",$H{r}*24)')
    dp.cell(row=r, column=24, value=f'=IF($H{r}="","",$I{r}*24+IF($I{r}<=$H{r},24,0))')
    dp.cell(row=r, column=25, value=f'=IF(OR($K{r}="",$K{r}=0),0,$L{r}/$K{r})')
    dp.cell(row=r, column=26, value=f'=IF($A{r}="",0,IF(WEEKDAY($A{r},2)=7,1,0))')
    dp.cell(row=r, column=27, value=f'=IF($A{r}="",0,IF(WEEKDAY($A{r}+1,2)=7,1,0))')
    dp.cell(row=r, column=28, value=f'=IF($A{r}="",0,MIN(1,COUNTIF(Feiertage!$A:$A,$A{r})))')
    dp.cell(row=r, column=29, value=f'=IF($A{r}="",0,MIN(1,COUNTIF(Feiertage!$A:$A,$A{r}+1)))')
    dp.cell(row=r, column=30, value=f'=IFERROR(INDEX(Feiertage!$D:$D,MATCH($A{r},Feiertage!$A:$A,0)),HFeier)')
    dp.cell(row=r, column=31, value=f'=IFERROR(INDEX(Feiertage!$D:$D,MATCH($A{r}+1,Feiertage!$A:$A,0)),HFeier)')
    dp.cell(row=r, column=32, value=f'=IF($W{r}="",0,MAX(0,MIN($X{r},NachtBis)-MAX($W{r},0))+MAX(0,MIN($X{r},24)-MAX($W{r},NachtVon)))')
    dp.cell(row=r, column=33, value=f'=IF($W{r}="",0,MAX(0,MIN($X{r},24+NachtBis)-MAX($W{r},24)))')
    dp.cell(row=r, column=34, value=f'=IF($W{r}="",0,MAX(0,MIN($X{r},24)-MAX($W{r},0)))')
    dp.cell(row=r, column=35, value=f'=IF($W{r}="",0,MAX(0,MIN($X{r},48)-MAX($W{r},24)))')
    dp.cell(row=r, column=36, value=f'=($Z{r}*$AB{r}*$AH{r}+$AA{r}*$AC{r}*$AI{r})*$Y{r}')
    dp.cell(row=r, column=37, value=f'=$O{r}-$AJ{r}')
    dp.cell(row=r, column=38, value=f'=$N{r}-$AJ{r}')
    dp.cell(row=r, column=39, value=f'=$M{r}-(MAX($Z{r},$AB{r})*$AF{r}+MAX($AA{r},$AC{r})*$AG{r})*$Y{r}')
    dp.cell(row=r, column=40, value=f'=IF($G{r}="j",ZFmit,IF($G{r}="n",ZFohne,IF(FZAStd="mit FZA",ZFmit,ZFohne)))')
    dp.cell(row=r, column=41, value=f'=IF($O{r}=0,HFeier,($AB{r}*$AH{r}*$AD{r}+$AC{r}*$AI{r}*$AE{r})*$Y{r}/$O{r})')
    dp.cell(row=r, column=13, value=f'=($AF{r}+$AG{r})*$Y{r}').number_format = '0.00'
    dp.cell(row=r, column=14, value=f'=($Z{r}*$AH{r}+$AA{r}*$AI{r})*$Y{r}').number_format = '0.00'
    dp.cell(row=r, column=15, value=f'=($AB{r}*$AH{r}+$AC{r}*$AI{r})*$Y{r}').number_format = '0.00'
    dp.cell(row=r, column=16, value=f'=IF($L{r}="","",MAX(0,$L{r}-$O{r}-$AL{r}-$AM{r}))').number_format = '0.00'
    dp.cell(row=r, column=17, value=(f'=ROUND(Lohn*($AK{r}*$AN{r}+$AJ{r}*MAX($AN{r},ZSonn)'
                                     f'+$AL{r}*ZSonn+$AM{r}*ZNacht),2)')).number_format = FMT_EUR
    dp.cell(row=r, column=18, value=(f'=IF($E{r}<>"Arbeit",0,ROUND(MIN(Lohn,GLSteuer)*'
                                     f'($AK{r}*MIN($AN{r},$AO{r})+$AJ{r}*MIN(MAX($AN{r},ZSonn),$AO{r})'
                                     f'+$AL{r}*MIN(ZSonn,HSonn)+$AM{r}*MIN(ZNacht,HNacht)),2))')).number_format = FMT_EUR
    dp.cell(row=r, column=19, value=(f'=IF($E{r}<>"Arbeit",0,ROUND(MIN(Lohn,GLSV)*'
                                     f'($AK{r}*MIN($AN{r},$AO{r})+$AJ{r}*MIN(MAX($AN{r},ZSonn),$AO{r})'
                                     f'+$AL{r}*MIN(ZSonn,HSonn)+$AM{r}*MIN(ZNacht,HNacht)),2))')).number_format = FMT_EUR
    if r == R0:
        dp.cell(row=r, column=20, value='')
    else:
        dp.cell(row=r, column=20, value=f'=IF(OR($W{r}="",$W{r-1}=""),"",24+$W{r}-$X{r-1})').number_format = '0.0'
    dp.cell(row=r, column=21, value=(
        f'=TRIM(IF($L{r}="","",IF($L{r}>10,"смена >10 ч (§ 3 ArbZG)  ","")'
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
            cell.alignment = Alignment(horizontal="left")
        else:
            cell.fill = FILL_CALC; cell.alignment = Alignment(horizontal="center")

TR = R1 + 1
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

dvD = DataValidation(type="list", formula1='"ND,TD,FREI"', allow_blank=True)
dvS = DataValidation(type="list", formula1='"Arbeit,Urlaub,Krank,Fortbildung"', allow_blank=True)
dvX = DataValidation(type="list", formula1='"x"', allow_blank=True)
dvF = DataValidation(type="list", formula1='"j,n"', allow_blank=True)
for dv, col in [(dvD, "D"), (dvS, "E"), (dvX, "F"), (dvF, "G")]:
    dp.add_data_validation(dv)
    dv.add(f"{col}{R0}:{col}{R1}")

dp.conditional_formatting.add(f"U{R0}:U{R1}",
    FormulaRule(formula=[f'LEN($U{R0})>0'], fill=FILL_WARN, font=Font(color="9C0006", bold=True)))
dp.conditional_formatting.add(f"T{R0}:T{R1}",
    CellIsRule(operator="lessThan", formula=["11"], fill=FILL_WARN, font=Font(color="9C0006", bold=True)))
dp.conditional_formatting.add(f"A{R0}:C{R1}",
    FormulaRule(formula=[f'WEEKDAY($A{R0},2)>5'], fill=PatternFill("solid", fgColor="FCE4D6")))

for c in range(22, 42):
    dp.column_dimensions[get_column_letter(c)].hidden = True
dp.freeze_panes = "D5"

# =====================================================================
# 5. РАСЧЁТ
# =====================================================================
ab = wb.create_sheet("Расчёт")
ab.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [2, 46, 40, 16, 58]):
    ab.column_dimensions[col].width = w
title_row(ab, 1, "РАСЧЁТ ЗАРПЛАТЫ ЗА МЕСЯЦ / GEHALTSABRECHNUNG", 5)
ab["B2"] = '=CONCATENATE("Месяц: ",TEXT(DATE(Jahr,Monat,1),"MMMM YYYY"),"   ·   Alisher Karimov   ·   Constantia Intensivpflege GmbH")'
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

def aline(ru, de, formula, key=None, fmt=FMT_EUR, note="", big=False):
    r = ptr[0]
    ab.cell(row=r, column=2, value=ru).font = F_RES if big else F_B
    ab.cell(row=r, column=3, value=de).font = F_N
    c3 = ab.cell(row=r, column=4, value=formula)
    if fmt: c3.number_format = fmt
    c3.font = F_RES if big else Font(name="Calibri", size=10, bold=True)
    c3.fill = FILL_RES if big else FILL_CALC
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

TRr = 36

asec("A. ЧАСЫ / ARBEITSZEIT")
aline("Отработано часов (факт)", "Ist-Stunden", f"=Dienstplan!L{TRr}", "ist", '0.00', "Сумма с листа Dienstplan (включая отпуск/больничный)")
aline("Плановая норма", "Sollstunden", "=ROUND(SollStd*Faktor,2)", "soll", '0.00', "оклад / ставка = 173,20 ч")
aline("Разница", "Differenz", f"=D{AR['ist']}-D{AR['soll']}", "diff", '0.00', "«+» переработка, «−» недоработка (оклад платится полностью в любом случае)")
aline("Из них включены в оклад", "pauschal abgegolten (§ 2 Abs. 3)", f"=MIN(MAX(0,D{AR['diff']}),UeFrei)", "abg", '0.00', "первые 8 ч/мес")
aline("Переработка к доплате", "auszuzahlende Überstunden", f'=IF(UeZahlen="ja",MAX(0,D{AR["diff"]}-UeFrei),0)', "ueh", '0.00', "")
aline("Ночных часов", "Nachtstunden", f"=Dienstplan!M{TRr}", "nh", '0.00', "")
aline("Воскресных часов", "Sonntagsstunden", f"=Dienstplan!N{TRr}", "sh", '0.00', "")
aline("Праздничных часов", "Feiertagsstunden", f"=Dienstplan!O{TRr}", "fh", '0.00', "")
aline("Смен «Eingesprungen»", "Einspringdienste", f"=Dienstplan!F{TRr}", "sprn", '0', "")
aspace()

asec("B. БРУТТО / BRUTTOENTGELT")
aline("Оклад", "Grundvergütung § 3 Abs. 1", "=ROUND(Gehalt*Faktor,2)", "grund", FMT_EUR, "4.676,40 € при полном месяце")
aline("Оплата переработки", "Überstundenvergütung", f"=ROUND(D{AR['ueh']}*Lohn*(1+UeZuschlag),2)", "uev", FMT_EUR, "часы × 27,00 € (надбавка за переработку договором не предусмотрена)")
aline("Надбавки всего", "Zeitzuschläge gesamt", f"=Dienstplan!Q{TRr}", "zus", FMT_EUR, "ночь / воскресенье / праздник; при совпадении только высшая")
aline("   из них без подоходного налога", "   steuerfrei § 3b EStG", f"=Dienstplan!R{TRr}", "zusstf", FMT_EUR, "")
aline("   из них без соцвзносов", "   SV-frei § 1 SvEV", f"=Dienstplan!S{TRr}", "zussvf", FMT_EUR, "меньше, т.к. база ограничена 25,00 €/ч")
aline("   облагается налогом", "   steuerpflichtiger Anteil", f"=D{AR['zus']}-D{AR['zusstf']}", "zusstp", FMT_EUR, "")
aline("   облагается взносами", "   SV-pflichtiger Anteil", f"=D{AR['zus']}-D{AR['zussvf']}", "zussvp", FMT_EUR, "")
aline("Премии за Eingesprungen", "Einspringprämien", f"=ROUND(D{AR['sprn']}*Praemie,2)", "prae", FMT_EUR, "облагается полностью; сумму впиши в «Настройки»")
aline("БРУТТО ВСЕГО", "GESAMTBRUTTO", f"=D{AR['grund']}+D{AR['uev']}+D{AR['zus']}+D{AR['prae']}", "brutto", FMT_EUR, "", big=True)
aline("Если переработку НЕ оплатят — потеря", "entgangene Überstundenvergütung", f"=D{AR['uev']}", "fehl", FMT_EUR,
      "Требовать письменно в течение 3 месяцев (§ 14 Ausschlussfrist). Утверждённый Dienstplan = angeordnete Überstunden")
aline("Налогооблагаемое брутто", "steuerpflichtiges Brutto", f"=D{AR['grund']}+D{AR['uev']}+D{AR['prae']}+D{AR['zusstp']}", "stbrutto", FMT_EUR, "база подоходного налога")
aline("База для соцвзносов", "SV-Brutto", f"=D{AR['grund']}+D{AR['uev']}+D{AR['prae']}+D{AR['zussvp']}", "svbrutto", FMT_EUR, "база KV/PV/RV/AV")
aspace()

asec("C. СОЦВЗНОСЫ (доля работника) / SOZIALVERSICHERUNG AN-ANTEIL")
aline("Медицинское страхование", "Krankenversicherung", f"=ROUND(MIN(D{AR['svbrutto']},BBGKV)*(KVSatz/2+KVZusatz/2),2)", "kv", FMT_EUR, "7,3 % + половина доп. взноса")
aline("Страхование по уходу", "Pflegeversicherung", f'=ROUND(MIN(D{AR["svbrutto"]},BBGKV)*(PVSatz/2+IF(Kinderlos="ja",PVZuschlag,0)),2)', "pv", FMT_EUR, "1,8 % + 0,6 % надбавка бездетным")
aline("Пенсионное страхование", "Rentenversicherung", f"=ROUND(MIN(D{AR['svbrutto']},BBGRV)*RVSatz/2,2)", "rv", FMT_EUR, "9,3 %")
aline("Страхование по безработице", "Arbeitslosenversicherung", f"=ROUND(MIN(D{AR['svbrutto']},BBGRV)*AVSatz/2,2)", "av", FMT_EUR, "1,3 %")
aline("СОЦВЗНОСЫ ВСЕГО", "SV gesamt", f"=D{AR['kv']}+D{AR['pv']}+D{AR['rv']}+D{AR['av']}", "svges", FMT_EUR, "")
aspace()

asec("D. НАЛОГИ / STEUERN")
aline("Подоходный налог", "Lohnsteuer", "=0", "lst", FMT_EUR, "оценка по методу ELStAM; налоговый класс с листа «Настройки»")
aline("Надбавка солидарности", "Solidaritätszuschlag", "=0", "soli", FMT_EUR, "при таком доходе = 0 (порог 20.350 €/год)")
aline("Церковный налог", "Kirchensteuer", "=0", "kist", FMT_EUR, "9 % от подоходного налога, если состоишь в церкви")
aline("НАЛОГИ ВСЕГО", "Steuern gesamt", f"=D{AR['lst']}+D{AR['soli']}+D{AR['kist']}", "stges", FMT_EUR, "")
aspace()

asec("E. НА РУКИ / AUSZAHLUNG")
aline("НЕТТО (на счёт)", "NETTOVERDIENST", f"=D{AR['brutto']}-D{AR['svges']}-D{AR['stges']}", "netto", FMT_EUR,
      "Выплата до 15-го числа следующего месяца (§ 3 Abs. 3)", big=True)
aline("Нетто за час фактической работы", "Nettostundenlohn", f'=IF(D{AR["ist"]}=0,0,ROUND(D{AR["netto"]}/D{AR["ist"]},2))', "nsl", FMT_EUR, "")
aline("Доля отчислений", "Abzugsquote", f"=IF(D{AR['brutto']}=0,0,(D{AR['svges']}+D{AR['stges']})/D{AR['brutto']})", "qu", FMT_PCT, "")
aspace()

asec("F. СПРАВОЧНО: расходы работодателя / Arbeitgeberkosten")
aline("KV работодатель", "AG-Anteil KV", f"=ROUND(MIN(D{AR['svbrutto']},BBGKV)*(KVSatz/2+KVZusatz/2),2)", "agkv", FMT_EUR, "")
aline("PV работодатель", "AG-Anteil PV", f"=ROUND(MIN(D{AR['svbrutto']},BBGKV)*PVSatz/2,2)", "agpv", FMT_EUR, "")
aline("RV работодатель", "AG-Anteil RV", f"=ROUND(MIN(D{AR['svbrutto']},BBGRV)*RVSatz/2,2)", "agrv", FMT_EUR, "")
aline("AV работодатель", "AG-Anteil AV", f"=ROUND(MIN(D{AR['svbrutto']},BBGRV)*AVSatz/2,2)", "agav", FMT_EUR, "")
aline("Общие расходы работодателя", "Gesamtkosten Arbeitgeber", f"=D{AR['brutto']}+D{AR['agkv']}+D{AR['agpv']}+D{AR['agrv']}+D{AR['agav']}", "agges", FMT_EUR, "без U1/U2 и страхования от несчастных случаев")
aspace()

asec("G. ПРОВЕРКА ПО ЗАКОНУ / PLAUSIBILITÄT")
aline("Ср. часов в неделю", "Ø Wochenarbeitszeit", f"=ROUND(D{AR['ist']}/(Tage/7),2)", "wk", '0.00',
      "§ 3 ArbZG: в среднем не более 48 ч/нед за 6 месяцев")
aline("Смен длиннее 10 ч", "Dienste > 10 h", '=COUNTIF(Dienstplan!L5:L35,">10")', "l10", '0',
      "§ 3 ArbZG: максимум 10 ч, и только с компенсацией до среднего 8 ч")
aline("Нарушений отдыха 11 ч", "Ruhezeit < 11 h", '=COUNTIFS(Dienstplan!T5:T35,"<11",Dienstplan!T5:T35,">0")', "ru", '0',
      "§ 5 ArbZG; в уходе допустимо 10 ч, если в течение месяца компенсируется сменой отдыха 12 ч")
aline("Смен, задевающих воскресенье", "Sonntage gearbeitet", '=SUMPRODUCT((Dienstplan!N5:N35>0)*1)', "so", '0',
      "§ 11 ArbZG: минимум 15 воскресений в году свободны; за отработанное воскресенье — Ersatzruhetag в течение 2 недель")
aline("Ставка ≥ Pflegemindestlohn?", "Pflegemindestlohn 21,03 €", '=IF(Lohn>=21.03,"да / ja","НЕТ — проверить!")', "pml", None,
      "7. PflegeArbbV, Pflegefachkraft с 01.07.2026")

for r in range(4, ptr[0]):
    ab.cell(row=r, column=2).alignment = Alignment(indent=1, vertical="center")

# =====================================================================
# 6. LOHNSTEUER
# =====================================================================
ls = wb.create_sheet("Lohnsteuer")
ls.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGH", [2, 44, 40, 16, 54, 3, 14, 16]):
    ls.column_dimensions[col].width = w
title_row(ls, 1, "ПОДОХОДНЫЙ НАЛОГ / LOHNSTEUERBERECHNUNG 2026 (оценка)", 5)
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

def tarif(x):
    return (f'IF({x}<=GFB,0,'
            f'IF({x}<=Z1,ROUNDDOWN(($G$4*(({x}-GFB)/10000)+$G$5)*(({x}-GFB)/10000),0),'
            f'IF({x}<=Z2,ROUNDDOWN(($G$6*(({x}-Z1)/10000)+$G$7)*(({x}-Z1)/10000)+$G$8,0),'
            f'IF({x}<=Z3,ROUNDDOWN($G$9*{x}-$G$10,0),'
            f'ROUNDDOWN($G$11*{x}-$G$12,0)))))')

lsec("1. ИСХОДНЫЕ ДАННЫЕ")
lline("Налогооблагаемое брутто за месяц", "steuerpflichtiger Monatslohn", f"='Расчёт'!D{AR['stbrutto']}", "mon", FMT_EUR,
      "оклад + переработка + премии + облагаемая часть надбавок")
lline("Прогноз на год", "hochgerechneter Jahresarbeitslohn", f"=ROUND(D{LR['mon']}*12,2)", "jahr", FMT_EUR,
      "как в ELStAM: текущий месяц × 12. В месяц с большой переработкой налог удерживают завышенно — часть вернётся по Steuererklärung")
lline("Налоговый класс", "Steuerklasse", "=StKl", "stkl", '0', "")
lptr[0] += 1

lsec("2. ВЫЧЕТЫ / FREIBETRÄGE (в год)")
lline("Вычет на работника", "Arbeitnehmer-Pauschbetrag", "=IF(StKl<6,ANPausch,0)", "anp", FMT_EUR, "§ 9a Nr. 1a EStG")
lline("Особые расходы", "Sonderausgaben-Pauschbetrag", "=IF(StKl<6,SAPausch,0)", "sap", FMT_EUR, "§ 10c EStG")
lline("Одинокому родителю", "Entlastungsbetrag Alleinerziehende", "=IF(StKl=2,EntlastAE,0)", "ent", FMT_EUR, "только класс 2")
lptr[0] += 1

lsec("3. VORSORGEPAUSCHALE (§ 39b Abs. 2 S. 5 Nr. 3 EStG)")
lline("Пенсионное страхование", "Teilbetrag Rentenversicherung", f"=ROUND(MIN(D{LR['jahr']},BBGRV*12)*RVSatz/2,2)", "vrv", FMT_EUR, "100 % доли работника")
lline("Медицинское страхование", "Teilbetrag Krankenversicherung", f"=ROUND(MIN(D{LR['jahr']},BBGKV*12)*(KVSatz/2+KVZusatz/2),2)", "vkv", FMT_EUR, "")
lline("Страхование по уходу", "Teilbetrag Pflegeversicherung", f'=ROUND(MIN(D{LR["jahr"]},BBGKV*12)*(PVSatz/2+IF(Kinderlos="ja",PVZuschlag,0)),2)', "vpv", FMT_EUR, "")
lline("Страхование по безработице", "Teilbetrag Arbeitslosenversicherung",
      f'=IF(AVinVP="ja",ROUND(MIN(MIN(D{LR["jahr"]},BBGRV*12)*AVSatz/2,MAX(0,1900-D{LR["vkv"]}-D{LR["vpv"]})),2),0)', "vav", FMT_EUR,
      "новое с 2026; засчитывается только в пределах 1.900 € вместе с KV/PV")
lline("Vorsorgepauschale всего", "Vorsorgepauschale gesamt", f"=D{LR['vrv']}+D{LR['vkv']}+D{LR['vpv']}+D{LR['vav']}", "vges", FMT_EUR, "")
lptr[0] += 1

lsec("4. ТАРИФ / TARIF")
lline("Налогооблагаемый доход", "zu versteuerndes Einkommen (zvE)",
      f"=MAX(0,ROUNDDOWN(D{LR['jahr']}-D{LR['anp']}-D{LR['sap']}-D{LR['ent']}-D{LR['vges']},0))", "zve", FMT_EUR,
      "округляется вниз до полного евро")
lline("Основной тариф", "Grundtarif § 32a Abs. 1", f"={tarif('D' + str(LR['zve']))}", "gt", FMT_EUR, "классы 1, 2, 4")
lline("Тариф со сплиттингом", "Splittingtarif § 32a Abs. 5", f"=2*({tarif('ROUNDDOWN(D' + str(LR['zve']) + '/2,0)')})", "sp", FMT_EUR, "класс 3")
lline("Подоходный налог за год", "Jahreslohnsteuer", f"=IF(StKl=3,D{LR['sp']},D{LR['gt']})", "jlst", FMT_EUR, "")
lline("ПОДОХОДНЫЙ НАЛОГ ЗА МЕСЯЦ", "Lohnsteuer monatlich", f"=ROUND(D{LR['jlst']}/12,2)", "mlst", FMT_EUR, "", big=True)
lptr[0] += 1

lsec("5. SOLI И ЦЕРКОВНЫЙ НАЛОГ")
lline("Soli за год", "Solidaritätszuschlag Jahr",
      f"=IF(D{LR['jlst']}<=SoliFrei,0,ROUND(MIN(0.055*D{LR['jlst']},0.119*(D{LR['jlst']}-SoliFrei)),2))", "soliJ", FMT_EUR,
      "порог 20.350 €, далее зона смягчения 11,9 %")
lline("Soli за месяц", "Solidaritätszuschlag monatlich", f"=ROUND(D{LR['soliJ']}/12,2)", "soliM", FMT_EUR, "")
lline("Церковный налог за месяц", "Kirchensteuer monatlich", f'=IF(KiSt="ja",ROUND(D{LR["mlst"]}*KiStSatz,2),0)', "kistM", FMT_EUR, "")
lptr[0] += 1

wr = lptr[0]
ls.merge_cells(start_row=wr, start_column=2, end_row=wr, end_column=5)
wc = ls.cell(row=wr, column=2, value=('=IF(OR(StKl=5,StKl=6),'
    '"ВНИМАНИЕ: для налоговых классов 5 и 6 действует другой алгоритм (§ 39b Abs. 2 S. 7 EStG) — этот расчёт НЕ подходит.",'
    '"Расчёт — оценка. Официальная сумма только в Lohnabrechnung работодателя. Сверить можно на bmf-steuerrechner.de")'))
wc.font = F_B; wc.fill = FILL_WARN
wc.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ls.row_dimensions[wr].height = 30

ab.cell(row=AR['lst'],  column=4, value=f"=Lohnsteuer!D{LR['mlst']}")
ab.cell(row=AR['soli'], column=4, value=f"=Lohnsteuer!D{LR['soliM']}")
ab.cell(row=AR['kist'], column=4, value=f"=Lohnsteuer!D{LR['kistM']}")

# =====================================================================
# 7. КОНТРОЛЬ ArbZG
# =====================================================================
kz = wb.create_sheet("Контроль ArbZG")
kz.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGH", [2, 16, 16, 12, 10, 14, 44, 20]):
    kz.column_dimensions[col].width = w
title_row(kz, 1, "КОНТРОЛЬ РАБОЧЕГО ВРЕМЕНИ / ARBEITSZEITGESETZ", 8)
kz["B2"] = "Считаются только смены, НАЧИНАЮЩИЕСЯ в этом месяце. Стыки месяцев проверь вручную."
kz["B2"].font = F_S

for i, h in enumerate(["Неделя с", "по", "Часов", "Смен", "Ночных смен", "Статус"], start=2):
    c = kz.cell(row=4, column=i, value=h); c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center")

for i in range(6):
    r = 5 + i
    kz.cell(row=r, column=2, value=f"=DATE(Jahr,Monat,1)-WEEKDAY(DATE(Jahr,Monat,1),2)+1+7*{i}").number_format = FMT_D
    kz.cell(row=r, column=3, value=f"=B{r}+6").number_format = FMT_D
    kz.cell(row=r, column=4, value=(f'=SUMIFS(Dienstplan!$L$5:$L$35,Dienstplan!$A$5:$A$35,">="&B{r},'
                                    f'Dienstplan!$A$5:$A$35,"<="&C{r})')).number_format = '0.00'
    kz.cell(row=r, column=5, value=(f'=COUNTIFS(Dienstplan!$A$5:$A$35,">="&B{r},Dienstplan!$A$5:$A$35,"<="&C{r},'
                                    f'Dienstplan!$L$5:$L$35,">0")')).number_format = '0'
    kz.cell(row=r, column=6, value=(f'=COUNTIFS(Dienstplan!$A$5:$A$35,">="&B{r},Dienstplan!$A$5:$A$35,"<="&C{r},'
                                    f'Dienstplan!$D$5:$D$35,"ND")')).number_format = '0'
    kz.cell(row=r, column=7, value=f'=IF(D{r}>48,"больше 48 ч — превышение (§ 3 ArbZG)",IF(D{r}>40,"выше нормы 40 ч","в норме"))')
    for c in range(2, 8):
        cc = kz.cell(row=r, column=c)
        cc.border = BOX; cc.font = F_N; cc.fill = FILL_CALC
        cc.alignment = Alignment(horizontal="center")
kz.conditional_formatting.add("D5:D10", CellIsRule(operator="greaterThan", formula=["48"], fill=FILL_WARN, font=Font(color="9C0006", bold=True)))
kz.conditional_formatting.add("G5:G10", FormulaRule(formula=["$D5>48"], fill=FILL_WARN, font=Font(color="9C0006", bold=True)))

r = 12
section(kz, r, "ЧТО ГОВОРИТ ЗАКОН", 8); r += 1
GESETZ = [
 ("§ 3 ArbZG", "Рабочий день — максимум 8 ч. До 10 ч можно, только если за 6 месяцев (или 24 недели) средняя не выше 8 ч в день, то есть около 48 ч в неделю."),
 ("§ 4 ArbZG", "Перерыв: при работе более 6 ч — 30 мин, более 9 ч — 45 мин. Можно дробить по 15 мин. Работать более 6 ч подряд без перерыва запрещено."),
 ("§ 5 ArbZG", "Отдых между сменами — 11 ч. В уходе допускается 10 ч, если в течение месяца это компенсируется другой сменой отдыха в 12 ч."),
 ("§ 6 Abs. 2 ArbZG", "Для ночных работников смена — максимум 8 ч; до 10 ч только если за 1 календарный месяц (или 4 недели) средняя не выше 8 ч."),
 ("§ 6 Abs. 3 ArbZG", "Право на бесплатное медобследование каждые 3 года (после 50 лет — ежегодно), за счёт работодателя."),
 ("§ 6 Abs. 5 ArbZG", "Право на соразмерную надбавку за ночную работу ИЛИ на оплачиваемые отгулы. Ориентир BAG — 25 %, при постоянной ночной работе суды присуждали до 30 %."),
 ("§ 11 ArbZG", "Минимум 15 воскресений в году должны быть свободны. За работу в воскресенье — Ersatzruhetag в течение 2 недель, за праздник — в течение 8 недель."),
 ("§ 16 ArbZG", "Работодатель обязан фиксировать всё время сверх 8 ч и хранить записи 2 года. Веди свой учёт — он пригодится при споре."),
 ("§ 3b EStG", "Надбавки за ночь/воскресенье/праздник не облагаются подоходным налогом (ночь 25 %, воскресенье 50 %, праздник 125/150 %), база — до 50 €/ч."),
 ("§ 1 SvEV", "От соцвзносов надбавки свободны только с базы до 25 €/ч. При ставке 27 € часть надбавок облагается взносами."),
 ("§ 11 BUrlG", "Отпускные считаются по среднему заработку за последние 13 недель, ВКЛЮЧАЯ надбавки (без сверхурочных). Проверяй, что надбавки учли."),
 ("§ 4 EFZG", "На больничном платят по принципу Lohnausfall — то, что заработал бы, включая надбавки за смены по графику."),
 ("§ 14 Arbeitsvertrag", "Ausschlussfrist 3 месяца. Недоплату надо потребовать письменно (e-mail достаточно) в течение 3 месяцев с даты выплаты, иначе право пропадает."),
]
for para, txt in GESETZ:
    kz.cell(row=r, column=2, value=para).font = F_B
    c = kz.cell(row=r, column=3, value=txt); c.font = F_N
    c.alignment = Alignment(wrap_text=True, vertical="top")
    kz.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
    kz.row_dimensions[r].height = max(15, 13 * (1 + len(txt) // 110))
    r += 1

# =====================================================================
# 8. ГОД
# =====================================================================
jr = wb.create_sheet("Год")
jr.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGH", [2, 16, 12, 14, 14, 14, 14, 14]):
    jr.column_dimensions[col].width = w
title_row(jr, 1, "СВОДКА ЗА ГОД / JAHRESÜBERSICHT", 8)
jr["B2"] = "Каждый месяц: сделай копию файла, посчитай и перенеси итоги сюда."
jr["B2"].font = F_S
for i, h in enumerate(["Месяц", "Часы", "Брутто", "Надбавки", "Соцвзносы", "Налоги", "Нетто"], start=2):
    c = jr.cell(row=4, column=i, value=h); c.font = F_H1; c.fill = FILL_HEAD; c.border = BOX
    c.alignment = Alignment(horizontal="center")
MON = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]
for i, m in enumerate(MON):
    r = 5 + i
    jr.cell(row=r, column=2, value=m).font = F_B
    jr.cell(row=r, column=2).border = BOX
    for c in range(3, 9):
        cc = jr.cell(row=r, column=c)
        cc.border = BOX; cc.fill = FILL_IN; cc.font = F_N
        cc.number_format = '0.00' if c == 3 else FMT_EUR
        cc.alignment = Alignment(horizontal="center")
r = 17
jr.cell(row=r, column=2, value="ИТОГО ЗА ГОД").font = F_H1
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
  "· Проверить, что срок Ausschlussfrist (3 месяца) нигде не пропущен.",
  "· Понять, выгодна ли налоговая декларация (Steuererklärung) — при переменном доходе обычно возвращают деньги.",
  "· Контроль среднего рабочего времени за 6 месяцев (§ 3 ArbZG: не более 48 ч/нед в среднем).",
  "· База для расчёта отпускных (§ 11 BUrlG — среднее за 13 недель, включая надбавки).",
]):
    jr.cell(row=20 + i, column=2, value=t).font = F_N
    jr.merge_cells(start_row=20 + i, start_column=2, end_row=20 + i, end_column=8)

wb.active = 0
wb.save(OUT)
print("OK ->", OUT)
