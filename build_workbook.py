# -*- coding: utf-8 -*-
"""Сборка расчётной записки (Excel) по кейсу «Карьерная вода без остановки».

Запуск:  python build_workbook.py Расчёт_водоотлив.xlsx
Все результаты в книге — формулы; значения появляются после пересчёта
(Excel пересчитывает сам при открытии, для проверки — LibreOffice).
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, ScatterChart, Series
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import absolute_coordinate
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "Расчёт_водоотлив.xlsx"

FN = "Arial"
F_NORM = Font(name=FN, size=10)
F_BOLD = Font(name=FN, size=10, bold=True)
F_TITLE = Font(name=FN, size=14, bold=True)
F_H2 = Font(name=FN, size=11, bold=True)
F_IN = Font(name=FN, size=10, color="0000FF")          # исходные данные
F_LINK = Font(name=FN, size=10, color="008000")        # ссылка на другой лист
F_NOTE = Font(name=FN, size=9, italic=True, color="595959")
FILL_DEC = PatternFill("solid", fgColor="FFFF00")      # решения команды
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")
FILL_SUM = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

SRC_PDF = "Постановка кейса (PDF)"
SRC_HYD = "PDF разд. 3; Гидравлическая_база.csv"
SRC_PR = "PDF разд. 5; Цены.csv"

wb = Workbook()
names_ws = {}


def sheet(title):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    return ws


wb.remove(wb.active)
ws_info = sheet("Описание")
ws_chk = sheet("Проверки")
ws_sch = sheet("Схема")
ws_in = sheet("Исходные")
ws_dec = sheet("Решения")
ws_p = sheet("Насосы")
ws_st = sheet("Накопитель")
ws_bA = sheet("Баланс_A")
ws_bB = sheet("Баланс_B")
ws_tr = sheet("Очистка")
ws_ec = sheet("Экономика")
ws_pl = sheet("Пилот")


def R(ws, coord):
    """Абсолютная ссылка на ячейку другого листа."""
    return f"'{ws.title}'!{absolute_coordinate(coord)}"


def defname(nm, ws, coord):
    wb.defined_names[nm] = DefinedName(nm, attr_text=R(ws, coord))


def put(ws, coord, value, font=F_NORM, fmt=None, fill=None, align=None, border=False):
    c = ws[coord]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    if border:
        c.border = BORDER
    return c


def header(ws, row, labels, col0=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col0 + i, value=lab)
        c.font = F_BOLD
        c.fill = FILL_HEAD
        c.alignment = CENTER
        c.border = BORDER


def title(ws, text, sub=None):
    put(ws, "A1", text, F_TITLE)
    if sub:
        put(ws, "A2", sub, F_NOTE)


def widths(ws, w):
    for col, val in w.items():
        ws.column_dimensions[col].width = val


# ---------------------------------------------------------------- Исходные
title(ws_in, "Исходные данные кейса",
      "Синий шрифт — значения из постановки (единые учебные условия, не измерения разреза). "
      "Каждой величине присвоено имя — по нему её используют формулы на других листах.")
header(ws_in, 4, ["Имя в формулах", "Показатель", "Значение", "Единица", "Источник / пояснение"])
inputs = [
    ("Нагрузки", None),
    ("Q_cur", "Текущий средний приток", 450, "м³/ч", SRC_HYD),
    ("Q_fut", "Будущий средний приток", 1000, "м³/ч", SRC_HYD),
    ("Q_pk", "Суточный пик: расход в пиковые часы", 1200, "м³/ч", SRC_HYD),
    ("T_pk", "Суточный пик: длительность пиковой части", 6, "ч", SRC_HYD),
    ("Q_opk", "Суточный пик: расход в остальные 18 ч", 933.333333, "м³/ч", SRC_HYD + "; среднее за сутки = 1000"),
    ("T_fail", "Отдельный полный отказ откачки: длительность", 2, "ч", "Не совмещается с пиком"),
    ("Q_fail", "Приток во время отказа и восстановления", 1000, "м³/ч", SRC_HYD),
    ("T_rec", "Допустимое время восстановления после отказа", 24, "ч", "PDF разд. 3"),
    ("Гидравлика и накопитель", None),
    ("H_st", "Статический напор сети", 60, "м", SRC_HYD),
    ("K_net", "Коэффициент сопротивления сети K", 0.000025, "м/(м³/ч)²", "H сети = H_st + K·Q²"),
    ("V_sl", "Объём осадка в накопителе", 400, "м³", SRC_HYD),
    ("V_dead", "Недоступный объём", 200, "м³", SRC_HYD),
    ("k_fb", "Свободный борт, доля полного объёма", 0.15, "доля", SRC_HYD),
    ("rho", "Плотность воды", 1000, "кг/м³", SRC_HYD),
    ("g_acc", "Ускорение свободного падения", 9.81, "м/с²", SRC_HYD),
    ("Area_lim", "Предел занимаемой площади накопителя", 1200, "м²", "PDF разд. 5"),
    ("Depth_lim", "Предел учебной глубины полного объёма", 5, "м", "PDF разд. 5"),
    ("Время и вода", None),
    ("H_yr", "Часов в году", 8760, "ч/год", SRC_HYD),
    ("H_cons", "Часы работы потребителя повторной воды", 6000, "ч/год", SRC_HYD),
    ("Q_reuse", "Спрос потребителя на повторную воду", 300, "м³/ч", SRC_HYD),
    ("C_reuse", "Цель по взвеси для повторной воды", 30, "мг/л", SRC_HYD),
    ("C_raw", "Сырая вода: взвешенные вещества", 500, "мг/л", SRC_HYD),
    ("pH_raw", "Сырая вода: pH", 6.8, "ед. pH", SRC_HYD),
    ("SO4_raw", "Сырая вода: сульфаты", 700, "мг/л", SRC_HYD),
    ("Hard_raw", "Сырая вода: жёсткость", 500, "мг/л (учебный базис)", SRC_HYD),
    ("C_dis", "Контроль взвеси оставшегося потока перед условным приёмником", 100, "мг/л",
     "PDF разд. 4: контроль сценария, не норматив сброса"),
    ("Экономический базис", None),
    ("tariff", "Электроэнергия", 7000, "руб./МВт·ч", SRC_PR),
    ("P_water", "Покупная свежая вода", 10, "руб./м³", SRC_PR),
    ("k_maint", "Обслуживание, доля всего выбранного CAPEX в год", 0.02, "1/год", SRC_PR),
    ("P_stor", "Полный накопитель", 0.002, "млн руб./м³ полного объёма", SRC_PR),
    ("CAPEX_main", "Магистраль/подключение повторной воды", 15, "млн руб.", SRC_PR),
    ("Pilot_lim", "Предел пилота", 100, "млн руб.", SRC_PR),
    ("Horizon", "Горизонт сравнения", 5, "лет", "Постоянные цены, без НДС и дисконтирования"),
]
r = 5
for item in inputs:
    if item[1] is None:
        put(ws_in, f"A{r}", item[0], F_H2)
        r += 1
        continue
    nm, lab, val, unit, src = item
    put(ws_in, f"A{r}", nm, F_NOTE, border=True)
    put(ws_in, f"B{r}", lab, border=True, align=WRAP)
    fmt = "0.000000" if nm == "K_net" else ("#,##0.######" if isinstance(val, float) else "#,##0")
    put(ws_in, f"C{r}", val, F_IN, fmt, border=True)
    put(ws_in, f"D{r}", unit, border=True)
    put(ws_in, f"E{r}", src, F_NOTE, border=True, align=WRAP)
    defname(nm, ws_in, f"C{r}")
    r += 1

r += 1
put(ws_in, f"A{r}", "Насосы (характеристика H = H₀ − a·(Q/n)²)", F_H2)
r += 1
header(ws_in, r, ["ID", "H₀, м", "a, м/(м³/ч)²", "q_min, м³/ч", "q_max, м³/ч", "КПД η", "Цена, млн руб."])
pump_rows = [("P1", 130, 8e-05, 200, 1000, 0.75, 8), ("P2", 165, 6e-05, 250, 1200, 0.8, 14)]
p0 = r + 1
for i, row in enumerate(pump_rows):
    rr = p0 + i
    for j, v in enumerate(row):
        col = "ABCDEFG"[j]
        fmt = "0.00000" if j == 2 else ("0.00" if j == 5 else "#,##0")
        put(ws_in, f"{col}{rr}", v, F_IN, fmt, border=True)
p1 = p0 + len(pump_rows) - 1
for col, nm in zip("ABCDEFG", ["Pump_ID", "Pump_H0", "Pump_a", "Pump_qmin", "Pump_qmax", "Pump_eta", "Pump_cost"]):
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{ws_in.title}'!${col}${p0}:${col}${p1}")
put(ws_in, f"A{p1 + 1}", "Источник: PDF разд. 4; Отдельные_насосы.csv. Диапазон q относится к одному включённому насосу.", F_NOTE)

r = p1 + 3
put(ws_in, f"A{r}", "Операции очистки (эффективность — от входа операции, потерь воды нет)", F_H2)
r += 1
header(ws_in, r, ["ID", "Операция", "Предел, м³/ч", "Удаление взвеси", "CAPEX, млн руб.", "кВт·ч/м³", "Прочие, руб./м³"])
tr_rows = [("S", "Отстаивание", 1500, 0.9, 20, 0.01, 2),
           ("F", "Фильтрация", 400, 0.8, 6, 0.04, 3),
           ("M", "Мембранная / тонкая очистка взвеси", 400, 0.96, 10, 0.15, 4)]
t0 = r + 1
for i, row in enumerate(tr_rows):
    rr = t0 + i
    for j, v in enumerate(row):
        col = "ABCDEFG"[j]
        fmt = "0.00" if j in (3, 5) else "#,##0"
        put(ws_in, f"{col}{rr}", v, F_IN if j != 1 else F_NORM, fmt if j != 1 else None, border=True)
t1 = t0 + len(tr_rows) - 1
for col, nm in zip("ACDEFG", ["Tr_ID", "Tr_lim", "Tr_eff", "Tr_capex", "Tr_kwh", "Tr_oth"]):
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{ws_in.title}'!${col}${t0}:${col}${t1}")
for col, nm in zip("CDEFG", ["S_lim", "S_eff", "S_capex", "S_kwh", "S_oth"]):
    defname(nm, ws_in, f"{col}{t0}")
put(ws_in, f"A{t1 + 1}", "Источник: PDF разд. 4; Операции_очистки.csv. S/F/M не меняют pH, сульфаты и жёсткость.", F_NOTE)
widths(ws_in, {"A": 12, "B": 52, "C": 14, "D": 22, "E": 46, "F": 12, "G": 14})

# ---------------------------------------------------------------- Решения
title(ws_dec, "Проектные решения команды",
      "Жёлтые ячейки — то, что выбирает команда. Меняйте их — вся книга пересчитается, "
      "а лист «Проверки» покажет, остаются ли выполненными требования кейса.")
header(ws_dec, 4, ["Имя в формулах", "Решение", "Значение", "Единица", "Обоснование"])
decisions = [
    ("Водоотлив: конфигурация A (основная)", None),
    ("A_type", "Тип насоса", "P2", "—", "Один P2 даёт 1111 м³/ч > 1000: резерв 100 % одним агрегатом"),
    ("A_N", "Установлено агрегатов", 2, "шт.", "1 рабочий + 1 резерв"),
    ("A_res", "Из них в резерве", 1, "шт.", ""),
    ("Водоотлив: конфигурация B (альтернатива)", None),
    ("B_type", "Тип насоса", "P1", "—", "Один P1 даёт 816 м³/ч < 1000, нужно два рабочих"),
    ("B_N", "Установлено агрегатов", 3, "шт.", "2 рабочих + 1 резерв"),
    ("B_res", "Из них в резерве", 1, "шт.", ""),
    ("Накопитель и уровни", None),
    ("V_use", "Полезный регулирующий объём", 2500, "м³",
     "Максимальное накопление по сценариям — 2000 м³ (отказ 2 ч × 1000 м³/ч), плюс запас 25 %"),
    ("Depth", "Глубина полного объёма", 4.5, "м", "Не более 5 м по условию; 0,5 м оставлено в запасе"),
    ("W_acc", "Полоса обслуживания/подъезда по периметру", 2, "м", "Доступ для очистки осадка и обслуживания"),
    ("V_on1", "Пуск базовой ступени: объём над уровнем «стоп»", 100, "м³",
     "Для алгоритма; часовая модель учитывает циклический режим долей работы"),
    ("V_L2", "Пуск повышенной ступени: объём над уровнем «стоп»", 500, "м³",
     "Выше пика конфигурации A (532 м³ не достигается на старте часа) и ниже накопления при отказе"),
    ("k_alarm", "Аварийный уровень, доля полезного объёма", 0.9, "доля", "Сигнал диспетчеру, проверка резерва"),
    ("Очистка и пилот", None),
    ("Reuse_tech", "Технология доочистки ветки повторной воды", "F",
     "—", "F: 50→10 мг/л при цели 30 — дешевле M при достаточном качестве (см. «Экономика»)"),
    ("sl_frac", "Доля твёрдого в уплотнённом осадке", 0.3, "доля",
     "Собственное допущение для оценки объёма шлама; уточнить на пилоте"),
    ("sl_rho", "Плотность уплотнённого осадка", 1200, "кг/м³", "Собственное допущение"),
    ("KIP", "Пилот: КИПиА (уровнемеры, расходомеры, мутномеры, связь)", 3, "млн руб.",
     "Собственная оценка, уточнить по коммерческим предложениям"),
    ("Test", "Пилот: испытания и лабораторный контроль", 2, "млн руб.", "Собственная оценка"),
]
r = 5
dec_cells = {}
for item in decisions:
    if item[1] is None:
        put(ws_dec, f"A{r}", item[0], F_H2)
        r += 1
        continue
    nm, lab, val, unit, why = item
    put(ws_dec, f"A{r}", nm, F_NOTE, border=True)
    put(ws_dec, f"B{r}", lab, border=True, align=WRAP)
    fmt = None if isinstance(val, str) else ("0.0#" if isinstance(val, float) else "#,##0")
    put(ws_dec, f"C{r}", val, F_IN, fmt, FILL_DEC, border=True)
    put(ws_dec, f"D{r}", unit, border=True)
    put(ws_dec, f"E{r}", why, F_NOTE, border=True, align=WRAP)
    defname(nm, ws_dec, f"C{r}")
    dec_cells[nm] = f"C{r}"
    r += 1
dv_p = DataValidation(type="list", formula1='"P1,P2"', allow_blank=False)
dv_t = DataValidation(type="list", formula1='"F,M"', allow_blank=False)
ws_dec.add_data_validation(dv_p)
ws_dec.add_data_validation(dv_t)
dv_p.add(dec_cells["A_type"])
dv_p.add(dec_cells["B_type"])
dv_t.add(dec_cells["Reuse_tech"])
widths(ws_dec, {"A": 12, "B": 50, "C": 12, "D": 12, "E": 70})

# ---------------------------------------------------------------- Насосы
title(ws_p, "Насосы: рабочие точки, мощность, резерв",
      "n одинаковых насосов параллельно: H₀ − a·(Q/n)² = 60 + 0,000025·Q². "
      "Мощность = ρ·g·Q·H/(3600·η). Синие/жёлтые входы — на листах «Исходные» и «Решения».")
PB = {}


def pump_block(top, cfg):
    t, n_, rs = f"{cfg}_type", f"{cfg}_N", f"{cfg}_res"
    put(ws_p, f"A{top}", f"Конфигурация {cfg}", F_H2)
    rows = [
        ("Тип насоса", f"={t}", None),
        ("H₀, м", f"=INDEX(Pump_H0,MATCH($B${top + 1},Pump_ID,0))", "0.0"),
        ("a, м/(м³/ч)²", f"=INDEX(Pump_a,MATCH($B${top + 1},Pump_ID,0))", "0.00000"),
        ("q_min одного насоса, м³/ч", f"=INDEX(Pump_qmin,MATCH($B${top + 1},Pump_ID,0))", "#,##0"),
        ("q_max одного насоса, м³/ч", f"=INDEX(Pump_qmax,MATCH($B${top + 1},Pump_ID,0))", "#,##0"),
        ("КПД η", f"=INDEX(Pump_eta,MATCH($B${top + 1},Pump_ID,0))", "0.00"),
        ("Цена агрегата, млн руб.", f"=INDEX(Pump_cost,MATCH($B${top + 1},Pump_ID,0))", "#,##0.0"),
        ("Установлено, шт.", f"={n_}", "0"),
        ("Из них резерв, шт.", f"={rs}", "0"),
    ]
    for i, (lab, f, fmt) in enumerate(rows):
        put(ws_p, f"A{top + 1 + i}", lab, border=True)
        put(ws_p, f"B{top + 1 + i}", f, F_LINK if i in (0, 7, 8) else F_NORM, fmt, border=True)
    h0, a, qmin, qmax, eta, cost, N, res = (f"$B${top + k}" for k in range(2, 10))
    hr = top + 11
    header(ws_p, hr, ["n рабочих", "Q общий, м³/ч", "q на насос, м³/ч", "H, м", "q в диапазоне",
                      "Мощность, кВт", "кВт·ч/м³", "n ≤ установленных"])
    for n in (1, 2, 3):
        rr = hr + n
        put(ws_p, f"A{rr}", n, fmt="0", border=True, align=CENTER)
        put(ws_p, f"B{rr}", f"=SQRT(({h0}-H_st)/({a}/A{rr}^2+K_net))", fmt="#,##0.0", border=True)
        put(ws_p, f"C{rr}", f"=B{rr}/A{rr}", fmt="#,##0.0", border=True)
        put(ws_p, f"D{rr}", f"=H_st+K_net*B{rr}^2", fmt="0.0", border=True)
        put(ws_p, f"E{rr}", f'=IF(AND(C{rr}>={qmin},C{rr}<={qmax}),"да","НЕТ")', border=True, align=CENTER)
        put(ws_p, f"F{rr}", f"=rho*g_acc*B{rr}*D{rr}/(3600*{eta})/1000", fmt="#,##0", border=True)
        put(ws_p, f"G{rr}", f"=F{rr}/B{rr}", fmt="0.000", border=True)
        put(ws_p, f"H{rr}", f'=IF(A{rr}<={N},"да","нет")', border=True, align=CENTER)
    q1, q2, q3 = (f"$B${hr + k}" for k in (1, 2, 3))
    rng_e = f"$E${hr + 1}:$E${hr + 3}"
    rng_q = f"$B${hr + 1}:$B${hr + 3}"
    s = hr + 5
    res_rows = [
        ("Базовое число рабочих при среднем будущем притоке (наименьшее n с Q(n) ≥ Q_fut)",
         f"=IF({q1}>=Q_fut,1,IF({q2}>=Q_fut,2,IF({q3}>=Q_fut,3,99)))", "0"),
        ("Базовое число рабочих при текущем притоке (наименьшее n с Q(n) ≥ Q_cur)",
         f"=IF({q1}>=Q_cur,1,IF({q2}>=Q_cur,2,IF({q3}>=Q_cur,3,99)))", "0"),
        ("Повышенная ступень: наибольшее n ≤ установленных с подачей ≤ пределу отстойника S",
         f"=IF(AND({N}>=3,{q3}<=S_lim),3,IF(AND({N}>=2,{q2}<=S_lim),2,1))", "0"),
        ("Базовая ступень обходится без резерва (n_баз ≤ N − резерв)",
         f'=IF(B{s}<={N}-{res},"да","НЕТ")', None),
        ("Подача при отказе одного агрегата Q(N−1), м³/ч",
         f"=IF({N}-1<1,0,INDEX({rng_q},MIN(3,{N}-1)))", "#,##0"),
        ("Резерв обеспечивает средний будущий приток (Q(N−1) ≥ Q_fut)",
         f'=IF(B{s + 4}>=Q_fut,"да","НЕТ")', None),
        ("Все используемые рабочие точки в диапазоне q_min…q_max",
         f'=IFERROR(IF(AND(INDEX({rng_e},B{s})="да",INDEX({rng_e},B{s + 1})="да",'
         f'INDEX({rng_e},MAX(B{s},B{s + 2}))="да"),"да","НЕТ"),"НЕТ")', None),
        ("CAPEX насосов, млн руб.", f"={N}*{cost}", "#,##0.0"),
    ]
    for i, (lab, f, fmt) in enumerate(res_rows):
        put(ws_p, f"A{s + i}", lab, border=True, align=WRAP)
        put(ws_p, f"B{s + i}", f, F_BOLD, fmt, FILL_SUM, border=True, align=CENTER)
    PB[cfg] = dict(Q=[f"B{hr + k}" for k in (1, 2, 3)], H=[f"D{hr + k}" for k in (1, 2, 3)],
                   e=[f"G{hr + k}" for k in (1, 2, 3)], nbf=f"B{s}", nbc=f"B{s + 1}", nh=f"B{s + 2}",
                   base_ok=f"B{s + 3}", res_ok=f"B{s + 5}", rng_ok=f"B{s + 6}", capex=f"B{s + 7}",
                   hr=hr)
    return s + 7


end_a = pump_block(4, "A")
end_b = pump_block(end_a + 3, "B")
put(ws_p, f"A{end_b + 2}",
    "Смешанная схема P1+P2 возможна (расходы складываются при общем напоре), но требует подбора напора "
    "итерацией; в базовом расчёте не используется.", F_NOTE)

# кривые для графика
cr = 4
header(ws_p, cr, ["Q, м³/ч", "Сеть", "P1 ×1", "P1 ×2", "P2 ×1", "P2 ×2"], col0=11)
for i in range(0, 21):
    rr = cr + 1 + i
    put(ws_p, f"K{rr}", i * 100, fmt="#,##0", border=True)
    put(ws_p, f"L{rr}", f"=H_st+K_net*K{rr}^2", fmt="0.0", border=True)
    for col, pid, n in (("M", "P1", 1), ("N", "P1", 2), ("O", "P2", 1), ("P", "P2", 2)):
        put(ws_p, f"{col}{rr}",
            f'=INDEX(Pump_H0,MATCH("{pid}",Pump_ID,0))-INDEX(Pump_a,MATCH("{pid}",Pump_ID,0))*(K{rr}/{n})^2',
            fmt="0.0", border=True)
ch = ScatterChart()
ch.title = "Характеристики насосов и сети"
ch.style = 13
ch.x_axis.title = "Q, м³/ч"
ch.y_axis.title = "H, м"
ch.y_axis.scaling.min = 40
ch.y_axis.scaling.max = 180
ch.x_axis.scaling.min = 0
ch.x_axis.scaling.max = 2000
ch.x_axis.delete = False
ch.y_axis.delete = False
xr = Reference(ws_p, min_col=11, min_row=cr + 1, max_row=cr + 21)
for col in range(12, 17):
    yr = Reference(ws_p, min_col=col, min_row=cr, max_row=cr + 21)
    se = Series(yr, xr, title_from_data=True)
    se.marker.symbol = "none"
    se.smooth = True
    ch.series.append(se)
for cfg in ("A", "B"):
    hr = PB[cfg]["hr"]
    se = Series(Reference(ws_p, min_col=4, min_row=hr + 1, max_row=hr + 3),
                Reference(ws_p, min_col=2, min_row=hr + 1, max_row=hr + 3),
                title=f"Рабочие точки {cfg} (n = 1…3)")
    se.marker.symbol = "circle" if cfg == "A" else "triangle"
    se.marker.size = 8
    se.graphicalProperties.line.noFill = True
    ch.series.append(se)
ch.height = 10
ch.width = 18
ws_p.add_chart(ch, "K27")
widths(ws_p, {"A": 58, "B": 14, "C": 14, "D": 10, "E": 13, "F": 13, "G": 11, "H": 15,
              "K": 10, "L": 9, "M": 9, "N": 9, "O": 9, "P": 9})

# ---------------------------------------------------------------- Баланс_A / Баланс_B
HOURS = 26
BL = {}  # адреса итогов


def balance_sheet(ws, cfg):
    title(ws, f"Почасовой баланс накопителя — конфигурация {cfg}",
          "Шаг 1 ч. Насосы откачивают не больше, чем есть в полезном объёме: при малом притоке это "
          "циклический режим, энергия — по фактически откачанному объёму (доля работы). "
          "Каждый сценарий начинается с пустого полезного объёма (уровень «стоп»).")
    pb = PB[cfg]
    put(ws, "A3", "n рабочих", F_BOLD)
    for i, col in enumerate("BCD"):
        put(ws, f"{col}3", i + 1, F_BOLD, "0", align=CENTER)
        put(ws, f"{col}4", f"={R(ws_p, pb['Q'][i])}", F_LINK, "#,##0.0")
        put(ws, f"{col}5", f"={R(ws_p, pb['e'][i])}", F_LINK, "0.000")
    links = [
        ("A4", "Подача Q(n), м³/ч"), ("A5", "Удельная энергия e(n), кВт·ч/м³"),
        ("A6", "Базовое n: при 1000 / при 450 м³/ч"), ("A7", "Повышенная ступень n"),
        ("A8", "Порог пуска повышенной ступени, м³"), ("A9", "Полезный объём, м³"),
        ("A10", "Площадь зеркала, м²"), ("A11", "Предел подачи на отстойник S, м³/ч"),
        ("A12", "Осадок + недоступный объём, м³"),
    ]
    for c, lab in links:
        put(ws, c, lab)
    put(ws, "B6", f"={R(ws_p, pb['nbf'])}", F_LINK, "0")
    put(ws, "C6", f"={R(ws_p, pb['nbc'])}", F_LINK, "0")
    put(ws, "B7", f"={R(ws_p, pb['nh'])}", F_LINK, "0")
    put(ws, "B8", "=V_L2", F_LINK, "#,##0")
    put(ws, "B9", "=V_use", F_LINK, "#,##0")
    put(ws, "B10", "=A_water", F_LINK, "#,##0.0")
    put(ws, "B11", "=S_lim", F_LINK, "#,##0")
    put(ws, "B12", "=V_sl+V_dead", F_LINK, "#,##0")

    scen = [
        ("peak", "Сценарий 1. Суточный пик: 1200 м³/ч × 6 ч, затем 933,3 м³/ч × 18 ч",
         "Требование: возврат к исходному уровню не позднее конца суток (час 24).",
         "=IF(MOD(A{r}-1,24)<T_pk,Q_pk,Q_opk)", "1", "$B$6", "T_pk"),
        ("fail", "Сценарий 2. Полный отказ откачки 2 ч при 1000 м³/ч, затем восстановление",
         "Требование: восстановление уровня за 24 ч после отказа при притоке 1000 м³/ч.",
         "=Q_fail", "=IF(A{r}<=T_fail,0,1)", "$B$6", "T_fail"),
        ("cur", "Сценарий 3. Текущий приток 450 м³/ч (проверка нынешнего режима)",
         "Требование: без перелива, циклический режим в допустимой рабочей точке.",
         "=Q_cur", "1", "$C$6", "0"),
    ]
    BL[cfg] = {}
    for k, (key, name, req, f_in, f_av, nb, t0) in enumerate(scen):
        s = 15 + 45 * k
        put(ws, f"A{s}", name, F_H2)
        put(ws, f"A{s + 1}", req, F_NOTE)
        hdr = s + 2
        header(ws, hdr, ["Час", "Приток, м³/ч", "Насосы доступны", "V нач., м³", "n рабочих",
                         "Подача при работе, м³/ч", "Откачано за час, м³", "V кон., м³", "Доля работы",
                         "Энергия, кВт·ч", "Уровень от дна, м", "Перелив", "Подача ≤ предела S",
                         "Час возврата", "Полезный объём, м³"])
        d0, d1 = hdr + 1, hdr + HOURS
        for h in range(1, HOURS + 1):
            rr = hdr + h
            put(ws, f"A{rr}", h, fmt="0", border=True, align=CENTER)
            put(ws, f"B{rr}", f_in.format(r=rr), fmt="#,##0.0", border=True)
            av = f_av.format(r=rr)
            put(ws, f"C{rr}", int(av) if av.isdigit() else av, fmt="0", border=True, align=CENTER)
            if h == 1:
                c = put(ws, f"D{rr}", 0, F_IN, "#,##0.0", border=True)
                c.comment = Comment("Старт сценария с пустого полезного объёма (условие кейса).", "расчёт")
            else:
                put(ws, f"D{rr}", f"=H{rr - 1}", fmt="#,##0.0", border=True)
            put(ws, f"E{rr}", f"=IF(C{rr}=0,0,IF(D{rr}>=$B$8,MAX($B$7,{nb}),{nb}))", fmt="0", border=True,
                align=CENTER)
            put(ws, f"F{rr}", f"=IF(E{rr}=0,0,INDEX($B$4:$D$4,E{rr}))", fmt="#,##0.0", border=True)
            put(ws, f"G{rr}", f"=MIN(F{rr},D{rr}+B{rr})", fmt="#,##0.0", border=True)
            put(ws, f"H{rr}", f"=D{rr}+B{rr}-G{rr}", fmt="#,##0.0", border=True)
            put(ws, f"I{rr}", f"=IF(F{rr}=0,0,G{rr}/F{rr})", fmt="0%", border=True)
            put(ws, f"J{rr}", f"=IF(E{rr}=0,0,G{rr}*INDEX($B$5:$D$5,E{rr}))", fmt="#,##0", border=True)
            put(ws, f"K{rr}", f"=($B$12+H{rr})/$B$10", fmt="0.00", border=True)
            put(ws, f"L{rr}", f'=IF(H{rr}>$B$9+0.001,"ПЕРЕЛИВ","ок")', border=True, align=CENTER)
            put(ws, f"M{rr}", f'=IF(F{rr}<=$B$11+0.001,"ок","ПРЕВЫШЕНИЕ")', border=True, align=CENTER)
            put(ws, f"N{rr}", f'=IF(AND(A{rr}>{t0},H{rr}<=0.001),A{rr},"")', fmt="0", border=True, align=CENTER)
            put(ws, f"O{rr}", "=$B$9", fmt="#,##0", border=True)
        r24 = hdr + 24
        rngs = {c: f"{c}{d0}:{c}{d1}" for c in "BGHIJKLMN"}
        sm = s + 30
        if key == "peak":
            ok = f'=IF(H{r24}<=0.001,"да","НЕТ")'
        elif key == "fail":
            ok = f'=IF(COUNT({rngs["N"]})=0,"НЕТ",IF(MIN({rngs["N"]})-T_fail<=T_rec,"да","НЕТ"))'
        else:
            ok = f'=IF(AND(H{r24}<=0.001,COUNTIF({rngs["L"]},"ПЕРЕЛИВ")=0),"да","НЕТ")'
        summ = [
            ("Приток за 26 ч, м³", f"=SUM({rngs['B']})", "#,##0"),
            ("Откачано за 26 ч, м³", f"=SUM({rngs['G']})", "#,##0"),
            ("Остаток в полезном объёме на конец, м³", f"=H{d1}", "#,##0.0"),
            ("Невязка баланса (приток − откачка − остаток), м³", f"=SUM({rngs['B']})-SUM({rngs['G']})-H{d1}-D{d0}",
             "0.000"),
            ("Максимальный объём в накопителе, м³", f"=MAX({rngs['H']})", "#,##0"),
            ("Максимальный уровень от дна, м", f"=MAX({rngs['K']})", "0.00"),
            ("Перелив за пределы полезного объёма", f'=IF(COUNTIF({rngs["L"]},"ПЕРЕЛИВ")=0,"нет","ЕСТЬ")', None),
            ("Подача на отстойник S в пределах во все часы", f'=IF(COUNTIF({rngs["M"]},"ПРЕВЫШЕНИЕ")=0,"да","НЕТ")',
             None),
            ("Час возврата к пустому полезному объёму", f'=IF(COUNT({rngs["N"]})=0,"не вернулся",MIN({rngs["N"]}))',
             "0"),
            ("Требование сценария выполнено", ok, None),
            ("Энергия насосов за сутки (часы 1–24), кВт·ч", f"=SUM(J{d0}:J{r24})", "#,##0"),
            ("Средняя доля работы за сутки", f"=AVERAGE(I{d0}:I{r24})", "0%"),
        ]
        if key == "fail":
            summ.append(("Время восстановления после окончания отказа, ч",
                         f'=IF(COUNT({rngs["N"]})=0,"—",MIN({rngs["N"]})-T_fail)', "0"))
        cells = {}
        for i, (lab, f, fmt) in enumerate(summ):
            put(ws, f"A{sm + i}", lab, border=True)
            ws.merge_cells(f"A{sm + i}:D{sm + i}")
            put(ws, f"E{sm + i}", f, F_BOLD, fmt, FILL_SUM, border=True, align=CENTER)
            cells[i] = f"E{sm + i}"
        BL[cfg][key] = dict(resid=cells[3], vmax=cells[4], overflow=cells[6], s_ok=cells[7],
                            ok=cells[9], energy=cells[10], start=f"D{d0}",
                            rec=cells.get(12))
        lc = LineChart()
        lc.title = f"{cfg}: объём в накопителе — {name.split('.')[1].split(':')[0].strip()}"
        lc.y_axis.title = "м³"
        lc.x_axis.title = "час"
        lc.x_axis.delete = False
        lc.y_axis.delete = False
        lc.add_data(Reference(ws, min_col=8, min_row=hdr, max_row=d1), titles_from_data=True)
        lc.add_data(Reference(ws, min_col=15, min_row=hdr, max_row=d1), titles_from_data=True)
        lc.set_categories(Reference(ws, min_col=1, min_row=d0, max_row=d1))
        lc.height = 7.5
        lc.width = 15
        ws.add_chart(lc, f"Q{s}")
    widths(ws, {"A": 9, "B": 12, "C": 11, "D": 11, "E": 13, "F": 13, "G": 13, "H": 11, "I": 9,
                "J": 11, "K": 11, "L": 10, "M": 13, "N": 9, "O": 11})
    ws.column_dimensions["A"].width = 12


balance_sheet(ws_bA, "A")
balance_sheet(ws_bB, "B")

# ---------------------------------------------------------------- Накопитель
title(ws_st, "Накопитель: объёмы, геометрия, отметки уровней",
      "V_полн = (V_полезн + V_осадка + V_недоступн) / (1 − 0,15). Ёмкость принята с вертикальными "
      "(облицованными) стенками; при откосах площадь нужно пересчитать.")
st = [
    ("Полезный регулирующий объём (решение), м³", "=V_use", "#,##0", F_LINK),
    ("Объём осадка, м³", "=V_sl", "#,##0", F_LINK),
    ("Недоступный объём, м³", "=V_dead", "#,##0", F_LINK),
    ("Свободный борт, доля полного", "=k_fb", "0%", F_LINK),
    ("Полный объём, м³", "=(B4+B5+B6)/(1-B7)", "#,##0", F_BOLD),
    ("Объём свободного борта, м³", "=B8*B7", "#,##0", F_NORM),
    ("Проверка: сумма частей − полный объём (должно быть 0)", "=B4+B5+B6+B9-B8", "0.000", F_NORM),
    (None, None, None, None),
    ("Требуемое накопление: максимум по всем сценариям и конфигурациям, м³",
     "=MAX(" + ",".join(R(ws_b, BL[c][k]["vmax"]) for c, ws_b in (("A", ws_bA), ("B", ws_bB))
                        for k in ("peak", "fail", "cur")) + ")", "#,##0", F_NORM),
    ("Запас полезного объёма над требуемым", "=IF(B12=0,0,B4/B12-1)", "0%", F_NORM),
    ("Полезный объём ≥ требуемого", '=IF(B4>=B12,"да","НЕТ")', None, F_BOLD),
    (None, None, None, None),
    ("Глубина полного объёма (решение), м", "=Depth", "0.0", F_LINK),
    ("Глубина ≤ предела 5 м", '=IF(B16<=Depth_lim,"да","НЕТ")', None, F_BOLD),
    ("Площадь зеркала, м²", "=B8/B16", "#,##0", F_NORM),
    ("Сторона квадратной ёмкости, м", "=SQRT(B18)", "0.0", F_NORM),
    ("Занимаемая площадь с полосой обслуживания, м²", "=(B19+2*W_acc)^2", "#,##0", F_NORM),
    ("Площадь ≤ предела 1200 м²", '=IF(B20<=Area_lim,"да","НЕТ")', None, F_BOLD),
    ("Резерв площади, м²", "=Area_lim-B20", "#,##0", F_NORM),
    ("CAPEX накопителя, млн руб.", "=B8*P_stor", "#,##0.00", F_BOLD),
]
for i, (lab, f, fmt, font) in enumerate(st):
    rr = 4 + i
    if lab is None:
        continue
    put(ws_st, f"A{rr}", lab, border=True, align=WRAP)
    put(ws_st, f"B{rr}", f, font, fmt, FILL_SUM if font is F_BOLD else None, border=True, align=CENTER)
defname("V_full", ws_st, "B8")
defname("A_water", ws_st, "B18")

put(ws_st, "A26", "Отметки уровней и алгоритм управления", F_H2)
header(ws_st, 27, ["Отметка", "Объём от дна, м³", "Высота от дна, м", "Действие"])
levels = [
    ("Дно", "=0", "Осадок удаляется при плановой очистке"),
    ("Верх зоны осадка", "=V_sl", "Контроль заиления (зондирование); очистка при превышении"),
    ("«Стоп» = верх недоступного объёма = низ полезного", "=V_sl+V_dead",
     "Все насосы отключаются (защита от сухого хода)"),
    ("«Пуск» базовой ступени", "=V_sl+V_dead+V_on1", "Включается базовое число насосов (A: 1×P2; B: 2×P1)"),
    ("«Пуск 2» повышенной ступени", "=V_sl+V_dead+V_L2",
     "Включается доп. насос, если подача остаётся ≤ предела отстойника (A: нельзя — 2×P2 = 1620 > 1500)"),
    ("«Авария»", "=V_sl+V_dead+k_alarm*V_use", "Сигнал диспетчеру, автоматический ввод резерва, проверка питания"),
    ("Верх полезного объёма", "=V_sl+V_dead+V_use", "Максимальный рабочий уровень; выше — только свободный борт"),
    ("Гребень (верх полного объёма)", "=V_full", "Свободный борт для регулирования не используется"),
]
for i, (lab, f, act) in enumerate(levels):
    rr = 28 + i
    put(ws_st, f"A{rr}", lab, border=True, align=WRAP)
    put(ws_st, f"B{rr}", f, fmt="#,##0", border=True)
    put(ws_st, f"C{rr}", f"=B{rr}/A_water", fmt="0.00", border=True)
    put(ws_st, f"D{rr}", act, F_NOTE, border=True, align=WRAP)
widths(ws_st, {"A": 62, "B": 16, "C": 16, "D": 70})

# ---------------------------------------------------------------- Очистка
title(ws_tr, "Очистка: водный и компонентный баланс по ветвям",
      "Схема: весь откачанный поток → отстойник S; из него 300 м³/ч (когда работает потребитель) → "
      "доочистка (F или M) → потребитель; остаток → условный приёмник. Отделённая взвесь → осадок.")
header(ws_tr, 4, ["Показатель", "Ед.", "Режим 1", "Режим 2", "Режим 3"])
put(ws_tr, "A5", "Описание режима", border=True)
put(ws_tr, "B5", "—", border=True)
for col, txt in zip("CDE", ["Будущий приток, потребитель работает", "Будущий приток, потребитель стоит",
                            "Текущий приток, потребитель работает"]):
    put(ws_tr, f"{col}5", txt, F_BOLD, align=CENTER, border=True)
tr = [
    ("Приток Q", "м³/ч", ["=Q_fut", "=Q_fut", "=Q_cur"], "#,##0"),
    ("Потребитель работает (1 — да, 0 — нет)", "—", [1, 0, 1], "0"),
    ("Часов в году в режиме", "ч/год", ["=H_cons", "=H_yr-H_cons", "=H_cons"], "#,##0"),
    ("Поток на отстойник S", "м³/ч", ["={c}6"] * 3, "#,##0"),
    ("Предел S", "м³/ч", ["=S_lim"] * 3, "#,##0"),
    ("Поток ≤ предела S", "—", ['=IF({c}9<={c}10,"да","НЕТ")'] * 3, None),
    ("Взвесь на входе S", "мг/л", ["=C_raw"] * 3, "#,##0"),
    ("Эффективность S", "доля", ["=S_eff"] * 3, "0%"),
    ("Взвесь после S", "мг/л", ["={c}12*(1-{c}13)"] * 3, "#,##0.0"),
    ("Снято на S (в осадок)", "кг/ч", ["={c}9*({c}12-{c}14)/1000"] * 3, "#,##0.0"),
    ("Поток в ветку повторной воды", "м³/ч", ["=IF({c}7=1,MIN(Q_reuse,{c}9),0)"] * 3, "#,##0"),
    ("Технология доочистки", "—", ["=Reuse_tech"] * 3, None),
    ("Предел доочистки", "м³/ч", ["=INDEX(Tr_lim,MATCH({c}17,Tr_ID,0))"] * 3, "#,##0"),
    ("Поток ≤ предела доочистки", "—", ['=IF({c}16<={c}18,"да","НЕТ")'] * 3, None),
    ("Эффективность доочистки", "доля", ["=INDEX(Tr_eff,MATCH({c}17,Tr_ID,0))"] * 3, "0%"),
    ("Взвесь в повторной воде", "мг/л", ["={c}14*(1-{c}20)"] * 3, "#,##0.0"),
    ("Цель ≤ 30 мг/л выполнена", "—", ['=IF({c}16=0,"не подаётся",IF({c}21<=C_reuse,"да","НЕТ"))'] * 3, None),
    ("Снято на доочистке (в осадок)", "кг/ч", ["={c}16*({c}14-{c}21)/1000"] * 3, "#,##0.0"),
    ("Поток к условному приёмнику", "м³/ч", ["={c}9-{c}16"] * 3, "#,##0"),
    ("Взвесь в потоке к приёмнику", "мг/л", ["={c}14"] * 3, "#,##0.0"),
    ("Контроль ≤ 100 мг/л выполнен", "—", ['=IF({c}25<=C_dis,"да","НЕТ")'] * 3, None),
    ("Водный баланс: приток − повторная − к приёмнику", "м³/ч", ["={c}6-{c}16-{c}24"] * 3, "0.000"),
    ("Взвесь на входе", "кг/ч", ["={c}6*{c}12/1000"] * 3, "#,##0.0"),
    ("Взвесь: в осадке + с повторной водой + к приёмнику", "кг/ч",
     ["={c}15+{c}23+{c}16*{c}21/1000+{c}24*{c}25/1000"] * 3, "#,##0.0"),
    ("Компонентный баланс по взвеси: невязка", "кг/ч", ["={c}28-{c}29"] * 3, "0.000"),
    ("Масса осадка (сухое вещество) всего", "кг/ч", ["={c}15+{c}23"] * 3, "#,##0.0"),
    ("Сульфаты в повторной воде и потоке к приёмнику", "мг/л", ["=SO4_raw"] * 3, "#,##0"),
    ("Жёсткость в повторной воде и потоке к приёмнику", "мг/л", ["=Hard_raw"] * 3, "#,##0"),
    ("pH в повторной воде и потоке к приёмнику", "ед.", ["=pH_raw"] * 3, "0.0"),
]
for i, (lab, unit, fs, fmt) in enumerate(tr):
    rr = 6 + i
    put(ws_tr, f"A{rr}", lab, border=True, align=WRAP)
    put(ws_tr, f"B{rr}", unit, border=True, align=CENTER)
    for col, f in zip("CDE", fs):
        v = f.format(c=col) if isinstance(f, str) else f
        font = F_IN if not isinstance(f, str) else F_NORM
        put(ws_tr, f"{col}{rr}", v, font, fmt, border=True, align=CENTER)
put(ws_tr, "A35",
    "S, F и M не меняют сульфаты, жёсткость и pH. Требования потребителя к ним в постановке не заданы — "
    "пригодность повторной воды по этим показателям нужно подтвердить отдельно (лабораторно на пилоте). "
    "Учебная очистка не является нормативным допуском к сбросу.", F_NOTE, align=WRAP)
ws_tr.merge_cells("A35:E36")
ws_tr.row_dimensions[35].height = 30

put(ws_tr, "A38", "Годовые итоги при будущем притоке (режимы 1 + 2)", F_H2)
ann = [
    ("Объём повторной воды", "м³/год", "=C16*C8+D16*D8", "#,##0"),
    ("Объём к условному приёмнику", "м³/год", "=C24*C8+D24*D8", "#,##0"),
    ("Проверка: сумма − годовой приток (должно быть 0)", "м³/год", "=C39+C40-Q_fut*H_yr", "0.0"),
    ("Осадок (сухое вещество) всего", "т/год", "=(C31*C8+D31*D8)/1000", "#,##0"),
    ("  в т. ч. с отстойника S", "т/год", "=(C15*C8+D15*D8)/1000", "#,##0"),
    ("  в т. ч. с доочистки", "т/год", "=(C23*C8+D23*D8)/1000", "#,##0"),
    ("Объём уплотнённого шлама (допущения на листе «Решения»)", "м³/год", "=C42/(sl_frac*sl_rho/1000)", "#,##0"),
    ("То же в среднем за сутки", "м³/сут", "=C45/365", "#,##0.0"),
]
for i, (lab, unit, f, fmt) in enumerate(ann):
    rr = 39 + i
    put(ws_tr, f"A{rr}", lab, border=True)
    put(ws_tr, f"B{rr}", unit, border=True, align=CENTER)
    put(ws_tr, f"C{rr}", f, F_BOLD, fmt, FILL_SUM, border=True, align=CENTER)
put(ws_tr, "A48", "Шламовое хозяйство (обезвоживание, размещение) в базовую модель не входит — направление развития.",
    F_NOTE)
widths(ws_tr, {"A": 56, "B": 9, "C": 22, "D": 22, "E": 22})

# ---------------------------------------------------------------- Экономика
title(ws_ec, "Экономика: водоотлив A/B, ветка повторной воды, база/проект",
      "Горизонт 5 лет, постоянные цены, без НДС и дисконтирования. Обслуживание — 2 % выбранного CAPEX в год. "
      "Эффект повторной воды считается отдельно от обязательных затрат водоотлива.")
put(ws_ec, "A3", "1. Водоотлив (обязателен в любой схеме — без него добыча невозможна)", F_H2)
header(ws_ec, 4, ["Статья", "Ед.", "Конфигурация A", "Конфигурация B"])
ec1 = [
    ("Состав насосной", "—", ['=A_N&" × "&A_type&" (резерв "&A_res&")"', '=B_N&" × "&B_type&" (резерв "&B_res&")"'],
     None),
    ("CAPEX насосов", "млн руб.", [f"={R(ws_p, PB['A']['capex'])}", f"={R(ws_p, PB['B']['capex'])}"], "#,##0.00"),
    ("CAPEX накопителя", "млн руб.", ["='Накопитель'!$B$23"] * 2, "#,##0.00"),
    ("CAPEX отстойника S (нужен для любого сброса: 500 > 100 мг/л)", "млн руб.", ["=S_capex"] * 2, "#,##0.00"),
    ("CAPEX итого", "млн руб.", ["=SUM({c}6:{c}8)"] * 2, "#,##0.00"),
    ("Годовой объём водоотлива", "м³/год", ["=Q_fut*H_yr"] * 2, "#,##0"),
    ("Энергия насосов (сутки с пиком, среднее 1000 м³/ч, × 365)", "МВт·ч/год",
     [f"={R(ws_bA, BL['A']['peak']['energy'])}*365/1000", f"={R(ws_bB, BL['B']['peak']['energy'])}*365/1000"],
     "#,##0"),
    ("Стоимость энергии насосов", "млн руб./год", ["={c}11*tariff/1000000"] * 2, "#,##0.00"),
    ("Энергия отстойника S", "млн руб./год", ["={c}10*S_kwh*tariff/1000/1000000"] * 2, "#,##0.00"),
    ("Прочие расходы отстойника S", "млн руб./год", ["={c}10*S_oth/1000000"] * 2, "#,##0.00"),
    ("Обслуживание (2 % CAPEX)", "млн руб./год", ["=k_maint*{c}9"] * 2, "#,##0.00"),
    ("OPEX итого", "млн руб./год", ["=SUM({c}12:{c}15)"] * 2, "#,##0.00"),
    ("Затраты за горизонт (CAPEX + 5 × OPEX)", "млн руб.", ["={c}9+Horizon*{c}16"] * 2, "#,##0.00"),
    ("Удельная стоимость водоотлива", "руб./м³", ["={c}17*1000000/({c}10*Horizon)"] * 2, "#,##0.00"),
    ("Проходит все проверки (лист «Проверки»)", "—", ["='Проверки'!$D$27", "='Проверки'!$E$27"], None),
]
for i, (lab, unit, fs, fmt) in enumerate(ec1):
    rr = 5 + i
    put(ws_ec, f"A{rr}", lab, border=True, align=WRAP)
    put(ws_ec, f"B{rr}", unit, border=True, align=CENTER)
    for col, f in zip("CD", fs):
        bold = rr in (9, 16, 17)
        put(ws_ec, f"{col}{rr}", f.format(c=col), F_BOLD if bold else F_NORM, fmt,
            FILL_SUM if bold else None, border=True, align=CENTER)
put(ws_ec, "A20", "Рекомендуемая конфигурация (дешевле за горизонт среди прошедших проверки)", F_BOLD, border=True)
put(ws_ec, "C20", '=IF(AND(C19="да",D19="да"),IF(C17<=D17,"A","B"),IF(C19="да","A",IF(D19="да","B","нет допустимой")))',
    F_BOLD, fill=FILL_SUM, border=True, align=CENTER)
put(ws_ec, "A21", "Экономия рекомендуемой против другой за горизонт, млн руб.", border=True)
put(ws_ec, "C21", "=ABS(C17-D17)", F_BOLD, "#,##0.00", FILL_SUM, border=True, align=CENTER)
put(ws_ec, "A22", "Справочно: энергия насосов при текущем притоке 450 м³/ч, МВт·ч/год", border=True)
put(ws_ec, "C22", f"={R(ws_bA, BL['A']['cur']['energy'])}*365/1000", fmt="#,##0", border=True, align=CENTER)
put(ws_ec, "D22", f"={R(ws_bB, BL['B']['cur']['energy'])}*365/1000", fmt="#,##0", border=True, align=CENTER)
defname("Sel_cfg", ws_ec, "C20")

put(ws_ec, "A24", "2. Ветка повторной воды (считается отдельно, CAPEX водоотлива на неё не относится)", F_H2)
header(ws_ec, 25, ["Статья", "Ед.", "Доочистка F", "Доочистка M"])
ec2 = [
    ("Технология", "—", ['="F"', '="M"'], None),
    ("Пропускная способность", "м³/ч", ["=INDEX(Tr_lim,MATCH({c}26,Tr_ID,0))"] * 2, "#,##0"),
    ("Хватает на спрос 300 м³/ч", "—", ['=IF({c}27>=Q_reuse,"да","НЕТ")'] * 2, None),
    ("Взвесь в повторной воде (после S и доочистки)", "мг/л",
     ["='Очистка'!$C$14*(1-INDEX(Tr_eff,MATCH({c}26,Tr_ID,0)))"] * 2, "#,##0.0"),
    ("Цель ≤ 30 мг/л выполнена", "—", ['=IF({c}29<=C_reuse,"да","НЕТ")'] * 2, None),
    ("CAPEX установки доочистки", "млн руб.", ["=INDEX(Tr_capex,MATCH({c}26,Tr_ID,0))"] * 2, "#,##0.00"),
    ("CAPEX магистрали/подключения", "млн руб.", ["=CAPEX_main"] * 2, "#,##0.00"),
    ("CAPEX ветки итого", "млн руб.", ["={c}31+{c}32"] * 2, "#,##0.00"),
    ("Объём повторной воды (300 м³/ч × 6000 ч)", "м³/год", ["=Q_reuse*H_cons"] * 2, "#,##0"),
    ("Замещённая покупная вода", "млн руб./год", ["={c}34*P_water/1000000"] * 2, "#,##0.00"),
    ("Энергия доочистки", "млн руб./год",
     ["={c}34*INDEX(Tr_kwh,MATCH({c}26,Tr_ID,0))*tariff/1000/1000000"] * 2, "#,##0.00"),
    ("Прочие расходы доочистки", "млн руб./год", ["={c}34*INDEX(Tr_oth,MATCH({c}26,Tr_ID,0))/1000000"] * 2,
     "#,##0.00"),
    ("Обслуживание (2 % CAPEX ветки)", "млн руб./год", ["=k_maint*{c}33"] * 2, "#,##0.00"),
    ("Дополнительные затраты ветки", "млн руб./год", ["=SUM({c}36:{c}38)"] * 2, "#,##0.00"),
    ("Чистый годовой эффект", "млн руб./год", ["={c}35-{c}39"] * 2, "#,##0.00"),
    ("Эффект за горизонт (5 × эффект − CAPEX)", "млн руб.", ["=Horizon*{c}40-{c}33"] * 2, "#,##0.00"),
    ("Простой срок окупаемости", "лет", ['=IF({c}40>0,{c}33/{c}40,"не окупается")'] * 2, "0.0"),
]
for i, (lab, unit, fs, fmt) in enumerate(ec2):
    rr = 26 + i
    put(ws_ec, f"A{rr}", lab, border=True, align=WRAP)
    put(ws_ec, f"B{rr}", unit, border=True, align=CENTER)
    for col, f in zip("CD", fs):
        bold = rr in (33, 40, 41)
        put(ws_ec, f"{col}{rr}", f.format(c=col), F_BOLD if bold else F_NORM, fmt,
            FILL_SUM if bold else None, border=True, align=CENTER)
put(ws_ec, "A43", "Лучшая технология ветки (больший эффект среди выполняющих цель)", F_BOLD, border=True)
put(ws_ec, "C43", '=IF(AND(C30="да",D30="да",C28="да",D28="да"),IF(C41>=D41,"F","M"),'
                  'IF(AND(C30="да",C28="да"),"F",IF(AND(D30="да",D28="да"),"M","нет")))',
    F_BOLD, fill=FILL_SUM, border=True, align=CENTER)
put(ws_ec, "A44", "Принята в проекте (лист «Решения») совпадает с лучшей", border=True)
put(ws_ec, "C44", '=IF(Reuse_tech=C43,"да","НЕТ")', F_BOLD, fill=FILL_SUM, border=True, align=CENTER)
put(ws_ec, "A45", "Допущение: энергия подачи повторной воды потребителю отдельно не задана в постановке и считается "
                  "входящей в магистраль/подключение; при известном напоре добавить по формуле ρgQH/(3600η).",
    F_NOTE, align=WRAP)
ws_ec.merge_cells("A45:D46")

put(ws_ec, "A48", "3. Сравнение «база / проект» (рекомендуемая конфигурация водоотлива)", F_H2)
header(ws_ec, 49, ["Статья", "Ед.", "База: водоотлив + покупная вода",
                   "Проект: водоотлив + повторная вода", "Эффект проекта (база − проект)"])
sel = 'IF(Sel_cfg="A",{a},{b})'
br = 'IF(Reuse_tech="F",{f},{m})'
ec3 = [
    ("CAPEX", "млн руб.", [f"={sel.format(a='C9', b='D9')}",
                           f"={sel.format(a='C9', b='D9')}+{br.format(f='C33', m='D33')}"], "#,##0.00"),
    ("OPEX водоотлива и отстойника", "млн руб./год", [f"={sel.format(a='C16', b='D16')}"] * 2, "#,##0.00"),
    ("Покупка свежей воды для потребителя", "млн руб./год", ["=Q_reuse*H_cons*P_water/1000000", "=0"], "#,##0.00"),
    ("Дополнительные затраты ветки", "млн руб./год", ["=0", f"={br.format(f='C39', m='D39')}"], "#,##0.00"),
    ("OPEX итого", "млн руб./год", ["=SUM({c}51:{c}53)"] * 2, "#,##0.00"),
    ("Затраты за горизонт", "млн руб.", ["={c}50+Horizon*{c}54"] * 2, "#,##0.00"),
]
for i, (lab, unit, fs, fmt) in enumerate(ec3):
    rr = 50 + i
    put(ws_ec, f"A{rr}", lab, border=True)
    put(ws_ec, f"B{rr}", unit, border=True, align=CENTER)
    for col, f in zip("CD", fs):
        put(ws_ec, f"{col}{rr}", f.format(c=col), F_BOLD if rr == 55 else F_NORM, fmt,
            FILL_SUM if rr == 55 else None, border=True, align=CENTER)
    put(ws_ec, f"E{rr}", f"=C{rr}-D{rr}", F_BOLD if rr == 55 else F_NORM, fmt,
        FILL_SUM if rr == 55 else None, border=True, align=CENTER)
widths(ws_ec, {"A": 62, "B": 13, "C": 22, "D": 22, "E": 22})

# ---------------------------------------------------------------- Пилот
title(ws_pl, "План пилота (≤ 100 млн руб.)",
      "Пилот — ветка повторной воды и система контроля; на нём же проверяются расчётные допущения "
      "водоотлива без риска для добычи.")
header(ws_pl, 4, ["№", "Статья", "млн руб.", "Источник"])
pil = [
    ("Установка доочистки (принятая технология)", "=INDEX(Tr_capex,MATCH(Reuse_tech,Tr_ID,0))", "Паспорт кейса"),
    ("Магистраль / подключение потребителя", "=CAPEX_main", "Паспорт кейса"),
    ("КИПиА: уровнемеры, расходомеры, мутномеры, связь", "=KIP", "Собственная оценка (лист «Решения»)"),
    ("Испытания и лабораторный контроль", "=Test", "Собственная оценка (лист «Решения»)"),
]
for i, (lab, f, src) in enumerate(pil):
    rr = 5 + i
    put(ws_pl, f"A{rr}", i + 1, border=True, align=CENTER)
    put(ws_pl, f"B{rr}", lab, border=True)
    put(ws_pl, f"C{rr}", f, F_LINK, "#,##0.00", border=True, align=CENTER)
    put(ws_pl, f"D{rr}", src, F_NOTE, border=True)
put(ws_pl, "B9", "Итого пилот", F_BOLD, border=True)
put(ws_pl, "C9", "=SUM(C5:C8)", F_BOLD, "#,##0.00", FILL_SUM, border=True, align=CENTER)
put(ws_pl, "B10", "Предел пилота", border=True)
put(ws_pl, "C10", "=Pilot_lim", F_LINK, "#,##0.00", border=True, align=CENTER)
put(ws_pl, "B11", "Пилот в пределах лимита", F_BOLD, border=True)
put(ws_pl, "C11", '=IF(C9<=C10,"да","НЕТ")', F_BOLD, fill=FILL_SUM, border=True, align=CENTER)
put(ws_pl, "B12", "Справочно: CAPEX всего проекта (водоотлив + ветка), млн руб.", border=True)
put(ws_pl, "C12", "='Экономика'!$D$50", F_LINK, "#,##0.00", border=True, align=CENTER)

put(ws_pl, "A14", "Программа пилота", F_H2)
header(ws_pl, 15, ["Этап", "Что делаем", "Срок", "Критерий успеха / что уточняем"])
prog = [
    ("1", "Обследование: замер фактического притока, снятие характеристики сети (H_st, K), пробы воды",
     "1–2 мес.", "Расхождение с учебными H_st и K ≤ 15 %; иначе — пересчёт рабочих точек"),
    ("2", "Монтаж КИПиА на действующем водоотливе, отладка алгоритма уровней",
     "1 мес.", "Уровни «стоп/пуск/авария» срабатывают; журнал часового баланса сходится с расчётом"),
    ("3", "Пуск ветки доочистки 300 м³/ч после отстойника",
     "3 мес.", "Взвесь ≤ 30 мг/л в 95 % проб; фактические кВт·ч/м³ и руб./м³ против паспорта"),
    ("4", "Лабораторная проверка сульфатов, жёсткости, pH у потребителя",
     "параллельно", "Согласованные с потребителем требования; при превышении — решение по доп. обработке"),
    ("5", "Учения по отказу: плановая остановка откачки на текущем притоке (450 м³/ч)",
     "1 день", "Рост уровня и время восстановления совпадают с расчётом ±20 %"),
    ("6", "Решение о масштабировании до 1000 м³/ч и выборе насосной конфигурации",
     "по итогам", "Обновлённая расчётная книга с фактическими данными"),
]
for i, row in enumerate(prog):
    rr = 16 + i
    for j, v in enumerate(row):
        put(ws_pl, f"{'ABCD'[j]}{rr}", v, border=True, align=WRAP)
    ws_pl.row_dimensions[rr].height = 30
widths(ws_pl, {"A": 6, "B": 62, "C": 14, "D": 64})

# ---------------------------------------------------------------- Проверки
title(ws_chk, "Проверки по разделу 9 постановки",
      "«да» — требование выполнено. Любое «НЕТ» означает, что решение надо менять (лист «Решения»).")
header(ws_chk, 4, ["№", "Требование", "Где считается", "Конфигурация A", "Конфигурация B"])


def both(fn):
    return [fn("A", ws_bA), fn("B", ws_bB)]


def yes(cond):
    return f'=IF({cond},"да","НЕТ")'


def bl(cfg, ws_b, key, field):
    return R(ws_b, BL[cfg][key][field])


checks = [
    ("Рабочие точки используемых насосов в диапазонах q_min…q_max", "Насосы",
     [f"={R(ws_p, PB[c]['rng_ok'])}" for c in "AB"]),
    ("Базовая ступень работает без задействования резерва", "Насосы",
     [f"={R(ws_p, PB[c]['base_ok'])}" for c in "AB"]),
    ("Резерв: при отказе одного агрегата подача ≥ 1000 м³/ч", "Насосы",
     [f"={R(ws_p, PB[c]['res_ok'])}" for c in "AB"]),
    ("Суточный пик: нет перелива", "Баланс",
     both(lambda c, w: yes(f'{bl(c, w, "peak", "overflow")}="нет"'))),
    ("Суточный пик: возврат к исходному уровню до конца суток", "Баланс",
     both(lambda c, w: f"={bl(c, w, 'peak', 'ok')}")),
    ("Отказ 2 ч: нет перелива", "Баланс",
     both(lambda c, w: yes(f'{bl(c, w, "fail", "overflow")}="нет"'))),
    ("Отказ 2 ч: уровень восстановлен ≤ 24 ч", "Баланс",
     both(lambda c, w: f"={bl(c, w, 'fail', 'ok')}")),
    ("Текущий приток 450 м³/ч: без перелива, уровень возвращается", "Баланс",
     both(lambda c, w: f"={bl(c, w, 'cur', 'ok')}")),
    ("Полный приток сохранён: невязка баланса = 0 во всех сценариях", "Баланс",
     both(lambda c, w: yes("AND(" + ",".join(f"ABS({bl(c, w, k, 'resid')})<0.01"
                                             for k in ("peak", "fail", "cur")) + ")"))),
    ("Подача насосов на отстойник S ≤ 1500 м³/ч во все часы", "Баланс",
     both(lambda c, w: yes("AND(" + ",".join(f'{bl(c, w, k, "s_ok")}="да"'
                                             for k in ("peak", "fail", "cur")) + ")"))),
    ("Каждый сценарий начинается с пустого полезного объёма", "Баланс",
     both(lambda c, w: yes("AND(" + ",".join(f"{bl(c, w, k, 'start')}=0"
                                             for k in ("peak", "fail", "cur")) + ")"))),
    ("Полезный объём ≥ максимального накопления этой конфигурации", "Баланс, Накопитель",
     both(lambda c, w: yes("V_use>=MAX(" + ",".join(bl(c, w, k, "vmax") for k in ("peak", "fail", "cur")) + ")"))),
    ("Полезный и полный объёмы различаются; осадок, недоступный объём и борт учтены по одному разу", "Накопитель",
     [yes("AND(V_full>V_use,ABS('Накопитель'!$B$10)<0.01)")] * 2),
    ("Глубина ≤ 5 м, площадь с обслуживанием ≤ 1200 м²", "Накопитель",
     [yes("AND('Накопитель'!$B$17=\"да\",'Накопитель'!$B$21=\"да\")")] * 2),
    ("Повторная вода: взвесь ≤ 30 мг/л", "Очистка",
     [yes("AND('Очистка'!$C$22=\"да\",'Очистка'!$E$22=\"да\")")] * 2),
    ("Оставшийся поток: взвесь ≤ 100 мг/л во всех режимах", "Очистка",
     [yes("COUNTIF('Очистка'!$C$26:$E$26,\"да\")=3")] * 2),
    ("Потоки в пределах S и доочистки", "Очистка",
     [yes("AND(COUNTIF('Очистка'!$C$11:$E$11,\"да\")=3,COUNTIF('Очистка'!$C$19:$E$19,\"да\")=3)")] * 2),
    ("Водный и компонентный балансы очистки сходятся", "Очистка",
     [yes("AND(SUMPRODUCT(ABS('Очистка'!$C$27:$E$27))<0.01,SUMPRODUCT(ABS('Очистка'!$C$30:$E$30))<0.01)")] * 2),
    ("Пилот в пределах 100 млн руб.", "Пилот", ["='Пилот'!$C$11"] * 2),
]
for i, (lab, where, fs) in enumerate(checks):
    rr = 5 + i
    put(ws_chk, f"A{rr}", i + 1, border=True, align=CENTER)
    put(ws_chk, f"B{rr}", lab, border=True, align=WRAP)
    put(ws_chk, f"C{rr}", where, F_NOTE, border=True)
    for col, f in zip("DE", fs):
        put(ws_chk, f"{col}{rr}", f, F_BOLD, border=True, align=CENTER)
last = 5 + len(checks) - 1
put(ws_chk, "B26", "Число невыполненных требований", F_BOLD, border=True)
put(ws_chk, "D26", f'=COUNTIF(D5:D{last},"НЕТ")', F_BOLD, "0", FILL_SUM, border=True, align=CENTER)
put(ws_chk, "E26", f'=COUNTIF(E5:E{last},"НЕТ")', F_BOLD, "0", FILL_SUM, border=True, align=CENTER)
put(ws_chk, "B27", "Конфигурация проходит все проверки", F_BOLD, border=True)
put(ws_chk, "D27", '=IF(D26=0,"да","НЕТ")', F_BOLD, fill=FILL_SUM, border=True, align=CENTER)
put(ws_chk, "E27", '=IF(E26=0,"да","НЕТ")', F_BOLD, fill=FILL_SUM, border=True, align=CENTER)
assert last <= 24, last
put(ws_chk, "B29", "Учебная очистка проверяет только взвесь по сценарию кейса и не является нормативным "
                   "допуском сброса; сульфаты, жёсткость и pH требуют отдельной проверки у потребителя.",
    F_NOTE, align=WRAP)
ws_chk.merge_cells("B29:E30")
widths(ws_chk, {"A": 5, "B": 70, "C": 20, "D": 16, "E": 16})

# ---------------------------------------------------------------- Схема
title(ws_sch, "Гидравлическая схема и алгоритм управления")
scheme = [
    "ПОТОКИ (рекомендуемая конфигурация A)",
    "",
    "Карьерный приток 450 → 1000 м³/ч (пик 1200)",
    "   │",
    "   ▼",
    "Накопитель-водосборник: осадок 400 м³ | недоступный 200 м³ | полезный 2500 м³ | свободный борт 15 %",
    "   │  (уровнемеры; отметки «стоп / пуск / пуск 2 / авария»)",
    "   ▼",
    "Насосная: P2 × 2 (1 рабочий 1111 м³/ч при 90,9 м + 1 резерв с АВР); напорный став H = 60 + 0,000025·Q²",
    "   │",
    "   ▼",
    "Отстойник S (весь поток, ≤ 1500 м³/ч): взвесь 500 → 50 мг/л ──► осадок (≈ 450 кг/ч сух.)",
    "   ├──► 300 м³/ч (6000 ч/год) ► доочистка F: 50 → 10 мг/л ► потребитель (замещает покупную воду)",
    "   │                                              └──► осадок (≈ 12 кг/ч сух.)",
    "   └──► остаток 700…1000 м³/ч, 50 мг/л ► условный приёмник (контроль ≤ 100 мг/л)",
    "",
    "АЛГОРИТМ",
    "1. Уровень ниже «стоп» — все насосы отключены (защита от сухого хода).",
    "2. Уровень выше «пуск» — работает базовая ступень (A: 1×P2; B: 2×P1) до уровня «стоп» (циклический режим).",
    "3. Уровень выше «пуск 2» — повышенная ступень, только если суммарная подача ≤ 1500 м³/ч (предел S): "
    "B — 3×P1 = 1437 м³/ч; A — второй P2 не включается (1620 > 1500), остаётся резервом.",
    "4. Отказ рабочего агрегата — автоматический ввод резерва (АВР); A: резервный P2 даёт те же 1111 м³/ч.",
    "5. Полная остановка (питание) — после восстановления каскадный пуск с задержкой, уровень снижается до «стоп».",
    "6. Уровень «авария» — сигнал диспетчеру, проверка резерва и питания; свободный борт для регулирования не используется.",
    "7. Ротация рабочих/резервных агрегатов по наработке (раз в смену/неделю).",
    "",
    "Числа на этой схеме — иллюстрация для рекомендуемых решений; расчётные значения см. на листах.",
]
for i, line in enumerate(scheme):
    put(ws_sch, f"A{3 + i}", line, F_BOLD if line in ("ПОТОКИ (рекомендуемая конфигурация A)", "АЛГОРИТМ")
        else Font(name="Consolas", size=10) if i < 16 else F_NORM)
ws_sch.column_dimensions["A"].width = 130

# ---------------------------------------------------------------- Описание
title(ws_info, "Карьерная вода без остановки — расчётная записка",
      "Инженерный кубок 2026–2027, кейс 16 (СУЭК). Учебная концепция водоотлива, накопления, очистки и "
      "повторного использования воды. Не проект строительства и не разрешение сброса.")
put(ws_info, "A4", "Ключевые результаты", F_H2)
header(ws_info, 5, ["Показатель", "Значение / конфигурация A", "Конфигурация B"])
summary = [
    ("Рекомендуемая конфигурация водоотлива", "='Экономика'!$C$20", None, None),
    ("Затраты на водоотлив за 5 лет, млн руб.", "='Экономика'!$C$17", "='Экономика'!$D$17", "#,##0.0"),
    ("Полезный объём накопителя, м³", "=V_use", None, "#,##0"),
    ("Полный объём накопителя, м³", "=V_full", None, "#,##0"),
    ("Максимальное накопление (отказ 2 ч), м³", "='Накопитель'!$B$12", None, "#,##0"),
    ("Восстановление после отказа, ч",
     f"={R(ws_bA, BL['A']['fail']['rec'])}", f"={R(ws_bB, BL['B']['fail']['rec'])}", "0"),
    ("Технология ветки повторной воды", "=Reuse_tech", None, None),
    ("Эффект повторной воды за 5 лет, млн руб.", "=IF(Reuse_tech=\"F\",'Экономика'!$C$41,'Экономика'!$D$41)",
     None, "#,##0.0"),
    ("Срок окупаемости ветки, лет", "=IF(Reuse_tech=\"F\",'Экономика'!$C$42,'Экономика'!$D$42)", None, "0.0"),
    ("Эффект проекта против базы за 5 лет, млн руб.", "='Экономика'!$E$55", None, "#,##0.0"),
    ("Бюджет пилота, млн руб.", "='Пилот'!$C$9", None, "#,##0.0"),
    ("Невыполненных требований", "='Проверки'!$D$26", "='Проверки'!$E$26", "0"),
]
for i, (lab, fa, fb, fmt) in enumerate(summary):
    rr = 6 + i
    put(ws_info, f"A{rr}", lab, border=True)
    put(ws_info, f"B{rr}", fa, F_BOLD, fmt, FILL_SUM, border=True, align=CENTER)
    put(ws_info, f"C{rr}", fb, F_BOLD, fmt, FILL_SUM if fb else None, border=True, align=CENTER)

put(ws_info, "A18", "Листы", F_H2)
sheets = [
    ("Проверки", "Сводка требований раздела 9 постановки: «да»/«НЕТ» по каждой конфигурации"),
    ("Схема", "Гидравлическая схема потоков и алгоритм уровней/переключений"),
    ("Исходные", "Паспортные данные кейса с именами для формул (синий шрифт)"),
    ("Решения", "Выбор команды (жёлтые ячейки): насосы, резерв, объём, уровни, технология доочистки"),
    ("Насосы", "Рабочие точки, мощность, удельная энергия, резерв; график характеристик"),
    ("Накопитель", "Полезный/полный объём, геометрия, площадь, отметки уровней"),
    ("Баланс_A, Баланс_B", "Почасовой баланс: суточный пик, отказ 2 ч + восстановление, текущий приток; графики"),
    ("Очистка", "Водный и компонентный (взвесь) баланс по ветвям, масса осадка"),
    ("Экономика", "CAPEX/OPEX A и B; ветка повторной воды F и M; база против проекта"),
    ("Пилот", "Бюджет пилота и программа испытаний"),
    ("Объект", "Открытые данные о Восточно-Бейском разрезе СУЭК и список источников"),
]
for i, (nm, txt) in enumerate(sheets):
    put(ws_info, f"A{19 + i}", nm, F_BOLD)
    put(ws_info, f"B{19 + i}", txt)

put(ws_info, "A30", "Условные обозначения", F_H2)
put(ws_info, "A31", "Синий шрифт — исходные данные кейса", F_IN)
put(ws_info, "A32", "Синий шрифт на жёлтом — решения команды (можно менять)", F_IN, fill=FILL_DEC)
put(ws_info, "A33", "Зелёный шрифт — ссылка на другой лист", F_LINK)
put(ws_info, "A34", "Чёрный шрифт — формула; серая заливка — итог", F_NORM, fill=FILL_SUM)

put(ws_info, "A36", "Основные допущения", F_H2)
assum = [
    "Все числа кейса — единые учебные условия, не измерения разреза. Источник: постановка (PDF) и CSV-приложения "
    "(Гидравлическая_база, Отдельные_насосы, Операции_очистки, Цены) — значения совпадают.",
    "Часовой шаг баланса: внутри часа насосы работают циклически, энергия считается по откачанному объёму "
    "и удельной энергии рабочей точки (кВт·ч/м³).",
    "Годовая энергия водоотлива = энергия суток с пиком (среднее 1000 м³/ч) × 365.",
    "Отказ — полная остановка откачки на 2 ч; после него доступны все исправные агрегаты, но подача на S ≤ 1500 м³/ч.",
    "Накопитель — ёмкость с вертикальными стенками; при откосах площадь растёт и её нужно пересчитать.",
    "Отстойник S обязателен для любого сброса (500 > 100 мг/л), поэтому его затраты — часть водоотлива, "
    "а не ветки повторной воды.",
    "Энергия подачи повторной воды потребителю не задана и считается входящей в магистраль/подключение.",
    "Доля твёрдого и плотность шлама, а также КИПиА и испытания пилота — собственные оценки (лист «Решения»).",
]
for i, t in enumerate(assum):
    put(ws_info, f"A{37 + i}", f"• {t}", align=WRAP)
    ws_info.merge_cells(f"A{37 + i}:B{37 + i}")
    ws_info.row_dimensions[37 + i].height = 28
widths(ws_info, {"A": 60, "B": 70, "C": 20})

# ---------------------------------------------------------------- Объект (открытые данные)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from object_facts import ACCESS, FACTS, SOURCES, refs  # noqa: E402

ws_obj = sheet("Объект")
title(ws_obj, "Объект: Восточно-Бейский разрез (СУЭК) — открытые данные",
      f"Факты из открытых источников (обращение {ACCESS}). Расчётный базис кейса не меняется — данные служат "
      "проверкой правдоподобия и выбором потребителя повторной воды. Номера в скобках — источники ниже.")
header(ws_obj, 4, ["Факт", "Значение", "Что это значит для проекта", "Источники"])
for i, (fct, val, mean, nums) in enumerate(FACTS):
    rr = 5 + i
    put(ws_obj, f"A{rr}", fct, F_BOLD, border=True, align=WRAP)
    put(ws_obj, f"B{rr}", val, border=True, align=WRAP)
    put(ws_obj, f"C{rr}", mean, border=True, align=WRAP)
    put(ws_obj, f"D{rr}", refs(nums) or "—", F_NOTE, border=True, align=CENTER)
r0 = 5 + len(FACTS) + 1
put(ws_obj, f"A{r0}", "Проверки правдоподобия", F_H2)
chk = [("Часы работы фабрики в год: 4,05 млн т / 750 т/ч, ч", "=4050000/750", "#,##0", "в кейсе — 6000 ч"),
       ("Удельный спрос на повторную воду: 300 м³/ч × 6000 ч / 4,05 млн т, м³/т", "=Q_reuse*H_cons/4050000",
        "0.00", "типично до 2 м³/т"),
       ("Будущий приток / перспективная мощность действующих очистных (670 м³/ч)", "=Q_fut/670", "0.0",
        "нужно расширение очистки")]
for i, (lab, f, fmt, note) in enumerate(chk):
    rr = r0 + 1 + i
    put(ws_obj, f"A{rr}", lab, border=True, align=WRAP)
    ws_obj.merge_cells(f"A{rr}:B{rr}")
    put(ws_obj, f"C{rr}", f, F_BOLD, fmt, FILL_SUM, border=True, align=CENTER)
    put(ws_obj, f"D{rr}", note, F_NOTE, border=True)
r1 = r0 + 6
put(ws_obj, f"A{r1}", "Список источников", F_H2)
for i, (txt, url) in enumerate(SOURCES, 1):
    rr = r1 + i
    put(ws_obj, f"A{rr}", f"[{i}]", F_NOTE)
    put(ws_obj, f"B{rr}", txt + (f" URL: {url} (дата обращения: {ACCESS})" if url else ""), F_NOTE, align=WRAP)
    ws_obj.merge_cells(f"B{rr}:D{rr}")
    ws_obj.row_dimensions[rr].height = 26
widths(ws_obj, {"A": 34, "B": 52, "C": 52, "D": 16})

# ---------------------------------------------------------------- оформление статусов
RED = PatternFill("solid", fgColor="F8CBAD")
GREEN = PatternFill("solid", fgColor="C6EFCE")
for ws in wb.worksheets:
    rng = f"A1:Z{max(ws.max_row, 10)}"
    for bad in ('"НЕТ"', '"ЕСТЬ"', '"ПЕРЕЛИВ"', '"ПРЕВЫШЕНИЕ"', '"не вернулся"'):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[bad], fill=RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"да"'], fill=GREEN))
    ws.freeze_panes = None
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and (c.font is None or c.font.name not in (FN, "Consolas")):
                c.font = Font(name=FN, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic,
                              color=c.font.color)

wb.calculation.fullCalcOnLoad = True   # Excel пересчитает все формулы при открытии
wb.save(OUT)
print("saved", OUT)
