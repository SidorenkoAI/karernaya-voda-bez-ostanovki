# -*- coding: utf-8 -*-
"""Презентация по кейсу «Карьерная вода без остановки».

Цифры берутся из пересчитанной книги Расчёт_водоотлив.xlsx (значения формул).
Запуск:  python build_presentation.py
Результат: <КОМАНДА>_Карьерная_вода_без_остановки.pptx (PDF делается LibreOffice).
"""
import math
import re

from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ------------------------------------------------------------------ данные титула (заполнить!)
TEAM = "КОМАНДА"                      # фамилия участника или название команды — идёт в имя файла
AUTHORS = "[ФИО участников]"
ORG = "[Образовательная организация / место работы]"
CITY = "[Город]"
YEAR = "2027"
PROJECT = "Карьерная вода без остановки"
OUT = f"{TEAM}_Карьерная_вода_без_остановки.pptx"

# ------------------------------------------------------------------ данные из расчёта
wb = load_workbook("Расчёт_водоотлив.xlsx", data_only=True)


def v(sheet, cell):
    val = wb[sheet][cell].value
    if val is None:
        raise ValueError(f"Нет значения {sheet}!{cell} — пересчитайте книгу")
    return val


def col(sheet, c, r0, r1):
    return [wb[sheet][f"{c}{r}"].value for r in range(r0, r1 + 1)]


def num(x, d=0):
    s = f"{x:,.{d}f}".replace(",", " ").replace(".", ",")
    return s


D = dict(
    qA=v("Насосы", "B16"), hA=v("Насосы", "D16"), eA=v("Насосы", "G16"), kwA=v("Насосы", "F16"),
    qA2=v("Насосы", "B17"),
    qB2=v("Насосы", "B43"), hB2=v("Насосы", "D43"), eB2=v("Насосы", "G43"), qB1=v("Насосы", "B42"),
    qB3=v("Насосы", "B44"),
    V_use=v("Накопитель", "B4"), V_full=v("Накопитель", "B8"), V_fb=v("Накопитель", "B9"),
    V_req=v("Накопитель", "B12"), margin=v("Накопитель", "B13"), depth=v("Накопитель", "B16"),
    foot=v("Накопитель", "B20"), st_capex=v("Накопитель", "B23"),
    lv=[(wb["Накопитель"][f"A{r}"].value, v("Накопитель", f"C{r}")) for r in range(28, 36)],
    pkA=col("Баланс_A", "H", 18, 41), pkB=col("Баланс_B", "H", 18, 41),
    pkA_max=v("Баланс_A", "E49"), pkA_ret=v("Баланс_A", "E53"),
    flA=col("Баланс_A", "H", 63, 88), flB=col("Баланс_B", "H", 63, 88),
    flA_max=v("Баланс_A", "E94"), recA=v("Баланс_A", "E102"), recB=v("Баланс_B", "E102"),
    cS=v("Очистка", "C14"), cF=v("Очистка", "C21"), sl_kgh=v("Очистка", "C31"), sl_t=v("Очистка", "C42"),
    reuse_m3=v("Очистка", "C39"),
    capA=v("Экономика", "C9"), capB=v("Экономика", "D9"),
    enA=v("Экономика", "C11"), enB=v("Экономика", "D11"),
    eCostA=v("Экономика", "C12"), eCostB=v("Экономика", "D12"),
    sCost=v("Экономика", "C13") + v("Экономика", "C14"),
    mA=v("Экономика", "C15"), mB=v("Экономика", "D15"),
    totA=v("Экономика", "C17"), totB=v("Экономика", "D17"), specA=v("Экономика", "C18"),
    econ=v("Экономика", "C21"), sel=v("Экономика", "C20"),
    capF=v("Экономика", "C33"), netF=v("Экономика", "C40"), effF=v("Экономика", "C41"),
    payF=v("Экономика", "C42"), effM=v("Экономика", "D41"), cM=v("Экономика", "D29"),
    pil=[(v("Пилот", f"B{r}"), v("Пилот", f"C{r}")) for r in range(5, 9)], pil_tot=v("Пилот", "C9"),
    failsA=v("Проверки", "D26"), failsB=v("Проверки", "E26"),
)
N_CHECKS = sum(1 for r in range(5, 24) if wb["Проверки"][f"D{r}"].value in ("да", "НЕТ"))
assert D["sel"] == "A"

# ------------------------------------------------------------------ дополнение и открытые данные о разрезе
import os  # noqa: E402
import sys  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
wa = load_workbook(os.path.join("Дополнение", "Дополнение_расчёты.xlsx"), data_only=True)


def va(sh, c):
    val = wa[sh][c].value
    if val is None:
        raise ValueError(f"Нет значения {sh}!{c} в книге дополнения — пересчитайте её")
    return val


AD = {k: dict(tot=va("Экономика_ABCD", f"{c}14"), diff=va("Экономика_ABCD", f"{c}15"),
              res=va("Экономика_ABCD", f"{c}16"), op=va("Экономика_ABCD", f"{c}17"),
              qmax=va("Стресс_тест", f"{c}7"), marg=va("Стресс_тест", f"{c}8"))
      for k, c in zip("ABCD", "BCDE")}
TREP = va("Стресс_тест", "B19")
VFD450, VFD1000, TAR_A = va("ЧРП", "B17"), va("ЧРП", "B30"), va("Тариф", "B15")
SRC = {4: "r-19.ru, 29.08.2018", 5: "РИА Новости, 29.08.2018", 6: "Neftegaz.RU, 04.08.2023",
       7: "Коммерсантъ, 26.05.2016", 8: "energybase.ru", 9: "suek.ru", 10: "СУЭК в Хакасии, итоги 2021",
       11: "Пульс Хакасии, 2022", 12: "«Дело. Промышленность»", 13: "list-org.com",
       14: "19rusinfo.ru, 05.02.2026 (приказ Госкомтарифа РХ № 7-э)", 15: "СанПиН 1.2.3685-21",
       16: "Горная энциклопедия"}


def src(slide, nums, extra=""):
    t = "Источники: " + "; ".join(f"[{n}] {SRC[n]}" for n in nums) + extra
    text(slide, 0.6, 6.9, 11.3, 0.4, t, size=12, color=GREY)

# ------------------------------------------------------------------ оформление
DARK = RGBColor(0x1F, 0x2A, 0x33)     # уголь
TEAL = RGBColor(0x0E, 0x7C, 0x86)     # вода
AMBER = RGBColor(0xE8, 0x9F, 0x00)    # акцент (горная техника)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TINT = RGBColor(0xE3, 0xF1, 0xF2)     # светлый оттенок бирюзового
GREY = RGBColor(0x5F, 0x6B, 0x73)     # вторичный текст
LGREY = RGBColor(0xA9, 0xB3, 0xBA)
FONT = "Arial"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
N = [0]


def text(slide, x, y, w, h, paras, size=18, color=DARK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, space=6, margin=0.05):
    """paras: строка, или список абзацев; абзац — строка или список (текст, {опции})."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, Inches(margin))
    if isinstance(paras, str):
        paras = [paras]
    for i, p in enumerate(paras):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.alignment = align
        par.space_after = Pt(space)
        runs = [(p, {})] if isinstance(p, str) else p
        for t, o in runs:
            r = par.add_run()
            r.text = re.sub(r"(\d) (?=[^\d\s—–=≤≥<>+×/])", "\\1\u00a0", t)
            f = r.font
            f.name = FONT
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.color.rgb = o.get("color", color)
    return tb


def box(slide, x, y, w, h, fill=TINT, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08, lw=1.5):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    s.shadow.inherit = False
    return s


def boxtext(slide, x, y, w, h, paras, fill=TINT, color=DARK, size=18, align=PP_ALIGN.CENTER, line=None,
            anchor=MSO_ANCHOR.MIDDLE):
    b = box(slide, x, y, w, h, fill, line)
    text(slide, x + 0.08, y + 0.05, w - 0.16, h - 0.1, paras, size=size, color=color, align=align,
         anchor=anchor, space=2)
    return b


def arrow(slide, x1, y1, x2, y2, color=DARK, w=2.25):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(w)
    ln = c.line._get_or_add_ln()
    tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
    ln.append(tail)
    return c


def new_slide(title, notes, dark=False):
    s = prs.slides.add_slide(BLANK)
    N[0] += 1
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = DARK if dark else WHITE
    if title:
        text(s, 0.6, 0.35, 12.1, 0.95, title, size=30, bold=True, color=WHITE if dark else DARK,
             anchor=MSO_ANCHOR.MIDDLE)
    text(s, 12.13, 6.95, 0.8, 0.4, str(N[0]), size=14, color=LGREY if dark else GREY, align=PP_ALIGN.RIGHT)
    s.notes_slide.notes_text_frame.text = notes
    return s


def style_chart(ch, legend=True):
    ch.font.name = FONT
    ch.font.size = Pt(14)
    ch.font.color.rgb = DARK
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(14)


def axis(ax, title=None, mn=None, mx=None, fmt=None, grid=True, major=None):
    ax.has_major_gridlines = grid
    if grid:
        ax.major_gridlines.format.line.color.rgb = RGBColor(0xDD, 0xE2, 0xE5)
        ax.major_gridlines.format.line.width = Pt(0.75)
    ax.format.line.color.rgb = LGREY
    ax.tick_labels.font.size = Pt(14)
    ax.tick_labels.font.color.rgb = GREY
    if mn is not None:
        ax.minimum_scale = mn
    if mx is not None:
        ax.maximum_scale = mx
    if major is not None:
        ax.major_unit = major
    if fmt:
        ax.tick_labels.number_format = fmt
        ax.tick_labels.number_format_is_linked = False
    if title:
        ax.has_title = True
        tf = ax.axis_title.text_frame
        tf.text = title
        r = tf.paragraphs[0].runs[0]
        r.font.size = Pt(14)
        r.font.bold = False
        r.font.color.rgb = GREY
        r.font.name = FONT


def stat(slide, x, y, w, big, small, color=TEAL, big_size=40, h=1.6, fill=TINT):
    box(slide, x, y, w, h, fill)
    text(slide, x + 0.2, y + 0.1, w - 0.4, 0.75, big, size=big_size, bold=True, color=color,
         anchor=MSO_ANCHOR.MIDDLE)
    text(slide, x + 0.2, y + 0.82, w - 0.4, h - 0.85, small, size=18, color=DARK)


def dot(slide, x, y, d=0.28, color=TEAL):
    s = box(slide, x, y, d, d, color, shape=MSO_SHAPE.OVAL)
    return s


# ================================================================== 1. Титул
s = new_slide(None, "Здравствуйте! Мы представляем проект «Карьерная вода без остановки» — учебную концепцию "
                    "водоотлива, накопления, очистки и повторного использования воды Восточно-Бейского разреза "
                    "СУЭК в Бейском районе Хакасии (кейс СУЭК, Инженерный кубок).", dark=True)
text(s, 0.9, 1.1, 11.5, 0.5, "Инженерный кубок 2026–2027 · кейс СУЭК", size=20, color=AMBER, bold=True)
text(s, 0.9, 1.7, 11.5, 1.3, PROJECT, size=48, bold=True, color=WHITE)
text(s, 0.9, 2.95, 11.2, 1.0, "Водоотлив, накопление, очистка и повторное использование воды "
                              "при росте притока до 1000 м³/ч", size=22, color=TINT)
text(s, 0.9, 3.9, 11.2, 0.6, "ООО «Восточно-Бейский разрез» (СУЭК) · Бейский район, Республика Хакасия",
     size=20, bold=True, color=WHITE)
text(s, 0.9, 4.85, 11.0, 1.5, [AUTHORS, ORG, CITY], size=20, color=WHITE, space=4)
text(s, 0.9, 6.45, 4, 0.5, YEAR, size=20, bold=True, color=AMBER)

# ================================================================== 2. Проблема
s = new_slide("Восточно-Бейский разрез: приток растёт до 1000 м³/ч, а очистные рассчитаны на 670",
              "Восточно-Бейский разрез СУЭК — Бейское месторождение в Хакасии, добыча 3,5–3,7 млн т в год с планом "
              "роста до 5 млн т. С 2018 года работают очистные карьерных вод на 350 м³/ч с перспективой до 670. "
              "По условиям кейса приток вырастет с 450 до 1000 м³/ч, есть пик 1200 м³/ч и возможен отказ на 2 часа. "
              "На разрезе строится обогатительная фабрика — потребитель воды. Задача: водоотлив и очистка на весь "
              "будущий приток.")
box(s, 0.6, 1.55, 5.75, 4.2, DARK)
text(s, 0.85, 1.65, 5.3, 4.0,
     [[("Объект", {"bold": True, "color": AMBER, "size": 20})],
      [("Где: ", {"bold": True, "color": AMBER}), ("Бейский район Хакасии, Бейское месторождение "
                                                   "Минусинского бассейна", {})],
      [("Добыча: ", {"bold": True, "color": AMBER}), ("3,5–3,7 млн т/год, план — до 5 млн т", {})],
      [("Очистные с 2018 г.: ", {"bold": True, "color": AMBER}), ("350 м³/ч, перспектива — 670 м³/ч", {})],
      [("Строится: ", {"bold": True, "color": AMBER}), ("обогатительная фабрика 750 т/ч", {})]],
     size=18, color=WHITE, space=8)
cards = [("450 → 1000", "м³/ч — средний приток по кейсу", TEAL),
         ("1200", "м³/ч — суточный пик, 6 часов", TEAL),
         ("2000 м³", "накопится за 2 ч отказа откачки", AMBER),
         ("300", "м³/ч — спрос фабрики на воду", TEAL)]
for i, (b, sm, c) in enumerate(cards):
    stat(s, 6.6 + (i % 2) * 3.15, 1.55 + (i // 2) * 2.15, 2.95, b, sm, color=c, big_size=30, h=2.0)
box(s, 0.6, 5.95, 12.1, 0.85, fill=None, line=TEAL, lw=1.5)
text(s, 0.85, 5.98, 11.6, 0.8,
     [[("Проблема: ", {"bold": True, "color": TEAL}),
       ("нужен водоотлив, который выдержит пик и отказ, и очистка на весь будущий приток — "
        "насосы, накопитель и очистка работают как один объект.", {})]],
     size=18, anchor=MSO_ANCHOR.MIDDLE)
src(s, [4, 6, 8, 10], "; [1] кейс")

# ================================================================== 3. Цель и задачи
s = new_slide("Цель и задачи проекта",
              "Цель — спроектировать систему на будущий приток с обязательной проверкой пика и отказа. "
              "Задачи идут в порядке расчёта: насосы, баланс, накопитель и очистка, экономика и пилот.")
box(s, 0.6, 1.6, 4.7, 4.9, TEAL)
text(s, 0.95, 1.85, 4.0, 0.6, "Цель", size=26, bold=True, color=WHITE)
text(s, 0.95, 2.55, 4.05, 3.8,
     "Спроектировать водоотлив, накопление, очистку и повторное использование карьерной воды "
     "Восточно-Бейского разреза на приток 1000 м³/ч с проверкой пика, отказа и экономики",
     size=22, color=WHITE)
tasks = ["Подобрать насосы и резерв: рабочие точки в допустимых диапазонах",
         "Рассчитать почасовой баланс накопителя: пик, отказ, восстановление",
         "Выбрать объём, уровни управления и схему очистки с балансом взвеси",
         "Сравнить экономику двух конфигураций, оценить повторную воду и пилот"]
for i, t in enumerate(tasks):
    y = 1.6 + i * 1.25
    box(s, 5.7, y, 0.8, 0.8, AMBER, shape=MSO_SHAPE.OVAL)
    text(s, 5.7, y, 0.8, 0.8, str(i + 1), size=26, bold=True, color=DARK, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, 6.75, y - 0.05, 6.0, 0.95, t, size=20, anchor=MSO_ANCHOR.MIDDLE)

# ================================================================== 4. Этапы
s = new_slide("Этапы проекта: от нагрузок к пилоту",
              "Мы шли шесть этапов. Все расчёты — в одной Excel-модели: каждое число получено формулой "
              "из исходных данных кейса, а 19 автоматических проверок повторяют критерии раздела 9 постановки.")
steps = ["Нагрузки и исходные данные", "Насосы: рабочие точки и резерв", "Почасовой баланс 3 сценариев",
         "Накопитель и уровни управления", "Очистка и качество воды", "Экономика и пилот"]
for i, t in enumerate(steps):
    x = 0.6 + i * 2.05
    box(s, x, 1.8, 1.85, 2.6, TINT)
    box(s, x + 0.55, 2.0, 0.75, 0.75, TEAL, shape=MSO_SHAPE.OVAL)
    text(s, x + 0.55, 2.0, 0.75, 0.75, str(i + 1), size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.1, 2.9, 1.65, 1.45, t, size=18, align=PP_ALIGN.CENTER)
    if i < 5:
        arrow(s, x + 1.87, 3.1, x + 2.03, 3.1, color=AMBER, w=2.5)
box(s, 0.6, 4.85, 12.1, 1.6, DARK)
text(s, 0.95, 4.95, 11.4, 1.4,
     [[("Инструмент: ", {"bold": True, "color": AMBER}),
       ("воспроизводимая Excel-модель — все результаты получены формулами из данных кейса; ", {"color": WHITE}),
       (f"{N_CHECKS} автоматических проверок", {"bold": True, "color": AMBER}),
       (" повторяют критерии приёмки кейса.", {"color": WHITE})]],
     size=20, anchor=MSO_ANCHOR.MIDDLE)

# ================================================================== 5. Схема
s = new_slide("Гидравлическая схема: накопитель, насосная, отстойник и две ветви",
              "Вода собирается в накопителе-водосборнике, насосная откачивает её в отстойник. После отстойника "
              "300 м³/ч идёт через фильтр к потребителю, остальное — к условному приёмнику. "
              "Отделённая взвесь уходит в осадок.")
row1 = [("Карьерный приток", f"450 → 1000 м³/ч\nпик 1200"),
        ("Накопитель", f"полезный {num(D['V_use'])} м³"),
        ("Насосная", "2 × P2:\n1 раб. + 1 резерв"),
        ("Отстойник S", f"500 → {num(D['cS'])} мг/л")]
for i, (h, t) in enumerate(row1):
    x = 0.6 + i * 3.15
    boxtext(s, x, 1.65, 2.6, 1.5, [[(h, {"bold": True})], t.split("\n")[0]] +
            ([t.split("\n")[1]] if "\n" in t else []), fill=TINT if i != 2 else TEAL,
            color=DARK if i != 2 else WHITE)
    if i < 3:
        arrow(s, x + 2.62, 2.4, x + 3.13, 2.4, color=DARK)
boxtext(s, 10.05, 3.85, 2.6, 1.2, [[("Фильтр F", {"bold": True})], f"{num(D['cS'])} → {num(D['cF'])} мг/л"])
arrow(s, 11.35, 3.17, 11.35, 3.83, color=TEAL)
text(s, 11.45, 3.25, 1.6, 0.5, "300 м³/ч", size=18, color=TEAL, bold=True)
boxtext(s, 10.05, 5.55, 2.6, 1.2, [[("Потребитель", {"bold": True})], f"{num(D['cF'])} мг/л ≤ 30"],
        fill=None, line=TEAL)
arrow(s, 11.35, 5.07, 11.35, 5.53, color=TEAL)
boxtext(s, 5.9, 3.85, 3.4, 1.2, [[("Условный приёмник", {"bold": True})], f"700–1000 м³/ч, {num(D['cS'])} мг/л"],
        fill=None, line=DARK)
arrow(s, 10.5, 3.17, 8.9, 3.83, color=DARK)
boxtext(s, 0.6, 3.85, 4.4, 1.2, [[("Осадок S + F", {"bold": True})],
                                 f"{num(D['sl_kgh'])} кг/ч сухого вещества"], fill=None, line=AMBER)
text(s, 0.6, 5.55, 8.4, 1.2,
     "Накопитель — регулирующий объём перед насосной: сглаживает пик и принимает воду при отказе, "
     "но не служит постоянным местом размещения притока.", size=18, color=GREY)

# ================================================================== 6. Насосы
s = new_slide("Один насос P2 перекрывает 1000 м³/ч — резерв обеспечивает один агрегат",
              "Рабочая точка — пересечение характеристики насосов и сети. Один P2 даёт 1111 м³/ч при напоре 91 м, "
              "это больше среднего притока, значит второй P2 — полноценный резерв. У P1 один насос даёт только "
              "816 м³/ч, поэтому в конфигурации B два рабочих и один резервный.")
cd = XyChartData()


def curve(name, h0, a, n, qmax):
    se = cd.add_series(name)
    q = 0
    while q <= qmax:
        h = h0 - a * (q / n) ** 2
        if h < 45:
            break
        se.add_data_point(q, h)
        q += 50


net = cd.add_series("Сеть: 60 + 0,000025·Q²")
for q in range(0, 1801, 50):
    net.add_data_point(q, 60 + 0.000025 * q * q)
curve("1 × P2 (A)", 165, 6e-5, 1, 1800)
curve("2 × P1 (B)", 130, 8e-5, 2, 1800)
pts = cd.add_series("Рабочие точки")
pts.add_data_point(D["qA"], D["hA"])
pts.add_data_point(D["qB2"], D["hB2"])
gf = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, Inches(0.5), Inches(1.45), Inches(7.6),
                        Inches(5.35), cd)
ch = gf.chart
style_chart(ch)
axis(ch.category_axis, "Q, м³/ч", 0, 1800, "#,##0", grid=False, major=300)
axis(ch.value_axis, "H, м", 40, 180, "0", major=20)
for se, c, w in zip(ch.series, (DARK, TEAL, AMBER, DARK), (2.5, 3.5, 3.5, 0)):
    se.smooth = True
    if w:
        se.format.line.color.rgb = c
        se.format.line.width = Pt(w)
    else:
        se.format.line.fill.background()
        se.marker.style = XL_MARKER_STYLE.CIRCLE
        se.marker.size = 12
        se.marker.format.fill.solid()
        se.marker.format.fill.fore_color.rgb = DARK
        se.marker.format.line.color.rgb = WHITE
ch.series[0].format.line.dash_style = MSO_LINE.DASH
box(s, 8.4, 1.55, 4.3, 2.45, TINT)
text(s, 8.65, 1.65, 3.9, 2.3,
     [[("A: 2 × P2 ", {"bold": True, "color": TEAL}), ("(1 + 1 резерв)", {})],
      f"{num(D['qA'])} м³/ч при {num(D['hA'], 1)} м",
      f"{num(D['eA'], 3)} кВт·ч/м³",
      f"отказ агрегата: {num(D['qA'])} м³/ч ≥ 1000"], size=18, space=4)
box(s, 8.4, 4.2, 4.3, 2.45, fill=None, line=LGREY)
text(s, 8.65, 4.3, 3.9, 2.3,
     [[("B: 3 × P1 ", {"bold": True, "color": AMBER}), ("(2 + 1 резерв)", {})],
      f"{num(D['qB2'])} м³/ч при {num(D['hB2'], 1)} м",
      f"{num(D['eB2'], 3)} кВт·ч/м³",
      f"отказ агрегата: {num(D['qB2'])} м³/ч ≥ 1000"], size=18, space=4)

# ================================================================== 7. Пик и отказ
s = new_slide(f"Пик сглаживается к {D['pkA_ret']}-му часу, после отказа уровень восстанавливается за {D['recA']} ч",
              "Слева — сутки с пиком: 6 часов приток 1200 м³/ч больше подачи одного P2 (1111), в накопителе "
              "собирается 531 м³, к 9-му часу он пуст; у B накопления нет. Справа — отказ на 2 часа: накапливается "
              "2000 м³, A восстанавливается одним P2 за 18 часов — второй P2 включать нельзя, вместе они дают "
              "1620 м³/ч при пределе отстойника 1500. B включает третий P1 и восстанавливается за 6 часов.")
for k, (ttl, cats, ser, ymax, step, xt) in enumerate((
        ("Сутки с пиком", [str(h) for h in range(1, 25)],
         [("A: 1 × P2", D["pkA"], TEAL, 3.5), ("B: 2 × P1", D["pkB"], AMBER, 3.5)], 600, 100, "час суток"),
        ("Отказ откачки 2 ч", [str(h) for h in range(1, 27)],
         [("A: 1 × P2", D["flA"], TEAL, 3.5), ("B: 2–3 × P1", D["flB"], AMBER, 3.5),
          ("Полезный объём", [D["V_use"]] * 26, DARK, 2)], 3000, 500, "час от начала отказа"))):
    cd = CategoryChartData()
    cd.categories = cats
    for nm, vals, _, _ in ser:
        cd.add_series(nm, [round(x, 1) for x in vals])
    gf = s.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(0.5 + k * 6.3), Inches(1.45), Inches(6.1), Inches(4.4), cd)
    ch = gf.chart
    style_chart(ch)
    ch.has_title = True
    ch.chart_title.text_frame.text = ttl
    tr = ch.chart_title.text_frame.paragraphs[0].runs[0]
    tr.font.size, tr.font.bold, tr.font.name = Pt(18), True, FONT
    tr.font.color.rgb = DARK
    axis(ch.category_axis, xt, grid=False)
    axis(ch.value_axis, "м³", 0, ymax, "#,##0", major=step)
    for se, (_, _, c, w) in zip(ch.series, ser):
        se.smooth = False
        se.format.line.color.rgb = c
        se.format.line.width = Pt(w)
        se.marker.style = XL_MARKER_STYLE.NONE
    if k == 1:
        ch.series[2].format.line.dash_style = MSO_LINE.DASH
text(s, 0.6, 5.95, 6.0, 0.85,
     [[("Пик: ", {"bold": True, "color": TEAL}),
       (f"максимум {num(D['pkA_max'])} м³ — {num(D['pkA_max'] / D['V_use'] * 100)} % полезного объёма; "
        "у B накопления нет", {})]], size=18)
text(s, 6.85, 5.95, 6.0, 0.85,
     [[("Отказ 2 ч: ", {"bold": True, "color": TEAL}),
       (f"{num(D['flA_max'])} м³; второй P2 не включаем — {num(D['qA2'])} > 1500 м³/ч", {})]], size=18)

# ================================================================== 8. Стресс-тест
s = new_slide(f"Стресс-тест: запас A по притоку — {num(AD['A']['marg'] * 100, 1)} %, у B и D — "
              f"около {num(AD['D']['marg'] * 100)} %",
              "Сверх требований кейса проверили запас прочности четырёх насосных. A выполняет все требования и "
              "дешевле, но восстановление после отказа за 24 часа перестаёт выполняться уже при притоке 1026 м³/ч, "
              "а отказ, совпавший с пиком, выводит уровень за полезный объём. B и D держат приток до ~1320 м³/ч. "
              "C — P2 плюс P1 — дешевле A, но P2 нужно ремонтировать не дольше 7 часов. Предложение: строим A и "
              "резервируем место под P1. Частотный регулятор окупается только при нынешних 450 м³/ч, ночной тариф "
              "Хакасии даёт всего 0,26 млн руб. в год.")
rows = [("", "A", "B", "C", "D"),
        ("Состав", "2×P2", "3×P1", "P2+P1", "2×P2+P1"),
        ("Затраты за 5 лет, млн руб.",) + tuple(num(AD[k]["tot"], 1) for k in "ABCD"),
        ("Выдерживает отказ в начале пика",) + tuple(AD[k]["op"] for k in "ABCD"),
        ("Предельный приток, м³/ч",) + tuple(num(AD[k]["qmax"]) for k in "ABCD"),
        ("Резерв без условий",) + tuple(AD[k]["res"] for k in "ABCD")]
tbl = s.shapes.add_table(len(rows), 5, Inches(0.6), Inches(1.6), Inches(7.6), Inches(4.1)).table
tbl.columns[0].width = Inches(3.2)
for j in range(1, 5):
    tbl.columns[j].width = Inches(1.1)
for i, row in enumerate(rows):
    for j, val in enumerate(row):
        cell = tbl.cell(i, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK if i == 0 else (TINT if i % 2 else WHITE)
        cell.text_frame.text = str(val).replace("НЕТ", "нет")
        cell.text_frame.word_wrap = True
        par = cell.text_frame.paragraphs[0]
        par.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
        r = par.runs[0] if par.runs else par.add_run()
        r.font.size, r.font.name = Pt(18 if i else 20), FONT
        r.font.bold = i == 0 or j == 0
        weak = j > 0 and str(val) in ("НЕТ",) or (j == 1 and i == 4)
        r.font.color.rgb = WHITE if i == 0 else (AMBER if weak else DARK)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
box(s, 8.5, 1.6, 4.2, 4.1, TINT)
text(s, 8.7, 1.7, 3.85, 3.95,
     [[("Решение", {"bold": True, "color": TEAL, "size": 20})],
      "Строим A: все требования кейса при меньших затратах.",
      f"Резервируем место под P1: если приток выше {num(AD['A']['qmax'])} м³/ч — переход к D за 8 млн руб.",
      f"C дешевле A на {num(-AD['C']['diff'], 1)} млн, но ремонт P2 — не дольше {num(TREP)} ч."],
     size=18, space=8)
text(s, 0.6, 5.95, 12.1, 0.85,
     [[("Расширения: ", {"bold": True, "color": TEAL}),
       (f"частотный регулятор окупается только при 450 м³/ч ({num(VFD450, 1)} млн руб./год экономии); "
        f"ночной тариф Хакасии (ночь в 0,46 раза дешевле) даёт лишь {num(TAR_A, 2)} млн руб./год.", {})]], size=18)
src(s, [14], "; расчёт — «Дополнение_расчёты.xlsx»")

# ================================================================== 9. Накопитель
s = new_slide(f"Полезный объём {num(D['V_use'])} м³ с запасом {num(D['margin'] * 100)} %, полный — {num(D['V_full'])} м³",
              "Полезный объём выбран по худшему сценарию — отказу, 2000 м³, плюс 25 % запаса. Осадок, недоступный "
              "объём и свободный борт добавлены сверху и учтены по одному разу: полный объём = (2500 + 400 + 200) / 0,85. "
              "При глубине 4,5 м ёмкость с полосой обслуживания занимает 1054 м² из 1200 допустимых.")
cd = CategoryChartData()
cd.categories = ["Полный объём"]
parts = [("Осадок", 400, DARK), ("Недоступный", 200, LGREY), ("Полезный", D["V_use"], TEAL),
         ("Свободный борт", D["V_fb"], AMBER)]
for nm, val, _ in parts:
    cd.add_series(nm, [round(val)])
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(0.5), Inches(1.5), Inches(12.3), Inches(2.3), cd)
ch = gf.chart
style_chart(ch)
ch.plots[0].gap_width = 40
ch.plots[0].overlap = 100
axis(ch.value_axis, None, 0, 3700, "#,##0", major=500)
ch.category_axis.visible = False
for se, (nm, val, c) in zip(ch.series, parts):
    se.format.fill.solid()
    se.format.fill.fore_color.rgb = c
    if val >= 400:
        dl = se.points[0].data_label
        dl.has_text_frame = True
        dl.text_frame.text = num(val)
        dl.position = XL_LABEL_POSITION.CENTER
        r = dl.text_frame.paragraphs[0].runs[0]
        r.font.size = Pt(18)
        r.font.bold = True
        r.font.color.rgb = WHITE if c != AMBER else DARK
cards = [(f"{num(D['V_req'])} м³", "максимум накопления (отказ)"),
         (f"{num(D['depth'], 1)} м", "глубина (предел 5 м)"),
         (f"{num(D['foot'])} м²", "площадь (предел 1200 м²)"),
         (f"{num(D['st_capex'], 1)}", "млн руб. — CAPEX накопителя")]
for i, (b, sm) in enumerate(cards):
    stat(s, 0.6 + i * 3.1, 4.25, 2.85, b, sm, big_size=34, h=1.9)
text(s, 0.6, 6.3, 11.4, 0.5, "V_полн = (V_полезн + V_осадка + V_недоступн) / (1 − 0,15)", size=18, color=GREY)

# ================================================================== 10. Алгоритм
s = new_slide("Насосами управляют уровни: стоп, пуск, пуск 2, авария",
              "Высоты — от дна ёмкости. Ниже уровня «стоп» насосы выключены — это защита от сухого хода. "
              "На «пуске» работает базовая ступень, на «пуске 2» включается дополнительный насос, но только если "
              "суммарная подача не превысит предел отстойника. «Авария» — сигнал диспетчеру и автоматический "
              "ввод резерва.")
H_TANK = 4.9
sc = H_TANK / D["depth"]
x0, y_bot, w_t = 0.9, 6.55, 1.6
zones = [("Осадок 400 м³", 0, D["lv"][1][1], DARK), ("Недоступный 200 м³", D["lv"][1][1], D["lv"][2][1], LGREY),
         (f"Полезный {num(D['V_use'])} м³", D["lv"][2][1], D["lv"][6][1], TEAL),
         ("Борт 15 %", D["lv"][6][1], D["lv"][7][1], WHITE)]
for nm, h1, h2, c in zones:
    box(s, x0, y_bot - h2 * sc, w_t, (h2 - h1) * sc, c, line=DARK if c == WHITE else None,
        shape=MSO_SHAPE.RECTANGLE, lw=1)
box(s, x0, y_bot - H_TANK, w_t, H_TANK, None, line=DARK, shape=MSO_SHAPE.RECTANGLE, lw=2)
lab_y = {0: y_bot - 0.32, 1: y_bot - 0.75, 2: y_bot - (D["lv"][2][1] + D["lv"][6][1]) / 2 * sc - 0.17,
         3: y_bot - (D["lv"][6][1] + D["lv"][7][1]) / 2 * sc - 0.17}
for i, (nm, *_rest) in enumerate(zones):
    text(s, x0 + w_t + 0.12, lab_y[i], 2.6, 0.36, nm, size=18, color=DARK, anchor=MSO_ANCHOR.MIDDLE)
rows = [("Отметка", "м", "Действие"),
        ("Стоп", D["lv"][2][1], "все насосы выключены"),
        ("Пуск", D["lv"][3][1], "базовая ступень (A: 1×P2; B: 2×P1)"),
        ("Пуск 2", D["lv"][4][1], "доп. насос при подаче ≤ 1500 м³/ч (B)"),
        ("Авария", D["lv"][5][1], "сигнал диспетчеру, ввод резерва (АВР)"),
        ("Максимум", D["lv"][6][1], "верх полезного объёма; выше — борт")]
tbl = s.shapes.add_table(len(rows), 3, Inches(5.3), Inches(1.6), Inches(7.4), Inches(4.2)).table
tbl.columns[0].width = Inches(1.6)
tbl.columns[1].width = Inches(0.9)
tbl.columns[2].width = Inches(4.9)
for i, row in enumerate(rows):
    for j, val in enumerate(row):
        cell = tbl.cell(i, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK if i == 0 else (TINT if i % 2 else WHITE)
        tf = cell.text_frame
        tf.text = num(val, 2) if isinstance(val, float) else str(val)
        p = tf.paragraphs[0]
        r = p.runs[0]
        r.font.size = Pt(18)
        r.font.name = FONT
        r.font.bold = i == 0 or j == 0
        r.font.color.rgb = WHITE if i == 0 else DARK
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
text(s, 5.3, 6.0, 7.4, 0.8, "Ротация агрегатов по наработке; после отключения питания — каскадный пуск.",
     size=18, color=GREY)

# ================================================================== 11. Очистка
s = new_slide(f"Отстойник для всего потока и фильтр для повторной воды: {num(D['cF'])} мг/л при цели 30",
              "Отстойник S обязателен в любой схеме: сырая вода содержит 500 мг/л взвеси, а для оставшегося потока "
              "нужно не больше 100. После S — 50 мг/л. Фильтр F доводит 300 м³/ч до 10 мг/л при цели 30. "
              "Водный и компонентный балансы сходятся. Сульфаты и жёсткость эти операции не меняют — их пригодность "
              "для потребителя проверим на пилоте.")
boxtext(s, 0.6, 2.55, 2.5, 1.5, [[("Сырая вода", {"bold": True})], "1000 м³/ч", "500 мг/л"])
arrow(s, 3.12, 3.3, 3.75, 3.3)
boxtext(s, 3.8, 2.55, 2.5, 1.5, [[("Отстойник S", {"bold": True})], "−90 % взвеси", f"→ {num(D['cS'])} мг/л"],
        fill=TEAL, color=WHITE)
arrow(s, 6.32, 3.0, 7.15, 2.15, color=TEAL)
arrow(s, 6.32, 3.6, 7.15, 4.45, color=DARK)
boxtext(s, 7.2, 1.5, 2.5, 1.4, [[("Фильтр F", {"bold": True})], "300 м³/ч, −80 %"])
arrow(s, 9.72, 2.2, 10.15, 2.2, color=TEAL)
boxtext(s, 10.2, 1.5, 2.5, 1.4, [[("Потребитель", {"bold": True})], f"{num(D['cF'])} мг/л ≤ 30"],
        fill=None, line=TEAL)
boxtext(s, 7.2, 3.8, 5.5, 1.4, [[("Условный приёмник", {"bold": True})],
                                f"700–1000 м³/ч, {num(D['cS'])} мг/л ≤ 100"], fill=None, line=DARK)
box(s, 0.6, 5.55, 12.1, 1.25, TINT)
text(s, 0.85, 5.6, 11.6, 1.15,
     [[("Практика разреза: ", {"bold": True, "color": TEAL}),
       ("действующие очистные (с 2018 г., 350 → 670 м³/ч) — двухступенчатая фильтрация АКВАФЛОУ и "
        "обеззараживание УФ; наша схема S + F — то же направление. ", {}),
       ("Осадок: ", {"bold": True, "color": TEAL}),
       (f"{num(D['sl_kgh'])} кг/ч ≈ {num(D['sl_t'])} т/год.", {})]],
     size=18, anchor=MSO_ANCHOR.MIDDLE)
src(s, [4, 5])

# ================================================================== 12. Экономика водоотлива
s = new_slide(f"Конфигурация A дешевле на {num(D['econ'], 1)} млн руб. за 5 лет за счёт энергии",
              "Сравниваем две работающие схемы за 5 лет в постоянных ценах. У A насосы дороже на 4 млн, но она "
              "тратит на 435 МВт·ч в год меньше энергии: один P2 работает с меньшим напором, чем два P1 параллельно. "
              "Затраты отстойника одинаковы — он нужен в обеих схемах.")
cd = CategoryChartData()
cd.categories = ["A: 2 × P2", "B: 3 × P1"]
series = [("CAPEX", D["capA"], D["capB"], DARK),
          ("Энергия насосов × 5 лет", D["eCostA"] * 5, D["eCostB"] * 5, TEAL),
          ("Отстойник S × 5 лет", D["sCost"] * 5, D["sCost"] * 5, LGREY),
          ("Обслуживание × 5 лет", D["mA"] * 5, D["mB"] * 5, AMBER)]
for nm, a, b, _ in series:
    cd.add_series(nm, [round(a, 2), round(b, 2)])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(0.5), Inches(1.45), Inches(7.6), Inches(5.35), cd)
ch = gf.chart
style_chart(ch)
ch.plots[0].gap_width = 60
ch.plots[0].overlap = 100
axis(ch.value_axis, "млн руб. за 5 лет", 0, 300, "#,##0", major=50)
axis(ch.category_axis, None, grid=False)
ch.category_axis.tick_labels.font.size = Pt(16)
ch.category_axis.tick_labels.font.color.rgb = DARK
for se, (nm, a, b, c) in zip(ch.series, series):
    se.format.fill.solid()
    se.format.fill.fore_color.rgb = c
    if nm.startswith("Обслуживание"):
        continue
    for k, val in enumerate((a, b)):
        dl = se.points[k].data_label
        dl.has_text_frame = True
        dl.text_frame.text = num(val, 1)
        dl.position = XL_LABEL_POSITION.CENTER
        r = dl.text_frame.paragraphs[0].runs[0]
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = WHITE if c in (DARK, TEAL) else DARK
stat(s, 8.4, 1.55, 4.3, f"{num(D['totA'], 1)} / {num(D['totB'], 1)}", "млн руб. за 5 лет: A / B", big_size=34)
stat(s, 8.4, 3.3, 4.3, f"{num(D['enA'])} / {num(D['enB'])}", "МВт·ч в год на откачку: A / B", big_size=34)
stat(s, 8.4, 5.05, 4.3, f"{num(D['specA'], 2)} руб./м³", "стоимость откачки 1 м³ (A)", big_size=34)

# ================================================================== 13. Повторная вода
s = new_slide(f"Повторная вода для фабрики разреза: +{num(D['effF'], 1)} млн руб. за 5 лет",
              "Потребитель — обогатительная фабрика, которая строится на территории Восточно-Бейского разреза: "
              "750 т/ч, 4,05 млн т в год. Наш спрос 300 м³/ч — это 0,44 м³ на тонну угля, для обогащения типично до "
              "2 м³/т. Эффект считаем отдельно: замещённая покупная вода минус затраты самой ветки. Окупаемость "
              "1,8 года. По взвеси вода подходит, сульфаты и жёсткость выше норм питьевой воды — требования фабрики "
              "проверим на пилоте.")
stat(s, 0.6, 1.6, 5.4, f"{num(D['reuse_m3'] / 1e6, 1)} млн м³/год", "покупной воды заменяется", big_size=34)
stat(s, 0.6, 3.35, 5.4, f"+{num(D['netF'], 1)} млн руб./год", f"чистый эффект; окупаемость {num(D['payF'], 1)} года",
     big_size=34)
stat(s, 0.6, 5.1, 5.4, "0,44 м³/т", "спрос фабрики 750 т/ч на тонну угля (типично до 2 м³/т)", big_size=34)
cd = CategoryChartData()
cd.categories = [f"F: фильтр, {num(D['cF'])} мг/л", f"M: мембрана, {num(D['cM'])} мг/л"]
cd.add_series("Эффект за 5 лет, млн руб.", [round(D["effF"], 1), round(D["effM"], 1)])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.4), Inches(1.45), Inches(6.4), Inches(3.75), cd)
ch = gf.chart
style_chart(ch, legend=False)
ch.has_title = True
ch.chart_title.text_frame.text = "Эффект за 5 лет, млн руб."
tr = ch.chart_title.text_frame.paragraphs[0].runs[0]
tr.font.size = Pt(18)
tr.font.bold = True
tr.font.color.rgb = DARK
tr.font.name = FONT
ch.plots[0].gap_width = 80
axis(ch.value_axis, None, 0, 45, "#,##0", major=10)
axis(ch.category_axis, None, grid=False)
ch.category_axis.tick_labels.font.size = Pt(16)
ch.category_axis.tick_labels.font.color.rgb = DARK
pl = ch.plots[0]
pl.has_data_labels = True
pl.data_labels.number_format = "#,##0.0"
pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
pl.data_labels.font.size = Pt(20)
pl.data_labels.font.bold = True
for k, c in enumerate((TEAL, LGREY)):
    pt = ch.series[0].points[k]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = c
text(s, 6.5, 5.3, 6.2, 1.3,
     [[("Качество: ", {"bold": True, "color": TEAL}),
       ("взвесь 10 мг/л подходит; сульфаты 700 мг/л и жёсткость ≈ 10 мг-экв/л выше норм питьевой воды — "
        "требования фабрики проверим на пилоте.", {})]], size=18)
src(s, [4, 6, 15, 16])

# ================================================================== 14. Ресурсы и пилот
s = new_slide(f"Пилот за {num(D['pil_tot'])} млн руб. из 100: доочистка, КИП и проверка расчёта",
              "Пилот берёт ветку повторной воды и систему контроля — это проверяет расчёт без риска для добычи. "
              "Бюджет 26 млн руб.: из них 5 млн — наши оценки на КИП и испытания. Программа идёт от обследования "
              "реальной сети к учениям по отказу на текущем притоке.")
rows = [("Статья", "млн руб.")] + [(nm.replace(" (принятая технология)", " F"), val) for nm, val in D["pil"]] + \
       [("Итого (лимит 100)", D["pil_tot"])]
short = {0: "Статья", 1: "Установка доочистки F", 2: "Магистраль к потребителю", 3: "КИПиА (оценка)",
         4: "Испытания, лаборатория (оценка)", 5: "Итого (лимит 100)"}
tbl = s.shapes.add_table(len(rows), 2, Inches(0.6), Inches(1.6), Inches(5.6), Inches(3.9)).table
tbl.columns[0].width = Inches(4.1)
tbl.columns[1].width = Inches(1.5)
for i, (nm, val) in enumerate(rows):
    for j, t in enumerate((short[i], val if i == 0 else num(val, 0))):
        cell = tbl.cell(i, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK if i in (0, len(rows) - 1) else (TINT if i % 2 else WHITE)
        cell.text_frame.text = str(t)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT if j == 1 else PP_ALIGN.LEFT
        r = p.runs[0]
        r.font.size = Pt(18)
        r.font.name = FONT
        r.font.bold = i in (0, len(rows) - 1)
        r.font.color.rgb = WHITE if i in (0, len(rows) - 1) else DARK
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
text(s, 0.6, 5.6, 5.8, 1.2, "Ресурсы: данные кейса и разреза, Excel-модель, команда. База пилота — "
                            "очистные разреза 350 м³/ч и строящаяся фабрика.", size=18, color=GREY)
src(s, [4, 6])
prog = [("Обследование", "фактический приток, характеристика сети"),
        ("КИП и уровни", "отладка алгоритма на действующей насосной"),
        ("Ветка F, 3 мес.", "взвесь ≤ 30 мг/л, фактические затраты"),
        ("Анализы у потребителя", "сульфаты, жёсткость, pH"),
        ("Учения по отказу", "на 450 м³/ч: рост и спад уровня")]
for i, (h, t) in enumerate(prog):
    y = 1.6 + i * 1.03
    box(s, 6.7, y + 0.1, 0.62, 0.62, TEAL, shape=MSO_SHAPE.OVAL)
    text(s, 6.7, y + 0.1, 0.62, 0.62, str(i + 1), size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, 7.5, y, 5.3, 0.95, [[(h, {"bold": True})], t], size=18, space=0, anchor=MSO_ANCHOR.MIDDLE)

# ================================================================== 15. Результаты
s = new_slide("Результат: все требования кейса выполнены",
              "Итог. Все 19 проверок пройдены, рекомендуем A с местом под P1. Стресс-тест показал узкое место A — "
              "запас по притоку 2,6 процента, поэтому пилот начинается с замера фактического притока. "
              "Ограничения: данные кейса учебные; реальные очистные разреза на 350 м³/ч стоили около 200 млн руб., "
              "а пригодность воды по сульфатам и жёсткости нужно подтвердить анализами.")
fails = D["failsA"]
box(s, 0.6, 1.6, 4.0, 4.0, TEAL)
text(s, 0.6, 1.9, 4.0, 1.6, f"{N_CHECKS - fails}/{N_CHECKS}", size=72, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.85, 3.6, 3.5, 1.8, "проверок по критериям кейса пройдено", size=20, color=WHITE,
     align=PP_ALIGN.CENTER)
res = ["Насосы в рабочих точках, резерв — 100 % одним агрегатом",
       f"Пик и отказ без перелива; восстановление {D['recA']} ч ≤ 24 ч",
       f"Объёмы по разу: полезный {num(D['V_use'])}, полный {num(D['V_full'])} м³",
       f"Взвесь {num(D['cF'])} ≤ 30 и {num(D['cS'])} ≤ 100 мг/л, балансы сходятся",
       f"A дешевле B на {num(D['econ'], 1)} млн; повторная вода +{num(D['effF'], 1)} млн",
       f"Стресс-тест: запас A — {num(AD['A']['marg'] * 100, 1)} %; путь к D заложен (+8 млн)"]
for i, t in enumerate(res):
    y = 1.6 + i * 0.68
    dot(s, 5.0, y + 0.15, 0.3)
    text(s, 5.5, y, 7.2, 0.6, t, size=18, anchor=MSO_ANCHOR.MIDDLE)
box(s, 0.6, 5.85, 12.1, 0.95, TINT)
text(s, 0.85, 5.88, 11.6, 0.9,
     [[("Ограничения: ", {"bold": True, "color": TEAL}),
       ("данные кейса учебные — реальные очистные разреза на 350 м³/ч стоили ≈ 200 млн руб.; "
        "приток и состав воды проверим на пилоте.", {})]],
     size=18, anchor=MSO_ANCHOR.MIDDLE)
src(s, [4])

# ================================================================== 16. Призыв
s = new_slide(None, "Обеспечьте водоотлив без остановки и сократите покупку воды — внедрите наше решение! "
                    "Спасибо за внимание, готовы ответить на вопросы.", dark=True)
text(s, 0.9, 1.4, 11.5, 2.4, "Обеспечьте водоотлив без остановки и сократите покупку воды. "
                             "Внедрите наше решение!", size=40, bold=True, color=WHITE)
cards = [(f"{D['recA']} ч", "восстановление после отказа"),
         (f"−{num(D['econ'], 1)} млн", "руб. за 5 лет на водоотливе"),
         (f"+{num(D['effF'], 1)} млн", "руб. за 5 лет от повторной воды")]
for i, (b, sm) in enumerate(cards):
    x = 0.9 + i * 3.9
    text(s, x, 4.1, 3.6, 0.9, b, size=40, bold=True, color=AMBER)
    text(s, x, 4.95, 3.6, 0.9, sm, size=18, color=TINT)
text(s, 0.9, 6.2, 11.0, 0.6, f"{AUTHORS} · {CITY}", size=18, color=LGREY)

prs.save(OUT)
print("saved", OUT, "slides:", N[0])
