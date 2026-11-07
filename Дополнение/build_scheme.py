"""Принципиальная гидравлическая схема водоотлива (дополнение к слайду 5).

Правила оформления — по ГОСТ 2.701, 2.704 и УГО по ГОСТ 2.780–2.785 в изложении пособия
Акуленко В.М. и др. «Практические работы по дисциплине «Гидропневмотранспортные машины...»
(ДВФУ, 2018), практическая работа № 1:
  * поток энергии снизу вверх и слева направо;
  * внизу — энергообеспечивающая часть (накопитель, насосы), в середине — направляющая
    и регулирующая арматура, вверху — исполнительная часть (очистка, потребители);
  * устройства — в исходном положении (готовность к пуску);
  * позиционные обозначения (буквенный код + номер) — справа и сверху от УГО;
  * законченные функциональные блоки — тонкой штрихпунктирной линией с обозначением блока.

Выход (папка «Дополнение»):
  Схема_гидравлическая_принципиальная.pdf — лист A3 со схемой, примечаниями и основной
      надписью + лист A4 с перечнем элементов;
  рисунки_дополнения/схема_лист_A3.png — лист A3, повёрнутый (для документов);
  рисунки_дополнения/схема_для_слайда.png — упрощённая схема крупно (слайд и документы);
  Слайд_5а_гидравлическая_схема.pptx / .png — вставка к слайду 5 в стиле презентации.

Числа берутся из основной книги «Расчёт_водоотлив.xlsx» (пересчитанной).
Запуск: python -X utf8 Дополнение/build_scheme.py
"""
import os
import re
import sys

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import Arc, Circle, Polygon, Rectangle
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
MM = 1 / 25.4                      # дюймов в миллиметре
plt.rcParams["font.family"] = "Arial"
plt.rcParams["pdf.fonttype"] = 42  # шрифт встраивается как TrueType (кириллица копируется)

# ----------------------------------------------------------------- данные из книги
wb = openpyxl.load_workbook(os.path.join(ROOT, "Расчёт_водоотлив.xlsx"), data_only=True)
_n, _o, _s, _r = wb["Насосы"], wb["Очистка"], wb["Накопитель"], wb["Решения"]
D = {
    "qA": _n["B16"].value, "hA": _n["D16"].value, "NA": _n["F16"].value, "q2": _n["B17"].value,
    "cS": _o["C14"].value, "cF": _o["C21"].value, "slS": _o["C15"].value, "slF": _o["C23"].value,
    "Vuse": _s["B4"].value, "Vfull": _s["B8"].value, "depth": _s["B16"].value,
}
LEV = {row[0].value: row[2].value for row in _s.iter_rows(min_row=28, max_row=35, max_col=3)}
D["h_stop"] = next(v for k, v in LEV.items() if k and k.startswith("«Стоп»"))
D["h_on"] = LEV["«Пуск» базовой ступени"]
D["h_alarm"] = LEV["«Авария»"]
D["h_top"] = LEV["Верх полезного объёма"]
for k in ("qA", "hA", "q2", "cS", "cF", "Vuse", "Vfull"):
    assert D[k], f"нет значения {k}: пересчитайте книгу (python build_all.py)"


def num(x, d=0):
    s = f"{x:,.{d}f}".replace(",", " ").replace(".", ",")
    return s


# Титульные данные — из build_presentation.py, чтобы правились в одном месте
_bp = open(os.path.join(ROOT, "build_presentation.py"), encoding="utf-8").read()
CONST = dict(re.findall(r'^(TEAM|AUTHORS|ORG|CITY|YEAR) = "([^"]*)"', _bp, re.M))
AUTHOR = "[Фамилия]" if CONST["AUTHORS"].startswith("[") else CONST["AUTHORS"].split(",")[0].split()[0]
ORGTXT = "[Организация]" if CONST["ORG"].startswith("[") else CONST["ORG"]
DESIGN = "ИК-2027.16.001"

# ----------------------------------------------------------------- примитивы УГО (мм)
K = "black"
LW_P, LW_S, LW_T = 1.5, 1.0, 0.55     # трубопровод, контур УГО, тонкая линия (pt)


class Draw:
    """Все размеры — в миллиметрах листа; fs — размер шрифта обозначений, pt."""

    def __init__(self, ax, fs=11, fs_small=8.5):
        self.ax, self.fs, self.fss = ax, fs, fs_small

    # --- линии
    def pipe(self, *pts, ls="-", lw=LW_P, z=1):
        xs, ys = zip(*pts)
        self.ax.plot(xs, ys, color=K, lw=lw, ls=ls, solid_capstyle="butt", solid_joinstyle="miter", zorder=z)

    def flow(self, x, y, d, size=2.6):
        """Стрелка направления потока на линии (залитый треугольник)."""
        L, W = size, size * 0.42
        v = {"u": ((0, L / 2), (-W, -L / 2), (W, -L / 2)), "d": ((0, -L / 2), (-W, L / 2), (W, L / 2)),
             "r": ((L / 2, 0), (-L / 2, W), (-L / 2, -W)), "l": ((-L / 2, 0), (L / 2, W), (L / 2, -W))}[d]
        self.ax.add_patch(Polygon([(x + a, y + b) for a, b in v], closed=True, fc=K, ec=K, lw=0.3, zorder=4))

    def node(self, x, y):
        self.ax.add_patch(Circle((x, y), 0.9, fc=K, ec=K, zorder=5))

    def label(self, x, y, s, size=None, ha="left", va="bottom", **kw):
        self.ax.text(x, y, s, fontsize=size or self.fs, ha=ha, va=va, zorder=6, **kw)

    def note(self, x, y, s, ha="left", va="bottom", **kw):
        self.ax.text(x, y, s, fontsize=self.fss, ha=ha, va=va, zorder=6, linespacing=1.15, **kw)

    def _c(self, x, y, r, fc="white", lw=LW_S):
        self.ax.add_patch(Circle((x, y), r, fc=fc, ec=K, lw=lw, zorder=3))

    # --- источник: место присоединения к источнику (залитый треугольник)
    def source(self, x, y):
        self.ax.add_patch(Polygon([(x, y - 2.2), (x, y + 2.2), (x + 4.5, y)], fc=K, ec=K, zorder=4))

    # --- бак открытый (общее обозначение)
    def tank(self, x0, x1, y0, y1):
        self.pipe((x0, y1), (x0, y0), (x1, y0), (x1, y1), lw=LW_S + 0.4)

    # --- насос нерегулируемый с постоянным направлением потока (выход вверх) + приводной двигатель
    def pump(self, cx, cy, r=7):
        self._c(cx, cy, r)
        h = r * 0.55
        self.ax.add_patch(Polygon([(cx - h * 0.62, cy + r - h), (cx + h * 0.62, cy + r - h), (cx, cy + r)],
                                  fc=K, ec=K, lw=0.3, zorder=4))
        for dy in (-1.1, 1.1):                                   # вал
            self.pipe((cx + r, cy + dy), (cx + r + 5, cy + dy), lw=LW_S)
        x0 = cx + r + 5
        self.ax.add_patch(Rectangle((x0, cy - 6), 12, 12, fc="white", ec=K, lw=LW_S, zorder=3))
        self.label(x0 + 6, cy, "М", ha="center", va="center")

    # --- запорная арматура (вентиль/задвижка); closed=True — закрытая (зачернённая)
    def valve(self, cx, cy, vertical=True, s=3.3, closed=False):
        if vertical:
            t1 = [(cx - s * 0.8, cy - s), (cx + s * 0.8, cy - s), (cx, cy)]
            t2 = [(cx - s * 0.8, cy + s), (cx + s * 0.8, cy + s), (cx, cy)]
        else:
            t1 = [(cx - s, cy - s * 0.8), (cx - s, cy + s * 0.8), (cx, cy)]
            t2 = [(cx + s, cy - s * 0.8), (cx + s, cy + s * 0.8), (cx, cy)]
        for t in (t1, t2):
            self.ax.add_patch(Polygon(t, closed=True, fc=K if closed else "white", ec=K, lw=LW_S, zorder=3))

    # --- обратный клапан (поток снизу вверх: седло-«V» снизу, шарик над ним)
    def check(self, cx, cy, r=2.4):
        self.ax.add_patch(Polygon([(cx - r - 0.9, cy + 0.6), (cx + r + 0.9, cy + 0.6), (cx, cy - r - 1.6)],
                                  closed=True, fc="white", ec="white", lw=0, zorder=2.5))
        self.pipe((cx - r - 0.9, cy + 0.6), (cx, cy - r - 1.6), (cx + r + 0.9, cy + 0.6), lw=LW_S, z=3)
        self._c(cx, cy + r * 0.55 + 0.6, r)

    # --- манометр: круг со стрелкой, ножка вниз
    def gauge(self, cx, cy, r=4.2):
        self._c(cx, cy, r)
        self.ax.annotate("", xy=(cx + r * 0.95, cy + r * 0.95), xytext=(cx - r * 0.6, cy - r * 0.6),
                         arrowprops=dict(arrowstyle="-|>,head_length=0.5,head_width=0.22", lw=LW_S, color=K),
                         zorder=4)

    # --- расходомер (круг с двумя дугами), встроенный в линию
    def flowmeter(self, cx, cy, vertical=False, r=4.2):
        self._c(cx, cy, r)
        if vertical:   # дуги слева и справа
            self.ax.add_patch(Arc((cx - r * 1.55, cy), r * 2, r * 2.2, theta1=-38, theta2=38, lw=LW_S, zorder=4))
            self.ax.add_patch(Arc((cx + r * 1.55, cy), r * 2, r * 2.2, theta1=142, theta2=218, lw=LW_S, zorder=4))
        else:          # дуги сверху и снизу
            self.ax.add_patch(Arc((cx, cy + r * 1.55), r * 2.2, r * 2, theta1=232, theta2=308, lw=LW_S, zorder=4))
            self.ax.add_patch(Arc((cx, cy - r * 1.55), r * 2.2, r * 2, theta1=52, theta2=128, lw=LW_S, zorder=4))

    # --- указатель уровня жидкости (только вертикально): круг с двумя горизонталями
    def level(self, cx, cy, r=4.2):
        self._c(cx, cy, r)
        for dy in (-0.7, 0.7):
            w = (r * r - dy * dy) ** 0.5
            self.pipe((cx - w, cy + dy), (cx + w, cy + dy), lw=LW_S, z=4)

    # --- место присоединения измерительного прибора (контрольная точка), отвод вниз
    def tap(self, x, y, down=True, L=11):
        sg = -1 if down else 1
        self.node(x, y)
        y1 = y + sg * (L - 6)
        self.pipe((x, y), (x, y1), lw=LW_S)
        cy = y1 + sg * 3.2
        self.pipe((x - 2.2, y1 + sg * 0.2), (x, y1 + sg * 1.4), (x + 2.2, y1 + sg * 0.2), lw=LW_S)
        self._c(x, cy, 2.0)
        self.pipe((x - 2.4, cy + sg * 2.6), (x + 2.4, cy + sg * 2.6), lw=LW_S)
        return cy + sg * 2.6

    # --- фильтр: ромб с пунктиром (поток слева направо)
    def filt(self, cx, cy, s=7):
        self.ax.add_patch(Polygon([(cx - s, cy), (cx, cy + s), (cx + s, cy), (cx, cy - s)], closed=True,
                                  fc="white", ec=K, lw=LW_S, zorder=3))
        self.pipe((cx, cy + s), (cx, cy - s), ls=(0, (2.2, 1.6)), lw=LW_T + 0.2, z=4)

    # --- регулятор расхода двухлинейный (поток слева направо)
    def regulator(self, cx, cy, w=16, h=11):
        self.ax.add_patch(Rectangle((cx - w / 2, cy - h / 2), w, h, fc="white", ec=K, lw=LW_S, zorder=3))
        self.pipe((cx - w / 2, cy), (cx + w / 2 - 3.2, cy), lw=LW_S, z=4)
        self.flow(cx + w / 2 - 1.8, cy, "r", size=2.8)
        for sg in (1, -1):
            self.ax.add_patch(Arc((cx - 1.5, cy + sg * 4.6), 7.5, 6, theta1=(250 if sg > 0 else 70),
                                  theta2=(290 if sg > 0 else 110), lw=LW_S, zorder=4))
        self.ax.annotate("", xy=(cx + 2.4, cy + h / 2 - 0.8), xytext=(cx - 5.4, cy - h / 2 + 0.8),
                         arrowprops=dict(arrowstyle="-|>,head_length=0.45,head_width=0.2", lw=LW_S, color=K),
                         zorder=5)

    # --- гаситель гидравлического удара (подключён снизу)
    def damper(self, cx, cy, w=18, h=10):
        self.ax.add_patch(Rectangle((cx - w / 2, cy - h / 2), w, h, fc="white", ec=K, lw=LW_S, zorder=3))
        self.ax.add_patch(Rectangle((cx - w / 2 + 3, cy - h / 2 + 2.2), w - 6, h - 4.4, fc="white", ec=K,
                                    lw=LW_T, zorder=4))
        self.pipe((cx - 3.5, cy), (cx + 2, cy), lw=LW_T, z=5)
        self.flow(cx + 2.3, cy, "r", size=2.2)

    # --- отстойник бассейновый; (x0, y0) — левый нижний угол, вход сверху слева, выход справа
    def settler(self, x0, y0, sc=2.4):
        p = [(0, 10), (14, 0), (22, 0), (22, 9), (28, 9), (28, 15)]
        self.pipe(*[(x0 + a * sc, y0 + b * sc) for a, b in p], lw=LW_S + 0.4)
        return (x0 + 28 * sc, y0 + 15 * sc), (x0 + 18 * sc, y0)   # выход перелива, выпуск осадка

    # --- конец трубопровода с фланцевой заглушкой (труба подходит сверху)
    def plug(self, x, y):
        for dy in (0, -1.4):
            self.pipe((x - 3, y + dy), (x + 3, y + dy), lw=LW_S + 0.3)

    # --- функциональный блок: тонкий штрихпунктир
    def block(self, x0, y0, x1, y1):
        self.ax.add_patch(Rectangle((x0, y0), x1 - x0, y1 - y0, fill=False, ec=K, lw=LW_T,
                                    ls=(0, (9, 2, 1.2, 2)), zorder=2))


# ----------------------------------------------------------------- сама схема
def draw_scheme(dr, texts=True):
    """Рисует схему в координатах листа A3 (мм). texts=False — без длинных подписей (для слайда)."""
    big = dr.fs
    # Б1 — накопитель-водосборник
    dr.tank(35, 178, 20, 58)
    dr.label(179, 59, "Б1")
    if texts:
        dr.note(100, 22.5, f"накопитель: полезный {num(D['Vuse'])} м³, полный {num(D['Vfull'])} м³",
                ha="center")
    # приток из карьера — место присоединения к источнику
    dr.source(23, 76)
    dr.pipe((27.5, 76), (50, 76), (50, 50))
    dr.flow(39, 76, "r")
    dr.flow(50, 63, "d")
    dr.note(23, 79.5, "Приток из карьера\n450 → 1000 м³/ч\nпик 1200" if texts else "Приток\nиз карьера")

    # А3 — уровнемер накопителя
    dr.pipe((171, 27), (171, 64.5), lw=LW_S)
    dr.level(171, 69)
    dr.label(176, 72, "А3")

    # А1 — насосная станция: Н1 (рабочий) и Н2 (резервный)
    yP, yC = 106, 168
    for i, cx in enumerate((75, 128)):
        n = i + 1
        dr.pipe((cx, 30), (cx, yP - 7))                       # всасывающая линия
        dr.flow(cx, 66, "u")
        dr.valve(cx, 89)
        dr.label(cx + 4.5, 91, f"ВН{n}")
        dr.pump(cx, yP)
        dr.label(cx + 3.5, yP + 7.5, f"Н{n}")
        dr.pipe((cx, yP + 7), (cx, yC))                       # напорная линия
        dr.check(cx, 125)
        dr.label(cx + 4.5, 127, f"КО{n}")
        dr.valve(cx, 139)
        dr.label(cx + 4.5, 141, f"ВН{n + 2}")
        dr.node(cx, 152)
        dr.pipe((cx, 152), (cx - 10, 152), (cx - 10, 155.3), lw=LW_S)
        dr.gauge(cx - 10, 159.5)
        dr.label(cx - 22, 163.5, f"МН{n}")
    if texts:
        dr.note(64, 103, "рабочий", ha="right")
        dr.note(117, 103, "резерв\n(АВР)", ha="right")
    # коллектор и резервный патрубок под P1
    dr.pipe((75, yC), (185, yC))
    dr.node(128, yC)
    dr.node(160, yC)
    dr.pipe((160, yC), (160, 146))
    dr.valve(160, 156, closed=True)
    dr.label(164.5, 157.5, "ВН5")
    dr.plug(160, 146)
    if texts:
        dr.note(160, 141.5, "место под P1", ha="center", va="top")
    dr.block(52, 81, 178, 177)
    dr.label(52, 178, "А1", ha="left")
    if texts:
        dr.note(63, 178.6, "насосная станция (конфигурация A: 2 × P2)")

    # напорный став
    dr.pipe((185, yC), (185, 268), (210, 268), (210, 247.5))
    dr.flow(185, 196, "u")
    dr.flow(185, 250, "u")
    dr.flow(197, 268, "r")
    dr.flow(210, 254, "d", size=3.4)
    dr.node(185, 182)
    dr.pipe((185, 182), (195, 182), (195, 187))
    dr.damper(195, 192)
    dr.label(205, 197.5, "А4")
    dr.flowmeter(185, 216, vertical=True)
    dr.label(190, 220, "А5")
    if texts:
        dr.ax.text(180.5, 224, "Напорный став, Hст = 60 м", rotation=90,
                   fontsize=dr.fss, ha="right", va="bottom", zorder=6)

    # С1 — отстойник
    out, sl = dr.settler(210, 224, sc=2.2)
    dr.label(244, 240, "С1")
    if texts:
        dr.note(214, 252, f"весь поток ≤ 1500 м³/ч\n500 → {num(D['cS'])} мг/л", va="bottom")
    dr.pipe(sl, (sl[0], 202))
    dr.flow(sl[0], 207, "d")
    dr.note(sl[0] + 2.5, 200, f"осадок ≈ {num(D['slS'])} кг/ч" if texts else "осадок", va="bottom")
    xj = 285
    dr.pipe(out, (xj, out[1]))
    dr.flow(275, out[1], "r")
    dr.node(xj, out[1])

    # А2 — узел доочистки повторной воды (верхняя ветвь)
    yR = 270
    dr.pipe((xj, out[1]), (xj, yR), (384, yR))
    dr.flow(384, yR, "r", size=3.4)
    dr.regulator(304, yR)
    dr.label(306, yR + 6.5, "РП1")
    dr.flowmeter(326, yR)
    dr.label(329, yR + 5.2, "А6")
    dr.filt(348, yR)
    dr.label(352, yR + 6, "Ф1")
    dr.pipe((348, yR - 7), (348, 244))
    dr.flow(348, 250, "d")
    dr.note(350.5, 244, f"осадок ≈ {num(D['slF'])} кг/ч" if texts else "осадок", va="bottom")
    dr.tap(366, yR)
    dr.label(369.5, yR - 13, "А7")
    dr.block(293, 254.5, 375, 284)
    dr.label(293, 284.8, "А2")
    if texts:
        dr.note(301, 285.3, f"доочистка: {num(D['cS'])} → {num(D['cF'])} мг/л")
    dr.note(388, yR, f"К фабрике\n300 м³/ч\n≤ 30 мг/л" if texts else "К фабрике\n300 м³/ч", va="center")

    # оставшийся поток — к условному приёмнику (нижняя ветвь)
    yD = 226
    dr.pipe((xj, out[1]), (xj, yD), (384, yD))
    dr.flow(xj, 240, "d")
    dr.flow(384, yD, "r", size=3.4)
    dr.flowmeter(316, yD)
    dr.label(319, yD + 5.2, "А8")
    dr.tap(340, yD)
    dr.label(343.5, yD - 13, "А9")
    dr.note(388, yD, "К приёмнику\n700–1000 м³/ч\n≤ 100 мг/л" if texts else "К приёмнику", va="center")


# ----------------------------------------------------------------- рамка и основная надпись
def frame(ax, W, H):
    ax.add_patch(Rectangle((20, 5), W - 25, H - 10, fill=False, ec=K, lw=2.0, zorder=7))


def stamp_form1(ax, W, name, design, sheet="1", sheets="1"):
    """Основная надпись по форме 1 ГОСТ 2.104 (185 × 55 мм) в правом нижнем углу рамки."""
    x0, y0 = W - 5 - 185, 5
    t, b = LW_T, 2.0

    def ln(x1, y1, x2, y2, w=t):
        ax.plot([x0 + x1, x0 + x2], [y0 + y1, y0 + y2], color=K, lw=w, zorder=7)

    def tx(x, y, s, size=7, ha="left", va="center", **kw):
        ax.text(x0 + x, y0 + y, s, fontsize=size, ha=ha, va=va, zorder=8, **kw)

    ax.add_patch(Rectangle((x0, y0), 185, 55, fill=False, ec=K, lw=b, zorder=7))
    # левая часть 65 мм: 11 строк по 5 мм
    for y in range(5, 55, 5):
        ln(0, y, 65, y, b if y in (30, 35) else t)
    for x in (7, 17, 40, 55):
        ln(x, 30, x, 55, b)
    for x in (17, 40, 55):
        ln(x, 0, x, 30, b)
    ln(65, 0, 65, 55, b)
    for i, s in enumerate(("Изм.", "Лист", "№ докум.", "Подп.", "Дата")):
        tx((0, 7, 17, 40, 55, 65)[i] + ((7, 10, 23, 15, 10)[i]) / 2, 32.5, s, size=6, ha="center")
    for i, s in enumerate(("Разраб.", "Пров.", "Т.контр.", "", "Н.контр.", "Утв.")):
        tx(1, 27.5 - 5 * i, s, size=6.5)
    tx(18, 27.5, AUTHOR, size=6.5)
    tx(56, 27.5, "11.26", size=6)
    # обозначение документа
    ln(65, 40, 185, 40, b)
    tx(125, 47.5, design, size=14, ha="center")
    # наименование и правая колонка
    ln(135, 0, 135, 40, b)
    ln(135, 35, 185, 35, b)
    ln(135, 20, 185, 20, b)
    ln(135, 15, 185, 15, b)
    ln(65, 15, 135, 15, b)
    for x in (140, 145):
        ln(x, 20, x, 35, t)
    for x in (150, 167):
        ln(x, 20, x, 40, b)
    ln(155, 15, 155, 20, b)
    tx(142.5, 37.5, "Лит.", size=6.5, ha="center")
    tx(158.5, 37.5, "Масса", size=6.5, ha="center")
    tx(176, 37.5, "Масштаб", size=6.5, ha="center")
    tx(176, 27.5, "—", size=9, ha="center")
    tx(136, 17.5, f"Лист {sheet}", size=6.5)
    tx(156, 17.5, f"Листов {sheets}", size=6.5)
    tx(100, 27.5, name, size=9.5, ha="center", linespacing=1.25)
    tx(160, 7.5, ORGTXT, size=8, ha="center")
    tx(100, 7.5, "Инженерный кубок 2026–2027, кейс 16", size=7, ha="center", color="#333333")


# ----------------------------------------------------------------- лист 1: A3 со схемой
NOTES = [
    "Схема выполнена по ГОСТ 2.701 и ГОСТ 2.704; условные графические обозначения — по ГОСТ 2.780,\n"
    "2.781, 2.782, 2.784, 2.785; буквенные коды элементов — по ГОСТ 2.704.",
    "Схема изображена в исходном положении: насосы Н1, Н2 остановлены, задвижки ВН1–ВН4 открыты,\n"
    "ВН5 закрыта (зачернена), уровень в Б1 — на отметке «стоп».",
    "Н1 — рабочий, Н2 — резервный с автоматическим вводом резерва (АВР); роли меняют по наработке.\n"
    f"Одновременная работа Н1 и Н2 запрещена блокировкой: 2 × P2 = {num(D['q2'])} м³/ч > 1 500 м³/ч (предел С1).",
    f"Рабочая точка одного насоса P2: {num(D['qA'])} м³/ч при {num(D['hA'], 1)} м, {num(D['NA'])} кВт;"
    " сеть H = 60 + 0,000025·Q², м.",
    f"Отметки уровня А3 от дна Б1: «стоп» {num(D['h_stop'], 2)} м; «пуск» {num(D['h_on'], 2)} м; "
    f"«авария» {num(D['h_alarm'], 2)} м;\nверх полезного объёма {num(D['h_top'], 2)} м; гребень "
    f"{num(D['depth'], 2)} м.",
    "Патрубок с ВН5 и заглушкой — место под насос P1 (переход к конфигурации D: 2 × P2 + P1).",
    "Сигналы А3, А5–А9 передаются в систему управления насосами; электрические цепи не показаны.",
]


def sheet_a3(pdf):
    W, H = 420, 297
    fig = plt.figure(figsize=(W * MM, H * MM))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.set_aspect("equal")
    ax.axis("off")
    frame(ax, W, H)
    dr = Draw(ax, fs=10.5, fs_small=8)
    draw_scheme(dr)
    y = 150
    ax.text(W - 5 - 185, y, "Примечания", fontsize=10, fontweight="bold", va="top")
    y -= 7
    for i, s in enumerate(NOTES, 1):
        ax.text(W - 5 - 185, y, f"{i}.", fontsize=7.6, va="top")
        ax.text(W - 5 - 180, y, s, fontsize=7.6, va="top", linespacing=1.3)
        y -= 4.2 + 3.55 * (s.count("\n") + 1)
    assert y > 61, "примечания залезают на основную надпись"
    stamp_form1(ax, W, "Водоотлив и очистка\nкарьерной воды\nСхема гидравлическая\nпринципиальная",
                f"{DESIGN} Г3")
    pdf.savefig(fig)
    return fig


# ----------------------------------------------------------------- лист 2: A4 перечень элементов
ITEMS = [
    ("А1", "Насосная станция водоотлива", "1", "конфигурация A: 2 × P2"),
    ("А2", "Узел доочистки повторной воды", "1", "300 м³/ч"),
    ("А3", "Уровнемер накопителя", "1", "отметки «стоп»,\n«пуск», «авария»"),
    ("А4", "Гаситель гидравлического удара", "1", "на напорном ставе"),
    ("А5", "Расходомер напорного става", "1", "до 1500 м³/ч"),
    ("А6", "Расходомер ветви повторной воды", "1", ""),
    ("А7", "Пробоотборник с мутномером", "1", "контроль ≤ 30 мг/л"),
    ("А8", "Расходомер потока к приёмнику", "1", ""),
    ("А9", "Пробоотборник с мутномером", "1", "контроль ≤ 100 мг/л"),
    ("Б1", "Накопитель-водосборник открытый", "1",
     f"полезный {num(D['Vuse'])} м³,\nполный {num(D['Vfull'])} м³, h = {num(D['depth'], 1)} м"),
    ("ВН1, ВН2", "Задвижка на всасывающей линии", "2", ""),
    ("ВН3, ВН4", "Задвижка на напорной линии", "2", ""),
    ("ВН5", "Задвижка резервного патрубка", "1", "закрыта; фланцевая\nзаглушка, место под P1"),
    ("КО1, КО2", "Клапан обратный", "2", "защита от обратного\nтока и гидроудара"),
    ("МН1, МН2", "Манометр", "2", ""),
    ("Н1", "Насосный агрегат P2 (рабочий)", "1",
     f"{num(D['qA'])} м³/ч, {num(D['hA'], 1)} м,\n{num(D['NA'])} кВт"),
    ("Н2", "Насосный агрегат P2 (резервный)", "1", "АВР"),
    ("РП1", "Регулятор расхода", "1", "300 м³/ч к фабрике"),
    ("С1", "Отстойник", "1", f"≤ 1500 м³/ч,\n500 → {num(D['cS'])} мг/л"),
    ("Ф1", "Фильтр", "1", f"≤ 400 м³/ч,\n{num(D['cS'])} → {num(D['cF'])} мг/л"),
]


def sheet_a4(pdf):
    W, H = 210, 297
    fig = plt.figure(figsize=(W * MM, H * MM))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.set_aspect("equal")
    ax.axis("off")
    frame(ax, W, H)
    x0, top = 20, 292
    cols = [20, 110, 10, 45]
    xs = [x0]
    for c in cols:
        xs.append(xs[-1] + c)
    hh, rh = 15, 8
    yb = top - hh - rh * len(ITEMS)
    for x in xs:
        ax.plot([x, x], [yb, top], color=K, lw=2.0 if x in (xs[0], xs[-1]) or True else LW_T, zorder=7)
    ax.plot([x0, xs[-1]], [top - hh, top - hh], color=K, lw=2.0)
    for i, s in enumerate(("Поз.\nобозна-\nчение", "Наименование", "Кол.", "Примечание")):
        ax.text((xs[i] + xs[i + 1]) / 2, top - hh / 2, s, fontsize=8.5, ha="center", va="center",
                linespacing=1.1)
    y = top - hh
    for code, name, qty, rem in ITEMS:
        ax.plot([x0, xs[-1]], [y - rh, y - rh], color=K, lw=LW_T)
        ax.text(xs[0] + 1.5, y - rh / 2, code, fontsize=8.5, va="center")
        ax.text(xs[1] + 1.5, y - rh / 2, name, fontsize=8.5, va="center")
        ax.text((xs[2] + xs[3]) / 2, y - rh / 2, qty, fontsize=8.5, ha="center", va="center")
        ax.text(xs[3] + 1.2, y - rh / 2, rem, fontsize=6.6 if "\n" in rem else 7.2, va="center",
                linespacing=1.05)
        y -= rh
    ax.text(x0 + 2, yb - 6, "Позиционные обозначения — по схеме «Водоотлив и очистка карьерной воды. Схема "
            "гидравлическая\nпринципиальная» (лист 1). Насосы P2 по каталогу кейса: H = 165 − 0,00006·Q², "
            "η = 0,80.", fontsize=7.5, va="top", linespacing=1.3)
    stamp_form1(ax, W, "Водоотлив и очистка\nкарьерной воды\nПеречень элементов", f"{DESIGN} ПГ3")
    pdf.savefig(fig)
    return fig


# ----------------------------------------------------------------- картинка для документов
def sheet_png(fig, path):
    """Лист A3 целиком, повёрнутый на 90° — чтобы встать на страницу A4 книжной ориентации."""
    from PIL import Image

    tmp = path + ".tmp.png"
    fig.savefig(tmp, dpi=200, facecolor="white")
    Image.open(tmp).rotate(90, expand=True).save(path, optimize=True)
    os.remove(tmp)


# ----------------------------------------------------------------- слайд 5а
def slide(png, pptx):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    DARK, TEAL = RGBColor(0x1F, 0x2A, 0x33), RGBColor(0x0E, 0x7C, 0x86)
    # Картинка — в миллиметрах слайда (338,7 × 165 мм под заголовком), своя упрощённая раскладка:
    # те же УГО и позиционные обозначения, без манометров, гасителя и пробоотборников.
    SW, SH = 338.7, 165
    fig = plt.figure(figsize=(SW * MM, SH * MM))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, SW)
    ax.set_ylim(0, SH)
    ax.set_aspect("equal")
    ax.axis("off")
    dr = Draw(ax, fs=18, fs_small=16)
    tc = "#0E7C86"
    # Б1 и приток
    dr.tank(14, 100, 6, 30)
    dr.label(102, 26, "Б1")
    dr.source(8, 46)
    dr.pipe((12.5, 46), (24, 46), (24, 22))
    dr.flow(24, 34, "d", size=3.4)
    dr.note(8, 49, "Приток")
    # А3 — уровнемер
    dr.pipe((62, 10), (62, 30.2), lw=LW_S)
    dr.level(62, 34, r=3.8)
    dr.label(66.5, 33, "А3", va="center")
    # А1 — насосная станция
    yP, yC = 60, 104
    for i, cx in enumerate((42, 82)):
        n = i + 1
        dr.pipe((cx, 10), (cx, yP - 6))
        dr.valve(cx, 46, s=2.8)
        dr.label(cx + 3.5, 46, f"ВН{n}", va="center")
        dr.pump(cx, yP, r=6)
        dr.label(cx - 7, yP + 6, f"Н{n}", ha="right", va="center")
        dr.pipe((cx, yP + 6), (cx, yC))
        dr.check(cx, 76, r=2.2)
        dr.label(cx + 3.5, 77, f"КО{n}", va="center")
        dr.valve(cx, 91, s=2.8)
        dr.label(cx + 3.5, 91, f"ВН{n + 2}", va="center")
    dr.pipe((42, yC), (138, yC))
    dr.node(82, yC)
    dr.node(116, yC)
    dr.pipe((116, yC), (116, 86))
    dr.valve(116, 95, s=2.8, closed=True)
    dr.label(119.5, 95, "ВН5", va="center")
    dr.plug(116, 86)
    dr.block(30, 39.5, 134, 112)
    dr.label(30, 113, "А1")
    # напорный став, А5
    dr.pipe((138, yC), (138, 148), (160, 148), (160, 127.5))
    dr.flow(138, 138, "u", size=3.4)
    dr.flow(160, 134, "d", size=3.4)
    dr.flowmeter(138, 121, vertical=True, r=3.8)
    dr.label(142.5, 121, "А5", va="center")
    # С1
    out, sl = dr.settler(158, 108, sc=1.75)
    dr.label(178, 118, "С1")
    dr.pipe(sl, (sl[0], 96))
    dr.flow(sl[0], 101, "d", size=3.4)
    dr.note(sl[0] + 2.5, 95, "осадок")
    xj = 222
    dr.pipe(out, (xj, out[1]))
    dr.node(xj, out[1])
    # А2 — доочистка, ветвь к фабрике
    yR = 142
    dr.pipe((xj, out[1]), (xj, yR), (300, yR))
    dr.flow(300, yR, "r", size=3.6)
    dr.regulator(239, yR, w=14, h=9)
    dr.label(239, yR + 6, "РП1", ha="center")
    dr.flowmeter(259, yR, r=3.8)
    dr.label(259, yR + 5.5, "А6", ha="center")
    dr.filt(278, yR, s=6)
    dr.label(278, yR + 7, "Ф1", ha="center")
    dr.pipe((278, yR - 6), (278, 128))
    dr.flow(278, 131, "d", size=3.4)
    dr.block(228, 132.5, 291, 156)
    dr.label(228, 156.5, "А2")
    dr.note(304, yR, "К фабрике\n300 м³/ч", va="center")
    # к приёмнику
    yD = 116
    dr.pipe((xj, out[1]), (xj, yD), (300, yD))
    dr.flow(300, yD, "r", size=3.6)
    dr.flowmeter(262, yD, r=3.8)
    dr.label(262, yD + 5.5, "А8", ha="center")
    dr.note(304, yD, "К приёмнику\n700–1000\nм³/ч", va="center")
    fig.savefig(png, dpi=250, facecolor="white")
    plt.close(fig)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6])

    def tb(x, y, w, h, paras, size=18, color=DARK, bold=False):
        box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        for i, p in enumerate(paras):
            par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            par.space_after = Pt(4)
            runs = p if isinstance(p, list) else [(p, {})]
            for t, o in runs:
                r = par.add_run()
                r.text = t
                r.font.name = "Arial"
                r.font.size = Pt(o.get("size", size))
                r.font.bold = o.get("bold", bold)
                r.font.color.rgb = o.get("color", color)
        return box

    tb(0.45, 0.2, 12.4, 0.8, ["Принципиальная гидравлическая схема водоотлива"], size=28, bold=True)
    s.shapes.add_picture(png, 0, Inches(7.5 - SH * MM), width=prs.slide_width)
    T = {"bold": True, "color": TEAL}
    NB = " "
    tb(5.55, 3.95, 7.6, 3.1, [
        [("Б1", T), (f" накопитель {num(D['Vuse'])} м³;  ", {}), ("А3", T), (" уровнемер", {})],
        [("Н1,", T), (NB + "Н2", T), (" насосы P2: рабочий и резерв (АВР)", {})],
        [("КО", T), (" обратные клапаны;  ", {}), ("ВН", T), (" задвижки", {})],
        [("ВН5", T), (" закрыта: место под насос P1", {})],
        [("С1", T), (f" отстойник 500 → {num(D['cS'])} мг/л;  ", {}), ("Ф1", T),
         (f" фильтр {num(D['cS'])} → {num(D['cF'])} мг/л", {})],
        [("РП1", T), (" регулятор 300 м³/ч;  ", {}), ("А5, А6, А8", T), (" расходомеры", {})],
        [("Поток снизу вверх, слева направо (ГОСТ 2.704)", {"color": TEAL})],
    ], size=18)
    tb(12.3, 6.95, 0.8, 0.4, ["5а"], size=14)
    prs.save(pptx)


if __name__ == "__main__":
    out_pdf = os.path.join(HERE, "Схема_гидравлическая_принципиальная.pdf")
    with PdfPages(out_pdf) as pdf:
        f1 = sheet_a3(pdf)
        f2 = sheet_a4(pdf)
        d = pdf.infodict()
        d["Title"] = "Водоотлив и очистка карьерной воды. Схема гидравлическая принципиальная"
    if "--preview" in sys.argv:
        pv = sys.argv[sys.argv.index("--preview") + 1]
        f1.savefig(os.path.join(pv, "a3.png"), dpi=110)
        f2.savefig(os.path.join(pv, "a4.png"), dpi=110)
    sheet_png(f1, os.path.join(HERE, "рисунки_дополнения", "схема_лист_A3.png"))
    plt.close("all")
    slide(os.path.join(HERE, "рисунки_дополнения", "схема_для_слайда.png"),
          os.path.join(HERE, "Слайд_5а_гидравлическая_схема.pptx"))
    print("готово:", out_pdf)
