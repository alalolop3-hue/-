"""Сборка защитной презентации в стиле v3 (плотные слайды, CAP-метки разделов,
тезисный подзаголовок, табличные блоки, мелкий курсивный итог снизу) с полным
содержанием v4 (финальный слайд, NPV/Payback, PESTEL, timeline пилота, SWOT).
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# === Цвета: 2 рабочих + нейтральные ===
RED = RGBColor(0xED, 0x00, 0x28)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x70, 0x70, 0x70)
GRAY_LIGHT = RGBColor(0x9A, 0x9A, 0x9A)
LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
VLIGHT = RGBColor(0xFA, 0xFA, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN_TXT = RGBColor(0x1E, 0x7A, 0x3C)
RED_TXT = RGBColor(0xC0, 0x1A, 0x1A)

FONT = 'Calibri'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

# ==================== ХЕЛПЕРЫ ====================

def add_slide():
    return prs.slides.add_slide(BLANK)

def T(s, text, left, top, width, height, *, size=14, bold=False, italic=False,
      color=DARK, align=PP_ALIGN.LEFT, ls=1.2, anchor=None):
    tb = s.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(2)
    if anchor:
        tf.vertical_anchor = anchor
    lines = text.split('\n') if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = ls
        r = p.add_run()
        r.text = line
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return tb

def CARD(s, left, top, width, height, *, fill=LIGHT, line=None, line_w=0.5):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top),
                             Inches(width), Inches(height))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_w)
    sh.shadow.inherit = False
    return sh

def slide_chrome(s, section_label, headline, page_num, total):
    """v3-стилизация: CAP-метка раздела + жирный тезисный подзаголовок + полоса под ним."""
    # CAP section label
    T(s, section_label.upper(), 0.5, 0.32, 5, 0.3, size=11, bold=True, color=RED,
      align=PP_ALIGN.LEFT, ls=1.0)
    # Page indicator справа сверху
    T(s, f'{page_num} / {total}', 12.4, 0.32, 0.7, 0.3, size=10, color=GRAY_LIGHT,
      align=PP_ALIGN.RIGHT, ls=1.0)
    # Headline (тезисный подзаголовок)
    T(s, headline, 0.5, 0.65, 12.3, 1.0, size=22, bold=True, color=DARK,
      align=PP_ALIGN.LEFT, ls=1.15)
    # Тонкая линия под заголовком
    ln = s.shapes.add_connector(1, Inches(0.5), Inches(1.6), Inches(12.83), Inches(1.6))
    ln.line.color.rgb = RED
    ln.line.width = Pt(1.5)

def punchline(s, text):
    """Мелкий курсивный итог снизу слайда (как в v3)."""
    T(s, text, 0.5, 6.95, 12.3, 0.35, size=10.5, italic=True, color=GRAY,
      align=PP_ALIGN.CENTER, ls=1.2)

def set_notes(s, text):
    s.notes_slide.notes_text_frame.text = text

def bignum(s, num, left, top, width, height, *, size=44, color=RED, align=PP_ALIGN.CENTER):
    T(s, num, left, top, width, height, size=size, bold=True, color=color,
      align=align, ls=1.0)

TOTAL = 16

# ==================== СЛАЙД 1: ТИТУЛЬНЫЙ ====================
s = add_slide()

# Верхний колонтитул
T(s, 'ВКР  |  Конкурентная стратегия МТС Юрент', 0.5, 0.35, 9, 0.3,
  size=11, bold=True, color=RED, ls=1.0)
T(s, 'Москва  ·  2026', 12.0, 0.35, 1, 0.3, size=10, color=GRAY, align=PP_ALIGN.RIGHT)

# Главный заголовок-вопрос (v3 wording)
T(s, 'Как МТС Юренту перейти от роста через парк\nк управлению отдачей регионального портфеля?',
  0.5, 1.1, 12.3, 1.8, size=32, bold=True, color=DARK, ls=1.15)

# Объект/предмет/проблема
T(s, 'Объект', 0.5, 3.15, 1.8, 0.3, size=11, color=GRAY)
T(s, 'МТС Юрент как федеральный оператор кикшеринга на портфеле региональных рынков',
  2.4, 3.15, 10.5, 0.3, size=12, color=DARK)

T(s, 'Предмет', 0.5, 3.55, 1.8, 0.3, size=11, color=GRAY)
T(s, 'Конкурентная стратегия на портфеле региональных рынков',
  2.4, 3.55, 10.5, 0.3, size=12, color=DARK)

T(s, 'Проблема', 0.5, 3.95, 1.8, 0.6, size=11, color=GRAY)
T(s, 'Экстенсивная модель роста через расширение парка и географии теряет '
     'эффективность на зрелом и неоднородном рынке. Лимиты парка, регуляторика и '
     'падающая предельная отдача требуют выбора разных режимов действий по городам.',
  2.4, 3.95, 10.5, 0.9, size=12, color=DARK, ls=1.3)

# 3 контекстных карточки
ctx = [
    ('Рынок зрелее',        'спрос в 2025 г.  −6 %',           '5 лет роста — закончились'),
    ('Парк ограничен',       'fleet cap и регуляторика',         'каждый новый самокат дороже'),
    ('Города разные',        'утилизация различается в 4×',     'единый режим не работает'),
]
y_c = 5.05; cw = 4.0; gap = 0.15
for i, (h, sub, ln) in enumerate(ctx):
    x = 0.5 + i*(cw + gap)
    CARD(s, x, y_c, cw, 1.2, fill=LIGHT)
    T(s, h, x+0.2, y_c+0.12, cw-0.4, 0.35, size=13, bold=True, color=DARK)
    T(s, sub, x+0.2, y_c+0.5, cw-0.4, 0.35, size=12, bold=True, color=RED)
    T(s, ln, x+0.2, y_c+0.85, cw-0.4, 0.3, size=10, color=GRAY)

# 2 больших числа справа от потока
T(s, '96,8 %', 0.5, 6.38, 1.8, 0.45, size=22, bold=True, color=RED, align=PP_ALIGN.RIGHT)
T(s, 'доля топ-3', 0.5, 6.83, 1.8, 0.2, size=9.5, color=GRAY, align=PP_ALIGN.RIGHT)

T(s, '187', 11.0, 6.38, 1.8, 0.45, size=22, bold=True, color=RED, align=PP_ALIGN.LEFT)
T(s, 'локаций МТС Юрент', 11.0, 6.83, 2.3, 0.2, size=9.5, color=GRAY, align=PP_ALIGN.LEFT)

# Логическая лента посередине
flow = ['Диагноз города', 'Выбор режима', 'Рычаги в сезоне', 'KPI', 'Эффект']
fy = 6.45; fw = 1.45; fgap = 0.1
total_fw = len(flow)*fw + (len(flow)-1)*fgap
fx0 = (13.333 - total_fw) / 2
for i, step in enumerate(flow):
    x = fx0 + i*(fw + fgap)
    CARD(s, x, fy, fw, 0.4, fill=DARK)
    T(s, step, x, fy+0.08, fw, 0.25, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Author + supervisor (внизу мелким)
T(s, 'Камашев Пётр  ·  Управление бизнесом  ·  научный руководитель  проф. И. А. Егоров',
  0.5, 7.18, 12.3, 0.2, size=10, color=GRAY, align=PP_ALIGN.CENTER)

set_notes(s, '«Добрый день, уважаемые члены государственной экзаменационной комиссии. '
              'Меня зовут Пётр Камашев, представляю выпускную квалификационную работу на тему '
              "«Конкурентная стратегия МТС Юрент на региональных рынках кикшеринга в России». "
              'Научный руководитель — профессор Иван Александрович Егоров.\n\n'
              'Центральный вопрос работы — как МТС Юренту перейти от роста через расширение парка '
              'к управлению отдачей уже сформированного регионального портфеля. '
              'Работа выстроена как цепочка из пяти шагов: диагноз города, выбор режима, рычаги в сезоне, KPI и эффект».')

# ==================== СЛАЙД 2: ЛОГИКА ====================
s = add_slide()
slide_chrome(s, 'ЛОГИКА',
             'Цель ВКР — разработать систему стратегического выбора для портфеля региональных рынков',
             2, TOTAL)

# Формулировка цели — отдельный блок (как v3 даёт мелким курсивом + крупным основным)
CARD(s, 0.5, 1.8, 12.3, 1.0, fill=VLIGHT)
T(s, 'Цель', 0.7, 1.9, 2, 0.3, size=10, color=GRAY)
T(s, 'Разработать аналитическую систему стратегического выбора оператора кикшеринга в портфеле '
     'региональных рынков и применить её к МТС Юрент: сопоставить три альтернативных сценария, '
     'обосновать выбор сценария, предложить проектный инструментарий и количественно оценить эффект '
     'на демонстрационном портфеле из пяти городов.',
  0.7, 2.15, 11.9, 0.65, size=11.5, color=DARK, ls=1.3)

# 5-шаговая логика работы (v3 style)
steps = [
    ('1', 'Проблема и рынок',  'переход от экстенсивного роста к зрелой стадии'),
    ('2', 'Диагностика Юрента', 'портфель городов, утилизация, выручка/самокат'),
    ('3', 'Выбор сценария',     'сравнение S1/S2/S3 по 5 критериям'),
    ('4', 'Режимы и УПС',       'R1/R2/R3 и инструмент управления покрытием'),
    ('5', 'Эффект и риски',     'финмодель, KPI, условия управляемости'),
]
y_st = 3.05; sw = 2.45; sh_h = 2.7; sgap = 0.1
for i, (n, h, body) in enumerate(steps):
    x = 0.5 + i*(sw + sgap)
    CARD(s, x, y_st, sw, sh_h, fill=LIGHT)
    # Number circle
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+0.2), Inches(y_st+0.25),
                                 Inches(0.55), Inches(0.55))
    circle.fill.solid(); circle.fill.fore_color.rgb = RED
    circle.line.fill.background()
    tb = circle.text_frame
    tb.margin_top = tb.margin_bottom = tb.margin_left = tb.margin_right = 0
    tb.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp = tb.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
    rr = pp.add_run(); rr.text = n
    rr.font.name = FONT; rr.font.size = Pt(17); rr.font.bold = True
    rr.font.color.rgb = WHITE

    T(s, h, x+0.2, y_st+0.95, sw-0.4, 0.55, size=12.5, bold=True, color=DARK, ls=1.15)
    T(s, body, x+0.2, y_st+1.55, sw-0.4, 1.1, size=10.5, color=GRAY, ls=1.3)

# 7 задач — компактным буллет-блоком
CARD(s, 0.5, 5.95, 12.3, 0.85, fill=VLIGHT)
T(s, 'Выход работы', 0.7, 6.02, 3, 0.25, size=10, color=GRAY)
T(s, 'Сценарий S3  ·  система режимов R1–R3  ·  проектный инструмент УПС  ·  '
     'финансовая модель на 5 городах  ·  KPI пилота',
  0.7, 6.3, 12, 0.45, size=12, bold=True, color=DARK, ls=1.3)

punchline(s, 'Цель работы — проектная: дать МТС Юрент аналитический язык выбора действий по региональным рынкам')
set_notes(s, 'Цель работы — разработать аналитическую систему стратегического выбора и применить к МТС Юрент. '
              'Логика — линейная: диагностика → 5 критериев → 3 сценария → выбор S3 → режимы + УПС → эффект. '
              'Выход работы — сценарий S3, система режимов R1–R3, инструмент УПС, финансовая модель и KPI.')

# ==================== СЛАЙД 3: МЕТОДЫ ====================
s = add_slide()
slide_chrome(s, 'МЕТОДЫ',
             'Выводы основаны на публичной аналитике, внутренних данных оператора и сценарной финмодели',
             3, TOTAL)

# Таблица методов (4×3)
rows = [
    ('Блок',              'Что использовано',                                              'Зачем это нужно'),
    ('Диагностика рынка', 'TrueSharing 2023–2025, отчётность Whoosh и МТС, деловая пресса', 'доказать переход к зрелости'),
    ('Внутренняя база',   'поквартальные данные МТС Юрент по 86 субъектам РФ за 2024 и 9м25', 'сопоставить города и построить финмодель'),
    ('Экспертиза',        'эксперт А. — M&A МТС;  эксперт Б. — калибровка логики',         'проверить управленческую применимость'),
    ('Методы',            'PESTEL, конкурентный анализ, рамка I-II-III, сценарное сравнение, финмодель', 'перейти от проблемы к выбору S3'),
]
y_t = 1.85
col_ws = [2.2, 6.5, 3.6]
row_h = 0.55
for r_i, row in enumerate(rows):
    y = y_t + r_i*row_h
    is_h = (r_i == 0)
    fill = DARK if is_h else (WHITE if r_i % 2 == 1 else VLIGHT)
    CARD(s, 0.5, y, sum(col_ws), row_h, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else DARK
        bold = is_h or c_i == 0
        size = 11 if is_h else 10.5
        T(s, cell, x+0.15, y, col_ws[c_i]-0.3, row_h,
          size=size, bold=bold, color=color, anchor=MSO_ANCHOR.MIDDLE, ls=1.15)
        x += col_ws[c_i]

# 4 ключевых числа в строку
nums = [
    ('86',          'субъектов РФ в агрегированной базе'),
    ('2024–9м25',   'период внутренней базы'),
    ('5',           'городов демонстрационного портфеля'),
    ('ИИ',          'обработка источников + редактура; решения и интерпретации авторские'),
]
y_n = 4.95; nw = 3.0; ngap = 0.1
for i, (n, body) in enumerate(nums):
    x = 0.5 + i*(nw + ngap)
    CARD(s, x, y_n, nw, 1.7, fill=LIGHT)
    bignum(s, n, x, y_n+0.2, nw, 0.85, size=30, color=RED, align=PP_ALIGN.CENTER)
    T(s, body, x+0.15, y_n+1.05, nw-0.3, 0.6, size=10, color=DARK,
      align=PP_ALIGN.CENTER, ls=1.25)

punchline(s, 'Эмпирическая база достаточна для предварительного выбора стратегии и дизайна пилота; апробация — этап пилота УПС')
set_notes(s, 'Доказательная база сочетает публичную аналитику, внутренние поквартальные данные МТС Юрент '
              'по 86 субъектам РФ за 2024 и 9 мес. 2025, и двух экспертов. '
              'Демонстрационный портфель из 5 городов покрывает 94 % выручки оператора. '
              'ИИ применялся для обработки источников и редактуры; стратегические решения, выбор сценария и '
              'параметры финмодели — авторские; декларация — в Приложении А ВКР.')

# ==================== СЛАЙД 4: РЫНОК ====================
s = add_slide()
slide_chrome(s, 'РЫНОК',
             'Российский рынок перешёл от роста парка к борьбе за эффективность — на зрелой стадии единая логика теряет силу',
             4, TOTAL)

# Левая колонка: что изменилось (буллеты)
T(s, 'Что изменилось', 0.5, 1.85, 6, 0.4, size=13, bold=True, color=DARK)
bullets = [
    ('2024',          '31,1 млрд руб. и 281,6 млн поездок · 96,8 % у топ-3 операторов'),
    ('2025',          'впервые отрицательная динамика спроса: −6 % поездок'),
    ('Whoosh',        '−7 % поездок, −13 % выручки, EBITDA-маржа 42,3 → 28,6 %'),
    ('Капитал',       'ключевая ставка ЦБ 21 % + 2–3 сезона срок службы парка — расширение дорожает'),
    ('Регуляторика',  'mos.ru в Москве, соглашения СПб, угрозы расторжения в Екб'),
    ('Сезон 2025',    'начался не в марте, а в июне — потери первого квартала'),
]
yb = 2.3
for hdr, body in bullets:
    T(s, hdr, 0.5, yb, 2.0, 0.3, size=11, bold=True, color=RED)
    T(s, body, 2.6, yb, 5.4, 0.55, size=11, color=DARK, ls=1.3)
    yb += 0.55

# Правая колонка: PESTEL-карточки 4 шт. (компактно 2×2)
T(s, 'Сводный PESTEL', 8.4, 1.85, 4.5, 0.4, size=13, bold=True, color=DARK)
pestel = [
    ('Рынок',         '−6 %',         'спрос РФ'),
    ('Капитал',       '21 %',         'ставка ЦБ'),
    ('Регуляторика',  '3 режима',     'M / СПб / Екб'),
    ('Сезон',         'март→июнь',    'сжатие на квартал'),
]
y_p = 2.3; pw = 2.25; ph = 1.1; pgap = 0.1
positions = [(8.4, y_p), (10.75, y_p), (8.4, y_p + ph + pgap), (10.75, y_p + ph + pgap)]
for (title, big, sub), (x, y) in zip(pestel, positions):
    CARD(s, x, y, pw, ph, fill=LIGHT)
    T(s, title, x+0.15, y+0.08, pw-0.3, 0.25, size=10, color=GRAY)
    bignum(s, big, x+0.15, y+0.32, pw-0.3, 0.5, size=20, color=RED, align=PP_ALIGN.LEFT)
    T(s, sub, x+0.15, y+0.83, pw-0.3, 0.22, size=9, color=GRAY)

# Бар стоимости капитала / маржи (мини-визуал в правой)
y_v = 4.65
T(s, 'Whoosh: маржа кикшеринга, %', 8.4, y_v, 5, 0.3, size=10, color=GRAY)
# Два бара
b_top = y_v + 0.35; b_h = 0.4
CARD(s, 8.4, b_top, 2.0, b_h, fill=GRAY_LIGHT)
T(s, '42,3 %', 8.4, b_top, 2.0, b_h, size=14, bold=True, color=WHITE,
  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
T(s, '2024', 8.4, b_top+b_h+0.05, 2, 0.25, size=9, color=GRAY)

CARD(s, 8.4, b_top + 0.95, 1.35, b_h, fill=RED)
T(s, '28,6 %', 8.4, b_top + 0.95, 1.35, b_h, size=14, bold=True, color=WHITE,
  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
T(s, '9 мес. 2025', 8.4, b_top + 0.95 + b_h + 0.05, 2, 0.25, size=9, color=GRAY)
T(s, '↓ −13,7 п.п.', 9.85, b_top + 1.0, 3, 0.35, size=13, bold=True, color=RED_TXT)

punchline(s, 'На зрелой стадии конкурентный результат зависит от отдачи уже размещённого парка, а не от его прироста')
set_notes(s, 'Внешняя среда поменялась: рынок впервые сжался, ставка ЦБ 21 % делает дороже расширение парка, '
              'регуляторика расходится по городам, сезон 2025 сжался с марта до июня. '
              'Маржа кикшеринга у Whoosh упала с 42,3 % до 28,6 % — это индикативный сигнал зрелости. '
              'Совокупный эффект — прежняя логика расширения парка и географии перестаёт быть достаточным основанием конкурентной борьбы.')

# ==================== СЛАЙД 5: ПОРТФЕЛЬ ====================
s = add_slide()
slide_chrome(s, 'ПОРТФЕЛЬ',
             'Даже внутри портфеля Юрента 5 городов расходятся в 4× по утилизации — единый режим невозможен',
             5, TOTAL)

# Таблица 5 городов
rows = [
    ('Город',           'Утилизация\nпоезд./самокат/год', 'YoY поездок\n9 мес. 2025', 'Регуляторика',     'Гипотеза режима'),
    ('Москва',          '1 073',                          '−6,6 %',                   'жёсткая',           'R3 (огранич.)'),
    ('Санкт-Петербург', '609',                            '+42,3 %',                  'умеренно жёсткая',  'R2 → R3'),
    ('Краснодар',       '564',                            '−10,2 %',                  'умеренная',         'R2 / R3'),
    ('Екатеринбург',    '319',                            '+4,1 %',                   'умеренная',         'R2 + пилот УПС'),
    ('Новосибирск',     '277',                            '+128,9 %',                 'умеренная',         'R2/R3 (эффект.)'),
]
y_t = 1.85
col_ws = [2.3, 2.3, 2.2, 2.6, 2.9]
total_w = sum(col_ws)
row_h = 0.55
for r_i, row in enumerate(rows):
    y = y_t + r_i*row_h
    is_h = (r_i == 0)
    fill = DARK if is_h else (WHITE if r_i % 2 == 1 else VLIGHT)
    CARD(s, 0.5, y, total_w, row_h, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        if is_h:
            color = WHITE; bold = True; size = 10.5
        else:
            color = DARK; bold = (c_i == 0 or c_i == 4); size = 12 if c_i in (0, 4) else 11.5
            if c_i == 4:
                color = RED
            elif c_i == 2:
                if cell.startswith('+'):
                    color = GREEN_TXT; bold = True
                elif cell.startswith('−'):
                    color = RED_TXT; bold = True
        align = PP_ALIGN.LEFT if c_i in (0, 3, 4) else PP_ALIGN.CENTER
        anchor = MSO_ANCHOR.MIDDLE
        T(s, cell, x+0.15, y, col_ws[c_i]-0.3, row_h,
          size=size, bold=bold, color=color, align=align, anchor=anchor, ls=1.1)
        x += col_ws[c_i]

# Внизу: 3 текстовых вывода
y_w = 5.4
findings = [
    ('Москва теряет объём',     '−14 % выручки 2025 при жёсткой регуляторике'),
    ('СПб и Новосибирск растут', '+42 % и +129 % YoY — резерв плотности'),
    ('Екб — низкая утилизация',  '319 п/с/год: идеальный полигон для пилота УПС'),
]
fw = 4.0; fgap = 0.15
for i, (h, b) in enumerate(findings):
    x = 0.5 + i*(fw + fgap)
    CARD(s, x, y_w, fw, 1.35, fill=LIGHT)
    T(s, h, x+0.2, y_w+0.15, fw-0.4, 0.4, size=12, bold=True, color=RED)
    T(s, b, x+0.2, y_w+0.55, fw-0.4, 0.7, size=11, color=DARK, ls=1.3)

punchline(s, 'Источник: внутренние поквартальные данные МТС Юрент за 2024 и 9 мес. 2025 (эксперт А.); портфель = 94 % выручки')
set_notes(s, 'На пяти городах демонстрационного портфеля параметры расходятся в разные стороны. '
              'Утилизация — почти четырёхкратный разрыв между Москвой и Новосибирском. Динамика — от −10 % в Краснодаре до +129 % в Новосибирске. '
              'Регуляторика — от жёсткой в Москве до умеренной в трёх других. Позиция МТС Юрент во всех — преследователь. '
              'Из этой неоднородности вытекает невозможность единого режима — портфельная логика становится единственным реалистичным ответом.')

# ==================== СЛАЙД 6: КОНКУРЕНЦИЯ ====================
s = add_slide()
slide_chrome(s, 'КОНКУРЕНЦИЯ',
             'КФУ на зрелой стадии — доступность покрытия; у МТС Юрент системный гэп 49 vs 67 тыс. ₽ с лидером',
             6, TOTAL)

# Таблица 3 операторов
rows = [
    ('Оператор',   'Парк',       'Логика преимущества',                          'Ограничение'),
    ('Whoosh',     '150 тыс.',   'масштаб парка + вертикальная интеграция',      'снижение предельной отдачи'),
    ('МТС Юрент',  '120 тыс.',   'экосистема МТС, рост, Eleven, широкая география', 'ниже утилизация и выручка/самокат'),
    ('Яндекс',     '30 тыс.',    'трафик суперприложения + транспорт. экосистема',  'меньший парк, непубличный фокус'),
]
y_t = 1.85
col_ws = [2.4, 1.7, 5.2, 3.0]
row_h = 0.6
for r_i, row in enumerate(rows):
    y = y_t + r_i*row_h
    is_h = (r_i == 0); is_mts = (r_i == 2)
    fill = DARK if is_h else (LIGHT if is_mts else (WHITE if r_i % 2 == 1 else VLIGHT))
    CARD(s, 0.5, y, sum(col_ws), row_h, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else (RED if (is_mts and c_i == 0) else DARK)
        bold = is_h or c_i == 0
        size = 11 if is_h else 11.5
        T(s, cell, x+0.15, y, col_ws[c_i]-0.3, row_h,
          size=size, bold=bold, color=color, anchor=MSO_ANCHOR.MIDDLE, ls=1.15)
        x += col_ws[c_i]

# Большой разрыв
y_g = 4.45
CARD(s, 0.5, y_g, 6.2, 2.3, fill=VLIGHT)
T(s, 'Гэп с лидером — в эффективности, не в размере парка',
  0.7, y_g+0.15, 5.8, 0.3, size=11.5, color=GRAY)
T(s, '49,4', 0.7, y_g+0.5, 2.0, 1.3, size=64, bold=True, color=RED,
  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
T(s, 'vs', 2.8, y_g+0.5, 0.6, 1.3, size=30, bold=True, color=DARK,
  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
T(s, '66,8', 3.5, y_g+0.5, 2.0, 1.3, size=64, bold=True, color=DARK,
  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
T(s, 'тыс. ₽ — выручка на самокат в год · МТС Юрент против Whoosh',
  0.7, y_g+1.85, 5.8, 0.3, size=10.5, color=DARK, align=PP_ALIGN.CENTER)

# КФУ блок
CARD(s, 6.95, y_g, 5.85, 2.3, fill=LIGHT)
T(s, 'Ключевые факторы успеха на зрелой стадии', 7.15, y_g+0.15, 5.6, 0.35,
  size=12, bold=True, color=DARK)
T(s, '1.  Фактическая доступность самоката в нужной точке',
  7.15, y_g+0.55, 5.6, 0.35, size=12, color=DARK)
T(s, '2.  Экономика на единицу парка',
  7.15, y_g+0.95, 5.6, 0.3, size=11, color=GRAY)
T(s, '3.  Операционная исполнимость',
  7.15, y_g+1.28, 5.6, 0.3, size=11, color=GRAY)
T(s, '→  Доступность покрытия — критический КФУ; к нему привязан УПС (Слайд 13)',
  7.15, y_g+1.7, 5.6, 0.45, size=11, bold=True, color=RED, ls=1.3)

punchline(s, 'Гэп закрывается не наращиванием парка, а повышением отдачи от уже размещённого парка')
set_notes(s, 'На рынке олигополия трёх: Whoosh — масштаб + интеграция; МТС Юрент — экосистема МТС; Яндекс — суперприложение. '
              'Ключевой фактор успеха на зрелой стадии — доступность покрытия. У МТС Юрент гэп: 49 vs 67 тыс. руб./самокат/год. '
              'Гэп возникает не из-за меньшего парка, а из-за более низкой эффективности уже размещённого парка. '
              'Закрытие гэпа требует дифференциации действий по типам региональных рынков и проектного инструмента, повышающего доступность точечно.')

# ==================== СЛАЙД 7: ВНУТРИ ====================
s = add_slide()
slide_chrome(s, 'ВНУТРИ',
             'Внутри сильная экосистема МТС и Eleven, но централизованная управленческая логика для 187 локаций',
             7, TOTAL)

# 2 колонки: сильные / слабые
y_int = 1.85; cwd = 6.05; ch = 4.35
CARD(s, 0.5, y_int, cwd, ch, fill=LIGHT)
CARD(s, 0.5, y_int, cwd, 0.5, fill=GREEN_TXT)
T(s, '+   СИЛЬНЫЕ СТОРОНЫ', 0.5, y_int+0.08, cwd, 0.35, size=12, bold=True,
  color=WHITE, align=PP_ALIGN.CENTER)
strong = [
    ('Экосистема МТС',         '80 млн абонентов · MTS ID · Premium · обезличенная геоаналитика'),
    ('Технол. платформа',      'после приобретения Eleven (декабрь 2024)'),
    ('Операционная база',      '120 тыс. СИМ в 187 локациях, отстроенные процедуры'),
    ('Гибкость пилотирования', 'компактная команда, короткий контур решений'),
]
y = y_int + 0.65
for h, b in strong:
    T(s, h, 0.7, y, cwd-0.4, 0.3, size=12, bold=True, color=DARK)
    T(s, b, 0.7, y+0.3, cwd-0.4, 0.5, size=10.5, color=GRAY, ls=1.25)
    y += 0.9

CARD(s, 6.78, y_int, cwd, ch, fill=LIGHT)
CARD(s, 6.78, y_int, cwd, 0.5, fill=RED_TXT)
T(s, '−   СЛАБЫЕ СТОРОНЫ', 6.78, y_int+0.08, cwd, 0.35, size=12, bold=True,
  color=WHITE, align=PP_ALIGN.CENTER)
weak = [
    ('Централизованная управл. логика', 'мало чувствительна к различиям регионов'),
    ('Нет процедуры назначения режимов', 'решения по городам — набор разовых кейсов'),
    ('Огранич. локальная экспертиза',    'дефицит выделенных функций региональной координации'),
    ('Отставание по эффективности',       'выручка/самокат и утилизация ниже Whoosh'),
]
y = y_int + 0.65
for h, b in weak:
    T(s, h, 6.98, y, cwd-0.4, 0.3, size=12, bold=True, color=DARK)
    T(s, b, 6.98, y+0.3, cwd-0.4, 0.5, size=10.5, color=GRAY, ls=1.25)
    y += 0.9

# Dynamic capabilities note
T(s, 'Dynamic capabilities:  операционный блок развит,  способность перестраивать ресурсы под изменения рынка — в начальной зрелости',
  0.5, 6.35, 12.3, 0.4, size=11, italic=True, color=DARK, align=PP_ALIGN.CENTER)

punchline(s, 'Проблема — не в отсутствии ресурсов, а в том, что текущая модель не превращает их в региональный результат')
set_notes(s, 'Внутри Юрента сильные стороны: экосистема МТС (80 млн абонентов, MTS ID, Premium), технологическая платформа Eleven, '
              'парк 120 тыс. СИМ в 187 локациях, гибкость пилотов. Слабые: централизованная управленческая логика, '
              'нет процедуры назначения режимов, ограниченная локальная экспертиза. В терминах динамических способностей: '
              'операционный блок развит, способность перестраивать ресурсы под изменения рынка — в начальной зрелости. '
              'Проблема — не в отсутствии потенциала, а в способе превращения ресурсов в региональный результат.')

# ==================== СЛАЙД 8: ДИАГНОЗ ====================
s = add_slide()
slide_chrome(s, 'ДИАГНОЗ',
             'Стратегический диагноз: портфель стал неоднородным быстрее, чем управленческая модель оператора',
             8, TOTAL)

# Главная схема: 4 блока с переходом
y_d = 1.85; box_h = 2.0
boxes = [
    ('Федеральный масштаб\nи зрелый рынок',  None, None),
    ('Разные города:\nспрос, регуляторика,\nсезон, конкуренция', None, None),
    ('Единая логика\nтеряет точность', None, None),
    ('Нужна система\n«город → режим → KPI»', None, None),
]
bw = 2.9; bgap = 0.25
total_bw = 4*bw + 3*bgap
start_bx = (13.333 - total_bw) / 2
for i, (text, _, _) in enumerate(boxes):
    x = start_bx + i*(bw + bgap)
    fill = RED if i == 3 else LIGHT
    txt = WHITE if i == 3 else DARK
    CARD(s, x, y_d, bw, box_h, fill=fill)
    T(s, text, x+0.1, y_d, bw-0.2, box_h, size=13, bold=True, color=txt,
      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, ls=1.25)
    if i < 3:
        # arrow
        ln = s.shapes.add_connector(1, Inches(x+bw), Inches(y_d+box_h/2),
                                      Inches(x+bw+bgap), Inches(y_d+box_h/2))
        ln.line.color.rgb = RED; ln.line.width = Pt(1.5)
        line_xml = ln.line._get_or_add_ln()
        tail = line_xml.find(qn('a:tailEnd'))
        if tail is None:
            from lxml import etree
            tail = etree.SubElement(line_xml, qn('a:tailEnd'))
        tail.set('type', 'triangle'); tail.set('w', 'med'); tail.set('len', 'med')

# 3 секции внизу: Причины / Последствия / Ответ
y_b = 4.2
sections = [
    ('Причины разрыва', RED, [
        'стагнация совокупного спроса',
        'регуляторная фрагментация по городам',
        'разная утилизация парка по городам',
    ]),
    ('Последствия (если не менять)', GRAY, [
        'ошибки в масштабе присутствия',
        'избыточный OpEx в тяжёлых рынках',
        'недоиспользование роста в R2-городах',
    ]),
    ('Ответ работы', GREEN_TXT, [
        'портфельная стратегия с разными режимами',
        'общий контур контроля и KPI',
        'дорожная карта пилота УПС',
    ]),
]
sw = 4.05; sgap = 0.1
for i, (h, color, items) in enumerate(sections):
    x = 0.5 + i*(sw + sgap)
    CARD(s, x, y_b, sw, 2.5, fill=VLIGHT)
    CARD(s, x, y_b, sw, 0.45, fill=color)
    T(s, h, x, y_b+0.07, sw, 0.3, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    body = '\n'.join(f'•  {it}' for it in items)
    T(s, body, x+0.25, y_b+0.6, sw-0.4, 1.85, size=11, color=DARK, ls=1.5)

punchline(s, 'Нужна управляемая система различения региональных рынков «город → режим → KPI»')
set_notes(s, 'Сводный стратегический диагноз: разрыв состоит в том, что МТС Юрент работает с неоднородной портфельной '
              'средой через единообразную управленческую логику. Причины: переход к зрелости, расхождение регионов, '
              'слабая формализация регионального управления. Если не менять — ухудшение экономики покрытия. '
              'Ответ — портфельная стратегия с дифференцированными режимами и общим контуром контроля.')

# ==================== СЛАЙД 9: РАМКА ====================
s = add_slide()
slide_chrome(s, 'РАМКА',
             'Рамка I-II-III переводит различия городов в параметры стратегического выбора',
             9, TOTAL)

# Левая колонка: таблица 3 уровней
T(s, 'Три уровня диагностики', 0.5, 1.85, 7, 0.4, size=13, bold=True, color=DARK)

rows = [
    ('Уровень', 'Что оцениваем', 'На что влияет'),
    ('I',  'ёмкость, регуляторика, конкуренция,\nположение МТС Юрент на рынке',
     'выбор режима\n(R1, R2 или R3)'),
    ('II', 'сезонность, плотность спроса,\nстоимость покрытия',
     'осуществимость\nрежима в сезоне'),
    ('III', 'цены, промо, партнёрства,\nпереразмещение парка, УПС',
     'рычаги внутри\nуже выбранного режима'),
]
y_t = 2.3; col_ws = [1.1, 4.1, 2.6]; row_h = 0.95
for r_i, row in enumerate(rows):
    y = y_t + r_i*row_h
    is_h = (r_i == 0)
    h_real = 0.45 if is_h else row_h
    if is_h:
        fill = DARK
    else:
        fill = WHITE if r_i % 2 == 1 else VLIGHT
    CARD(s, 0.5, y, sum(col_ws), h_real, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else (RED if c_i == 0 and not is_h else DARK)
        bold = is_h or c_i == 0
        size = 22 if (c_i == 0 and not is_h) else (10.5 if is_h else 11)
        align = PP_ALIGN.CENTER if c_i == 0 else PP_ALIGN.LEFT
        T(s, cell, x+0.15, y, col_ws[c_i]-0.3, h_real,
          size=size, bold=bold, color=color, align=align, anchor=MSO_ANCHOR.MIDDLE, ls=1.2)
        x += col_ws[c_i]
    if is_h:
        y_t += -row_h + 0.45  # adjust next row position

# Правая колонка: 5 критериев оценки сценариев (на основе теории)
T(s, '5 критериев оценки сценариев', 8.5, 1.85, 4.5, 0.4, size=13, bold=True, color=DARK)
criteria = [
    ('1', 'Ресурсно-теоретическая защитимость'),
    ('2', 'Согласованность с динамическими способностями'),
    ('3', 'Операционная исполнимость'),
    ('4', 'Финансово-операционный эффект'),
    ('5', 'Профиль риска'),
]
y_c = 2.3
for n, name in criteria:
    CARD(s, 8.5, y_c, 4.4, 0.6, fill=LIGHT)
    T(s, n, 8.5, y_c, 0.6, 0.6, size=18, bold=True, color=RED,
      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    T(s, name, 9.1, y_c, 3.7, 0.6, size=11, color=DARK,
      anchor=MSO_ANCHOR.MIDDLE, ls=1.2)
    y_c += 0.65

# Логика связки
T(s, 'I-II дают диагноз города → выбор режима;  III показывает доступные рычаги управления',
  0.5, 6.4, 12.3, 0.35, size=11, italic=True, color=DARK, align=PP_ALIGN.CENTER)

punchline(s, 'Рамка готовит сравнение S1/S2/S3 и карту действий по городам')
set_notes(s, 'Главный методический результат — две конструкции. '
              'Трёхуровневая иерархия факторов различения рынков: I — структурная конфигурация (выбор режима), '
              'II — операционная исполнимость (проверка реализуемости), III — тактические рычаги. '
              'Пять критериев оценки сценариев, выведенных из теории: ресурсная защитимость, динамические способности, '
              'операционная исполнимость, финансовый эффект, профиль риска.')

# ==================== СЛАЙД 10: АЛЬТЕРНАТИВЫ ====================
s = add_slide()
slide_chrome(s, 'АЛЬТЕРНАТИВЫ',
             'Из трёх альтернатив оптимальна диверсификация режимов по портфелю — S3',
             10, TOTAL)

# Большая таблица 6×4
rows = [
    ('Критерий',                                'S1: единая логика',                      'S2: единый режим + адаптация',    'S3: диверсификация режимов'),
    ('Ресурсная защитимость',                   'низкая: ресурсы не дают дифференциации', 'средняя: дисциплина исполнения',  'высокая: ресурсы МТС работают там, где дают эффект'),
    ('Dynamic capabilities',                    'низкая: слабое различение рынков',       'средняя: адаптация внутри 1 режима', 'высокая: воспроизводимое различение рынков'),
    ('Операционная исполнимость',               'высокая по факту (инерция)',             'высокая',                          'средняя: нужны компенсаторы'),
    ('Финансовый эффект',                       'низкий',                                  'средний',                          'высокий при пилоте и контроле OpEx'),
    ('Профиль риска',                           'концентрирован',                          'концентрирован в R3',              'распределён; требует координации'),
]
y_t = 1.85
col_ws = [3.3, 3.3, 3.3, 3.0]
row_h = 0.65
for r_i, row in enumerate(rows):
    y = y_t + r_i*row_h
    is_h = (r_i == 0)
    fill = DARK if is_h else (WHITE if r_i % 2 == 1 else VLIGHT)
    CARD(s, 0.5, y, sum(col_ws), row_h, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else DARK
        bold = is_h or c_i == 0
        if not is_h and c_i == 3:
            color = RED; bold = True
        size = 11 if is_h else (10.5 if c_i > 0 else 11)
        align = PP_ALIGN.LEFT
        T(s, cell, x+0.15, y, col_ws[c_i]-0.3, row_h,
          size=size, bold=bold, color=color, align=align, anchor=MSO_ANCHOR.MIDDLE, ls=1.2)
        x += col_ws[c_i]

# Вывод-полоса
y_w = 6.0
CARD(s, 0.5, y_w, 12.3, 0.85, fill=RED)
T(s, 'Выбор: S3', 0.7, y_w+0.15, 2.5, 0.4, size=14, bold=True, color=WHITE)
T(s, 'S3 решает исходный разрыв при управляемых рисках. Уступает только по операционной '
     'исполнимости в горизонте перестройки — адресовано компенсаторами (Слайд 14)',
  3.0, y_w+0.15, 9.0, 0.6, size=11, color=WHITE, ls=1.3)

punchline(s, 'S3 связывает региональные различия с ресурсами и эффектом ценой повышенных требований к координации')
set_notes(s, 'Три альтернативы: S1 — единая логика, S2 — единый режим R3 + адаптация, S3 — диверсификация режимов по портфелю. '
              'Сравнение по пяти критериям. S3 сильнее по четырём из пяти. Уступает только по операционной исполнимости в '
              'горизонте перестройки — но этот проигрыш адресован компенсаторами. Выбран S3.')

# ==================== СЛАЙД 11: СТРАТЕГИЯ ====================
s = add_slide()
slide_chrome(s, 'СТРАТЕГИЯ',
             'S3 — управление портфелем режимов, назначенных конкретным городам с регулярным пересмотром по KPI',
             11, TOTAL)

# 3 карточки режимов (компактные)
modes = [
    ('R1', 'Плацдарм',  'закрепление',
     'минимальная плотность парка; операционные стандарты; контакты с городом; ограниченное промо',
     'KPI:  покрытие ключевых зон, базовая утилизация, отсутствие конфликтов'),
    ('R2', 'Атака',     'рост доли',
     'расширение и переразмещение парка; плотность ключевых зон; клиентская и ценовая работа; УПС',
     'KPI:  поездки, доля относительно лидера, активная аудитория, выручка/самокат'),
    ('R3', 'Удержание', 'защита экономики',
     'дисциплина расстановки; снижение издержек обслуживания; защита микролокаций; удержание пользователей',
     'KPI:  доля поездок, выручка/самокат, ребалансировка, EBITDA-маржа'),
]
y_m = 1.85; mh = 2.3; mw = 4.05; mgap = 0.1
for i, (code, name, purpose, lev, kpi) in enumerate(modes):
    x = 0.5 + i*(mw + mgap)
    CARD(s, x, y_m, mw, mh, fill=LIGHT)
    # Top red strip with code + name
    CARD(s, x, y_m, mw, 0.7, fill=RED)
    T(s, code, x+0.15, y_m+0.13, 0.9, 0.45, size=22, bold=True, color=WHITE,
      align=PP_ALIGN.CENTER)
    T(s, name, x+1.1, y_m+0.1, mw-1.25, 0.3, size=14, bold=True, color=WHITE)
    T(s, purpose, x+1.1, y_m+0.4, mw-1.25, 0.25, size=10, color=WHITE)
    # Body
    T(s, lev, x+0.2, y_m+0.85, mw-0.4, 0.95, size=10.5, color=DARK, ls=1.3)
    T(s, kpi, x+0.2, y_m+1.85, mw-0.4, 0.4, size=10, bold=True, color=RED, ls=1.2)

# Таблица городов 5×3
T(s, 'Назначение режимов по городам портфеля', 0.5, 4.4, 12, 0.35, size=12, bold=True, color=DARK)
cities = [
    ('Город',            'Режим S3',            'Что делаем'),
    ('Москва',           'R3 (огранич.)',       'оптимизируем парк, концентрируемся в устойчивых зонах, соблюдаем регуляторные ограничения'),
    ('Санкт-Петербург',  'R2 → R3',             'атакуем при запасе плотности; готовим переход к удержанию при ужесточении регуляторики'),
    ('Екатеринбург',     'R2 + пилот УПС',      'пилотируем УПС, переразмещаем парк в ключевых зонах, растим утилизацию'),
    ('Краснодар',        'R2 / R3',             'используем длинный сезон, удерживаем плотность, контролируем стоимость покрытия'),
    ('Новосибирск',      'R2/R3 эффект.',       'растём с низкой базы, контролируем утилизацию и стоимость покрытия в коротком сезоне'),
]
y_c = 4.8
col_ws = [2.5, 2.5, 7.3]
row_h = 0.35
for r_i, row in enumerate(cities):
    y = y_c + r_i*row_h
    is_h = (r_i == 0); is_pilot = (r_i == 3)
    fill = DARK if is_h else (LIGHT if is_pilot else (WHITE if r_i % 2 == 1 else VLIGHT))
    CARD(s, 0.5, y, sum(col_ws), row_h, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else DARK
        if not is_h and c_i == 1:
            color = RED
        bold = is_h or c_i == 0 or c_i == 1
        size = 10 if is_h else 10
        T(s, cell, x+0.15, y, col_ws[c_i]-0.3, row_h,
          size=size, bold=bold, color=color, anchor=MSO_ANCHOR.MIDDLE, ls=1.15)
        x += col_ws[c_i]

punchline(s, 'Стратегия задаёт правила назначения режима и переходов между режимами по фактическим KPI каждый сезон')
set_notes(s, 'Содержательное наполнение S3 — три режима. R1 «Плацдарм» — закрепление, R2 «Атака» — рост доли, '
              'R3 «Удержание» — защита маржи. Каждый описан через однородные поля. Назначение по 5 городам: '
              'Москва — R3 (огранич.), СПб — R2 → R3, Екб — R2 + пилот УПС, Краснодар — R2/R3, Новосибирск — R2/R3 (эффект.). '
              'Это не статическая кластеризация, а воспроизводимая процедура с пересмотром каждый сезон.')

# ==================== СЛАЙД 12: ИНСТРУМЕНТ ====================
s = add_slide()
slide_chrome(s, 'ИНСТРУМЕНТ',
             'УПС превращает локальный дефицит доступности в управляемое действие, дешевле физической ребалансировки в 4,5×',
             12, TOTAL)

# Процесс-лента
flow = ['Сигнал дефицита', 'Отбор точки\nи окна', 'Выбор стимула', 'Доступность\nвосстановлена']
fy = 1.85; fw = 2.85; fgap = 0.2
total_fw = 4*fw + 3*fgap
fx0 = (13.333 - total_fw) / 2
for i, step in enumerate(flow):
    x = fx0 + i*(fw + fgap)
    CARD(s, x, fy, fw, 0.7, fill=LIGHT)
    T(s, step, x, fy, fw, 0.7, size=12, bold=True, color=DARK,
      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, ls=1.2)
    if i < 3:
        ln = s.shapes.add_connector(1, Inches(x+fw), Inches(fy+0.35),
                                      Inches(x+fw+fgap), Inches(fy+0.35))
        ln.line.color.rgb = RED; ln.line.width = Pt(2)
        line_xml = ln.line._get_or_add_ln()
        tail = line_xml.find(qn('a:tailEnd'))
        if tail is None:
            from lxml import etree
            tail = etree.SubElement(line_xml, qn('a:tailEnd'))
        tail.set('type', 'triangle'); tail.set('w', 'med'); tail.set('len', 'med')

# 3 механизма (под шагом "Выбор стимула")
T(s, 'Три механизма УПС (шаг «Выбор стимула»)', 0.5, 2.85, 12.3, 0.35,
  size=12, bold=True, color=DARK)

mechs = [
    ('М1', 'Стимул на стороне точки высадки',     'бонус за завершение поездки в зоне дефицита'),
    ('М2', 'Стимул на стороне точки посадки',     'стимул взять самокат из зоны переизбытка'),
    ('М3', 'Микрозадание внешнему исполнителю',   'внешний исполнитель перемещает самокат между точками'),
]
y_me = 3.3; mew = 4.05; meh = 1.55; megap = 0.1
for i, (code, name, body) in enumerate(mechs):
    x = 0.5 + i*(mew + megap)
    CARD(s, x, y_me, mew, meh, fill=LIGHT)
    # Code badge
    CARD(s, x, y_me, 0.9, meh, fill=RED)
    T(s, code, x, y_me, 0.9, meh, size=22, bold=True, color=WHITE,
      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    T(s, name, x+1.0, y_me+0.15, mew-1.1, 0.45, size=11.5, bold=True, color=DARK, ls=1.2)
    T(s, body, x+1.0, y_me+0.7, mew-1.1, 0.8, size=10.5, color=GRAY, ls=1.3)

# 4 числа эффекта внизу
nums = [
    ('120 ₽',     'физическая\nребалансировка',  DARK),
    ('26,5 ₽',    'средняя стоимость\nстимула',   RED),
    ('≈ 11,7 млн ₽', 'денежный\nэффект сезона',     DARK),
    ('≈ 6 млн ₽', 'EBITDA-вклад\nгод 1',         RED),
]
y_n = 5.05; nw = 3.0; ngap = 0.1
for i, (n, body, color) in enumerate(nums):
    x = 0.5 + i*(nw + ngap)
    CARD(s, x, y_n, nw, 1.7, fill=VLIGHT)
    bignum(s, n, x, y_n+0.25, nw, 0.7, size=24, color=color, align=PP_ALIGN.CENTER)
    T(s, body, x, y_n+1.0, nw, 0.6, size=10.5, color=GRAY,
      align=PP_ALIGN.CENTER, ls=1.3)

punchline(s, 'KPI пилота:  прирост поездок в пилотных зонах · снижение физической ребалансировки · стоимость стимула на предотвращённую потерю · отклик пользователей')
set_notes(s, 'УПС адресует ключевой фактор успеха — доступность покрытия. Четырёхшаговая логика: сигнал дефицита, '
              'отбор точки и окна, выбор стимула, доступность восстановлена. Три механизма: М1 — стимул на точке высадки, '
              'М2 — стимул на точке посадки, М3 — оплачиваемое микрозадание внешнему исполнителю. '
              'Ключевая экономика: 120 руб. физической ребалансировки vs 26,5 руб. стимула — в 4,5 раза дешевле. '
              'Сезонный денежный эффект пилота в Екб — около 11,7 млн руб., в EBITDA-выражении около 6 млн руб.')

# ==================== СЛАЙД 13: ФИНМОДЕЛЬ ====================
s = add_slide()
slide_chrome(s, 'ФИНМОДЕЛЬ',
             'S3 окупается за счёт перенастройки уже существующего портфеля: +25 млн ₽ EBITDA год 1, NPV +134 млн при WACC 20 %',
             13, TOTAL)

# Левая колонка: финансовая таблица
T(s, 'Эффект S3 vs S1 на 5 городах (4,17 млрд ₽ выручки), млн руб.',
  0.5, 1.85, 7.3, 0.35, size=11, color=GRAY)

rows = [
    ('Статья',                                                  'Млн ₽'),
    ('Затраты S3:  управление, маркетинг, УПС, накладные',      '−48'),
    ('Эффект R2:  прирост выручки СПб, Екб, Крд, Нск',          '+12'),
    ('УПС-пилот в Екатеринбурге',                                '+6'),
    ('R3 Москва:  оптимизация OpEx (−2 % от выручки)',           '+52'),
    ('R1/R2 малые города:  экономия капвложений',                '+3'),
    ('Чистый эффект год 1',                                      '+25'),
]
y_t = 2.3
col_ws = [5.8, 1.5]
row_h = 0.45
for r_i, row in enumerate(rows):
    y = y_t + r_i*row_h
    is_h = (r_i == 0); is_total = (r_i == len(rows)-1)
    fill = DARK if is_h else (RED if is_total else (WHITE if r_i % 2 == 1 else VLIGHT))
    CARD(s, 0.5, y, sum(col_ws), row_h, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        if is_h:
            color = WHITE
        elif is_total:
            color = WHITE
        elif c_i == 1:
            color = RED_TXT if cell.startswith('−') else GREEN_TXT
        else:
            color = DARK
        bold = is_h or is_total or c_i == 1
        align = PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.RIGHT
        size = 11.5
        T(s, cell, x+0.15, y, col_ws[c_i]-0.3, row_h,
          size=size, bold=bold, color=color, align=align, anchor=MSO_ANCHOR.MIDDLE)
        x += col_ws[c_i]

# 3 сценария чувствительности (бары-индикаторы)
y_s = 5.65
T(s, 'Чувствительность год 1', 0.5, y_s, 7.3, 0.3, size=11, bold=True, color=DARK)
scenarios = [
    ('Пессимистично', '−14 млн', RED_TXT),
    ('База',          '+25 млн', RED),
    ('Оптимистично',  '+64 млн', GREEN_TXT),
]
sw_s = 2.4; sgap_s = 0.05
for i, (lbl, val, color) in enumerate(scenarios):
    x = 0.5 + i*(sw_s + sgap_s)
    CARD(s, x, y_s+0.35, sw_s, 0.85, fill=LIGHT)
    T(s, lbl, x, y_s+0.4, sw_s, 0.3, size=10, color=GRAY, align=PP_ALIGN.CENTER)
    T(s, val, x, y_s+0.65, sw_s, 0.5, size=18, bold=True, color=color,
      align=PP_ALIGN.CENTER)

# Правая колонка: ключевые финансовые показатели
T(s, 'Показатели проекта', 8.1, 1.85, 4.7, 0.35, size=11, color=GRAY)

metrics = [
    ('+25 млн ₽',     'чистый эффект',  'год 1',                            RED),
    ('+80…+100',      'млн ₽',          'год 2 при тиражировании УПС',      RED),
    ('+134 млн ₽',    'NPV',            'WACC 20 %, горизонт 3 года',       DARK),
    ('≈ 8 мес.',      'Payback',        'один пилотный сезон',              DARK),
    ('52 % / 243 %',  'ROI',            'год 1 / год 2',                    DARK),
]
y_m = 2.3; mh = 0.85; mgap = 0.05
for big, mid, sub, color in metrics:
    CARD(s, 8.1, y_m, 4.7, mh, fill=LIGHT)
    T(s, big, 8.25, y_m+0.05, 2.4, mh-0.1, size=18, bold=True, color=color,
      anchor=MSO_ANCHOR.MIDDLE)
    T(s, mid, 10.7, y_m+0.07, 2.0, 0.35, size=11, bold=True, color=DARK)
    T(s, sub, 10.7, y_m+0.42, 2.0, 0.35, size=9.5, color=GRAY, ls=1.2)
    y_m += mh + mgap

punchline(s, 'Главный драйвер модели — перераспределение режимов, снижение OpEx в тяжёлых рынках и пилот УПС там, где есть локальный дефицит')
set_notes(s, 'Финансовая модель: затраты года 1 — 48 млн руб., эффекты — 73 млн EBITDA, чистый эффект +25 млн. '
              'Чувствительность −14 / +25 / +64 млн. Год 2 при тиражировании УПС — +80–100 млн. '
              'NPV проекта при WACC 20 % на 3-летнем горизонте — +134 млн руб. Payback — около 8 месяцев. '
              'ROI год 1 — 52 %, год 2 — 243 %. Capex проекта всего 8 млн на IT-инфраструктуру УПС — низкая капитальная интенсивность.')

# ==================== СЛАЙД 14: ПЛАН ====================
s = add_slide()
slide_chrome(s, 'ПЛАН',
             'S3 запускается через пилот в Екатеринбурге, не через одномоментную перестройку',
             14, TOTAL)

# Timeline 5 месяцев
T(s, 'Дорожная карта пилота УПС, 12 месяцев', 0.5, 1.85, 12.3, 0.35, size=11, color=GRAY)

months = [
    ('М1', 'Диагностика',       'базовые KPI,\nчувствительные микролокации'),
    ('М2', 'Запуск М1+М2',     'стимулы\nпользователям'),
    ('М3', 'Запуск М3',         'микрозадания\nвнешним исполнителям'),
    ('М4', 'Калибровка',        'параметры стимулов,\nрасширение зоны'),
    ('М5', 'Итоги',             'решение\nо тиражировании'),
]
y_t = 2.4; tw = 2.4; tgap = 0.1
total_tw = 5*tw + 4*tgap
tx0 = (13.333 - total_tw) / 2
# Linking line
ln = s.shapes.add_connector(1, Inches(tx0+0.4), Inches(y_t+0.8),
                             Inches(tx0+total_tw-0.4), Inches(y_t+0.8))
ln.line.color.rgb = RED; ln.line.width = Pt(2)
for i, (code, name, body) in enumerate(months):
    x = tx0 + i*(tw + tgap)
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+tw/2-0.5),
                                 Inches(y_t+0.4), Inches(1.0), Inches(1.0))
    circle.fill.solid(); circle.fill.fore_color.rgb = RED
    circle.line.fill.background()
    tb = circle.text_frame
    tb.margin_top = tb.margin_bottom = tb.margin_left = tb.margin_right = 0
    tb.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp = tb.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
    rr = pp.add_run(); rr.text = code
    rr.font.name = FONT; rr.font.size = Pt(22); rr.font.bold = True
    rr.font.color.rgb = WHITE
    T(s, name, x, y_t+1.5, tw, 0.4, size=12, bold=True, color=DARK,
      align=PP_ALIGN.CENTER)
    T(s, body, x, y_t+1.9, tw, 0.8, size=10, color=GRAY,
      align=PP_ALIGN.CENTER, ls=1.3)

# Bottom — план после пилота
y_after = 5.0
CARD(s, 0.5, y_after, 12.3, 0.6, fill=VLIGHT)
T(s, 'После пилота', 0.7, y_after+0.05, 2, 0.25, size=10, color=GRAY)
T(s, 'Год 2 →  тиражирование УПС на СПб, Краснодар и Новосибирск (R2-кластер);  пересмотр режима каждый сезон по фактическим KPI',
  0.7, y_after+0.3, 12, 0.3, size=11, bold=True, color=DARK)

# KPI пилота — 4 индикатора
T(s, '4 KPI пилота', 0.5, 5.85, 12.3, 0.3, size=12, bold=True, color=DARK)
kpis = [
    ('Прирост поездок',         'в пилотных зонах vs база'),
    ('Снижение ребалансировки', 'доля в общем объёме операций'),
    ('Стоимость стимула',       'на 1 предотвращённую потерю'),
    ('Отклик пользователей',    'доля принятых стимулов'),
]
y_k = 6.2; kw = 3.0; kgap = 0.1
for i, (h, b) in enumerate(kpis):
    x = 0.5 + i*(kw + kgap)
    CARD(s, x, y_k, kw, 0.6, fill=LIGHT)
    T(s, h, x+0.15, y_k+0.05, kw-0.3, 0.3, size=11, bold=True, color=DARK)
    T(s, b, x+0.15, y_k+0.32, kw-0.3, 0.25, size=9, color=GRAY, ls=1.1)

punchline(s, 'Пилот калибрует допущения финмодели; решение о тиражировании принимается на конец сезона по факту KPI')
set_notes(s, 'Дорожная карта пилота УПС в Екатеринбурге: М1 — диагностика, М2 — М1+М2, М3 — М3, М4 — калибровка, М5 — итоги. '
              'С года 2 — тиражирование на СПб, Краснодар, Новосибирск. KPI пилота: прирост поездок в зонах, снижение доли физической '
              'ребалансировки, стоимость стимула на предотвращённую потерю, отклик пользователей на предложения.')

# ==================== СЛАЙД 15: РИСКИ ====================
s = add_slide()
slide_chrome(s, 'РИСКИ',
             'S3 управляем при наличии отдельного портфельного контура и регулярного пересмотра режимов по KPI',
             15, TOTAL)

# Таблица рисков 6×3 (риск/компенсатор/что контролировать)
rows = [
    ('Риск',                              'Компенсатор',                                              'Что контролировать'),
    ('Рассинхронизация режимов',          'функция портфельного управления',                          'ежемесячная сверка «город → режим»'),
    ('Потеря единого UX',                 'единый продуктовый слой при разной операционной модели',   'приложение, базовая цена, сервисные правила'),
    ('Кадровая нагрузка',                 'поэтапный запуск: 5 городов → расширенный портфель',       'региональные менеджеры и аналитики'),
    ('Регуляторные шоки',                 'правила перехода R2 → R3 / R3 ограниченное',               'изменение условий присутствия'),
    ('Ошибка диагностики',                'пересмотр режима по KPI',                                   'выручка/самокат, доля поездок, стоимость покрытия'),
    ('Универсализация УПС',               'фиксация условий применимости (только R2 и R3)',           'каталог сценариев применения УПС'),
]
y_t = 1.85
col_ws = [3.4, 5.3, 4.1]
row_h = 0.52
for r_i, row in enumerate(rows):
    y = y_t + r_i*row_h
    is_h = (r_i == 0)
    fill = DARK if is_h else (WHITE if r_i % 2 == 1 else VLIGHT)
    CARD(s, 0.5, y, sum(col_ws), row_h, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else (RED if c_i == 0 and not is_h else DARK)
        bold = is_h or c_i == 0
        size = 11 if is_h else 10.5
        T(s, cell, x+0.15, y, col_ws[c_i]-0.3, row_h,
          size=size, bold=bold, color=color, anchor=MSO_ANCHOR.MIDDLE, ls=1.15)
        x += col_ws[c_i]

# Финальная полоса с темой "S3 — управляемая система"
y_w = 5.85
CARD(s, 0.5, y_w, 12.3, 1.0, fill=RED)
T(s, 'Главное', 0.7, y_w+0.12, 2, 0.3, size=11, color=WHITE)
T(s, 'S3 не требует одномоментной перестройки всей компании. Запускается через пилот и расширяется '
     'при подтверждении экономики. Каждый риск адресован конкретным управленческим механизмом.',
  0.7, y_w+0.38, 11.5, 0.55, size=12, bold=True, color=WHITE, ls=1.3)

punchline(s, 'Распределённый профиль риска даётся ценой повышенных требований к координации — управленческий компенсатор обязателен, а не опционален')
set_notes(s, 'Шесть рисков, у каждого свой компенсатор и измеримая величина для контроля. '
              'S3 не требует одномоментной перестройки. Запускается через пилот. Каждый риск адресован конкретным управленческим механизмом.')

# ==================== СЛАЙД 16: ВЫВОД ====================
s = add_slide()
slide_chrome(s, 'ВЫВОД',
             'Управленческий вывод: МТС Юрент получает рабочий инструментарий для решений по портфелю',
             16, TOTAL)

# 4 действия
verbs = [
    ('1', 'Диагностирована',     'управленческая проблема\nрегионального портфеля\nиз 187 локаций'),
    ('2', 'Сопоставлены',         'три стратегические альтернативы\nпо пяти критериям\nиз теоретической основы'),
    ('3', 'Предложен инструмент', 'УПС — управление покрытием\nчерез стимулы,\nв 4,5× дешевле ребалансировки'),
    ('4', 'Выполнена оценка',     'финмодель:  +25 млн ₽ год 1,\nNPV +134 млн при WACC 20 %,\npayback ≈ 8 мес.'),
]
y_v = 1.95; vw = 2.95; vh = 2.5; vgap = 0.2
for i, (n, head, body) in enumerate(verbs):
    x = 0.5 + i*(vw + vgap)
    is_first = (i == 0)
    bg = RED if is_first else LIGHT
    txt = WHITE if is_first else DARK
    sub = WHITE if is_first else GRAY
    head_color = WHITE if is_first else RED
    CARD(s, x, y_v, vw, vh, fill=bg)
    T(s, n, x+0.2, y_v+0.15, vw-0.4, 0.65, size=36, bold=True, color=txt)
    T(s, head, x+0.2, y_v+0.9, vw-0.4, 0.4, size=13, bold=True, color=head_color)
    T(s, body, x+0.2, y_v+1.35, vw-0.4, 1.1, size=11, color=sub, ls=1.3)

# Главный вклад
CARD(s, 0.5, 4.7, 12.3, 1.6, fill=VLIGHT)
T(s, 'Главный практический вклад работы', 0.85, 4.85, 12, 0.3, size=11, color=GRAY)
T(s, 'Рабочий инструментарий для подготовки стратегических решений МТС Юрент: '
     'диагностический каркас «город → режим → KPI», система режимов R1–R3, '
     'проектный инструмент УПС и шаблон финансовой модели для расширения '
     'на остальные города присутствия.',
  0.85, 5.2, 11.5, 1.0, size=14, bold=True, color=DARK, ls=1.35)

# "Спасибо"
T(s, 'Спасибо за внимание · готов к вопросам',
  0.5, 6.6, 12.3, 0.4, size=16, bold=True, color=RED, align=PP_ALIGN.CENTER)

set_notes(s, 'Финальная фраза: «Если кратко подвести итог работы. Управленческая проблема МТС Юрент состояла в том, '
              'что для портфеля из 187 локаций нет формализованной процедуры выбора режима присутствия. '
              'Проведённая диагностика показала, что причина — в исчерпании источника новой ценности от единой логики расширения. '
              'Сопоставление трёх сценариев по пяти критериям обосновало выбор стратегии диверсификации режимов. '
              'В её рамках предложен проектный инструмент УПС и построена портфельная финансовая модель с эффектом +25 млн в год 1 '
              'и +80–100 млн в год 2 при тиражировании. Главный практический вклад — рабочий инструментарий для МТС Юрент. '
              'Спасибо за внимание, готов к вопросам».')


# ==================== БЭКАП-СЛАЙДЫ ====================

def backup_chrome(s, label, num, total_b, headline):
    # CAP-метка серая для backup
    T(s, f'BACKUP  ·  {label}', 0.5, 0.32, 5, 0.3, size=11, bold=True, color=GRAY,
      align=PP_ALIGN.LEFT)
    T(s, f'{num} / {total_b}', 12.4, 0.32, 0.7, 0.3, size=10, color=GRAY_LIGHT,
      align=PP_ALIGN.RIGHT)
    T(s, headline, 0.5, 0.65, 12.3, 1.0, size=20, bold=True, color=DARK, ls=1.15)
    ln = s.shapes.add_connector(1, Inches(0.5), Inches(1.55), Inches(12.83), Inches(1.55))
    ln.line.color.rgb = GRAY; ln.line.width = Pt(1.5)

BTOTAL = 5

# B1: Пять критериев подробно
s = add_slide()
backup_chrome(s, 'КРИТЕРИИ', 1, BTOTAL,
              'Пять критериев оценки стратегических сценариев — теоретическая основа и измерение')
crit = [
    ('1. Ресурсно-теоретическая защитимость',         'Wernerfelt 1984; Barney 1991',
     'Опирается ли сценарий на актуальную ресурсную базу оператора и не требует радикальной перестройки'),
    ('2. Согласованность с динамическими способностями','Teece, Pisano, Shuen 1997; Teece 2007',
     'Способность сценария воспроизводимо перестраивать конфигурацию ресурсов под меняющиеся условия'),
    ('3. Операционная исполнимость',                  'Sareen, Remme, Haarstad 2021; Moran 2021',
     'Реализуемость сценария при текущей операционной зрелости оператора и регуляторных ограничениях'),
    ('4. Финансово-операционный эффект',              'Aarhaug et al. 2023; Shah et al. 2022; Cennamo 2021',
     'Ожидаемый прирост EBITDA на портфеле городов в горизонте 12 месяцев'),
    ('5. Профиль риска',                               'Groth et al. 2025; Coenegrachts et al. 2024',
     'Структура и величина рисков (поведенческих, технологических, регуляторных, конкурентных)'),
]
yc = 1.85
for h, src, body in crit:
    CARD(s, 0.5, yc, 12.3, 0.95, fill=LIGHT)
    T(s, h, 0.7, yc+0.1, 6.0, 0.4, size=12.5, bold=True, color=DARK)
    T(s, src, 0.7, yc+0.5, 6.0, 0.35, size=10, italic=True, color=GRAY)
    T(s, body, 6.95, yc+0.15, 5.65, 0.7, size=10.5, color=DARK, ls=1.3)
    yc += 1.05

# B2: Сценарная матрица с обоснованием
s = add_slide()
backup_chrome(s, 'МАТРИЦА СЦЕНАРИЕВ', 2, BTOTAL,
              'Полная оценка S1, S2, S3 по пяти критериям с обоснованием каждой ячейки')
matrix = [
    ('Критерий', 'S1 (единая логика)', 'S2 (единый режим + адаптация)', 'S3 (диверсификация)'),
    ('1. Ресурсная защитимость',
     'Использует ресурсы, не делает источником преимущества',
     'Опирается на операционный блок, недоиспользует экосистему',
     'Дифференцированно задействует все блоки способностей'),
    ('2. Динамические способности',
     'Не требует различения рынков — способности не задействованы',
     'Адаптация внутри одного режима, не на портфельном уровне',
     'Требует и тренирует воспроизводимую процедуру различения'),
    ('3. Операционная исполнимость',
     'Уже работает (инерция), но эффективность снижается',
     'Высокая (один режим везде, ниже требования к координации)',
     'Средняя — требует поддержки нескольких режимов параллельно'),
    ('4. Финансовый эффект',
     'Стагнация выручки/самокат, ухудшение в жёсткой регуляторике',
     'Удержание маржи, упущенный рост в R2-городах',
     'Максимизация отдачи в условиях неоднородности'),
    ('5. Профиль риска',
     'Концентрация: ужесточение в нескольких городах = удар по портфелю',
     'Концентр. в R3: невозможность атаковать при возможностях',
     'Распределён между режимами; добавляется риск координации'),
]
y_m = 1.85; rw = [3.0, 2.9, 3.2, 3.2]; mrh = 0.78
for r_i, row in enumerate(matrix):
    y = y_m + r_i*mrh
    is_h = (r_i == 0)
    fill = DARK if is_h else (WHITE if r_i % 2 == 1 else VLIGHT)
    CARD(s, 0.5, y, sum(rw), mrh, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else DARK
        if not is_h and c_i == 3:
            color = RED
        bold = is_h or c_i == 0
        size = 11 if is_h else (10 if not is_h else 11)
        T(s, cell, x+0.1, y+0.1, rw[c_i]-0.2, mrh-0.2,
          size=size, bold=bold, color=color, ls=1.2)
        x += rw[c_i]

# B3: Финмодель детально
s = add_slide()
backup_chrome(s, 'ФИНМОДЕЛЬ — ДЕТАЛИ', 3, BTOTAL,
              'Финансовая модель: разбивка затрат и эффектов, источник каждого числа')

T(s, 'Затраты года 1', 0.5, 1.85, 5.5, 0.4, size=13, bold=True, color=DARK)
cost = [
    ('Блок',                            'Млн ₽', 'Источник'),
    ('A. Региональное управление',      '−22',   '3 менеджера + 3 аналитика'),
    ('B. Маркетинг S3',                 '−7',    'A/B-тесты, креатив по 3 режимам'),
    ('C. УПС (вкл. capex 8)',           '−16',   'IT + команда пилота + стимулы'),
    ('D. Накладные',                    '−3',    'обучение, методич. сопровождение'),
    ('Итого',                           '−48',   ''),
]
yc = 2.3; row_h = 0.5
for r_i, row in enumerate(cost):
    is_h = (r_i == 0); is_t = (r_i == len(cost)-1)
    fill = DARK if is_h else (LIGHT if is_t else (WHITE if r_i % 2 == 1 else VLIGHT))
    CARD(s, 0.5, yc, 6.0, row_h, fill=fill)
    cw_c = [3.0, 0.9, 2.1]
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else (RED if c_i == 1 and not is_h else DARK)
        bold = is_h or is_t or c_i == 1
        align = PP_ALIGN.LEFT if c_i in (0, 2) else PP_ALIGN.RIGHT
        size = 10.5 if c_i < 2 else 9.5
        T(s, cell, x+0.1, yc, cw_c[c_i]-0.2, row_h,
          size=size, bold=bold, color=color, align=align, anchor=MSO_ANCHOR.MIDDLE)
        x += cw_c[c_i]
    yc += row_h

T(s, 'Эффекты года 1 (EBITDA)', 6.83, 1.85, 5.5, 0.4, size=13, bold=True, color=DARK)
eff = [
    ('Эффект',                          'Млн ₽', 'Расчёт'),
    ('1. R2-прирост выручки',           '+12',   '+3 п.п. × 1 544 × 25 % маржа'),
    ('2. УПС-пилот Екб',                '+6',    'экономия 3,5 + выручка 8,2 × 25 %'),
    ('3. R3-Москва OpEx',               '+52',   '−2 % от выручки Москвы 2 626 млн'),
    ('4. Малые города',                 '+3',    'отказ от убыточных локаций'),
    ('Итого',                           '+73',   ''),
]
yc = 2.3
for r_i, row in enumerate(eff):
    is_h = (r_i == 0); is_t = (r_i == len(eff)-1)
    fill = DARK if is_h else (LIGHT if is_t else (WHITE if r_i % 2 == 1 else VLIGHT))
    CARD(s, 6.83, yc, 6.0, row_h, fill=fill)
    cw_c = [2.4, 0.9, 2.7]
    x = 6.83
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else (GREEN_TXT if c_i == 1 and not is_h else DARK)
        bold = is_h or is_t or c_i == 1
        align = PP_ALIGN.LEFT if c_i in (0, 2) else PP_ALIGN.RIGHT
        size = 10.5 if c_i < 2 else 9.5
        T(s, cell, x+0.1, yc, cw_c[c_i]-0.2, row_h,
          size=size, bold=bold, color=color, align=align, anchor=MSO_ANCHOR.MIDDLE)
        x += cw_c[c_i]
    yc += row_h

# Net + sensitivity
CARD(s, 0.5, 5.55, 12.33, 0.85, fill=RED)
T(s, 'Чистый эффект год 1', 0.7, 5.65, 4, 0.3, size=11, color=WHITE)
T(s, '+25 млн ₽', 0.7, 5.92, 4, 0.45, size=22, bold=True, color=WHITE)
T(s, 'Чувствительность', 5.5, 5.65, 4, 0.3, size=11, color=WHITE)
T(s, '−14  /  +25  /  +64', 5.5, 5.92, 4, 0.45, size=22, bold=True, color=WHITE)
T(s, 'Год 2 при тиражировании УПС', 9.5, 5.65, 4, 0.3, size=11, color=WHITE)
T(s, '+80…+100 млн ₽', 9.5, 5.92, 4, 0.45, size=20, bold=True, color=WHITE)

# NPV/Payback/ROI
CARD(s, 0.5, 6.55, 4.0, 0.7, fill=LIGHT)
T(s, 'NPV  (WACC 20 %, 3 года)', 0.7, 6.6, 3.7, 0.25, size=10, color=GRAY)
T(s, '+134 млн ₽', 0.7, 6.85, 3.7, 0.35, size=16, bold=True, color=RED)
CARD(s, 4.65, 6.55, 4.0, 0.7, fill=LIGHT)
T(s, 'Payback (простой)', 4.85, 6.6, 3.7, 0.25, size=10, color=GRAY)
T(s, '≈ 8 месяцев', 4.85, 6.85, 3.7, 0.35, size=16, bold=True, color=RED)
CARD(s, 8.8, 6.55, 4.0, 0.7, fill=LIGHT)
T(s, 'ROI год 1 / год 2', 9.0, 6.6, 3.7, 0.25, size=10, color=GRAY)
T(s, '52 %  /  243 %', 9.0, 6.85, 3.7, 0.35, size=16, bold=True, color=RED)

# B4: УПС расчётная цепочка
s = add_slide()
backup_chrome(s, 'УПС — РАСЧЁТНАЯ ЦЕПОЧКА', 4, BTOTAL,
              'УПС — расчёт по Екатеринбургу (базовый сценарий B) и 4 KPI пилота')

T(s, 'Расчётная цепочка', 0.5, 1.85, 12.3, 0.35, size=13, bold=True, color=DARK)
chain = [
    ('Выручка сезона',     '136 млн ₽',  '1 819 014 поездок × 75 ₽ средний чек'),
    ('OpEx сезона',        '75 млн ₽',   'выручка × 55 % OpEx-доля'),
    ('Ребалансировка',     '22,5 млн ₽', 'OpEx × 30 % доля ребалансировки'),
    ('Операций ребаланс.', '187,6 тыс.', '22 500 000 / 120 ₽ за операцию'),
    ('Замещение (20 %)',   '37,5 тыс.',  'базовый сценарий B'),
    ('Чистая экономия',    '3,5 млн ₽',  '(120 − 26,5) × 37,5 тыс.'),
    ('Доп. выручка',       '8,2 млн ₽',  '1 819 014 × 6 % × 75 ₽'),
    ('EBITDA-вклад',       '≈ 6 млн ₽',  '3,5 + 8,2 × 25 % маржа'),
]
yc = 2.3; row_h = 0.42
for i, (lbl, val, src) in enumerate(chain):
    is_final = (i == len(chain)-1)
    fill = RED if is_final else (WHITE if i % 2 == 0 else VLIGHT)
    CARD(s, 0.5, yc, 12.3, row_h, fill=fill)
    color = WHITE if is_final else DARK
    T(s, lbl, 0.7, yc, 3.5, row_h, size=11, bold=True, color=color, anchor=MSO_ANCHOR.MIDDLE)
    T(s, val, 4.4, yc, 2.5, row_h, size=13, bold=True,
      color=(WHITE if is_final else RED), align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    T(s, src, 7.2, yc, 5, row_h, size=10,
      color=(WHITE if is_final else GRAY), align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    yc += row_h

# 4 KPI
T(s, '4 KPI пилота', 0.5, 6.0, 12.3, 0.35, size=13, bold=True, color=DARK)
kpis = [
    ('Прирост поездок',         'в пилотных зонах vs база'),
    ('Снижение ребалансировки', 'доля в общем объёме операций'),
    ('Стоимость стимула',       'на 1 предотвращённую потерю'),
    ('Отклик пользователей',    'доля принятых стимулов'),
]
y_k = 6.45; kw = 3.0; kgap = 0.1
for i, (h, b) in enumerate(kpis):
    x = 0.5 + i*(kw + kgap)
    CARD(s, x, y_k, kw, 0.6, fill=LIGHT)
    T(s, h, x+0.15, y_k+0.05, kw-0.3, 0.3, size=11, bold=True, color=DARK)
    T(s, b, x+0.15, y_k+0.32, kw-0.3, 0.25, size=9.5, color=GRAY)

# B5: Декларация ИИ
s = add_slide()
backup_chrome(s, 'ДЕКЛАРАЦИЯ ИИ', 5, BTOTAL,
              'Декларация использования ИИ — что и где применялось при подготовке ВКР')

T(s, 'Инструменты: ChatGPT (OpenAI, GPT-4o, GPT-5) и Claude (Anthropic, Sonnet 4, Opus 4.7) через стандартные веб-интерфейсы и Cursor IDE.',
  0.5, 1.85, 12.3, 0.5, size=12, color=DARK, ls=1.3)

CARD(s, 0.5, 2.55, 6.1, 4.3, fill=LIGHT)
CARD(s, 0.5, 2.55, 6.1, 0.55, fill=GREEN_TXT)
T(s, '✓   ГДЕ ИСПОЛЬЗОВАЛСЯ ИИ', 0.5, 2.62, 6.1, 0.4, size=12, bold=True, color=WHITE,
  align=PP_ALIGN.CENTER)
ai_use = [
    ('Поиск и оформление литературы',     'фильтрация источников, ГОСТ-форматирование'),
    ('Структурирование черновиков',        'формулировки задач, структура глав'),
    ('Редактура и стилистика',             'академический стиль, грамматика'),
    ('Перевод аннотации',                  'русский → английский'),
    ('Автоматизация расчётов',             'Python для CSV и Excel-модели'),
    ('Сборка финального .docx и .xlsx',    'форматирование, проверка консистентности'),
]
y = 3.25
for h, b in ai_use:
    T(s, '✓  ' + h, 0.75, y, 5.6, 0.3, size=11, bold=True, color=DARK)
    T(s, b, 1.05, y+0.3, 5.3, 0.25, size=10, color=GRAY)
    y += 0.6

CARD(s, 6.75, 2.55, 6.1, 4.3, fill=LIGHT)
CARD(s, 6.75, 2.55, 6.1, 0.55, fill=RED_TXT)
T(s, '✗   ГДЕ ИИ НЕ ПРИМЕНЯЛСЯ', 6.75, 2.62, 6.1, 0.4, size=12, bold=True, color=WHITE,
  align=PP_ALIGN.CENTER)
ai_no = [
    ('Постановка проблемы и цели',         'авторская формулировка'),
    ('Интерпретация интервью с экспертами','обобщённые тезисы передавались, не расшифровки'),
    ('Выбор сценария S3 и режимов R1–R3', 'авторские стратегические решения'),
    ('Концепция и параметры УПС',           'авторская разработка'),
    ('Допущения и параметры финмодели',     'выбор автора, обоснованы в тексте'),
    ('Внутренние данные МТС Юрент',         'только в обезличенном агрегированном виде'),
]
y = 3.25
for h, b in ai_no:
    T(s, '✗  ' + h, 7.0, y, 5.6, 0.3, size=11, bold=True, color=DARK)
    T(s, b, 7.3, y+0.3, 5.3, 0.25, size=10, color=GRAY)
    y += 0.6

T(s, 'Все сгенерированные с участием ИИ фрагменты проверены и отредактированы автором. Ответственность за корректность и интерпретацию — автора ВКР.',
  0.5, 7.0, 12.3, 0.3, size=10, italic=True, color=GRAY, align=PP_ALIGN.CENTER)


prs.save('/workspace/vkr/VKR_presentation.pptx')
print(f'Saved VKR_presentation.pptx ({len(prs.slides)} slides)')
