"""Сборка защитной презентации МТС Юрент (16 слайдов + 5 backup).

Принципы оформления:
- 16:9, белый фон, тонкая красная полоса сверху
- Два цвета: МТС red (#ED0028) + dark gray (#1A1A1A) + светло-серый для фона блоков
- Один шрифт (Calibri): bold для заголовков, regular для тела
- На каждом слайде доминирующий визуальный элемент (цифра / схема / таблица)
- Спикер-ноуты с тезисами того, что говорить
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from copy import deepcopy

# === Цвета ===
RED = RGBColor(0xED, 0x00, 0x28)        # МТС-style accent
DARK = RGBColor(0x1A, 0x1A, 0x1A)       # основной текст
GRAY = RGBColor(0x59, 0x59, 0x59)       # вторичный текст
LIGHT = RGBColor(0xF2, 0xF2, 0xF2)      # фон карточек
VLIGHT = RGBColor(0xFA, 0xFA, 0xFA)     # фон секций
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN_TXT = RGBColor(0x1E, 0x7A, 0x3C)  # для положительных эффектов
RED_TXT = RGBColor(0xC0, 0x1A, 0x1A)    # для затрат/минусов

FONT = 'Calibri'

# === Презентация 16:9 ===
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

# ============================== ХЕЛПЕРЫ ==============================

def add_slide():
    s = prs.slides.add_slide(BLANK)
    # Top red bar
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(0.18))
    bar.fill.solid(); bar.fill.fore_color.rgb = RED
    bar.line.fill.background()
    return s

def add_footer(s, num, total):
    # Project name left
    tb = s.shapes.add_textbox(Inches(0.4), Inches(7.05), Inches(8), Inches(0.35))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = 'МТС Юрент · Конкурентная стратегия на региональных рынках кикшеринга'
    r.font.name = FONT; r.font.size = Pt(9); r.font.color.rgb = GRAY
    # Page number right
    tb2 = s.shapes.add_textbox(Inches(12.3), Inches(7.05), Inches(0.9), Inches(0.35))
    tf2 = tb2.text_frame
    p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run(); r2.text = f'{num}/{total}'
    r2.font.name = FONT; r2.font.size = Pt(9); r2.font.color.rgb = GRAY

def add_title(s, text, top=0.45, left=0.5, width=12.3, height=1.0, size=24):
    tb = s.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = text
    r.font.name = FONT; r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = DARK
    p.line_spacing = 1.1
    return tb

def add_text(s, text, left, top, width, height, size=14, bold=False, italic=False, color=DARK, align=PP_ALIGN.LEFT, ls=1.2):
    tb = s.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(2)
    lines = text.split('\n') if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = ls
        r = p.add_run(); r.text = line
        r.font.name = FONT; r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic; r.font.color.rgb = color
    return tb

def add_card(s, left, top, width, height, fill=LIGHT, line=None, radius=0):
    shape = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid(); shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line; shape.line.width = Pt(0.5)
    if radius:
        # adjust corner radius via XML hack
        try:
            shape.adjustments[0] = 0.05
        except Exception:
            pass
    shape.shadow.inherit = False
    return shape

def add_bignum(s, text, left, top, width, height, size=54, color=RED, align=PP_ALIGN.CENTER):
    return add_text(s, text, left, top, width, height, size=size, bold=True, color=color, align=align, ls=1.0)

def add_label(s, text, left, top, width, height, size=11, color=GRAY, align=PP_ALIGN.CENTER):
    return add_text(s, text, left, top, width, height, size=size, color=color, align=align, ls=1.1)

def add_arrow(s, x1, y1, x2, y2, color=RED, weight=1.5):
    line = s.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    line.line.color.rgb = color; line.line.width = Pt(weight)
    # Add arrowhead via XML
    line_xml = line.line._get_or_add_ln()
    tail = line_xml.find(qn('a:tailEnd'))
    if tail is None:
        from lxml import etree
        tail = etree.SubElement(line_xml, qn('a:tailEnd'))
    tail.set('type', 'triangle'); tail.set('w', 'med'); tail.set('len', 'med')
    return line

def add_subtitle(s, text, top=1.5, left=0.5, width=12.3, size=14, color=GRAY):
    return add_text(s, text, left, top, width, 0.5, size=size, color=color, ls=1.2)

def set_notes(s, text):
    s.notes_slide.notes_text_frame.text = text

def add_punchline(s, text):
    """Однострочная подсказка-итог внизу слайда (над футером).
    Делает структуру 'заголовок-вывод сверху + доказательство в теле + усилитель снизу'."""
    add_card(s, 0.5, 6.65, 12.33, 0.4, fill=DARK)
    add_text(s, text, 0.5, 6.7, 12.33, 0.3, size=11, bold=True,
             color=WHITE, align=PP_ALIGN.CENTER)

TOTAL = 16  # основных слайдов, бэкап нумеруется отдельно

# ============================== СЛАЙД 1: ТИТУЛЬНЫЙ ==============================
s = add_slide()
# Большой акцентный блок слева
add_card(s, 0, 0.18, 5.5, 7.32, fill=RED)
add_text(s, 'ВКР', 0.5, 0.6, 4.5, 0.6, size=18, bold=True, color=WHITE)
add_text(s, 'Управление\nбизнесом', 0.5, 5.8, 4.5, 1.2, size=14, color=WHITE)
add_text(s, 'Москва · 2026', 0.5, 6.6, 4.5, 0.4, size=12, color=WHITE)

# Title
add_text(s, 'Конкурентная стратегия\nМТС Юрент на региональных\nрынках кикшеринга в России',
         6.0, 0.9, 7.0, 2.5, size=28, bold=True, color=DARK, ls=1.15)

# Подзаголовок-вопрос (из v3)
add_text(s, 'Как МТС Юренту перейти от роста через парк к управлению отдачей регионального портфеля?',
         6.0, 3.6, 7.0, 1.0, size=15, italic=False, color=GRAY, ls=1.3)

# Логическая лента (из v3, доработана)
flow = ['Диагноз города', 'Выбор режима', 'Рычаги в сезоне', 'KPI', 'Эффект']
fy = 5.0; fw = 1.3; fgap = 0.05
fx = 6.0
for i, step in enumerate(flow):
    add_card(s, fx, fy, fw, 0.5, fill=LIGHT)
    add_text(s, step, fx, fy+0.13, fw, 0.3, size=9, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    fx += fw + fgap
    if i < len(flow) - 1:
        # Arrow gap is just spacing
        pass

# Author + supervisor
add_text(s, 'Автор', 6.0, 5.8, 3.5, 0.3, size=11, color=GRAY)
add_text(s, 'Камашев Пётр', 6.0, 6.1, 3.5, 0.4, size=15, bold=True, color=DARK)
add_text(s, 'Научный руководитель', 9.5, 5.8, 4, 0.3, size=11, color=GRAY)
add_text(s, 'профессор И. А. Егоров', 9.5, 6.1, 4, 0.4, size=15, bold=True, color=DARK)

set_notes(s, 'Стартовая фраза (учить наизусть):\n\n'
              '«Добрый день, уважаемые члены государственной экзаменационной комиссии. '
              'Меня зовут Пётр Камашев, я представляю выпускную квалификационную работу на тему '
              '"Конкурентная стратегия МТС Юрент на региональных рынках кикшеринга в России". '
              'Научный руководитель — профессор Иван Александрович Егоров.\n\n'
              'Центральный вопрос работы — как МТС Юренту перейти от роста через расширение парка '
              'к управлению отдачей уже сформированного регионального портфеля. '
              'Работа выстроена как цепочка из пяти шагов: диагноз города, выбор режима, рычаги в сезоне, KPI и эффект».')

# ============================== СЛАЙД 2: УПРАВЛЕНЧЕСКАЯ ПРОБЛЕМА ==============================
s = add_slide()
add_title(s, 'Рынок впервые сжался, а регионы расходятся — единая логика МТС Юрент по 187 локациям перестаёт работать', size=22)

# 3 big stat cards
y_top = 1.9
card_w = 4.05; gap = 0.15

for i, (num, num_sub, label) in enumerate([
    ('−6 %', 'поездок', 'Рынок РФ в 2025 — впервые отрицательная динамика после 5 лет роста'),
    ('42 → 28 %', 'EBITDA-маржа', 'Лидер Whoosh: выручка −13 %, маржа кикшеринга обвалилась'),
    ('4×', 'разрыв', 'Утилизация парка между городами Юрента: Москва 1 073 vs Новосибирск 277 поездок/самокат/год'),
]):
    x = 0.5 + i*(card_w + gap)
    add_card(s, x, y_top, card_w, 3.0, fill=LIGHT)
    add_bignum(s, num, x, y_top+0.3, card_w, 1.2, size=56, color=RED)
    add_label(s, num_sub, x, y_top+1.55, card_w, 0.3, size=14, color=DARK)
    add_text(s, label, x+0.3, y_top+2.0, card_w-0.6, 1.0, size=11.5, color=GRAY, align=PP_ALIGN.CENTER, ls=1.25)

# Низ: объект + проблема
add_card(s, 0.5, 5.3, 12.3, 1.55, fill=VLIGHT)
add_text(s, 'Объект исследования', 0.85, 5.45, 4, 0.3, size=11, color=GRAY)
add_text(s, 'МТС Юрент — оператор кикшеринга на портфеле региональных рынков РФ',
         0.85, 5.72, 11.5, 0.5, size=14, bold=True, color=DARK)
add_text(s, 'Управленческая проблема', 0.85, 6.15, 4, 0.3, size=11, color=GRAY)
add_text(s, 'Формализованной процедуры выбора режима присутствия по регионам нет — решения принимаются как набор разовых кейсов',
         0.85, 6.42, 11.5, 0.5, size=14, bold=True, color=RED_TXT)

add_footer(s, 2, TOTAL)
set_notes(s, 'Ключевая мысль: текущая логика расширения исчерпала источник нового эффекта. '
              '«Объект исследования — МТС Юрент. Управленческая проблема состоит в том, что российский рынок кикшеринга в 2025 году '
              'впервые показал отрицательную динамику — поездки сократились на 6 %, выручка лидера Whoosh упала на 13 %, '
              'а EBITDA-маржа кикшеринга снизилась с 42 до 28 %. На этом фоне региональные рынки расходятся по структурным условиям. '
              'Единая логика присутствия в портфеле теряет состоятельность, а формализованной процедуры выбора режима по городам '
              'у компании сегодня нет».')

# ============================== СЛАЙД 3: ЦЕЛЬ И ЛОГИКА ==============================
s = add_slide()
add_title(s, 'Цель — построить систему стратегического выбора и применить её к МТС Юрент', size=22)
add_subtitle(s,
             'Разработать аналитическую систему стратегического выбора оператора кикшеринга '
             'в портфеле региональных рынков и применить её к МТС Юрент',
             top=1.45, size=14, color=GRAY)

# Workflow 6 blocks with arrows
labels = [
    ('Диагностика\nрынка', 'PESTEL,\nконкурентный анализ'),
    ('5 критериев\nоценки', 'из теоретико-\nметодологической\nосновы'),
    ('3 сценария\nS1–S3', 'сопоставление\nальтернатив'),
    ('Выбор\nS3', 'диверсификация\nрежимов'),
    ('Режимы R1–R3\n+ УПС', 'проектное\nнаполнение'),
    ('Финансовый\nэффект', 'демо-портфель\n5 городов'),
]
n = len(labels); block_w = 1.85; arrow_gap = 0.12
total_w = n*block_w + (n-1)*arrow_gap
start_x = (13.333 - total_w) / 2
y = 2.5

for i, (top, bot) in enumerate(labels):
    x = start_x + i*(block_w + arrow_gap)
    fill = RED if i == 3 else LIGHT
    txtcolor = WHITE if i == 3 else DARK
    subcolor = WHITE if i == 3 else GRAY
    add_card(s, x, y, block_w, 1.7, fill=fill, radius=1)
    add_text(s, top, x+0.05, y+0.25, block_w-0.1, 0.7, size=13, bold=True, color=txtcolor, align=PP_ALIGN.CENTER, ls=1.1)
    add_text(s, bot, x+0.05, y+1.05, block_w-0.1, 0.6, size=10, color=subcolor, align=PP_ALIGN.CENTER, ls=1.1)
    if i < n-1:
        add_arrow(s, x+block_w, y+0.85, x+block_w+arrow_gap, y+0.85, color=GRAY, weight=1.2)

# 7 задач
add_text(s, '7 задач работы', 0.5, 4.85, 12.3, 0.4, size=12, bold=True, color=DARK)
tasks = ('1. Систематизация литературы и обоснование критериев · '
         '2. Диагностика российского рынка кикшеринга 2022–2025 · '
         '3. Иерархия факторов регионального различения\n'
         '4. Каркас режимов R1–R3 · '
         '5. Сопоставление сценариев S1–S3 · '
         '6. Проектный инструмент УПС · '
         '7. Финансовая модель сценария S3')
add_text(s, tasks, 0.5, 5.25, 12.3, 1.3, size=11.5, color=GRAY, ls=1.4)

add_punchline(s, 'Выход работы — система S3, каркас режимов R1–R3, инструмент УПС и финансовая модель на демонстрационном портфеле')
add_footer(s, 3, TOTAL)
set_notes(s, 'Ключевая мысль: вся работа выстроена как линейная цепочка от диагностики к измеримому эффекту. '
              '«Цель работы — разработать аналитическую систему стратегического выбора и применить её к МТС Юрент. '
              'Логика работы: внешняя среда → внутренняя диагностика → стратегические альтернативы → выбор сценария '
              '→ проектное наполнение → оценка эффекта. За следующие тринадцать минут я покажу, как семь задач складываются в одно решение».')

# ============================== СЛАЙД 4: МЕТОДОЛОГИЯ ==============================
s = add_slide()
add_title(s, 'Методология: 5 теоретических линий и закрытые данные по 86 субъектам РФ', size=22)

# 3 columns
cols = [
    ('Методы', [
        'PESTEL-анализ',
        'Конкурентный анализ',
        'Иерархия факторов (3 уровня)',
        'Критериальное сопоставление',
        'Финансовое моделирование',
        'Сценарный анализ',
    ]),
    ('Источники', [
        'Отчётность ПАО «МТС», ПАО «ВУШ Холдинг»',
        'Отраслевая аналитика (TrueSharing, NACTO)',
        'Деловая пресса (Коммерсантъ, Эксперт)',
        '37 академических источников',
        'Эксперт А. — M&A МТС',
        'Эксперт Б. — научный руководитель',
    ]),
    ('Эмпирика', [
        '86 субъектов РФ — внутренние\nпоквартальные данные МТС Юрент',
        'Период: 2024 + 9 мес. 2025',
        'Демонстрационный портфель:\n5 городов = 94 % выручки оператора',
        'Финмодель сценария S3:\n5 эффектов × 3 сценария чувствительности',
    ]),
]
col_w = 4.0; col_gap = 0.15; y0 = 1.7
for i, (head, items) in enumerate(cols):
    x = 0.5 + i*(col_w + col_gap)
    add_card(s, x, y0, col_w, 4.7, fill=LIGHT)
    # header
    add_card(s, x, y0, col_w, 0.55, fill=RED)
    add_text(s, head, x, y0+0.07, col_w, 0.45, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # items
    body = '\n'.join(f'•  {it}' for it in items)
    add_text(s, body, x+0.25, y0+0.7, col_w-0.4, 4.0, size=11.5, color=DARK, ls=1.35)

# Bottom: AI note
add_text(s, 'Использование ИИ: ChatGPT и Claude — обработка источников, форматирование, редактура. '
            'Стратегические решения, выбор сценария, интерпретация данных и параметры финмодели — авторские. '
            'Декларация — Приложение А ВКР.',
         0.5, 6.27, 12.3, 0.35, size=10, color=GRAY, ls=1.2)

add_punchline(s, 'База достаточна для предварительного выбора стратегии и дизайна пилота; апробация — этап пилота')
add_footer(s, 4, TOTAL)
set_notes(s, 'Ключевая мысль: выводы опираются на сочетание академической рамки, экспертной перспективы и внутренней операционной отчётности. '
              '«Теоретическая основа — пять линий: ресурсная теория фирмы, концепция динамических способностей, теория платформенной конкуренции, '
              'эмпирические исследования рынков шеринговой микромобильности и системная экономика Клейнера. '
              'Главный закрытый источник — внутренние поквартальные данные МТС Юрент по 86 субъектам РФ за 2024 и девять месяцев 2025 года. '
              'Отдельно отмечу: для обработки источников, форматирования и редактуры я использовал ChatGPT и Claude; '
              'стратегические решения, интерпретация и параметры финмодели — авторские; декларация — в Приложении А работы».')

# ============================== СЛАЙД 5: ВНЕШНЯЯ СРЕДА ==============================
s = add_slide()
add_title(s, 'Внешняя среда: рынок сжался впервые за 5 лет, сезон сократился, капитал подорожал', size=22)

# 4 cards 2×2 with PESTEL summary
pestel = [
    ('Рынок', '−6 %', 'поездок в 2025 (Платформа ОФД)\nWhoosh: парк +17 %, поездки −7 %\nвыручка −13 %, маржа 42 → 28 %'),
    ('Капитал', '21 %', 'ключевая ставка ЦБ в 2025\nсрок службы парка 2–3 сезона\nстоимость единицы 60–90 тыс. руб.'),
    ('Регуляторика', '3 режима', 'Москва: верификация mos.ru\nСПб: соглашения с операторами\nЕкб: угрозы расторжения'),
    ('Сезон', 'март → июнь', 'сезон 2025 начался на 3 мес позже\nпотери первого квартала\nрост чувствительности к погоде'),
]
cw, ch, gap = 6.0, 2.35, 0.2
positions = [(0.7, 1.85), (6.9, 1.85), (0.7, 4.4), (6.9, 4.4)]
for (title, num, body), (x, y) in zip(pestel, positions):
    add_card(s, x, y, cw, ch, fill=LIGHT)
    add_text(s, title, x+0.3, y+0.2, 2.5, 0.4, size=12, color=GRAY)
    add_bignum(s, num, x+0.3, y+0.55, 2.7, 0.9, size=32, color=RED, align=PP_ALIGN.LEFT)
    add_text(s, body, x+3.3, y+0.25, cw-3.5, ch-0.4, size=11.5, color=DARK, ls=1.35)

add_punchline(s, 'Прежняя логика расширения парка и географии перестаёт быть достаточным основанием конкурентной борьбы')
add_footer(s, 5, TOTAL)
set_notes(s, 'Ключевая мысль: экстенсивная модель расширения не выдерживает экономики покрытия в новых условиях. '
              '«На объект одновременно воздействуют четыре фактора. Рынок: поездки на рынке снизились на 6 %, у лидера Whoosh — на 7 %, '
              'маржа кикшеринга упала с 42 до 28 %. Капитал: ключевая ставка ЦБ 21 % существенно повышает стоимость финансирования парка '
              'со сроком службы два-три сезона. Регуляторика: режимы расходятся от mos.ru-верификации в Москве до угроз расторжения в Екатеринбурге. '
              'Сезон 2025 сжался — начался не в марте, а в июне. Совокупный эффект: прежняя логика расширения парка и географии перестаёт быть '
              'достаточным основанием конкурентной борьбы».')

# ============================== СЛАЙД 6: РЕГИОНАЛЬНАЯ НЕОДНОРОДНОСТЬ ==============================
s = add_slide()
add_title(s, 'Региональная неоднородность: пять рынков расходятся по структурным условиям', size=22)
add_subtitle(s, 'Демонстрационный портфель = 94 % выручки МТС Юрент в России. '
                'Параметры I и II уровней дают принципиально разные сочетания на реальной географии.',
             top=1.45, size=13, color=GRAY)

# Table 5×5 (custom)
headers = ['Город', 'Утилизация\nпоездок/самокат/год', 'YoY поездок\n9 мес. 2025', 'Регуляторика', 'Гипотеза режима']
rows = [
    ('Москва',          '1 073', '−6,6 %',   'жёсткая',           'R3 (огранич.)'),
    ('Санкт-Петербург', '609',   '+42,3 %',  'умеренно жёсткая',  'R2 → R3'),
    ('Краснодар',       '564',   '−10,2 %',  'умеренная',         'R2 / R3'),
    ('Екатеринбург',    '319',   '+4,1 %',   'умеренная',         'R2 + пилот УПС'),
    ('Новосибирск',     '277',   '+128,9 %', 'умеренная',         'R2/R3 (эффект.)'),
]

table_left = 0.7; table_top = 2.2; table_w = 11.95
col_widths = [2.5, 2.3, 2.2, 2.45, 2.5]

# Header
x = table_left
h_h = 0.6
add_card(s, table_left, table_top, table_w, h_h, fill=DARK)
for i, h in enumerate(headers):
    add_text(s, h, x+0.1, table_top+0.05, col_widths[i]-0.2, h_h-0.1,
             size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, ls=1.1)
    x += col_widths[i]

# Rows
row_h = 0.7
for r_i, row in enumerate(rows):
    y = table_top + h_h + r_i*row_h
    fill_color = WHITE if r_i % 2 == 0 else VLIGHT
    add_card(s, table_left, y, table_w, row_h, fill=fill_color)
    x = table_left
    for c_i, cell in enumerate(row):
        bold = (c_i == 0) or (c_i == 4)
        color = DARK
        if c_i == 4:
            color = RED
        align = PP_ALIGN.LEFT if c_i in (0, 3, 4) else PP_ALIGN.CENTER
        if c_i == 2 and cell.startswith('+'):
            color = GREEN_TXT; bold = True
        elif c_i == 2 and cell.startswith('−'):
            color = RED_TXT; bold = True
        add_text(s, cell, x+0.15, y+0.18, col_widths[c_i]-0.3, row_h-0.2,
                 size=12, bold=bold, color=color, align=align, ls=1.1)
        x += col_widths[c_i]

# Footnote
add_text(s, 'Источник: внутренние поквартальные данные МТС Юрент 2024 / 9 мес. 2025 (эксперт А.). Позиция МТС Юрент во всех 5 городах — преследователь.',
         0.7, 6.27, 12, 0.3, size=10, color=GRAY)

add_punchline(s, 'Параметры расходятся в 4× по утилизации и в широком диапазоне по динамике — единый режим по портфелю невозможен')
add_footer(s, 6, TOTAL)
set_notes(s, 'Ключевая мысль: единая стратегия по такому портфелю физически невозможна — параметры расходятся в разные стороны. '
              '«На демонстрационном портфеле из пяти городов разрыв утилизации между крайними точками — почти в четыре раза, '
              'динамика YoY от −10 до +129 %, регуляторика от мягкой до жёсткой. На той же географии и предварительная гипотеза режима '
              'оказывается разной — это и есть основание для портфельной логики, которую я разворачиваю дальше».')

# ============================== СЛАЙД 7: КОНКУРЕНТНАЯ ПОЗИЦИЯ ==============================
s = add_slide()
add_title(s, 'Конкурентная позиция: КФУ на зрелой стадии — доступность покрытия, у Юрента системный гэп', size=22)

# 3 operator cards
ops = [
    ('Whoosh', 'лидер', '150 тыс.', 'масштаб парка +\nвертикальная интеграция', WHITE),
    ('МТС Юрент', 'преследователь', '120 тыс.', 'экосистема МТС +\nплатформа Eleven', WHITE),
    ('Яндекс', 'третий', '30 тыс.', 'трафик\nсуперприложения', WHITE),
]
y_op = 1.85; op_w = 3.6; op_h = 1.8; gap = 0.15
total_op = 3*op_w + 2*gap
start_op = (13.333 - total_op) / 2
for i, (name, role, fleet, logic, fill) in enumerate(ops):
    x = start_op + i*(op_w + gap)
    is_mts = (name == 'МТС Юрент')
    bg = RED if is_mts else LIGHT
    txt = WHITE if is_mts else DARK
    sub = WHITE if is_mts else GRAY
    add_card(s, x, y_op, op_w, op_h, fill=bg)
    add_text(s, name, x+0.2, y_op+0.15, op_w-0.4, 0.4, size=18, bold=True, color=txt)
    add_text(s, role, x+0.2, y_op+0.6, op_w-0.4, 0.3, size=11, color=sub)
    add_text(s, f'{fleet} парк', x+0.2, y_op+0.95, op_w-0.4, 0.35, size=13, bold=True, color=txt)
    add_text(s, logic, x+0.2, y_op+1.32, op_w-0.4, 0.45, size=10.5, color=sub, ls=1.2)

# Big number — гэп
add_card(s, 0.7, 4.0, 6.0, 2.5, fill=VLIGHT)
add_text(s, 'Гэп с лидером — в эффективности, не в размере парка',
         0.95, 4.1, 5.5, 0.35, size=12, color=GRAY)
add_bignum(s, '49  vs  67', 0.95, 4.45, 5.5, 1.3, size=58, color=RED, align=PP_ALIGN.LEFT)
add_text(s, 'тыс. руб. — выручка на 1 самокат в год · МТС Юрент против Whoosh',
         0.95, 5.85, 5.5, 0.5, size=11.5, color=DARK, ls=1.2)

# КФУ block
add_card(s, 6.95, 4.0, 5.75, 2.5, fill=LIGHT)
add_text(s, 'Ключевые факторы успеха на зрелой стадии', 7.2, 4.1, 5.5, 0.35, size=12, bold=True, color=DARK)
add_text(s, '1.  Фактическая доступность самоката в нужной точке и в нужное время',
         7.2, 4.45, 5.4, 0.55, size=12, color=DARK, ls=1.2)
add_text(s, '2.  Экономика на единицу парка', 7.2, 5.05, 5.4, 0.3, size=11.5, color=GRAY)
add_text(s, '3.  Операционная исполнимость', 7.2, 5.4, 5.4, 0.3, size=11.5, color=GRAY)
add_text(s, '→ доступность покрытия — критический КФУ; к нему привязан УПС (Слайд 13)',
         7.2, 5.8, 5.4, 0.5, size=10.5, bold=True, color=RED, ls=1.2)

add_punchline(s, 'Гэп закрывается не наращиванием парка, а повышением отдачи от уже размещённого парка')
add_footer(s, 7, TOTAL)
set_notes(s, 'Ключевая мысль: гэп с Whoosh возник не из-за меньшего парка, а из-за более низкой эффективности уже размещённого парка. '
              '«Конкурентный анализ показал, что на российском рынке к 2025 году сложилась олигополия трёх операторов с долей 96,8 %. '
              'Игроки различаются стратегической логикой. У МТС Юрент системный гэп: выручка на самокат 49 тысяч рублей в год против 67 тысяч у Whoosh. '
              'В разрезе городов разрыв различается в три раза — от 1 073 поездок на самокат в год в Москве до 277 в Новосибирске. '
              'Закрытие гэпа требует не универсального наращивания парка, а дифференциации действий по типам региональных рынков».')

# ============================== СЛАЙД 8: ВНУТРЕННЯЯ ДИАГНОСТИКА ==============================
s = add_slide()
add_title(s, 'Внутри сильная экосистема МТС, но централизованная управленческая логика для 187 локаций', size=22)

# 2 columns: Strong / Weak
y_int = 1.85; col_w_int = 6.0; col_h = 4.3
add_card(s, 0.7, y_int, col_w_int, col_h, fill=LIGHT)
add_card(s, 0.7, y_int, col_w_int, 0.55, fill=GREEN_TXT)
add_text(s, '+  Сильные стороны', 0.7, y_int+0.08, col_w_int, 0.4, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
strong = [
    ('Экосистема МТС', '80 млн абонентов · MTS ID · Premium · обезличенная геоаналитика'),
    ('Технологическая платформа', 'после приобретения Eleven (декабрь 2024)'),
    ('Операционная база', '120 тыс. СИМ в 187 локациях, отстроенные процедуры'),
    ('Гибкость пилотирования', 'компактная команда, короткий контур решений'),
]
for i, (h, b) in enumerate(strong):
    y = y_int + 0.75 + i*0.85
    add_text(s, h, 0.95, y, col_w_int-0.4, 0.35, size=12.5, bold=True, color=DARK)
    add_text(s, b, 0.95, y+0.35, col_w_int-0.4, 0.45, size=11, color=GRAY, ls=1.2)

add_card(s, 6.95, y_int, col_w_int, col_h, fill=LIGHT)
add_card(s, 6.95, y_int, col_w_int, 0.55, fill=RED_TXT)
add_text(s, '−  Слабые стороны', 6.95, y_int+0.08, col_w_int, 0.4, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
weak = [
    ('Централизованная управленческая логика', 'мало чувствительна к различиям регионов'),
    ('Нет процедуры назначения режимов', 'решения по городам — набор разовых кейсов'),
    ('Ограниченная локальная экспертиза', 'дефицит выделенных функций региональной координации'),
    ('Систематическое отставание по эффективности', 'выручка/самокат и утилизация ниже Whoosh'),
]
for i, (h, b) in enumerate(weak):
    y = y_int + 0.75 + i*0.85
    add_text(s, h, 7.2, y, col_w_int-0.4, 0.35, size=12.5, bold=True, color=DARK)
    add_text(s, b, 7.2, y+0.35, col_w_int-0.4, 0.45, size=11, color=GRAY, ls=1.2)

# Dynamic capabilities note (передвинут выше)
add_text(s, 'Dynamic capabilities: операционный блок развит; способность перестраивать ресурсы под изменения рынка — в начальной зрелости',
         0.5, 6.25, 12.3, 0.3, size=11, italic=True, color=GRAY, align=PP_ALIGN.CENTER)

add_punchline(s, 'Проблема — не в отсутствии ресурсов, а в том, что текущая модель не превращает их в региональный результат')
add_footer(s, 8, TOTAL)
set_notes(s, 'Ключевая мысль: ресурсы есть; проблема — в способе их превращения в результат на неоднородной географии. '
              '«Внутренний анализ показал, что Юрент уже обладает ресурсами для дифференцированной стратегии: экосистема МТС, '
              'технологическая платформа Eleven, парк 120 тыс. СИМ в 187 локациях. Однако управленческая логика остаётся централизованной, '
              'формализованной процедуры назначения режимов городам нет. В терминах динамических способностей — операционный блок развит, '
              'способность воспроизводимо перестраивать ресурсы под изменения рынка находится в начальной зрелости. '
              'Проблема не в отсутствии потенциала, а в том, что текущая модель не превращает имеющиеся ресурсы в устойчивый региональный результат».')

# ============================== СЛАЙД 9: СВОДНЫЙ ДИАГНОЗ ==============================
s = add_slide()
add_title(s, 'Стратегический диагноз: единая модель присутствия исчерпала источник новой ценности', size=22)

# Schema: Текущая ситуация → Причины → Последствия → Требуемый ответ
y_d = 1.85; box_h = 4.7
boxes = [
    ('1', 'Текущая ситуация', 'Принципиально неоднородная портфельная среда\n+ единообразная управленческая логика', LIGHT, DARK),
    ('2', 'Причины разрыва', '• Переход рынка к зрелой стадии\n• Расхождение региональных условий\n• Слабая формализация регионального\n   управления внутри компании', LIGHT, DARK),
    ('3', 'Если не менять', 'Ухудшение экономики покрытия\n+ потеря источника новой ценности\n   от расширения парка и географии', LIGHT, DARK),
    ('4', 'Требуемый ответ', 'Иерархия факторов\nразличения рынков (3 уровня)\n+\n5 критериев оценки\nстратегических сценариев', RED, WHITE),
]
bw = 2.9; bg = 0.2
total_b = 4*bw + 3*bg
start_b = (13.333 - total_b) / 2
for i, (n, h, body, fill, txt) in enumerate(boxes):
    x = start_b + i*(bw + bg)
    add_card(s, x, y_d, bw, box_h, fill=fill)
    # Number circle
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+0.15), Inches(y_d+0.15), Inches(0.45), Inches(0.45))
    circle.fill.solid(); circle.fill.fore_color.rgb = (WHITE if fill == RED else RED)
    circle.line.fill.background()
    tb = circle.text_frame; tb.margin_top = tb.margin_bottom = 0
    pp = tb.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
    rr = pp.add_run(); rr.text = n
    rr.font.name = FONT; rr.font.size = Pt(14); rr.font.bold = True
    rr.font.color.rgb = (RED if fill == RED else WHITE)
    # Heading
    add_text(s, h, x+0.7, y_d+0.18, bw-0.85, 0.5, size=13, bold=True, color=txt)
    # Body
    add_text(s, body, x+0.2, y_d+0.95, bw-0.4, box_h-1.1, size=11.5, color=txt, ls=1.35)
    # arrow
    if i < 3:
        add_arrow(s, x+bw, y_d+box_h/2, x+bw+bg, y_d+box_h/2, color=GRAY, weight=1.5)

add_punchline(s, 'Нужна управляемая система «город → режим → KPI», а не разовые корректировки по каждому случаю')
add_footer(s, 9, TOTAL)
set_notes(s, 'Ключевая мысль: разрыв между неоднородностью рынков и единообразием управленческой логики требует системного, а не локального ответа. '
              '«На основании внешнего и внутреннего анализа стратегический диагноз формулируется следующим образом. '
              'Ключевой разрыв состоит в том, что МТС Юрент работает с принципиально неоднородной портфельной средой через единообразную управленческую логику. '
              'Этот разрыв вызван тремя причинами: переходом рынка к зрелой стадии, расхождением региональных условий по структурным факторам '
              'и слабой формализацией регионального управления внутри компании. Если сохранить текущую модель, оператор столкнётся с ухудшением '
              'экономики покрытия. Для систематической работы с этой неоднородностью я ввёл два инструмента: трёхуровневую иерархию факторов '
              'различения рынков и пять критериев оценки стратегических сценариев».')

# ============================== СЛАЙД 10: ТРИ СЦЕНАРИЯ ==============================
s = add_slide()
add_title(s, 'Три сценария по 5 критериям — S3 побеждает по четырём из пяти', size=22)

# Top: 3 scenario cards
y_sc = 1.7; sc_h = 1.95; sc_w = 4.05; gap_sc = 0.15
scenarios = [
    ('S1', 'Сохранение единой логики', 'Все регионы — по одной модели расширения. Слабое место: не различает рынки, теряет источник новой ценности.'),
    ('S2', 'Единый режим R3 + операционная адаптация', 'Всему портфелю — режим удержания. Слабое место: не использует резерв плотности в R2-городах.'),
    ('S3', 'Диверсификация режимов по портфелю', 'Каждой группе рынков — свой режим. Сложнее в координации, но даёт максимальный эффект.'),
]
for i, (code, name, body) in enumerate(scenarios):
    x = 0.5 + i*(sc_w + gap_sc)
    is_winner = (code == 'S3')
    bg = RED if is_winner else LIGHT
    txt = WHITE if is_winner else DARK
    sub = WHITE if is_winner else GRAY
    add_card(s, x, y_sc, sc_w, sc_h, fill=bg)
    add_text(s, code, x+0.25, y_sc+0.15, 1.5, 0.5, size=24, bold=True, color=txt)
    add_text(s, name, x+1.5, y_sc+0.2, sc_w-1.65, 0.55, size=12, bold=True, color=txt, ls=1.15)
    add_text(s, body, x+0.25, y_sc+0.85, sc_w-0.4, 1.1, size=10.5, color=sub, ls=1.3)

# Table 3×5 of criteria
y_tb = 3.85
crit_table = [
    ('Критерий', 'S1', 'S2', 'S3'),
    ('1. Ресурсно-теоретическая защитимость', 'низкая', 'средняя', 'высокая'),
    ('2. Согласованность с динамическими способностями', 'низкая', 'средняя', 'высокая'),
    ('3. Операционная исполнимость', 'высокая (инерция)', 'высокая', 'средняя'),
    ('4. Финансово-операционный эффект', 'низкий', 'средний', 'высокий'),
    ('5. Профиль риска', 'концентрирован', 'концентр. в R3', 'распределён'),
]
tb_left = 0.5; tb_w = 12.33
col_ws = [5.5, 2.2, 2.2, 2.43]
for r_i, row in enumerate(crit_table):
    y = y_tb + r_i*0.42
    is_header = (r_i == 0)
    fill = DARK if is_header else (WHITE if r_i % 2 == 0 else VLIGHT)
    add_card(s, tb_left, y, tb_w, 0.42, fill=fill)
    x = tb_left
    for c_i, cell in enumerate(row):
        bold = is_header or c_i == 3 or c_i == 0
        color = WHITE if is_header else (RED if (c_i == 3 and not is_header) else DARK)
        if not is_header and c_i == 3 and cell in ('высокая', 'высокий', 'распределён'):
            color = GREEN_TXT
        align = PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.CENTER
        add_text(s, cell, x+0.15, y+0.08, col_ws[c_i]-0.3, 0.3,
                 size=11, bold=bold, color=color, align=align)
        x += col_ws[c_i]

add_punchline(s, 'S3 уступает только по критерию операционной исполнимости в горизонте перестройки — адресовано компенсаторами (Слайд 15)')
add_footer(s, 10, TOTAL)
set_notes(s, 'Ключевая мысль: выбор сценария — не голосование по одному критерию, а согласованное сопоставление по пяти. '
              '«На основе диагноза я рассмотрел три стратегических альтернативы — S1, S2, S3. Сравнение проводилось по пяти критериям, '
              'выведенным из теоретической основы. Результаты показали, что S3 сильнее альтернатив по четырём критериям из пяти. '
              'Уступает только по операционной исполнимости в горизонте перестройки, поскольку требует одновременной поддержки нескольких режимов. '
              'Этот проигрыш не блокирует выбор: для него предусмотрены управленческие компенсаторы — поэтапное внедрение, '
              'функция портфельного управления и единообразие на уровне продукта».')

# ============================== СЛАЙД 11: ВЫБРАН S3 ==============================
s = add_slide()
add_title(s, 'Выбран S3 — диверсификация режимов по сценарно-однородным группам рынков', size=22)
add_subtitle(s, 'Оператор переходит от единой логики ко всем рынкам к назначению каждому рынку режима по фактическим параметрам и регулярному пересмотру каждый сезон',
             top=1.45, size=13, color=GRAY)

# 5 city cards with mode + key reason
cities = [
    ('Москва',          'R3 (огранич.)',   'Жёсткая регуляторика, удержание маржи без расширения парка'),
    ('Санкт-Петербург', 'R2 → R3',          'Запас плотности; переход к R3 при ужесточении регуляторики'),
    ('Краснодар',       'R2 / R3',          'Длинный сезон, переменная позиция; гибкое назначение по сезону'),
    ('Екатеринбург',    'R2 + пилот УПС',   'Резерв плотности + позиция преследователя — идеальный полигон'),
    ('Новосибирск',     'R2/R3 эффект.',    'Короткий сезон, протяжённая геометрия — приоритет эффективности'),
]
y_c = 2.2; c_h = 1.7; c_w = 2.45; c_gap = 0.1
total_c = 5*c_w + 4*c_gap
start_c = (13.333 - total_c) / 2
for i, (city, mode, reason) in enumerate(cities):
    x = start_c + i*(c_w + c_gap)
    is_pilot = 'УПС' in mode
    bg = RED if is_pilot else LIGHT
    txt = WHITE if is_pilot else DARK
    sub = WHITE if is_pilot else GRAY
    add_card(s, x, y_c, c_w, c_h, fill=bg)
    add_text(s, city, x+0.15, y_c+0.2, c_w-0.3, 0.4, size=13, bold=True, color=txt)
    add_text(s, mode, x+0.15, y_c+0.62, c_w-0.3, 0.4, size=14, bold=True, color=(WHITE if is_pilot else RED))
    add_text(s, reason, x+0.15, y_c+1.05, c_w-0.3, 0.6, size=9.5, color=sub, ls=1.25)

# Bottom block: что отличает S3 от обычной кластеризации (компактнее)
add_card(s, 0.5, 4.3, 12.33, 2.25, fill=VLIGHT)
add_text(s, 'Что отличает S3 от обычной кластеризации городов', 0.85, 4.45, 12, 0.35, size=13, bold=True, color=DARK)

points = [
    ('Воспроизводимая процедура',
     'Назначение режима — по фактическим KPI; пересмотр каждый сезон, а не статическая фиксация'),
    ('Однородные поля описания',
     'Каждый режим описан через цель, условия, рычаги, KPI, ограничения — это делает сопоставимыми решения по разным городам'),
    ('Защита от копирования',
     'Конкуренту недостаточно воспроизвести матрицу — нужна функция портфельного управления и локальной адаптации (Reed, DeFillippi, 1990)'),
]
for i, (h, b) in enumerate(points):
    x = 0.85 + i*4.05
    add_text(s, h, x, 4.85, 3.8, 0.4, size=11.5, bold=True, color=RED)
    add_text(s, b, x, 5.25, 3.8, 1.25, size=10.5, color=DARK, ls=1.3)

add_punchline(s, 'S3 — не статическая кластеризация, а воспроизводимая процедура с пересмотром каждый сезон')
add_footer(s, 11, TOTAL)
set_notes(s, 'Ключевая мысль: S3 — не статическая карта городов, а воспроизводимая процедура; именно это сложно скопировать конкуренту. '
              '«Выбранная стратегия — диверсификация режимов по портфелю. Её суть в том, что оператор переходит от единой логики ко всем рынкам '
              'к назначению каждому рынку режима по фактическим параметрам и регулярному пересмотру этого назначения. '
              'На демонстрационном портфеле из пяти городов S3 операционализируется так: Москва — режим удержания в ограниченной форме, '
              'Санкт-Петербург — атака с переходом к удержанию, Екатеринбург — атака с пилотом проектного инструмента, '
              'Краснодар и Новосибирск — гибкие комбинации R2 и R3 по фактической позиции».')

# ============================== СЛАЙД 12: КАРКАС РЕЖИМОВ ==============================
s = add_slide()
add_title(s, 'Каркас режимов R1–R3 покрывает жизненный цикл присутствия оператора на рынке', size=22)

modes = [
    ('R1', 'Плацдарм', 'Закрепиться на новом или слабо освоенном рынке',
     [('Условия', 'доля < 15 %, новый рынок'),
      ('Рычаги', 'минимально достаточный парк, отстройка процессов, базовый бренд'),
      ('KPI', 'покрытие ключевых зон, базовая утилизация'),
      ('Применение', 'малые города, ранний вход')]),
    ('R2', 'Атака', 'Рост доли поездок в позиции преследователя',
     [('Условия', 'доля 15–35 %, есть запас плотности'),
      ('Рычаги', 'расширение покрытия, ценовая политика, УПС'),
      ('KPI', 'рост поездок, утилизация парка, выручка/самокат'),
      ('Применение', 'СПб, Екатеринбург, Краснодар, Новосибирск')]),
    ('R3', 'Удержание', 'Защита позиции при приоритете эффективности',
     [('Условия', 'доля > 35 % / жёсткая регуляторика'),
      ('Рычаги', 'дисциплина расстановки, OpEx-оптимизация, УПС'),
      ('KPI', 'EBITDA-маржа, доля физической ребалансировки'),
      ('Применение', 'Москва, зрелые рынки')]),
]
y_m = 1.85; m_h = 4.95; m_w = 4.05; m_gap = 0.15
for i, (code, name, purpose, fields) in enumerate(modes):
    x = 0.5 + i*(m_w + m_gap)
    add_card(s, x, y_m, m_w, m_h, fill=LIGHT)
    # Header bar (увеличена высота)
    head_h = 1.35
    add_card(s, x, y_m, m_w, head_h, fill=RED)
    add_text(s, code, x+0.2, y_m+0.25, 1.2, 0.7, size=30, bold=True, color=WHITE)
    add_text(s, name, x+1.5, y_m+0.18, m_w-1.65, 0.45, size=15, bold=True, color=WHITE)
    add_text(s, purpose, x+1.5, y_m+0.65, m_w-1.65, 0.65, size=10, color=WHITE, ls=1.2)
    # Fields
    fy = y_m + head_h + 0.15
    for f_h, f_b in fields:
        add_text(s, f_h, x+0.25, fy, m_w-0.5, 0.3, size=10.5, bold=True, color=RED)
        add_text(s, f_b, x+0.25, fy+0.3, m_w-0.5, 0.55, size=11, color=DARK, ls=1.25)
        fy += 0.85

add_punchline(s, 'Однородный набор полей описания делает диагностики разных городов сопоставимыми')

add_footer(s, 12, TOTAL)
set_notes(s, 'Ключевая мысль: три режима покрывают полный жизненный цикл; однородные поля описания делают решения сопоставимыми. '
              '«Содержательное наполнение S3 — система трёх режимов конкурентных действий. R1 "Плацдарм" — закрепление на новом рынке без атаки. '
              'R2 "Атака" — рост доли в позиции преследователя при наличии резерва плотности. R3 "Удержание" — защита позиции при приоритете эффективности. '
              'Каждый режим описан через однородный набор полей: цель, условия применения, рычаги, KPI, ограничения. '
              'Это и делает сопоставимыми диагностики разных городов».')

# ============================== СЛАЙД 13: УПС ==============================
s = add_slide()
add_title(s, 'Проектный инструмент УПС: стимул в 4,5× дешевле физической ребалансировки', size=22)

# Process flow strip (заимствование подачи v3)
flow_steps = ['Сигнал дефицита', 'Отбор точки и окна', 'Выбор стимула', 'Доступность восстановлена']
fy = 1.45; fh = 0.4
fw = (12.33 - 3*0.15) / 4
fx_start = 0.5
for i, step in enumerate(flow_steps):
    x = fx_start + i*(fw + 0.15)
    add_card(s, x, fy, fw, fh, fill=LIGHT)
    add_text(s, step, x, fy+0.1, fw, 0.25, size=10, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    if i < len(flow_steps)-1:
        # Small arrow between cards
        add_text(s, '→', x+fw+0.01, fy+0.07, 0.13, 0.3, size=14, bold=True, color=RED, align=PP_ALIGN.CENTER)

# Left: big number comparison (compressed)
add_card(s, 0.5, 2.0, 6.5, 4.55, fill=VLIGHT)
add_text(s, 'Экономика операции по перемещению парка', 0.7, 2.15, 6.2, 0.4, size=12, color=GRAY)

# 120 руб
add_text(s, 'Физическая ребалансировка', 0.7, 2.6, 6.2, 0.4, size=12, bold=True, color=DARK)
add_bignum(s, '120 ₽', 0.7, 2.95, 6.2, 1.0, size=58, color=DARK, align=PP_ALIGN.LEFT)
add_text(s, 'за операцию, данные эксперта МТС', 0.7, 3.95, 6.2, 0.3, size=10, color=GRAY)

# vs
add_text(s, '↓ замещение стимулом', 0.7, 4.3, 6.2, 0.3, size=11, color=GRAY, align=PP_ALIGN.CENTER)

# 26,5 руб
add_text(s, 'Средневзвешенная стоимость стимула УПС', 0.7, 4.65, 6.2, 0.4, size=12, bold=True, color=DARK)
add_bignum(s, '26,5 ₽', 0.7, 5.0, 6.2, 1.0, size=58, color=RED, align=PP_ALIGN.LEFT)
add_text(s, 'за операцию · в 4,5× дешевле', 0.7, 6.0, 6.2, 0.3, size=11, bold=True, color=RED)

# Right: 3 mechanisms scheme
add_text(s, 'Три механизма УПС (шаг «Выбор стимула»)', 7.4, 2.0, 5.5, 0.3, size=12, bold=True, color=DARK)

mechs = [
    ('1', 'Стимул на стороне точки высадки',
     'Пользователю предлагается выгода за завершение поездки в зоне дефицита'),
    ('2', 'Стимул на стороне точки посадки',
     'Скидка за самокат из зоны переизбытка — снижает плотность в перенасыщенных зонах'),
    ('3', 'Микрозадание внешнему исполнителю',
     'Оплачиваемое перемещение в случаях, где пользовательских стимулов недостаточно'),
]
my = 2.45
for code, name, body in mechs:
    add_card(s, 7.4, my, 5.5, 1.25, fill=LIGHT)
    # Code circle (M + number on two lines via newline)
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.55), Inches(my+0.25), Inches(0.75), Inches(0.75))
    circle.fill.solid(); circle.fill.fore_color.rgb = RED
    circle.line.fill.background()
    tb = circle.text_frame
    tb.margin_top = tb.margin_bottom = tb.margin_left = tb.margin_right = 0
    tb.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp = tb.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
    rr = pp.add_run(); rr.text = 'М' + code
    rr.font.name = FONT; rr.font.size = Pt(18); rr.font.bold = True
    rr.font.color.rgb = WHITE
    add_text(s, name, 8.5, my+0.15, 4.3, 0.4, size=12, bold=True, color=DARK)
    add_text(s, body, 8.5, my+0.55, 4.3, 0.65, size=10, color=GRAY, ls=1.3)
    my += 1.35

add_punchline(s, 'УПС применяется в режимах R2 и R3, снижает физическую ребалансировку, не отменяя её полностью')
add_footer(s, 13, TOTAL)
set_notes(s, 'Ключевая мысль: проектный инструмент адресует ключевой фактор успеха — доступность покрытия — и при этом дешевле физической ребалансировки. '
              '«Проектный инструмент я назвал "Управление покрытием через стимулы". Это механизм, который замещает часть физической ребалансировки '
              'тремя типами стимулов: на стороне точки высадки, на стороне точки посадки и оплачиваемого микрозадания внешнему исполнителю. '
              'Ключевая цифра: средневзвешенная стоимость стимула — 26,5 рубля за операцию, стоимость физической ребалансировки — 120 рублей '
              'за операцию по данным эксперта МТС. Инструмент в 4,5 раза дешевле и применяется в режимах атаки и удержания».')

# ============================== СЛАЙД 14: ФИНАНСОВАЯ МОДЕЛЬ ==============================
s = add_slide()
add_title(s, 'Финансовая модель: +25 млн ₽ EBITDA в год 1, NPV +134 млн при WACC 20 %', size=22)

# Waterfall в виде столбиков
y_wf = 1.85
wf_h = 3.4
add_text(s, 'Год 1, эффект сценария S3 vs S1 (млн руб. EBITDA)', 0.5, y_wf, 12.3, 0.35, size=12, color=GRAY)
# 8 bars: затраты A/B/C/D, эффекты 1/2/3/4, чистый
bars = [
    ('Управление',     -22, RED_TXT),
    ('Маркетинг',      -7,  RED_TXT),
    ('УПС',            -16, RED_TXT),
    ('Накладные',      -3,  RED_TXT),
    ('R2-выручка',     +12, GREEN_TXT),
    ('УПС-пилот',      +6,  GREEN_TXT),
    ('OpEx Москва',    +52, GREEN_TXT),
    ('Малые города',   +3,  GREEN_TXT),
    ('ЧИСТЫЙ',         +25, RED),
]
b_y_top = y_wf + 0.5
b_y_bot = y_wf + wf_h
b_total_h = b_y_bot - b_y_top  # доступная высота
b_zero = b_y_top + b_total_h*0.55  # zero-линия

# zero line
zline = s.shapes.add_connector(1, Inches(0.5), Inches(b_zero), Inches(8.7), Inches(b_zero))
zline.line.color.rgb = GRAY; zline.line.width = Pt(0.75)

max_abs = max(abs(v) for _, v, _ in bars)  # 52
unit_h = (b_total_h*0.45) / max_abs  # max возможная высота

bar_w = 0.83; bar_gap = 0.07
x0 = 0.5
for i, (lbl, val, color) in enumerate(bars):
    x = x0 + i*(bar_w + bar_gap)
    h_in = abs(val) * unit_h
    if val >= 0:
        y = b_zero - h_in
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(bar_w), Inches(h_in))
    else:
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(b_zero), Inches(bar_w), Inches(h_in))
    bar.fill.solid()
    if lbl == 'ЧИСТЫЙ':
        bar.fill.fore_color.rgb = DARK
    elif val < 0:
        bar.fill.fore_color.rgb = RGBColor(0xD0, 0x40, 0x40)
    else:
        bar.fill.fore_color.rgb = RGBColor(0x47, 0x9A, 0x5F)
    bar.line.fill.background()
    # Value label
    val_str = f'{val:+d}'
    val_y = (b_zero - h_in - 0.32) if val >= 0 else (b_zero + h_in + 0.05)
    add_text(s, val_str, x-0.1, val_y, bar_w+0.2, 0.3, size=11, bold=True,
             color=color, align=PP_ALIGN.CENTER, ls=1.0)
    # Label below — узкая колонка, мелкий шрифт, чтобы не наезжать
    add_text(s, lbl, x-0.15, b_y_bot+0.1, bar_w+0.3, 0.5, size=8.5,
             color=(RED if val == 25 else DARK), bold=(val == 25),
             align=PP_ALIGN.CENTER, ls=1.15)

# Right metrics block — увеличен интервал, чтобы строки не перекрывались
mx = 9.6; mw = 3.5
add_card(s, mx, y_wf+0.2, mw, 4.05, fill=VLIGHT)
add_text(s, 'Ключевые показатели', mx+0.25, y_wf+0.32, mw-0.5, 0.3, size=11, bold=True, color=GRAY)

metrics = [
    ('Год 2',          '+80…+100 млн ₽', 'при тиражировании УПС'),
    ('NPV',            '+134 млн ₽',     'WACC 20 %, 3 года'),
    ('Payback',        '≈ 8 мес.',       'один пилотный сезон'),
    ('ROI год 1',      '52 %',           'чистый эффект / затраты'),
]
my2 = y_wf + 0.75
for h, big, sub in metrics:
    add_text(s, h, mx+0.3, my2, mw-0.4, 0.25, size=9.5, color=GRAY)
    add_text(s, big, mx+0.3, my2+0.23, mw-0.4, 0.4, size=16, bold=True, color=RED)
    add_text(s, sub, mx+0.3, my2+0.65, mw-0.4, 0.22, size=8.5, color=GRAY)
    my2 += 0.83

# Bottom: sensitivity (compressed) + punchline
add_text(s, 'Чувствительность год 1:  −14 / +25 / +64 млн ₽',
         0.5, 6.3, 12.3, 0.3, size=11, color=GRAY, align=PP_ALIGN.CENTER)

add_punchline(s, 'S3 окупается за счёт перенастройки уже существующего портфеля — capex проекта всего 8 млн ₽')
add_footer(s, 14, TOTAL)
set_notes(s, 'Ключевая мысль: проект окупается за один сезон при минимальной капитальной интенсивности. '
              '«Затраты года 1 — около 48 миллионов рублей; capex составляет всего 8 миллионов на IT-инфраструктуру УПС, остальное операционные. '
              'Эффекты — 73 миллиона EBITDA: 12 от прироста выручки в R2-городах, 6 от пилота УПС в Екатеринбурге, '
              '52 от оптимизации OpEx в Москве и 3 от отказа от инвестиций в нерентабельные локации. '
              'Чистый эффект года 1 — плюс 25 миллионов; чувствительность от минус 14 до плюс 64. '
              'Год 2 при тиражировании УПС — плюс 80–100 миллионов. NPV при WACC 20 % на трёхлетнем горизонте — плюс 134 миллиона, '
              'период окупаемости около восьми месяцев. Это сезон одного пилота».')

# ============================== СЛАЙД 15: ПЛАН И РИСКИ ==============================
s = add_slide()
add_title(s, 'Поэтапное внедрение: пилот в Екатеринбурге → масштабирование на R2-кластер в год 2', size=22)

# Timeline of 5 months
add_text(s, 'Дорожная карта пилота УПС, 12 месяцев', 0.5, 1.6, 12.3, 0.3, size=12, color=GRAY)

months = [
    ('М1', 'Диагностика', 'базовые KPI,\nчувствительные\nмикролокации'),
    ('М2', 'Запуск М1 + М2', 'стимулы\nпользователям'),
    ('М3', 'Запуск М3', 'микрозадания\nвнешним\nисполнителям'),
    ('М4', 'Калибровка', 'параметры стимулов,\nрасширение зоны'),
    ('М5', 'Итоги', 'решение\nо тиражировании'),
]
y_t = 2.0; t_h = 1.95; t_w = 2.45; t_gap = 0.1
total_t = 5*t_w + 4*t_gap
start_t = (13.333 - total_t) / 2
# Connecting line
ln = s.shapes.add_connector(1, Inches(start_t+0.4), Inches(y_t+1.0),
                             Inches(start_t+total_t-0.4), Inches(y_t+1.0))
ln.line.color.rgb = RED; ln.line.width = Pt(2)
for i, (code, name, body) in enumerate(months):
    x = start_t + i*(t_w + t_gap)
    # Circle on line — увеличен размер чтобы М+цифра поместилось в одну строку
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+t_w/2-0.5), Inches(y_t+0.5), Inches(1.0), Inches(1.0))
    circle.fill.solid(); circle.fill.fore_color.rgb = RED
    circle.line.fill.background()
    tb = circle.text_frame
    tb.margin_top = tb.margin_bottom = tb.margin_left = tb.margin_right = 0
    tb.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp = tb.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
    rr = pp.add_run(); rr.text = code
    rr.font.name = FONT; rr.font.size = Pt(22); rr.font.bold = True
    rr.font.color.rgb = WHITE
    # Name and body
    add_text(s, name, x, y_t+1.55, t_w, 0.35, size=12, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    add_text(s, body, x, y_t+1.9, t_w, 1.0, size=10, color=GRAY, align=PP_ALIGN.CENTER, ls=1.25)

# Bottom: risks vs compensators
y_rk = 4.55
add_text(s, 'Риски и управленческие компенсаторы', 0.5, y_rk, 12.3, 0.35, size=14, bold=True, color=DARK)

risks = [
    ('Рассинхронизация режимов',                'функция портфельного управления + цикл сверки'),
    ('Потеря единого пользовательского опыта',  'единообразие на уровне продукта при дифференциации'),
    ('Кадровая нагрузка координации',           'поэтапное внедрение: 5 городов → расширение'),
    ('Регуляторный шок в нескольких рынках',    'встроенные правила перехода R2 → R3'),
    ('Ошибка диагностики рынка',                'регулярный пересмотр режима по KPI'),
    ('Универсализация УПС',                     'фиксация условий применимости (R2, R3)'),
]
r_y = y_rk + 0.5
for i, (risk, comp) in enumerate(risks):
    col = i % 2; row = i // 2
    x = 0.5 + col*6.15
    y = r_y + row*0.65
    add_card(s, x, y, 6.0, 0.55, fill=LIGHT)
    add_text(s, risk, x+0.2, y+0.05, 2.7, 0.45, size=10.5, bold=True, color=DARK, ls=1.1)
    # arrow
    add_text(s, '→', x+2.95, y+0.1, 0.35, 0.45, size=14, bold=True, color=RED, align=PP_ALIGN.CENTER)
    add_text(s, comp, x+3.35, y+0.05, 2.55, 0.45, size=10, color=GRAY, ls=1.1)

add_punchline(s, 'S3 не требует одномоментной перестройки — запускается через пилот за 5 месяцев')
add_footer(s, 15, TOTAL)
set_notes(s, 'Ключевая мысль: S3 не требует одномоментной перестройки; запускается через пилот, у каждого риска есть управленческий ответ. '
              '«Внедрение S3 идёт через пилот УПС в Екатеринбурге — пять месяцев от диагностики до итогов. С года 2 — тиражирование на четыре города R2-кластера. '
              'Шесть рисков, каждый с управленческим ответом. Главный тезис: S3 не требует одномоментной перестройки всей компании, '
              'запускается через пилот и расширяется при подтверждении экономики».')

# ============================== СЛАЙД 16: ФИНАЛ ==============================
s = add_slide()
add_title(s, 'Управленческий вывод: неоднородность регионов превращается из проблемы в инструмент выбора действий', size=22)

# 4 verbs of contribution
y_f = 1.95; f_h = 2.6; f_w = 2.95; f_gap = 0.2
verbs = [
    ('1', 'Диагностирована', 'управленческая проблема\nрегионального портфеля\nиз 187 локаций'),
    ('2', 'Сопоставлены', 'три стратегические альтернативы\nпо пяти критериям, выводимым\nиз теоретической основы'),
    ('3', 'Предложен инструмент', 'УПС — управление покрытием\nчерез стимулы, в 4,5× дешевле\nфизической ребалансировки'),
    ('4', 'Выполнена оценка', 'финмодель: +25 млн ₽ год 1,\nNPV +134 млн при WACC 20 %,\npayback ≈ 8 мес.'),
]
for i, (n, head, body) in enumerate(verbs):
    x = 0.5 + i*(f_w + f_gap)
    is_first = (i == 0)
    bg = RED if is_first else LIGHT
    txt = WHITE if is_first else DARK
    sub = WHITE if is_first else GRAY
    head_color = WHITE if is_first else RED
    add_card(s, x, y_f, f_w, f_h, fill=bg)
    add_bignum(s, n, x+0.15, y_f+0.2, f_w-0.3, 0.7, size=36, color=txt, align=PP_ALIGN.LEFT)
    add_text(s, head, x+0.25, y_f+0.95, f_w-0.4, 0.45, size=14, bold=True, color=head_color)
    add_text(s, body, x+0.25, y_f+1.4, f_w-0.4, 1.1, size=11, color=sub, ls=1.3)

# Final phrase block
add_card(s, 0.5, 4.85, 12.3, 1.85, fill=VLIGHT)
add_text(s, 'Главный практический вклад работы', 0.85, 4.95, 12, 0.35, size=12, color=GRAY)
add_text(s, 'Рабочий инструментарий для подготовки стратегических решений МТС Юрент: диагностический каркас, система режимов R1–R3, '
            'проектный инструмент УПС и шаблон финансовой модели для расширения на остальные города присутствия.',
         0.85, 5.3, 12, 1.35, size=14, bold=True, color=DARK, ls=1.35)

# Spasibo
add_text(s, 'Спасибо за внимание', 0.5, 6.85, 12.3, 0.45, size=18, bold=True, color=RED, align=PP_ALIGN.CENTER)

set_notes(s, 'Финальная фраза: «Если кратко подвести итог работы. Управленческая проблема МТС Юрент состояла в том, '
              'что для портфеля из 187 локаций нет формализованной процедуры выбора режима присутствия. '
              'Проведённая диагностика показала, что причина — в исчерпании источника новой ценности от единой логики расширения. '
              'Сопоставление трёх сценариев по пяти критериям обосновало выбор стратегии диверсификации режимов. '
              'В её рамках предложен проектный инструмент УПС и построена портфельная финансовая модель с эффектом +25 миллионов в год 1 '
              'и +80–100 миллионов в год 2. Главный практический вклад — рабочий инструментарий для МТС Юрент. Спасибо за внимание, готов к вопросам».')

# ============================== БЭКАП-СЛАЙДЫ ==============================

def backup_header(s, num, total_b, title):
    # Replace top bar with gray to distinguish backup
    # Already has red bar, add a "Backup" badge
    badge = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(11.0), Inches(0.35), Inches(2.0), Inches(0.45))
    badge.fill.solid(); badge.fill.fore_color.rgb = GRAY
    badge.line.fill.background()
    tb = badge.text_frame; tb.margin_top = tb.margin_bottom = 0
    p = tb.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = f'BACKUP  {num}/{total_b}'
    r.font.name = FONT; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = WHITE
    add_title(s, title, top=0.95, size=22)

BTOTAL = 5

# B1: 5 критериев подробно
s = add_slide()
backup_header(s, 1, BTOTAL, 'Пять критериев оценки стратегических сценариев — теоретическая основа')
crit = [
    ('1. Ресурсно-теоретическая защитимость', 'Wernerfelt 1984; Barney 1991',
     'Насколько сценарий опирается на актуальную ресурсную базу оператора и не требует её радикальной перестройки'),
    ('2. Согласованность с динамическими способностями', 'Teece, Pisano, Shuen 1997; Teece 2007',
     'Воспроизводимость перестройки конфигурации ресурсов под меняющиеся условия рынка'),
    ('3. Операционная исполнимость', 'Sareen, Remme, Haarstad 2021; Moran 2021',
     'Реализуемость сценария при текущей операционной зрелости оператора и регуляторных ограничениях'),
    ('4. Финансово-операционный эффект', 'Aarhaug et al. 2023; Shah et al. 2022; Cennamo 2021',
     'Ожидаемый прирост EBITDA на портфеле городов в горизонте 12 месяцев'),
    ('5. Профиль риска', 'Groth et al. 2025; Coenegrachts et al. 2024',
     'Структура и величина рисков (поведенческих, технологических, регуляторных, конкурентных)'),
]
yc = 2.05
for h, src, body in crit:
    add_card(s, 0.5, yc, 12.3, 0.85, fill=LIGHT)
    add_text(s, h, 0.7, yc+0.1, 6.5, 0.4, size=13, bold=True, color=DARK)
    add_text(s, src, 0.7, yc+0.48, 6.5, 0.3, size=10, color=GRAY, ls=1.0)
    add_text(s, body, 7.4, yc+0.15, 5.2, 0.6, size=10.5, color=DARK, ls=1.25)
    yc += 0.95

# B2: полная матрица S1/S2/S3
s = add_slide()
backup_header(s, 2, BTOTAL, 'Сценарная матрица: полная оценка S1, S2, S3 по пяти критериям')
matrix = [
    ('Критерий', 'S1 (единая логика)', 'S2 (единый режим + адаптация)', 'S3 (диверсификация)'),
    ('1. Ресурсная защитимость',
     'Использует ресурсы, но не делает их источником преимущества',
     'Опирается на операционный блок, недоиспользует экосистему',
     'Дифференцированно задействует все блоки способностей'),
    ('2. Динамические способности',
     'Не требует различения рынков — способности не задействованы',
     'Адаптация в рамках одного режима, не на портфельном уровне',
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
     'Концентр. в R3: невозможность атаковать при появлении возможностей',
     'Распределён между режимами; добавляется риск координации'),
]
y_m = 1.95; rw = [3.0, 2.9, 3.2, 3.2]
for r_i, row in enumerate(matrix):
    y = y_m + r_i*0.8
    is_h = (r_i == 0)
    fill = DARK if is_h else (WHITE if r_i % 2 == 0 else VLIGHT)
    add_card(s, 0.5, y, 12.3, 0.8, fill=fill)
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else DARK
        if not is_h and c_i == 3:
            color = RED
        add_text(s, cell, x+0.1, y+0.1, rw[c_i]-0.2, 0.6,
                 size=10 if not is_h else 11, bold=is_h or c_i == 0,
                 color=color, ls=1.2)
        x += rw[c_i]

# B3: финмодель детально
s = add_slide()
backup_header(s, 3, BTOTAL, 'Финансовая модель: детализация затрат и эффектов')

# Two tables
add_text(s, 'Затраты года 1', 0.5, 1.95, 5.5, 0.4, size=14, bold=True, color=DARK)
cost = [
    ('Блок', 'Млн ₽', 'Источник'),
    ('A. Региональное управление', '−22', '3 менеджера + 3 аналитика'),
    ('B. Маркетинг S3', '−7', 'A/B-тесты, креатив по 3 режимам'),
    ('C. УПС (вкл. capex 8)', '−16', 'IT + команда пилота + стимулы'),
    ('D. Накладные', '−3', 'обучение, методич. сопровождение'),
    ('Итого', '−48', ''),
]
yc = 2.4
row_h = 0.5
for r_i, row in enumerate(cost):
    is_h = (r_i == 0); is_total = (r_i == len(cost)-1)
    fill = DARK if is_h else (LIGHT if is_total else (WHITE if r_i % 2 == 0 else VLIGHT))
    add_card(s, 0.5, yc, 6.0, row_h, fill=fill)
    cw_c = [3.0, 0.9, 2.1]
    x = 0.5
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else (RED if c_i == 1 and not is_h else DARK)
        bold = is_h or is_total or c_i == 1
        align = PP_ALIGN.LEFT if c_i in (0, 2) else PP_ALIGN.RIGHT
        size = 10.5 if c_i < 2 else 9.5
        add_text(s, cell, x+0.1, yc+0.13, cw_c[c_i]-0.2, 0.3,
                 size=size, bold=bold, color=color, align=align)
        x += cw_c[c_i]
    yc += row_h

add_text(s, 'Эффекты года 1 (EBITDA)', 6.83, 1.95, 5.5, 0.4, size=14, bold=True, color=DARK)
eff = [
    ('Эффект', 'Млн ₽', 'Расчёт'),
    ('1. R2-прирост выручки', '+12', '+3 п.п. × 1 544 × 25 % маржа'),
    ('2. УПС-пилот Екб', '+6', 'экономия 3,5 + выручка 8,2 × 25 %'),
    ('3. R3-Москва OpEx', '+52', '−2 % от выручки Москвы 2 626 млн'),
    ('4. Малые города', '+3', 'отказ от убыточных локаций'),
    ('Итого', '+73', ''),
]
yc = 2.4
for r_i, row in enumerate(eff):
    is_h = (r_i == 0); is_total = (r_i == len(eff)-1)
    fill = DARK if is_h else (LIGHT if is_total else (WHITE if r_i % 2 == 0 else VLIGHT))
    add_card(s, 6.83, yc, 6.0, row_h, fill=fill)
    cw_c = [2.4, 0.9, 2.7]
    x = 6.83
    for c_i, cell in enumerate(row):
        color = WHITE if is_h else (GREEN_TXT if c_i == 1 and not is_h else DARK)
        bold = is_h or is_total or c_i == 1
        align = PP_ALIGN.LEFT if c_i in (0, 2) else PP_ALIGN.RIGHT
        size = 10.5 if c_i < 2 else 9.5
        add_text(s, cell, x+0.1, yc+0.13, cw_c[c_i]-0.2, 0.3,
                 size=size, bold=bold, color=color, align=align)
        x += cw_c[c_i]
    yc += row_h

# Net
add_card(s, 0.5, 5.55, 12.33, 0.85, fill=RED)
add_text(s, 'Чистый эффект год 1 vs S1', 0.7, 5.65, 4, 0.3, size=11, color=WHITE)
add_text(s, '+25 млн ₽', 0.7, 5.95, 4, 0.4, size=22, bold=True, color=WHITE)
add_text(s, 'Чувствительность', 5.5, 5.65, 4, 0.3, size=11, color=WHITE)
add_text(s, '−14  /  +25  /  +64', 5.5, 5.95, 4, 0.4, size=20, bold=True, color=WHITE)
add_text(s, 'Год 2 при тиражировании УПС', 9.5, 5.65, 4, 0.3, size=11, color=WHITE)
add_text(s, '+80…+100 млн ₽', 9.5, 5.95, 4, 0.4, size=20, bold=True, color=WHITE)

# Bottom row of NPV/Payback
add_card(s, 0.5, 6.55, 4.0, 0.6, fill=LIGHT)
add_text(s, 'NPV (WACC 20 %, 3 года)', 0.65, 6.6, 3.7, 0.25, size=10, color=GRAY)
add_text(s, '+134 млн ₽', 0.65, 6.83, 3.7, 0.3, size=14, bold=True, color=RED)
add_card(s, 4.65, 6.55, 4.0, 0.6, fill=LIGHT)
add_text(s, 'Payback (простой)', 4.8, 6.6, 3.7, 0.25, size=10, color=GRAY)
add_text(s, '≈ 8 месяцев', 4.8, 6.83, 3.7, 0.3, size=14, bold=True, color=RED)
add_card(s, 8.8, 6.55, 4.0, 0.6, fill=LIGHT)
add_text(s, 'ROI год 1 / год 2', 8.95, 6.6, 3.7, 0.25, size=10, color=GRAY)
add_text(s, '52 %  /  243 %', 8.95, 6.83, 3.7, 0.3, size=14, bold=True, color=RED)

# B4: УПС расчёт
s = add_slide()
backup_header(s, 4, BTOTAL, 'УПС — расчётная цепочка по Екатеринбургу и 4 KPI пилота')

# Calculation chain
add_text(s, 'Расчётная цепочка (базовый сценарий B)', 0.5, 1.95, 12.3, 0.4, size=14, bold=True, color=DARK)
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
yc = 2.45
for i, (lbl, val, src) in enumerate(chain):
    is_final = (i == len(chain)-1)
    add_card(s, 0.5, yc, 12.3, 0.42,
             fill=RED if is_final else (WHITE if i % 2 == 0 else VLIGHT))
    color = WHITE if is_final else DARK
    add_text(s, lbl, 0.7, yc+0.08, 3.5, 0.3, size=11, bold=True, color=color)
    add_text(s, val, 4.4, yc+0.08, 2.5, 0.3, size=13, bold=True,
             color=(WHITE if is_final else RED), align=PP_ALIGN.RIGHT)
    add_text(s, src, 7.2, yc+0.08, 5, 0.3, size=10,
             color=(WHITE if is_final else GRAY), align=PP_ALIGN.LEFT)
    yc += 0.42

# 4 KPIs
add_text(s, '4 KPI пилота', 0.5, 6.0, 12.3, 0.4, size=14, bold=True, color=DARK)
kpis = [
    ('Прирост поездок', 'в пилотных зонах vs база'),
    ('Снижение ребалансировки', 'доля в общем объёме операций'),
    ('Стоимость стимула', 'на 1 предотвращённую потерю'),
    ('Отклик пользователей', 'доля принятых стимулов'),
]
y_k = 6.45; kw = 2.95; kgap = 0.15
for i, (h, b) in enumerate(kpis):
    x = 0.5 + i*(kw + kgap)
    add_card(s, x, y_k, kw, 0.6, fill=LIGHT)
    add_text(s, h, x+0.2, y_k+0.05, kw-0.4, 0.3, size=11, bold=True, color=DARK)
    add_text(s, b, x+0.2, y_k+0.32, kw-0.4, 0.25, size=9.5, color=GRAY)

# B5: декларация ИИ
s = add_slide()
backup_header(s, 5, BTOTAL, 'Декларация использования ИИ — что и где применялось при подготовке ВКР')

add_text(s, 'Инструменты: ChatGPT (OpenAI, GPT-4o, GPT-5) и Claude (Anthropic, Sonnet 4, Opus 4.7) через стандартные веб-интерфейсы и Cursor IDE.',
         0.5, 1.95, 12.3, 0.5, size=12, color=DARK, ls=1.3)

# 2 columns
add_card(s, 0.5, 2.65, 6.1, 4.3, fill=LIGHT)
add_card(s, 0.5, 2.65, 6.1, 0.55, fill=GREEN_TXT)
add_text(s, '✓  Где использовался ИИ', 0.5, 2.73, 6.1, 0.4, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
ai_use = [
    ('Поиск и оформление литературы', 'фильтрация источников, ГОСТ-форматирование'),
    ('Структурирование черновиков', 'формулировки задач, структура глав'),
    ('Редактура и стилистика', 'академический стиль, грамматика'),
    ('Перевод аннотации', 'русский → английский'),
    ('Автоматизация расчётов', 'Python для CSV и Excel-модели'),
    ('Сборка финального .docx и .xlsx', 'форматирование, проверка консистентности'),
]
y = 3.35
for h, b in ai_use:
    add_text(s, '✓  ' + h, 0.75, y, 5.6, 0.3, size=11, bold=True, color=DARK)
    add_text(s, b, 1.05, y+0.3, 5.3, 0.25, size=10, color=GRAY)
    y += 0.6

add_card(s, 6.75, 2.65, 6.1, 4.3, fill=LIGHT)
add_card(s, 6.75, 2.65, 6.1, 0.55, fill=RED_TXT)
add_text(s, '✗  Где ИИ не применялся', 6.75, 2.73, 6.1, 0.4, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
ai_no = [
    ('Постановка проблемы и цели', 'авторская формулировка'),
    ('Интерпретация интервью с экспертами', 'обобщённые тезисы передавались, не расшифровки'),
    ('Выбор сценария S3 и режимов R1–R3', 'авторские стратегические решения'),
    ('Концепция и параметры УПС', 'авторская разработка'),
    ('Допущения и параметры финмодели', 'выбор автора, обоснованы в тексте'),
    ('Внутренние данные МТС Юрент', 'только в обезличенном агрегированном виде'),
]
y = 3.35
for h, b in ai_no:
    add_text(s, '✗  ' + h, 7.0, y, 5.6, 0.3, size=11, bold=True, color=DARK)
    add_text(s, b, 7.3, y+0.3, 5.3, 0.25, size=10, color=GRAY)
    y += 0.6

add_text(s, 'Все сгенерированные с участием ИИ фрагменты проверены и отредактированы автором. Ответственность за корректность и интерпретацию — автора ВКР.',
         0.5, 7.05, 12.3, 0.3, size=10, color=GRAY, align=PP_ALIGN.CENTER)

# === Сохранение ===
prs.save('/workspace/vkr/VKR_presentation.pptx')
print(f'Saved VKR_presentation.pptx')
print(f'Total slides: {len(prs.slides)}')
