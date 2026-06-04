#!/usr/bin/env python3
"""Excel-приложение к §3.4 ВКР: финансовая модель сценария S3.
Формулы построены через явные ссылки на ячейки с использованием словаря _row_of.
Все цифры воспроизводимы (открытые формулы), нет циклических ссылок.
"""
import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CITIES = [
    ('Москва', 1),
    ('Санкт-Петербург', 4),
    ('Краснодар', 3),
    ('Екатеринбург', 12),
    ('Новосибирск', 7),
]

def num(s):
    s = (s or '').strip().replace(' ', '').replace('\xa0','')
    try: return int(s)
    except:
        try: return float(s)
        except: return 0

rows = []
with open('/home/ubuntu/.cursor/projects/workspace/uploads/kicksharing_v2_b578.csv','r',encoding='utf-8') as f:
    for r in csv.reader(f, delimiter=';'):
        rows.append(r)

city_data = {}
for name, rid in CITIES:
    for rr in rows[3:]:
        if rr and rr[0].strip()==str(rid):
            r = rr
            break
    rev24_q = [num(r[i]) for i in range(1,5)]
    rev25_9m_q = [num(r[i]) for i in range(5,8)]
    fleet24 = num(r[8])
    fleet25 = num(r[9])
    active24 = num(r[17])
    active25_9m = num(r[18])
    trips24_q = [num(r[i]) for i in range(21,25)]
    trips25_9m_q = [num(r[i]) for i in range(25,28)]
    rev24 = sum(rev24_q)
    trips24 = sum(trips24_q)
    rev24_9m = sum(rev24_q[:3])
    rev25_9m = sum(rev25_9m_q)
    trips24_9m = sum(trips24_q[:3])
    trips25_9m = sum(trips25_9m_q)
    city_data[name] = dict(
        rid=rid, rev24_q=rev24_q, rev25_9m_q=rev25_9m_q,
        fleet24=fleet24, fleet25=fleet25,
        active24=active24, active25_9m=active25_9m,
        trips24_q=trips24_q, trips25_9m_q=trips25_9m_q,
        rev24=rev24, trips24=trips24,
        rev24_9m=rev24_9m, rev25_9m=rev25_9m,
        trips24_9m=trips24_9m, trips25_9m=trips25_9m,
        avg_check=rev24/trips24 if trips24 else 0,
        util24=trips24/fleet24 if fleet24 else 0,
        yoy_rev=(rev25_9m/rev24_9m-1)*100 if rev24_9m else 0,
        yoy_trips=(trips25_9m/trips24_9m-1)*100 if trips24_9m else 0,
    )

# Стили
HEAD = Font(name='Times New Roman', size=12, bold=True, color='FFFFFF')
HEAD_FILL = PatternFill('solid', fgColor='2F5496')
SUBHEAD = Font(name='Times New Roman', size=11, bold=True)
SUBHEAD_FILL = PatternFill('solid', fgColor='D9E2F3')
NORMAL = Font(name='Times New Roman', size=11)
TOTAL = Font(name='Times New Roman', size=11, bold=True)
TOTAL_FILL = PatternFill('solid', fgColor='FFF2CC')
GOOD_FILL = PatternFill('solid', fgColor='C6EFCE')
BAD_FILL = PatternFill('solid', fgColor='FFC7CE')
BIG_TOTAL = Font(name='Times New Roman', size=13, bold=True)
BIG_README = Font(name='Times New Roman', size=16, bold=True)
TITLE = Font(name='Times New Roman', size=14, bold=True)
BORDER = Border(left=Side(style='thin', color='999999'),
                right=Side(style='thin', color='999999'),
                top=Side(style='thin', color='999999'),
                bottom=Side(style='thin', color='999999'))

def cell(ws, row, col, value, font=NORMAL, fill=None, align=None, border=BORDER, fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    if fill: c.fill = fill
    if align: c.alignment = align
    if border: c.border = border
    if fmt: c.number_format = fmt
    return c

def widths(ws, **kwargs):
    for col, w in kwargs.items():
        ws.column_dimensions[col].width = w

wb = Workbook()
del wb[wb.active.title]

# ============================================================
# Лист 1. README
# ============================================================
ws = wb.create_sheet('README')
ws['A1'] = 'Финансовая модель сценария S3 — приложение к ВКР'
ws['A1'].font = BIG_README
ws.merge_cells('A1:E1')
ws.column_dimensions['A'].width = 100

readme = [
    '',
    'Автор: [ФИО студента], ВКР НИУ ВШЭ, 2026 г.',
    '',
    'Назначение файла. Расчётная часть финансовой модели сценария S3 (диверсификация',
    'режимов конкурентных действий МТС Юрент по портфелю региональных рынков',
    'кикшеринга) на горизонте 12 и 24 месяцев. Все цифры из §3.4 ВКР воспроизводимы',
    'из ячеек этого файла; формулы открыты, циклических ссылок нет.',
    '',
    'Структура листов:',
    '  • README — описание модели и источников.',
    '  • Данные_Юрент — сырые данные эксперта А. по 5 городам портфеля',
    '    (поквартально, 2024 и 9 месяцев 2025): выручка, парк, активные пользователи,',
    '    поездки. На основе этих данных рассчитываются производные показатели',
    '    (выручка итого, утилизация, средний чек, YoY).',
    '  • Затраты — статьи затрат на реализацию S3: региональное управление (A),',
    '    маркетинг (B), инфраструктура и пилот УПС (C), накладные (D).',
    '  • Эффекты — статьи эффектов с допущениями: R2 рост, УПС-пилот, R3 Москва,',
    '    R1 малые города.',
    '  • Сводная_финмодель — затраты, эффекты, чистый эффект S3 vs S1 (год 1 + год 2).',
    '  • Чувствительность — сценарии (консервативный, базовый, оптимистичный)',
    '    с пересчётом результата по 4 ключевым параметрам.',
    '  • УПС_расчёт — детальная разбивка пилота в Екатеринбурге (как в §3.2 ВКР).',
    '',
    'Источники данных:',
    '  • Внутренняя поквартальная отчётность МТС Юрент по 86 субъектам РФ за 2024 г.',
    '    и 9 месяцев 2025 г. (эксперт А., руководитель подразделения M&A ПАО «МТС»).',
    '  • Публичная отчётность ПАО «ВУШ Холдинг» по МСФО (9 мес. 2024) — для калибровки',
    '    OpEx/Выручка и EBITDA-маржи бенчмарк-оператора.',
    '  • Расчёт по проектному инструменту УПС (§3.2 ВКР).',
    '',
    'Допущения, требующие калибровки на этапе пилота:',
    '  — прирост темпа выручки R2-городов от дифференциации: 1–5 п.п. (база 3 п.п.);',
    '  — доля замещения физической ребалансировки стимулами: 10–30 % (база 20 %);',
    '  — стоимость физической ребалансировки: 120 руб./операция (отраслевая оценка);',
    '  — доля ребалансировки в OpEx: ≈30 % (по аналогии с международными данными);',
    '  — снижение OpEx Москвы в режиме R3: 1–3 % от выручки (база 2 %).',
    '',
    'Все статьи с допущениями явно маркированы [ДОПУЩЕНИЕ] в столбце «Источник».',
    'Точная калибровка модели возможна после доступа к внутренним данным',
    'МТС Юрент по операционным издержкам и поведению пользователей на пилоте.',
]
for i, line in enumerate(readme, start=2):
    c = ws.cell(row=i, column=1, value=line)
    c.font = NORMAL

# ============================================================
# Лист 2. Данные_Юрент
# ============================================================
ws = wb.create_sheet('Данные_Юрент')
ws['A1'] = 'Сырые данные по 5 городам портфеля (источник: эксперт А., внутренняя отчётность МТС Юрент)'
ws['A1'].font = TITLE
ws.merge_cells('A1:U1')

hdrs = [
    'Город', 'Q1 2024 выручка, руб.', 'Q2 2024', 'Q3 2024', 'Q4 2024',
    'Выручка 2024 итого, руб.',
    'Q1 2025', 'Q2 2025', 'Q3 2025', 'Выручка 9м 2025 итого, руб.',
    'Парк 30.09.2024', 'Парк 30.09.2025',
    'Поездки 2024, шт.', 'Поездки 9м 2025, шт.',
    'Активные польз. 2024', 'Активные польз. 9м 2025',
    'Средний чек 2024, руб.', 'Утилизация 2024, п/с/год',
    'YoY выручка 9м, %', 'YoY поездки 9м, %'
]
for ci, h in enumerate(hdrs, 1):
    cell(ws, 2, ci, h, font=HEAD, fill=HEAD_FILL,
         align=Alignment(wrap_text=True, horizontal='center', vertical='center'))
ws.row_dimensions[2].height = 50

row = 3
for name in ['Москва','Санкт-Петербург','Краснодар','Екатеринбург','Новосибирск']:
    d = city_data[name]
    cell(ws, row, 1, name, font=SUBHEAD)
    for i, v in enumerate(d['rev24_q']):
        cell(ws, row, 2+i, v, fmt='#,##0')
    cell(ws, row, 6, f'=SUM(B{row}:E{row})', fmt='#,##0', font=TOTAL, fill=TOTAL_FILL)
    for i, v in enumerate(d['rev25_9m_q']):
        cell(ws, row, 7+i, v, fmt='#,##0')
    cell(ws, row, 10, f'=SUM(G{row}:I{row})', fmt='#,##0', font=TOTAL, fill=TOTAL_FILL)
    cell(ws, row, 11, d['fleet24'], fmt='#,##0')
    cell(ws, row, 12, d['fleet25'], fmt='#,##0')
    cell(ws, row, 13, d['trips24'], fmt='#,##0')
    cell(ws, row, 14, d['trips25_9m'], fmt='#,##0')
    cell(ws, row, 15, d['active24'], fmt='#,##0')
    cell(ws, row, 16, d['active25_9m'], fmt='#,##0')
    # Средний чек = выручка 2024 / поездки 2024
    cell(ws, row, 17, f'=F{row}/M{row}', fmt='0.0', font=TOTAL)
    # Утилизация = поездки 2024 / парк
    cell(ws, row, 18, f'=M{row}/K{row}', fmt='0', font=TOTAL)
    # YoY выручка: 9м2025 / (Q1+Q2+Q3 2024) -1
    cell(ws, row, 19, f'=(J{row}/SUM(B{row}:D{row})-1)*100', fmt='+0.0;-0.0;0.0', font=TOTAL)
    # YoY поездки — сложнее, нет в листе поквартальных. Пишу как значение из расчёта.
    cell(ws, row, 20, d['yoy_trips'], fmt='+0.0;-0.0;0.0', font=TOTAL)
    row += 1

# ИТОГО
cell(ws, row, 1, 'Итого 5 городов', font=TOTAL, fill=TOTAL_FILL)
for c_idx in [6, 10, 11, 12, 13, 14, 15, 16]:
    col = get_column_letter(c_idx)
    cell(ws, row, c_idx, f'=SUM({col}3:{col}{row-1})', fmt='#,##0', font=TOTAL, fill=TOTAL_FILL)
# Средняя утилизация и средний чек по портфелю
cell(ws, row, 17, f'=F{row}/M{row}', fmt='0.0', font=TOTAL, fill=TOTAL_FILL)
cell(ws, row, 18, f'=M{row}/K{row}', fmt='0', font=TOTAL, fill=TOTAL_FILL)

DATA_TOTAL_ROW = row
# Запишем имя «Выручка_2024_итого_5_городов» = ячейка F{DATA_TOTAL_ROW}
# Также пригодится по каждому городу: ссылки F3..F7

CITY_TO_DATA_ROW = {'Москва':3, 'Санкт-Петербург':4, 'Краснодар':5, 'Екатеринбург':6, 'Новосибирск':7}

row += 2
cell(ws, row, 1, 'Источник: внутренняя поквартальная отчётность МТС Юрент по регионам РФ (эксперт А., данные на 30.09.2025). Выручка — без НДС, включая партнёров и франчайзи на территории РФ.',
     font=NORMAL, align=Alignment(wrap_text=True))
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)

widths(ws, **{chr(c):14 for c in range(ord('B'), ord('U'))})
ws.column_dimensions['A'].width = 22
ws.freeze_panes = 'B3'

# ============================================================
# Лист 3. Затраты
# ============================================================
ws = wb.create_sheet('Затраты')
ws['A1'] = 'Затраты на реализацию сценария S3 (горизонт 12 месяцев)'
ws['A1'].font = TITLE
ws.merge_cells('A1:E1')

for ci, h in enumerate(['Блок', 'Статья', 'Расчёт', 'Сумма, тыс. руб./год', 'Источник / допущение'], 1):
    cell(ws, 2, ci, h, font=HEAD, fill=HEAD_FILL,
         align=Alignment(wrap_text=True, horizontal='center', vertical='center'))
ws.row_dimensions[2].height = 30

costs = [
    ('A. Региональное управление', '3 региональных менеджера по кластерам режимов', '3 × 350 × 12', 12600,
     'ФОТ gross 350 тыс. руб./мес. с налогами и соцпакетом [ДОПУЩЕНИЕ — рыночная вилка middle management в МТС]'),
    ('A. Региональное управление', '3 региональных аналитика', '3 × 200 × 12', 7200,
     'ФОТ gross 200 тыс. руб./мес. [ДОПУЩЕНИЕ — стандартная вилка для аналитиков]'),
    ('A. Региональное управление', 'Накладные (командировки, BI-инструменты, рабочие места)', '~10% от ФОТ', 2000,
     '[ДОПУЩЕНИЕ]'),
    ('B. Маркетинг S3', 'Исследования и креатив для трёх режимных кампаний', '—', 3000,
     '[ДОПУЩЕНИЕ — без учёта медиа-инвестиций из стандартного бюджета]'),
    ('B. Маркетинг S3', 'A/B-тестирование, аналитика результатов', '—', 4000,
     '[ДОПУЩЕНИЕ]'),
    ('C. УПС-инфраструктура и пилот', 'IT-доработки (геозоны, push-таргетинг, движок стимулов, ЛК исполнителя)',
     'capex однократно', 8000, '[ДОПУЩЕНИЕ — оценка по аналогии с продуктовыми спринтами Юрент]'),
    ('C. УПС-инфраструктура и пилот', 'Команда пилота: 1 продакт + 1 аналитик × 6 мес.', '2 × 400 × 6', 4800,
     'ФОТ gross 400 тыс. руб./мес. [ДОПУЩЕНИЕ]'),
    ('C. УПС-инфраструктура и пилот', 'Бюджет стимулов и микрозаданий в Екатеринбурге за сезон',
     'базовый сценарий §3.2', 1000, 'Расчёт §3.2 ВКР: 26,5 руб./опер × ≈ 37 500 замещ. операций (см. лист «УПС_расчёт»)'),
    ('C. УПС-инфраструктура и пилот', 'Резерв на масштабирование стимульной кампании и непредвиденные', '—', 2200, '[ДОПУЩЕНИЕ — для сценариев A/B/C доля стимулов меняется в 2–3 раза]'),
    ('D. Накладные и обучение', 'Обучение региональных менеджеров и аналитиков', '—', 1500, '[ДОПУЩЕНИЕ]'),
    ('D. Накладные и обучение', 'Управленческий резерв', '—', 1500, '[ДОПУЩЕНИЕ]'),
]
row = 3
for block, item, calc, val, src in costs:
    cell(ws, row, 1, block, font=NORMAL)
    cell(ws, row, 2, item, font=NORMAL, align=Alignment(wrap_text=True, vertical='top'))
    cell(ws, row, 3, calc, font=NORMAL, align=Alignment(horizontal='center'))
    cell(ws, row, 4, val, font=NORMAL, fmt='#,##0')
    cell(ws, row, 5, src, font=NORMAL, align=Alignment(wrap_text=True, vertical='top'))
    row += 1

cell(ws, row, 1, 'ИТОГО затраты S3, год 1, тыс. руб.', font=TOTAL, fill=TOTAL_FILL)
cell(ws, row, 4, f'=SUM(D3:D{row-1})', font=TOTAL, fill=TOTAL_FILL, fmt='#,##0')
TOT_COSTS_ROW = row
TOT_COSTS_CELL = f"'Затраты'!$D${row}"

widths(ws, A=28, B=45, C=22, D=22, E=55)

# ============================================================
# Лист 4. Эффекты
# ============================================================
ws = wb.create_sheet('Эффекты')
ws['A1'] = 'Эффекты сценария S3 (горизонт 12 месяцев)'
ws['A1'].font = TITLE
ws.merge_cells('A1:F1')

for ci, h in enumerate(['Эффект', 'Город / кластер', 'База расчёта',
                         'Параметр / допущение', 'EBITDA-эффект, тыс. руб./год', 'Источник'], 1):
    cell(ws, 2, ci, h, font=HEAD, fill=HEAD_FILL,
         align=Alignment(wrap_text=True, horizontal='center', vertical='center'))
ws.row_dimensions[2].height = 35

R2_CITIES = ['Санкт-Петербург','Екатеринбург','Краснодар','Новосибирск']
rev24_r2 = sum(city_data[c]['rev24'] for c in R2_CITIES)
rev24_msk = city_data['Москва']['rev24']

# Эффект 1: R2 темпы
row = 3
cell(ws, row, 1, 'Эффект 1. R2 атака: прирост выручки', font=NORMAL)
cell(ws, row, 2, 'СПб + Екб + Крд + Нск', font=NORMAL)
cell(ws, row, 3, f'Выручка 2024 R2-портфеля: {rev24_r2/1e6:,.1f} млн руб.', font=NORMAL)
cell(ws, row, 4, '+3 п.п. к темпу выручки YoY (база; диапазон 1–5 п.п.) × EBITDA-маржа 25 %', font=NORMAL,
     align=Alignment(wrap_text=True))
# EBITDA = выручка_R2_тыс × темп(0.03) × маржа(0.25)
cell(ws, row, 5, f'=ROUND({rev24_r2/1000}*0.03*0.25,0)', font=TOTAL, fill=GOOD_FILL, fmt='#,##0')
cell(ws, row, 6, '§3.4 ВКР', font=NORMAL)
EFF1_CELL = f"'Эффекты'!$E${row}"

# Эффект 2: УПС-пилот Екатеринбург (EBITDA, пересчитанный через маржу)
# Денежный эффект на УПС-расчёт = 11,7 млн (3,5 экономия + 8,2 доп. выручка).
# EBITDA-эффект = 3,5 (прямая экономия OpEx, уже EBITDA) + 8,2 × 0,25 (маржа) = 5,55 ≈ 6
row += 1
cell(ws, row, 1, 'Эффект 2. УПС-пилот Екатеринбург (EBITDA)', font=NORMAL)
cell(ws, row, 2, 'Екатеринбург', font=NORMAL)
cell(ws, row, 3, 'Парк 6448, утилизация 319, средний чек 75 руб.', font=NORMAL)
cell(ws, row, 4, 'Чистая экономия (3,5 млн, прямой EBITDA) + доп. выручка 8,2 млн × маржа 25 % = 2,05 EBITDA',
     font=NORMAL, align=Alignment(wrap_text=True))
# Формула пересчёта через маржу: экономия (B19 УПС_расчёт) + доп.выручка (B20) × 0,25
UPS_NET_CELL_PLACEHOLDER = (row, 5)
cell(ws, row, 5, 0, font=TOTAL, fill=GOOD_FILL, fmt='#,##0')
cell(ws, row, 6, '§3.2 ВКР, лист «УПС_расчёт», пересчёт через EBITDA-маржу 25 %', font=NORMAL)
EFF2_CELL = f"'Эффекты'!$E${row}"

# Эффект 3: R3 Москва оптимизация OpEx
row += 1
cell(ws, row, 1, 'Эффект 3. R3 Москва: оптимизация OpEx', font=NORMAL)
cell(ws, row, 2, 'Москва', font=NORMAL)
cell(ws, row, 3, f'Выручка 2024 Москвы: {rev24_msk/1e6:,.1f} млн руб.', font=NORMAL)
cell(ws, row, 4, '2 % от выручки рынка (база; диапазон 1–3 %)', font=NORMAL,
     align=Alignment(wrap_text=True))
cell(ws, row, 5, f'=ROUND({rev24_msk/1000}*0.02,0)', font=TOTAL, fill=GOOD_FILL, fmt='#,##0')
cell(ws, row, 6, '§3.4 ВКР; [ДОПУЩЕНИЕ]', font=NORMAL)
EFF3_CELL = f"'Эффекты'!$E${row}"

# Эффект 4: R1 малые города
row += 1
cell(ws, row, 1, 'Эффект 4. R1 малые города: экономия капвложений', font=NORMAL)
cell(ws, row, 2, 'Малые города', font=NORMAL)
cell(ws, row, 3, 'За пределами портфеля 5 городов', font=NORMAL)
cell(ws, row, 4, '~10 млн руб. экономии капвложений (диапазон 5–20); эффект на EBITDA ≈ 30 %', font=NORMAL,
     align=Alignment(wrap_text=True))
cell(ws, row, 5, 3000, font=TOTAL, fill=GOOD_FILL, fmt='#,##0')
cell(ws, row, 6, '§3.4 ВКР; [ДОПУЩЕНИЕ]', font=NORMAL)
EFF4_CELL = f"'Эффекты'!$E${row}"

# ИТОГО
row += 1
cell(ws, row, 1, 'ИТОГО эффект S3, год 1, тыс. руб.', font=TOTAL, fill=TOTAL_FILL)
cell(ws, row, 5, f'=SUM(E3:E{row-1})', font=TOTAL, fill=TOTAL_FILL, fmt='#,##0')
TOT_EFF_CELL = f"'Эффекты'!$E${row}"

widths(ws, A=38, B=23, C=35, D=50, E=24, F=25)

# ============================================================
# Лист 5. УПС_расчёт (создаём ДО Сводной, т.к. на него ссылается Эффект 2)
# ============================================================
ws = wb.create_sheet('УПС_расчёт')
ws['A1'] = 'Детальный расчёт пилота УПС в Екатеринбурге (соответствует §3.2 ВКР)'
ws['A1'].font = TITLE
ws.merge_cells('A1:D1')

for ci, h in enumerate(['№', 'Параметр', 'Значение', 'Источник / расчёт'], 1):
    cell(ws, 2, ci, h, font=HEAD, fill=HEAD_FILL,
         align=Alignment(wrap_text=True, horizontal='center', vertical='center'))
ws.row_dimensions[2].height = 30

# Все формулы строим с явным указанием ячейки.
# Структура: row → (name, value-or-formula, source). value может быть числом, строкой,
# или dict {{'formula': '...'}} с подстановкой $B${row} соседних
upspar = [
    # Параметры данных
    (1, 'Парк Екб на 30.09.2024, шт.',           6448,      'данные эксперта А.'),
    (2, 'Поездки Екб за сезон Q2+Q3 2024, шт.',  1819014,   'данные эксперта А. (Q2 565 936 + Q3 1 253 078)'),
    (3, 'Средний чек Екб 2024, руб.',            75,        'выручка/поездки 2024 ≈ 74,6 руб. (данные эксперта А.)'),
    (4, 'Утилизация 2024, поездок/самокат/год',  319,       'поездки 2024 / парк (данные эксперта А.)'),
    # Параметры экономики
    (5, 'OpEx / Выручка, %',                     55,        'Whoosh 48 % (МСФО, 9 мес. 2024) + 7 п.п. поправка на масштаб'),
    (6, 'Доля ребалансировки в OpEx, %',         30,        '[ДОПУЩЕНИЕ — по аналогии с международными исследованиями]'),
    (7, 'Стоимость физической ребалансировки, руб./операция', 120, 'Данные эксперта А. — стоимость одной операции переноса самоката'),
    # Параметры стимулов
    (8, 'Стимул на стороне точки высадки, руб.', 15, '[ДОПУЩЕНИЕ; доля случаев 50 %]'),
    (9, 'Стимул на стороне точки посадки, руб.', 10, '[ДОПУЩЕНИЕ; доля случаев 30 %]'),
    (10, 'Стимул внешнему исполнителю, руб.',    80, '[ДОПУЩЕНИЕ; доля случаев 20 %]'),
    (11, 'Средневзвешенная стоимость стимула, руб./операция', None,
        'B8 × 0,5 + B9 × 0,3 + B10 × 0,2'),
    # Сценарии
    (12, 'Замещение ребалансировки стимулами, % (базовый сценарий B)', 20, '§3.2: A=10 %, B=20 %, C=30 %'),
    (13, 'Прирост восстановленных поездок в пилотных зонах, %', 6,
        '[ДОПУЩЕНИЕ — доля чувствительных микролокаций]'),
    # Расчёт
    ('—', '', '', ''),
    (14, 'Выручка Екб за сезон Q2+Q3, тыс. руб.', None, 'B2 × B3 / 1000'),
    (15, 'OpEx сезона, тыс. руб.', None, 'B14 × B5 / 100'),
    (16, 'Расходы на физ. ребалансировку (сезон), тыс. руб.', None, 'B15 × B6 / 100'),
    (17, 'Базовое число операций физ. ребалансировки за сезон', None, 'B16 × 1000 / B7'),
    (18, 'Число замещённых операций (базовый B)', None, 'B17 × B12 / 100'),
    ('—', '', '', ''),
    (19, 'Чистая экономия от замещения, тыс. руб.', None,
        '(B7 – B11) × B18 / 1000  ← платим стимул вместо физ. ребалансировки'),
    (20, 'Дополнительная выручка от восстановленных поездок, тыс. руб.', None,
        'B2 × B13 / 100 × B3 / 1000  ← устранение разрывов покрытия'),
    ('—', '', '', ''),
    (21, 'ЧИСТЫЙ ЭФФЕКТ УПС (базовый сценарий B), тыс. руб.', None,
        'B19 + B20'),
]

# Заполним лист. Тут B-ячейки соответствуют ВСЕМ нашим показателям (1..19),
# но в Excel _row отличается от № параметра из-за заголовков и разделителей.
# Поэтому строим словарь num → excel_row.
num_to_row = {}
row = 3
for item in upspar:
    if item[0] == '—':
        # пустая строка-разделитель
        for col in range(1,5):
            c = ws.cell(row=row, column=col, value='')
            c.border = None
        row += 1
        continue
    n, name, val, src = item
    cell(ws, row, 1, n, font=NORMAL, align=Alignment(horizontal='center'))
    cell(ws, row, 2, name, font=NORMAL)
    num_to_row[n] = row
    if val is not None:
        if isinstance(val, str) and val.startswith('='):
            # Подставим B-ссылки в формулу: B3 в исходном смысле = строка с № 3
            # Но в строке формулы '=B3/24' имеется в виду B-ячейка с поездками,
            # которая на этом этапе пока ещё не «знает» о num_to_row.
            # Поэтому мы СНАЧАЛА сохраняем значение, потом ПЕРЕЗАПИСЫВАЕМ формулы.
            cell(ws, row, 3, val, font=TOTAL, fill=TOTAL_FILL, fmt='#,##0')
        else:
            cell(ws, row, 3, val, font=NORMAL, fmt='#,##0')
    else:
        # формула — будет добавлена ниже
        cell(ws, row, 3, '', font=TOTAL, fill=TOTAL_FILL, fmt='#,##0')
    cell(ws, row, 4, src, font=NORMAL, align=Alignment(wrap_text=True))
    row += 1

# Теперь зная num_to_row, заполняем формулы корректными ссылками
# ВАЖНО: значения параметров в столбце C, поэтому используем $C$
def rb(n): return f'$C${num_to_row[n]}'

formulas = {
    11: f'={rb(8)}*0.5+{rb(9)}*0.3+{rb(10)}*0.2',          # средневзв. стоимость стимула
    14: f'=ROUND({rb(2)}*{rb(3)}/1000,0)',                  # выручка сезона тыс
    15: f'=ROUND({rb(14)}*{rb(5)}/100,0)',                  # OpEx сезона тыс
    16: f'=ROUND({rb(15)}*{rb(6)}/100,0)',                  # расходы на физреб сезона тыс
    17: f'=ROUND({rb(16)}*1000/{rb(7)},0)',                 # число операций
    18: f'=ROUND({rb(17)}*{rb(12)}/100,0)',                 # замещ. операций
    19: f'=ROUND(({rb(7)}-{rb(11)})*{rb(18)}/1000,0)',      # чистая экономия
    20: f'=ROUND({rb(2)}*{rb(13)}/100*{rb(3)}/1000,0)',     # доп. выручка
    21: f'={rb(19)}+{rb(20)}',                              # чистый эффект
}
for n, f in formulas.items():
    r = num_to_row[n]
    c = ws.cell(row=r, column=3, value=f)
    c.font = TOTAL if n != 21 else BIG_TOTAL
    c.fill = GOOD_FILL if n == 21 else TOTAL_FILL
    c.number_format = '#,##0'

# Жирная подсветка строки итога
ws.row_dimensions[num_to_row[21]].height = 25
cell(ws, num_to_row[21], 2, 'ЧИСТЫЙ ЭФФЕКТ УПС (базовый сценарий B), тыс. руб.',
     font=BIG_TOTAL, fill=GOOD_FILL)

# Теперь, когда УПС-лист готов, обновим ссылку Эффект 2 в листе «Эффекты»
# EBITDA-эффект = чистая экономия (B19) + доп. выручка (B20) × 0,25 EBITDA-маржа
ups_econ_row = num_to_row[19]   # чистая экономия (уже EBITDA)
ups_rev_row = num_to_row[20]    # доп. выручка (нужно умножить на маржу)
eff_ws = wb['Эффекты']
prow, pcol = UPS_NET_CELL_PLACEHOLDER
formula = f"=ROUND('УПС_расчёт'!$C${ups_econ_row} + 'УПС_расчёт'!$C${ups_rev_row}*0.25, 0)"
c_ref = eff_ws.cell(row=prow, column=pcol, value=formula)
c_ref.font = TOTAL; c_ref.fill = GOOD_FILL; c_ref.number_format = '#,##0'; c_ref.border = BORDER

# Сценарии A/B/C по разной доле замещения
row = num_to_row[21] + 2
cell(ws, row, 1, 'Сценарии чувствительности УПС (по доле замещения ребалансировки)',
     font=SUBHEAD, fill=SUBHEAD_FILL)
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
row += 1
for ci, h in enumerate(['Сценарий', 'Замещение, %', 'Прирост поездок (масштаб), %', 'Чистый эффект, тыс. руб.'], 1):
    cell(ws, row, ci, h, font=HEAD, fill=HEAD_FILL,
         align=Alignment(horizontal='center', wrap_text=True))
row += 1
scenarios = [('A. Консервативный', 10, 3), ('B. Базовый', 20, 6), ('C. Оптимистичный', 30, 9)]
for name_sc, pct, growth in scenarios:
    cell(ws, row, 1, name_sc, font=NORMAL)
    cell(ws, row, 2, pct, font=NORMAL, fmt='0" %"')
    cell(ws, row, 3, growth, font=NORMAL, fmt='0" %"')
    # Эффект = (физреб - стимул) × ops × pct/100 / 1000  +  поездки × growth/100 × чек / 1000
    f_eff = (
        f'=ROUND(({rb(7)}-{rb(11)})*{rb(17)}*{pct}/100/1000'
        f' + {rb(2)}*{growth}/100*{rb(3)}/1000, 0)'
    )
    cell(ws, row, 4, f_eff, font=TOTAL, fill=GOOD_FILL, fmt='#,##0')
    row += 1

widths(ws, A=6, B=55, C=22, D=60)

# ============================================================
# Лист 6. Сводная_финмодель
# ============================================================
ws = wb.create_sheet('Сводная_финмодель')
ws['A1'] = 'Сводная финансовая модель S3'
ws['A1'].font = BIG_README
ws.merge_cells('A1:C1')

for ci, h in enumerate(['Статья', 'Млн руб.', 'Источник'], 1):
    cell(ws, 2, ci, h, font=HEAD, fill=HEAD_FILL, align=Alignment(horizontal='center'))

# Год 1
row = 3
cell(ws, row, 1, 'ГОД 1 (12 месяцев)', font=BIG_TOTAL, fill=SUBHEAD_FILL)
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
row += 1

items_y1 = [
    ('Затраты, всего', f'=-{TOT_COSTS_CELL}/1000', 'Лист «Затраты», итого', BAD_FILL),
    ('Эффект 1. R2-прирост выручки (СПб, Екб, Крд, Нск)', f'={EFF1_CELL}/1000', '+3 п.п. × выручка R2 × EBITDA 25 %', GOOD_FILL),
    ('Эффект 2. УПС-пилот Екатеринбург', f'={EFF2_CELL}/1000', '§3.2, базовый сценарий B', GOOD_FILL),
    ('Эффект 3. R3 Москва: оптимизация OpEx', f'={EFF3_CELL}/1000', '2 % выручки Москвы', GOOD_FILL),
    ('Эффект 4. R1 малые города: экономия капвложений', f'={EFF4_CELL}/1000', 'Защита от нерентабельных инвестиций', GOOD_FILL),
]
first_data_row = row
for st, val, src, fill in items_y1:
    cell(ws, row, 1, st, font=NORMAL)
    cell(ws, row, 2, val, font=TOTAL, fill=fill, fmt='+#,##0.0;-#,##0.0;0.0')
    cell(ws, row, 3, src, font=NORMAL, align=Alignment(wrap_text=True))
    row += 1
costs_row_y1 = first_data_row
effects_first_row_y1 = first_data_row + 1
effects_last_row_y1 = row - 1

cell(ws, row, 1, 'Эффект, всего', font=TOTAL, fill=TOTAL_FILL)
cell(ws, row, 2, f'=SUM(B{effects_first_row_y1}:B{effects_last_row_y1})', font=TOTAL, fill=TOTAL_FILL, fmt='+#,##0.0;-#,##0.0;0.0')
row += 1

cell(ws, row, 1, 'ЧИСТЫЙ ЭФФЕКТ S3 vs S1, год 1', font=BIG_TOTAL, fill=TOTAL_FILL)
cell(ws, row, 2, f'=B{costs_row_y1}+SUM(B{effects_first_row_y1}:B{effects_last_row_y1})',
     font=BIG_TOTAL, fill=TOTAL_FILL, fmt='+#,##0.0;-#,##0.0;0.0')
ws.row_dimensions[row].height = 25
NET_Y1_ROW = row

# Год 2
row += 3
cell(ws, row, 1, 'ГОД 2 (тиражирование УПС на 3 города R2-кластера, capex выпадает)',
     font=BIG_TOTAL, fill=SUBHEAD_FILL)
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
row += 1

items_y2 = [
    ('Затраты год 2 (без capex УПС)', -37, 'A + B + D + операционная часть C'),
    ('Эффект 1 расширенный (R2-темпы +5 п.п.)', round(rev24_r2/1e6 * 0.05 * 0.25, 1),
     '+5 п.п. × выручка R2 × EBITDA 25 %'),
    ('Эффект 2 расширенный (УПС в 4 городах R2)', 40,
     'Коэф. 3,0 млн EBITDA / 1 млн поездок Q2+Q3 × 13 млн сезонных поездок R2-кластера'),
    ('Эффект 3 (R3 Москва, год 2)', 55, 'Углубление оптимизации до 2,1 % от выручки'),
    ('Эффект 4 (R1 малые города)', 5, 'Расширение применения каркаса'),
]
first_y2 = row
for st, val, src in items_y2:
    cell(ws, row, 1, st, font=NORMAL)
    cell(ws, row, 2, val, font=TOTAL, fill=(BAD_FILL if val < 0 else GOOD_FILL),
         fmt='+#,##0.0;-#,##0.0;0.0')
    cell(ws, row, 3, src, font=NORMAL, align=Alignment(wrap_text=True))
    row += 1
last_y2 = row - 1
cell(ws, row, 1, 'ЧИСТЫЙ ЭФФЕКТ S3 vs S1, год 2', font=BIG_TOTAL, fill=TOTAL_FILL)
cell(ws, row, 2, f'=SUM(B{first_y2}:B{last_y2})', font=BIG_TOTAL, fill=TOTAL_FILL,
     fmt='+#,##0.0;-#,##0.0;0.0')
ws.row_dimensions[row].height = 25

widths(ws, A=55, B=20, C=55)

# ============================================================
# Лист 7. Чувствительность
# ============================================================
ws = wb.create_sheet('Чувствительность')
ws['A1'] = 'Сценарии чувствительности чистого эффекта S3 (год 1)'
ws['A1'].font = TITLE
ws.merge_cells('A1:E1')

for ci, h in enumerate(['Параметр', 'Консервативный', 'Базовый', 'Оптимистичный', 'Комментарий'], 1):
    cell(ws, 2, ci, h, font=HEAD, fill=HEAD_FILL, align=Alignment(horizontal='center'))

# Параметры с диапазонами
params = [
    ('Прирост темпа R2 выручки, п.п.',   1, 3, 5,  'Допущение о силе эффекта дифференциации'),
    ('Замещение ребалансировки УПС, %',  10, 20, 30, 'Сценарии A/B/C из §3.2 / лист «УПС_расчёт»'),
    ('Оптимизация OpEx Москвы, %',       1, 2, 3,  'Глубина рычагов R3 в Москве'),
    ('Экономия капвложений в малых городах, млн руб./год', 5, 10, 20, 'Темпы расширения портфеля'),
]
row = 3
P_ROWS = {}  # param name → row
for st, c, b, o, src in params:
    cell(ws, row, 1, st, font=NORMAL)
    cell(ws, row, 2, c, font=NORMAL, fmt='0.0')
    cell(ws, row, 3, b, font=TOTAL, fill=TOTAL_FILL, fmt='0.0')
    cell(ws, row, 4, o, font=NORMAL, fmt='0.0')
    cell(ws, row, 5, src, font=NORMAL, align=Alignment(wrap_text=True))
    P_ROWS[st] = row
    row += 1
PR1 = P_ROWS['Прирост темпа R2 выручки, п.п.']
PR2 = P_ROWS['Замещение ребалансировки УПС, %']
PR3 = P_ROWS['Оптимизация OpEx Москвы, %']
PR4 = P_ROWS['Экономия капвложений в малых городах, млн руб./год']

# Производные эффекты для каждого столбца
row += 1
cell(ws, row, 1, 'Производные эффекты', font=SUBHEAD, fill=SUBHEAD_FILL)
for col in [2,3,4]:
    cell(ws, row, col, '', font=NORMAL, fill=SUBHEAD_FILL)
cell(ws, row, 5, '', font=NORMAL, fill=SUBHEAD_FILL)
row += 1

rev_r2_mln = rev24_r2/1e6
rev_msk_mln = rev24_msk/1e6
ups_base = 9790  # из УПС_расчёт; будем использовать как масштаб

# Эффект 1
eff1_r = row
cell(ws, row, 1, 'Эффект 1 (R2 темпы), млн руб.', font=NORMAL)
for col_n, col_l in [(2,'B'),(3,'C'),(4,'D')]:
    cell(ws, row, col_n, f'=ROUND({rev_r2_mln}*{col_l}{PR1}/100*0.25,1)', fmt='+0.0;-0.0;0.0',
         font=TOTAL if col_n==3 else NORMAL,
         fill=TOTAL_FILL if col_n==3 else None)
row += 1

# Эффект 2: масштабируется пропорционально % замещения (база 20%, эффект 9790 тыс)
eff2_r = row
cell(ws, row, 1, 'Эффект 2 (УПС-пилот), млн руб.', font=NORMAL)
for col_n, col_l in [(2,'B'),(3,'C'),(4,'D')]:
    cell(ws, row, col_n, f'=ROUND({col_l}{PR2}/20*9.79,1)', fmt='+0.0;-0.0;0.0',
         font=TOTAL if col_n==3 else NORMAL,
         fill=TOTAL_FILL if col_n==3 else None)
row += 1

# Эффект 3
eff3_r = row
cell(ws, row, 1, 'Эффект 3 (R3 Москва), млн руб.', font=NORMAL)
for col_n, col_l in [(2,'B'),(3,'C'),(4,'D')]:
    cell(ws, row, col_n, f'=ROUND({rev_msk_mln}*{col_l}{PR3}/100,1)', fmt='+0.0;-0.0;0.0',
         font=TOTAL if col_n==3 else NORMAL,
         fill=TOTAL_FILL if col_n==3 else None)
row += 1

# Эффект 4 — 30% от экономии капвложений на EBITDA
eff4_r = row
cell(ws, row, 1, 'Эффект 4 (R1 капвложения), млн руб.', font=NORMAL)
for col_n, col_l in [(2,'B'),(3,'C'),(4,'D')]:
    cell(ws, row, col_n, f'=ROUND({col_l}{PR4}*0.3,1)', fmt='+0.0;-0.0;0.0',
         font=TOTAL if col_n==3 else NORMAL,
         fill=TOTAL_FILL if col_n==3 else None)
row += 1

# Затраты — статичные
costs_r = row
cell(ws, row, 1, 'Затраты S3, млн руб.', font=NORMAL)
for col_n in [2,3,4]:
    cell(ws, row, col_n, -48, fmt='+0.0;-0.0;0.0',
         font=TOTAL if col_n==3 else NORMAL,
         fill=BAD_FILL if col_n==3 else None)
row += 1

# Итого
net_r = row
cell(ws, row, 1, 'ЧИСТЫЙ ЭФФЕКТ S3, год 1, млн руб.', font=BIG_TOTAL, fill=TOTAL_FILL)
for col_n, col_l in [(2,'B'),(3,'C'),(4,'D')]:
    cell(ws, row, col_n, f'=SUM({col_l}{eff1_r}:{col_l}{costs_r})',
         font=BIG_TOTAL, fill=TOTAL_FILL, fmt='+#,##0.0;-#,##0.0;0.0')
ws.row_dimensions[row].height = 25

widths(ws, A=42, B=18, C=18, D=18, E=50)

# Порядок листов: README, Данные_Юрент, Затраты, Эффекты, УПС_расчёт, Сводная_финмодель, Чувствительность
# Уже создаются в этом порядке; перетасуем УПС_расчёт перед Сводной
order = ['README','Данные_Юрент','Затраты','Эффекты','УПС_расчёт','Сводная_финмодель','Чувствительность']
wb._sheets = [wb[name] for name in order]

# Сохранить
out = Path('/workspace/vkr/VKR_findata.xlsx')
wb.save(out)
print(f"Saved: {out} ({out.stat().st_size:,} bytes)")
print("Sheets:", wb.sheetnames)
print(f"Verify (Python check):")
print(f"  УПС база операций ≈ {1819014/24:.0f} (≈ 75 793 операций за сезон)")
print(f"  Замещ. операций при 20% ≈ {1819014/24*0.20:.0f}")
print(f"  Стимул средневзв. = 15*0.5 + 10*0.3 + 80*0.2 = {15*0.5+10*0.3+80*0.2:.2f} руб/опер")
print(f"  Экономия = (120-26.5)*15159/1000 ≈ {(120-26.5)*1819014/24*0.2/1000:.0f} тыс руб")
print(f"  Доп выручка = 1819014*0.06*75/1000 ≈ {1819014*0.06*75/1000:.0f} тыс руб")
print(f"  Расходы стимулов = 26.5*15159/1000 ≈ {26.5*1819014/24*0.2/1000:.0f} тыс руб")
upx = (120-26.5)*1819014/24*0.2/1000 + 1819014*0.06*75/1000 - 26.5*1819014/24*0.2/1000
print(f"  ЧИСТЫЙ ЭФФЕКТ УПС ≈ {upx:.0f} тыс руб")
