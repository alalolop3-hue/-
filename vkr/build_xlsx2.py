#!/usr/bin/env python3
"""
Excel-приложение к §3.4 — версия 2 с предвычисленными значениями.
Вместо формул в ячейках хранятся РАССЧИТАННЫЕ ЗНАЧЕНИЯ (numbers).
Формула как текст показана в столбце «Расчёт / источник» для прозрачности.
Это устраняет проблемы с локалью Excel.
"""
import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CITIES = [('Москва',1),('Санкт-Петербург',4),('Краснодар',3),('Екатеринбург',12),('Новосибирск',7)]
def num(s):
    s = (s or '').strip().replace(' ','').replace('\xa0','')
    try: return int(s)
    except:
        try: return float(s)
        except: return 0

rows=[]
with open('/home/ubuntu/.cursor/projects/workspace/uploads/kicksharing_v2_b578.csv','r',encoding='utf-8') as f:
    for r in csv.reader(f, delimiter=';'): rows.append(r)

city_data={}
for name, rid in CITIES:
    for rr in rows[3:]:
        if rr and rr[0].strip()==str(rid):
            r=rr; break
    rev24_q=[num(r[i]) for i in range(1,5)]
    rev25_9m_q=[num(r[i]) for i in range(5,8)]
    fleet24=num(r[8]); fleet25=num(r[9])
    active24=num(r[17]); active25_9m=num(r[18])
    trips24_q=[num(r[i]) for i in range(21,25)]
    trips25_9m_q=[num(r[i]) for i in range(25,28)]
    rev24=sum(rev24_q); trips24=sum(trips24_q)
    rev24_9m=sum(rev24_q[:3]); rev25_9m=sum(rev25_9m_q)
    trips24_9m=sum(trips24_q[:3]); trips25_9m=sum(trips25_9m_q)
    city_data[name]=dict(
        rev24_q=rev24_q,rev25_9m_q=rev25_9m_q,fleet24=fleet24,fleet25=fleet25,
        active24=active24,active25_9m=active25_9m,trips24_q=trips24_q,trips25_9m_q=trips25_9m_q,
        rev24=rev24,trips24=trips24,
        avg_check=rev24/trips24 if trips24 else 0,
        util24=trips24/fleet24 if fleet24 else 0,
        yoy_rev=(rev25_9m/rev24_9m-1)*100 if rev24_9m else 0,
        yoy_trips=(trips25_9m/trips24_9m-1)*100 if trips24_9m else 0,
    )

# Стили
HEAD=Font(name='Times New Roman',size=12,bold=True,color='FFFFFF')
HEAD_FILL=PatternFill('solid',fgColor='2F5496')
SUBHEAD=Font(name='Times New Roman',size=11,bold=True)
SUBHEAD_FILL=PatternFill('solid',fgColor='D9E2F3')
NORMAL=Font(name='Times New Roman',size=11)
TOTAL=Font(name='Times New Roman',size=11,bold=True)
TOTAL_FILL=PatternFill('solid',fgColor='FFF2CC')
GOOD_FILL=PatternFill('solid',fgColor='C6EFCE')
BAD_FILL=PatternFill('solid',fgColor='FFC7CE')
BIG_TOTAL=Font(name='Times New Roman',size=13,bold=True)
TITLE=Font(name='Times New Roman',size=14,bold=True)
BORDER=Border(left=Side(style='thin',color='999999'),right=Side(style='thin',color='999999'),
              top=Side(style='thin',color='999999'),bottom=Side(style='thin',color='999999'))

def cell(ws,row,col,value,font=NORMAL,fill=None,align=None,border=BORDER,fmt=None):
    # Если value — строка, начинающаяся с '=', openpyxl запишет её как формулу.
    # Чтобы строки-объяснения с '=' оставались просто текстом, заменим на префикс «расчёт: »
    if isinstance(value, str) and value.startswith('='):
        value = 'расчёт: ' + value[1:].lstrip()
    c=ws.cell(row=row,column=col,value=value)
    c.font=font
    if fill: c.fill=fill
    if align: c.alignment=align
    if border: c.border=border
    if fmt: c.number_format=fmt
    return c

def widths(ws,**kwargs):
    for col,w in kwargs.items():
        ws.column_dimensions[col].width=w

wb=Workbook()
del wb[wb.active.title]

# === Лист README ===
ws=wb.create_sheet('README')
ws['A1']='Финансовая модель сценария S3 — приложение к ВКР'
ws['A1'].font=TITLE
ws.merge_cells('A1:E1')
ws.column_dimensions['A'].width=100
descr=['','Автор: [ФИО студента], ВКР НИУ ВШЭ, 2026 г.','',
       'Финансовая модель сценария S3 (диверсификация режимов МТС Юрент) на горизонте 12 и 24 мес.',
       'Все значения предвычислены и записаны в ячейках напрямую — формулы в колонке «Расчёт» как текст',
       'для прозрачности (избегаем проблем с русской локалью Excel).','',
       'Источники данных:',
       '  • Внутренняя поквартальная отчётность МТС Юрент по 86 субъектам РФ за 2024 г. и 9 мес. 2025 г.',
       '    (эксперт А., руководитель подразделения M&A ПАО «МТС»).',
       '  • Публичная отчётность ПАО «ВУШ Холдинг» по МСФО (9 мес. 2024) — бенчмарк OpEx и EBITDA.',
       '  • Расчёт по проектному инструменту УПС (§3.2 ВКР).','',
       'Ключевые показатели сценария S3 (см. лист «Сводная_финмодель»):',
       '  • Год 1: затраты –48 млн руб., эффекты +73 млн руб., чистый эффект +25 млн руб.',
       '  • Год 2: затраты –37 млн руб., эффекты +115–135 млн руб., чистый эффект +80…+100 млн руб.',
       '  • Сценарии чувствительности год 1: –10 / +25 / +85 млн руб.',
       '','Допущения, требующие калибровки на этапе пилота:',
       '  — прирост темпа выручки R2-городов от дифференциации: 1–5 п.п. (база 3);',
       '  — доля замещения физической ребалансировки стимулами: 10–30 % (база 20 %);',
       '  — стоимость физической ребалансировки 120 руб./операция (данные эксперта А.);',
       '  — доля ребалансировки в OpEx ≈ 30 % (по аналогии с международными данными);',
       '  — снижение OpEx Москвы в R3: 1–3 % от выручки (база 2 %).',
       ]
for i,line in enumerate(descr,start=2):
    c=ws.cell(row=i,column=1,value=line); c.font=NORMAL

# === Лист Данные_Юрент ===
ws=wb.create_sheet('Данные_Юрент')
ws['A1']='Сырые данные по 5 городам портфеля (источник: эксперт А., внутренняя отчётность МТС Юрент)'
ws['A1'].font=TITLE
ws.merge_cells('A1:U1')
hdrs=['Город','Q1 2024 выручка, руб.','Q2 2024','Q3 2024','Q4 2024','Выручка 2024 итого',
      'Q1 2025','Q2 2025','Q3 2025','Выручка 9м 2025 итого','Парк 30.09.2024','Парк 30.09.2025',
      'Поездки 2024','Поездки 9м 2025','Активные 2024','Активные 9м 2025',
      'Средний чек 2024, руб.','Утилизация 2024, п/с/год','YoY выручка 9м, %','YoY поездки 9м, %']
for ci,h in enumerate(hdrs,1):
    cell(ws,2,ci,h,font=HEAD,fill=HEAD_FILL,align=Alignment(wrap_text=True,horizontal='center',vertical='center'))
ws.row_dimensions[2].height=50
row=3
for name in ['Москва','Санкт-Петербург','Краснодар','Екатеринбург','Новосибирск']:
    d=city_data[name]
    cell(ws,row,1,name,font=SUBHEAD)
    for i,v in enumerate(d['rev24_q']): cell(ws,row,2+i,v,fmt='#,##0')
    cell(ws,row,6,d['rev24'],fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
    for i,v in enumerate(d['rev25_9m_q']): cell(ws,row,7+i,v,fmt='#,##0')
    cell(ws,row,10,sum(d['rev25_9m_q']),fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
    cell(ws,row,11,d['fleet24'],fmt='#,##0')
    cell(ws,row,12,d['fleet25'],fmt='#,##0')
    cell(ws,row,13,d['trips24'],fmt='#,##0')
    cell(ws,row,14,sum(d['trips25_9m_q']),fmt='#,##0')
    cell(ws,row,15,d['active24'],fmt='#,##0')
    cell(ws,row,16,d['active25_9m'],fmt='#,##0')
    cell(ws,row,17,round(d['avg_check'],1),fmt='0.0',font=TOTAL)
    cell(ws,row,18,round(d['util24'],0),fmt='0',font=TOTAL)
    cell(ws,row,19,round(d['yoy_rev'],1),fmt='+0.0;-0.0;0.0',font=TOTAL)
    cell(ws,row,20,round(d['yoy_trips'],1),fmt='+0.0;-0.0;0.0',font=TOTAL)
    row+=1
# Итого
cell(ws,row,1,'Итого 5 городов',font=TOTAL,fill=TOTAL_FILL)
total_rev=sum(city_data[c]['rev24'] for c in city_data)
total_fleet=sum(city_data[c]['fleet24'] for c in city_data)
total_trips=sum(city_data[c]['trips24'] for c in city_data)
cell(ws,row,6,total_rev,fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,10,sum(sum(city_data[c]['rev25_9m_q']) for c in city_data),fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,11,total_fleet,fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,12,sum(city_data[c]['fleet25'] for c in city_data),fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,13,total_trips,fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,14,sum(sum(city_data[c]['trips25_9m_q']) for c in city_data),fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,15,sum(city_data[c]['active24'] for c in city_data),fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,16,sum(city_data[c]['active25_9m'] for c in city_data),fmt='#,##0',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,17,round(total_rev/total_trips,1),fmt='0.0',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,18,round(total_trips/total_fleet,0),fmt='0',font=TOTAL,fill=TOTAL_FILL)
widths(ws,**{chr(c):14 for c in range(ord('B'),ord('U'))})
ws.column_dimensions['A'].width=22
ws.freeze_panes='B3'

# === Лист Затраты ===
ws=wb.create_sheet('Затраты')
ws['A1']='Затраты на реализацию сценария S3 (горизонт 12 месяцев)'
ws['A1'].font=TITLE
ws.merge_cells('A1:E1')
for ci,h in enumerate(['Блок','Статья','Расчёт','Сумма, тыс. руб./год','Источник / допущение'],1):
    cell(ws,2,ci,h,font=HEAD,fill=HEAD_FILL,align=Alignment(wrap_text=True,horizontal='center',vertical='center'))
ws.row_dimensions[2].height=30
costs=[
    ('A. Региональное управление','3 региональных менеджера по кластерам режимов','3 × 350 × 12',12600,'ФОТ gross 350 тыс. руб./мес. [ДОПУЩЕНИЕ — рыночная вилка]'),
    ('A. Региональное управление','3 региональных аналитика','3 × 200 × 12',7200,'ФОТ gross 200 тыс. руб./мес. [ДОПУЩЕНИЕ]'),
    ('A. Региональное управление','Накладные (командировки, BI-инструменты, рабочие места)','~10% от ФОТ',2000,'[ДОПУЩЕНИЕ]'),
    ('B. Маркетинг S3','Исследования и креатив для трёх режимных кампаний','—',3000,'[ДОПУЩЕНИЕ — без медиа-инвестиций]'),
    ('B. Маркетинг S3','A/B-тестирование, аналитика результатов','—',4000,'[ДОПУЩЕНИЕ]'),
    ('C. УПС-инфраструктура и пилот','IT-доработки (геозоны, push-таргетинг, движок стимулов, ЛК исполнителя)','capex однократно',8000,'[ДОПУЩЕНИЕ — оценка по аналогии с продуктовыми спринтами]'),
    ('C. УПС-инфраструктура и пилот','Команда пилота: 1 продакт + 1 аналитик × 6 мес.','2 × 400 × 6',4800,'ФОТ gross 400 тыс. руб./мес. [ДОПУЩЕНИЕ]'),
    ('C. УПС-инфраструктура и пилот','Бюджет стимулов и микрозаданий в Екб за сезон','базовый сценарий §3.2',1000,'§3.2: 26,5 руб./опер × 37 500 замещ. операций'),
    ('C. УПС-инфраструктура и пилот','Резерв на масштабирование стимульной кампании','—',2200,'[ДОПУЩЕНИЕ]'),
    ('D. Накладные и обучение','Обучение региональных менеджеров и аналитиков','—',1500,'[ДОПУЩЕНИЕ]'),
    ('D. Накладные и обучение','Управленческий резерв','—',1500,'[ДОПУЩЕНИЕ]'),
]
row=3; tot_costs=0
for block,item,calc,val,src in costs:
    cell(ws,row,1,block,font=NORMAL)
    cell(ws,row,2,item,font=NORMAL,align=Alignment(wrap_text=True,vertical='top'))
    cell(ws,row,3,calc,font=NORMAL,align=Alignment(horizontal='center'))
    cell(ws,row,4,val,font=NORMAL,fmt='#,##0')
    cell(ws,row,5,src,font=NORMAL,align=Alignment(wrap_text=True,vertical='top'))
    tot_costs+=val; row+=1
cell(ws,row,1,'ИТОГО затраты S3, год 1, тыс. руб.',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,4,tot_costs,font=TOTAL,fill=TOTAL_FILL,fmt='#,##0')
widths(ws,A=28,B=45,C=22,D=22,E=55)

# === Лист УПС_расчёт — С ВЫЧИСЛЕННЫМИ ЗНАЧЕНИЯМИ ===
ws=wb.create_sheet('УПС_расчёт')
ws['A1']='Детальный расчёт пилота УПС в Екатеринбурге (соответствует §3.2 ВКР)'
ws['A1'].font=TITLE
ws.merge_cells('A1:D1')
for ci,h in enumerate(['№','Параметр','Значение','Расчёт / источник'],1):
    cell(ws,2,ci,h,font=HEAD,fill=HEAD_FILL,align=Alignment(wrap_text=True,horizontal='center',vertical='center'))
ws.row_dimensions[2].height=30

# Все значения вычислены заранее
fleet_ekb=6448; trips_ekb=1819014; check_ekb=75; util_ekb=319
opex_share=55; reb_share=30; reb_cost=120
sub_drop=15; sub_pickup=10; sub_external=80
sub_weights=[0.5,0.3,0.2]
sub_avg=sub_drop*sub_weights[0]+sub_pickup*sub_weights[1]+sub_external*sub_weights[2]  # 26.5
pct_base=20; growth_base=6
rev_season=trips_ekb*check_ekb/1000  # 136 426 тыс
opex_season=rev_season*opex_share/100  # 75 034
reb_expense=opex_season*reb_share/100  # 22 510
base_ops=reb_expense*1000/reb_cost  # 187 583
repl_ops=base_ops*pct_base/100  # 37 517
net_economy=(reb_cost-sub_avg)*repl_ops/1000  # (120-26.5)*37517/1000 = 3 508
extra_revenue=trips_ekb*growth_base/100*check_ekb/1000  # 8 186
net_ups=net_economy+extra_revenue  # 11 694

data_lines=[
    (1,'Парк Екб на 30.09.2024, шт.',fleet_ekb,'данные эксперта А.'),
    (2,'Поездки Екб за сезон Q2+Q3 2024, шт.',trips_ekb,'данные эксперта А. (Q2 565 936 + Q3 1 253 078)'),
    (3,'Средний чек Екб 2024, руб.',check_ekb,'выручка/поездки 2024 ≈ 74,6 руб. (данные эксперта А.)'),
    (4,'Утилизация 2024, поездок/самокат/год',util_ekb,'поездки 2024 / парк (данные эксперта А.)'),
    (5,'OpEx / Выручка, %',opex_share,'Whoosh 48 % (МСФО, 9 мес. 2024) + 7 п.п. поправка на масштаб'),
    (6,'Доля ребалансировки в OpEx, %',reb_share,'[ДОПУЩЕНИЕ — по аналогии с международными исследованиями]'),
    (7,'Стоимость физической ребалансировки, руб./операция',reb_cost,'Данные эксперта А. — стоимость одной операции переноса'),
    (8,'Стимул на стороне точки высадки, руб.',sub_drop,'[ДОПУЩЕНИЕ; доля случаев 50 %]'),
    (9,'Стимул на стороне точки посадки, руб.',sub_pickup,'[ДОПУЩЕНИЕ; доля случаев 30 %]'),
    (10,'Стимул внешнему исполнителю, руб.',sub_external,'[ДОПУЩЕНИЕ; доля случаев 20 %]'),
    (11,'Средневзвешенная стоимость стимула, руб./операция',round(sub_avg,1),f'= {sub_drop}×0,5 + {sub_pickup}×0,3 + {sub_external}×0,2 = {sub_avg:.1f}'),
    (12,'Замещение ребалансировки стимулами, % (базовый сценарий B)',pct_base,'§3.2: A=10%, B=20%, C=30%'),
    (13,'Прирост восстановленных поездок в пилотных зонах, %',growth_base,'[ДОПУЩЕНИЕ — доля чувствительных микролокаций]'),
    ('—','','',''),
    (14,'Выручка Екб за сезон Q2+Q3, тыс. руб.',round(rev_season,0),f'= поездки × чек / 1000 = {trips_ekb} × {check_ekb} / 1000 ≈ {rev_season:,.0f}'),
    (15,'OpEx сезона, тыс. руб.',round(opex_season,0),f'= {round(rev_season,0):,.0f} × {opex_share}% ≈ {opex_season:,.0f}'),
    (16,'Расходы на физ. ребалансировку (сезон), тыс. руб.',round(reb_expense,0),f'= {round(opex_season,0):,.0f} × {reb_share}% ≈ {reb_expense:,.0f}'),
    (17,'Базовое число операций физ. ребалансировки за сезон',round(base_ops,0),f'= {round(reb_expense,0):,.0f}×1000 / {reb_cost} ≈ {base_ops:,.0f}'),
    (18,'Число замещённых операций (базовый B)',round(repl_ops,0),f'= {round(base_ops,0):,.0f} × {pct_base}% ≈ {repl_ops:,.0f}'),
    ('—','','',''),
    (19,'Чистая экономия от замещения, тыс. руб.',round(net_economy,0),f'= ({reb_cost}–{sub_avg:.1f}) × {round(repl_ops,0):,.0f} / 1000 ≈ {net_economy:,.0f}'),
    (20,'Дополнительная выручка от восстановленных поездок, тыс. руб.',round(extra_revenue,0),f'= {trips_ekb} × {growth_base}% × {check_ekb} / 1000 ≈ {extra_revenue:,.0f}'),
    ('—','','',''),
    (21,'ЧИСТЫЙ ЭФФЕКТ УПС (денежный, базовый B), тыс. руб.',round(net_ups,0),f'= {net_economy:,.0f} + {extra_revenue:,.0f} ≈ {net_ups:,.0f}'),
    (22,'EBITDA-эффект УПС (для финмодели §3.4), тыс. руб.',round(net_economy+extra_revenue*0.25,0),
     f'= экономия (уже EBITDA) + доп. выручка × 25% маржа = {net_economy:,.0f} + {extra_revenue*0.25:,.0f} ≈ {net_economy+extra_revenue*0.25:,.0f}'),
]

row=3
for item in data_lines:
    if item[0]=='—':
        row+=1
        continue
    n,name,val,src=item
    cell(ws,row,1,n,font=NORMAL,align=Alignment(horizontal='center'))
    is_calc=str(n) in {'11','14','15','16','17','18','19','20'}
    is_final=str(n) in {'21','22'}
    f=BIG_TOTAL if is_final else (TOTAL if is_calc else NORMAL)
    fill=GOOD_FILL if is_final else (TOTAL_FILL if is_calc else None)
    cell(ws,row,2,name,font=BIG_TOTAL if is_final else NORMAL,fill=fill if is_final else None)
    cell(ws,row,3,val,font=f,fill=fill,fmt='#,##0')
    cell(ws,row,4,src,font=NORMAL,align=Alignment(wrap_text=True))
    if is_final: ws.row_dimensions[row].height=22
    row+=1

# Сценарии A/B/C
row+=1
cell(ws,row,1,'Сценарии чувствительности УПС',font=SUBHEAD,fill=SUBHEAD_FILL)
ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=4)
row+=1
for ci,h in enumerate(['Сценарий','Замещение, %','Прирост поездок, %','Чистый эффект (денежный), тыс. руб.'],1):
    cell(ws,row,ci,h,font=HEAD,fill=HEAD_FILL,align=Alignment(horizontal='center',wrap_text=True))
row+=1
for sc_name,pct,growth in [('A. Консервативный',10,3),('B. Базовый',20,6),('C. Оптимистичный',30,9)]:
    eff_value=(reb_cost-sub_avg)*base_ops*pct/100/1000 + trips_ekb*growth/100*check_ekb/1000
    cell(ws,row,1,sc_name,font=NORMAL)
    cell(ws,row,2,pct,font=NORMAL,fmt='0" %"')
    cell(ws,row,3,growth,font=NORMAL,fmt='0" %"')
    cell(ws,row,4,round(eff_value,0),font=TOTAL,fill=GOOD_FILL,fmt='#,##0')
    row+=1

widths(ws,A=6,B=58,C=22,D=70)

# === Лист Эффекты ===
ws=wb.create_sheet('Эффекты')
ws['A1']='Эффекты сценария S3 (горизонт 12 месяцев)'
ws['A1'].font=TITLE
ws.merge_cells('A1:F1')
for ci,h in enumerate(['Эффект','Город / кластер','База расчёта','Параметр','EBITDA-эффект, тыс. руб./год','Источник'],1):
    cell(ws,2,ci,h,font=HEAD,fill=HEAD_FILL,align=Alignment(wrap_text=True,horizontal='center',vertical='center'))
ws.row_dimensions[2].height=35
R2_CITIES=['Санкт-Петербург','Екатеринбург','Краснодар','Новосибирск']
rev24_r2=sum(city_data[c]['rev24'] for c in R2_CITIES)
rev24_msk=city_data['Москва']['rev24']

eff1=round(rev24_r2/1000*0.03*0.25,0)
eff2=round(net_economy+extra_revenue*0.25,0)
eff3=round(rev24_msk/1000*0.02,0)
eff4=3000

effects_data=[
    ('Эффект 1. R2 атака: прирост выручки','СПб + Екб + Крд + Нск',
     f'Выручка R2 2024: {rev24_r2/1e6:,.1f} млн','+3 п.п. × EBITDA-маржа 25 %',eff1,
     f'= {rev24_r2/1000:.0f} × 0,03 × 0,25 = {eff1:,.0f}'),
    ('Эффект 2. УПС-пилот Екатеринбург (EBITDA)','Екатеринбург',
     'Парк 6448, утилизация 319','3,5 (экономия) + 8,2 × 0,25 (маржа)',eff2,
     f'= {net_economy:.0f} + {extra_revenue:.0f} × 0,25 ≈ {eff2:.0f} тыс. руб.'),
    ('Эффект 3. R3 Москва: оптимизация OpEx','Москва',
     f'Выручка Москвы 2024: {rev24_msk/1e6:,.1f} млн','2 % от выручки',eff3,
     f'= {rev24_msk/1000:.0f} × 0,02 = {eff3:,.0f}'),
    ('Эффект 4. R1 малые города: экономия капвложений','Малые города',
     '~10 млн капвложений, ~30 % EBITDA-вклад','~30 % от 10 млн',eff4,
     '[ДОПУЩЕНИЕ]'),
]
row=3
for eff_name,city,base,param,val,src in effects_data:
    cell(ws,row,1,eff_name,font=NORMAL)
    cell(ws,row,2,city,font=NORMAL)
    cell(ws,row,3,base,font=NORMAL)
    cell(ws,row,4,param,font=NORMAL,align=Alignment(wrap_text=True))
    cell(ws,row,5,val,font=TOTAL,fill=GOOD_FILL,fmt='#,##0')
    cell(ws,row,6,src,font=NORMAL,align=Alignment(wrap_text=True))
    row+=1
total_eff=eff1+eff2+eff3+eff4
cell(ws,row,1,'ИТОГО эффект S3, год 1, тыс. руб.',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,5,total_eff,font=TOTAL,fill=TOTAL_FILL,fmt='#,##0')
widths(ws,A=38,B=23,C=35,D=40,E=24,F=50)

# === Лист Сводная_финмодель ===
ws=wb.create_sheet('Сводная_финмодель')
ws['A1']='Сводная финансовая модель S3'
ws['A1'].font=TITLE
ws.merge_cells('A1:C1')
for ci,h in enumerate(['Статья','Млн руб.','Источник'],1):
    cell(ws,2,ci,h,font=HEAD,fill=HEAD_FILL,align=Alignment(horizontal='center'))

# Год 1
row=3
cell(ws,row,1,'ГОД 1 (12 месяцев)',font=BIG_TOTAL,fill=SUBHEAD_FILL)
ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=3)
row+=1
y1_items=[
    ('Затраты, всего',-round(tot_costs/1000,1),'Лист «Затраты», итого',BAD_FILL),
    ('Эффект 1. R2-прирост выручки',round(eff1/1000,1),'+3 п.п. × выручка R2 × 25 % маржа',GOOD_FILL),
    ('Эффект 2. УПС-пилот (EBITDA)',round(eff2/1000,1),'§3.2; экономия 3,5 + доп.выручка × 0,25',GOOD_FILL),
    ('Эффект 3. R3 Москва: оптимизация',round(eff3/1000,1),'2 % выручки Москвы',GOOD_FILL),
    ('Эффект 4. R1 малые города',round(eff4/1000,1),'Экономия капвложений',GOOD_FILL),
]
for st,val,src,fill in y1_items:
    cell(ws,row,1,st,font=NORMAL)
    cell(ws,row,2,val,font=TOTAL,fill=fill,fmt='+#,##0.0;-#,##0.0;0.0')
    cell(ws,row,3,src,font=NORMAL,align=Alignment(wrap_text=True))
    row+=1
total_y1=round(total_eff/1000,1)
net_y1=round((total_eff-tot_costs)/1000,1)
cell(ws,row,1,'Эффект, всего',font=TOTAL,fill=TOTAL_FILL)
cell(ws,row,2,total_y1,font=TOTAL,fill=TOTAL_FILL,fmt='+#,##0.0;-#,##0.0;0.0')
row+=1
cell(ws,row,1,'ЧИСТЫЙ ЭФФЕКТ S3 vs S1, год 1',font=BIG_TOTAL,fill=TOTAL_FILL)
cell(ws,row,2,net_y1,font=BIG_TOTAL,fill=TOTAL_FILL,fmt='+#,##0.0;-#,##0.0;0.0')
ws.row_dimensions[row].height=25

# Год 2
row+=3
cell(ws,row,1,'ГОД 2 (тиражирование УПС на 4 R2-города; capex выпадает)',font=BIG_TOTAL,fill=SUBHEAD_FILL)
ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=3)
row+=1
y2_items=[
    ('Затраты год 2',-37,'A + B + D + операционная C, без capex'),
    ('Эффект 1 расширенный (+5 п.п.)',round(rev24_r2/1e6*0.05*0.25,1),'+5 п.п. × выручка R2 × 25 %'),
    ('Эффект 2 (УПС в 4 городах R2)',40,'Коэф. 3,0 млн EBITDA / 1 млн поездок × 13 млн сез. поездок'),
    ('Эффект 3 (R3 Москва, год 2)',55,'Углубление оптимизации до 2,1 %'),
    ('Эффект 4 (R1 малые)',5,'Расширение каркаса'),
]
first_y2=row
for st,val,src in y2_items:
    cell(ws,row,1,st,font=NORMAL)
    cell(ws,row,2,val,font=TOTAL,fill=BAD_FILL if val<0 else GOOD_FILL,fmt='+#,##0.0;-#,##0.0;0.0')
    cell(ws,row,3,src,font=NORMAL,align=Alignment(wrap_text=True))
    row+=1
y2_sum=sum(val for _,val,_ in y2_items)
cell(ws,row,1,'ЧИСТЫЙ ЭФФЕКТ S3 vs S1, год 2',font=BIG_TOTAL,fill=TOTAL_FILL)
cell(ws,row,2,y2_sum,font=BIG_TOTAL,fill=TOTAL_FILL,fmt='+#,##0.0;-#,##0.0;0.0')
ws.row_dimensions[row].height=25
widths(ws,A=55,B=20,C=55)

# === Лист Чувствительность ===
ws=wb.create_sheet('Чувствительность')
ws['A1']='Сценарии чувствительности чистого эффекта S3 (год 1)'
ws['A1'].font=TITLE
ws.merge_cells('A1:E1')
for ci,h in enumerate(['Параметр','Консерв.','Базовый','Оптим.','Комментарий'],1):
    cell(ws,2,ci,h,font=HEAD,fill=HEAD_FILL,align=Alignment(horizontal='center'))

params=[
    ('Прирост темпа R2 выручки, п.п.',1,3,5,'Сила эффекта дифференциации'),
    ('Замещение ребалансировки УПС, %',10,20,30,'Сценарии A/B/C из §3.2'),
    ('Прирост восст. поездок УПС, %',3,6,9,'Доля чувствительных микролокаций'),
    ('Оптимизация OpEx Москвы, %',1,2,3,'Глубина рычагов R3'),
    ('Экономия капвложений малые города, млн',5,10,20,'Темпы расширения портфеля'),
]
row=3
for st,c,b,o,src in params:
    cell(ws,row,1,st,font=NORMAL)
    cell(ws,row,2,c,font=NORMAL,fmt='0.0')
    cell(ws,row,3,b,font=TOTAL,fill=TOTAL_FILL,fmt='0.0')
    cell(ws,row,4,o,font=NORMAL,fmt='0.0')
    cell(ws,row,5,src,font=NORMAL,align=Alignment(wrap_text=True))
    row+=1

# Производные эффекты (вычислены заранее)
row+=1
cell(ws,row,1,'Производные эффекты, млн руб.',font=SUBHEAD,fill=SUBHEAD_FILL)
ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=5)
row+=1

# Эффект 1: rev_r2 × pp/100 × 0.25
def eff1_calc(pp): return round(rev24_r2/1e6 * pp/100 * 0.25, 1)
# Эффект 2 (EBITDA): экономия + доп.выручка × 0,25, масштабируется через pct/20 и growth/6
def eff2_calc(pct, growth):
    econ = (reb_cost-sub_avg) * base_ops * pct/100 / 1000  # тыс. руб
    rev = trips_ekb * growth/100 * check_ekb / 1000  # тыс. руб
    return round((econ + rev*0.25)/1000, 1)  # млн руб
# Эффект 3: msk_rev × pct/100
def eff3_calc(pct): return round(rev24_msk/1e6 * pct/100, 1)
# Эффект 4: capex × 0,3
def eff4_calc(capex): return round(capex * 0.3, 1)

scen_p = [(1,10,3,1,5),(3,20,6,2,10),(5,30,9,3,20)]
labels = ['Эффект 1 (R2 темпы)','Эффект 2 (УПС-пилот)','Эффект 3 (R3 Москва)','Эффект 4 (R1 малые)','Затраты S3']
eff1_v = [eff1_calc(1), eff1_calc(3), eff1_calc(5)]
eff2_v = [eff2_calc(10,3), eff2_calc(20,6), eff2_calc(30,9)]
eff3_v = [eff3_calc(1), eff3_calc(2), eff3_calc(3)]
eff4_v = [eff4_calc(5), eff4_calc(10), eff4_calc(20)]
costs_v = [-48,-48,-48]

for label, vals in [('Эффект 1 (R2 темпы)',eff1_v),('Эффект 2 (УПС-пилот EBITDA)',eff2_v),
                    ('Эффект 3 (R3 Москва)',eff3_v),('Эффект 4 (R1 малые)',eff4_v),
                    ('Затраты S3',costs_v)]:
    cell(ws,row,1,label,font=NORMAL)
    cell(ws,row,2,vals[0],font=NORMAL,fmt='+0.0;-0.0;0.0')
    cell(ws,row,3,vals[1],font=TOTAL,fill=TOTAL_FILL,fmt='+0.0;-0.0;0.0')
    cell(ws,row,4,vals[2],font=NORMAL,fmt='+0.0;-0.0;0.0')
    row+=1

# Итого
nets = [round(eff1_v[i]+eff2_v[i]+eff3_v[i]+eff4_v[i]+costs_v[i],1) for i in range(3)]
cell(ws,row,1,'ЧИСТЫЙ ЭФФЕКТ S3, год 1, млн руб.',font=BIG_TOTAL,fill=TOTAL_FILL)
cell(ws,row,2,nets[0],font=BIG_TOTAL,fill=TOTAL_FILL,fmt='+#,##0.0;-#,##0.0;0.0')
cell(ws,row,3,nets[1],font=BIG_TOTAL,fill=TOTAL_FILL,fmt='+#,##0.0;-#,##0.0;0.0')
cell(ws,row,4,nets[2],font=BIG_TOTAL,fill=TOTAL_FILL,fmt='+#,##0.0;-#,##0.0;0.0')
ws.row_dimensions[row].height=25

widths(ws,A=42,B=18,C=18,D=18,E=50)

# Порядок листов
order=['README','Данные_Юрент','Затраты','УПС_расчёт','Эффекты','Сводная_финмодель','Чувствительность']
wb._sheets=[wb[name] for name in order]
out=Path('/workspace/vkr/VKR_findata.xlsx')
wb.save(out)
print(f"Saved: {out} ({out.stat().st_size:,} bytes)")
print(f"Sheets: {wb.sheetnames}")
print()
print('=== ВЕРИФИКАЦИЯ ЗНАЧЕНИЙ ===')
print(f'УПС-пилот денежный = {net_ups:,.0f} тыс. руб. ({net_ups/1000:.1f} млн)')
print(f'УПС-пилот EBITDA   = {eff2:,.0f} тыс. руб. ({eff2/1000:.1f} млн)')
print(f'Эффект 1 (R2)      = {eff1:,.0f} тыс. ({eff1/1000:.1f} млн)')
print(f'Эффект 3 (R3 МСК)  = {eff3:,.0f} тыс. ({eff3/1000:.1f} млн)')
print(f'Эффект 4 (R1)      = {eff4:,.0f} тыс. ({eff4/1000:.1f} млн)')
print(f'ИТОГО эффект       = {total_eff:,.0f} тыс. ({total_eff/1000:.1f} млн)')
print(f'ИТОГО затраты      = {tot_costs:,.0f} тыс. ({tot_costs/1000:.1f} млн)')
print(f'ЧИСТЫЙ ЭФФЕКТ      = {total_eff-tot_costs:,.0f} тыс. ({(total_eff-tot_costs)/1000:.1f} млн)')
print()
print('Чувствительность год 1 (млн руб.):')
print(f'  Консервативный: {nets[0]:+.1f}')
print(f'  Базовый:        {nets[1]:+.1f}')
print(f'  Оптимистичный:  {nets[2]:+.1f}')
