#!/usr/bin/env python3
"""Сборка .docx из vkr_final.txt по требованиям ГОСТ 7.32–2017:
- Times New Roman 14, интервал 1,5
- Поля: левое 3 см, правое 1 см, верх/низ 2 см
- Заголовки жирным по левому краю
- Нумерация страниц снизу по центру
- Таблицы из строк с TAB-разделителем
- Маркированные списки из строк, начинающихся с '●'
"""
import re
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

src = Path('/workspace/vkr/vkr_final.txt').read_text(encoding='utf-8')

doc = Document()

# --- Поля и шрифт ---
section = doc.sections[0]
section.top_margin = Cm(2)
section.bottom_margin = Cm(2)
section.left_margin = Cm(3)
section.right_margin = Cm(1)

# Default style
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(14)
# Кириллица
rPr = style.element.get_or_add_rPr()
rFonts = rPr.find(qn('w:rFonts'))
if rFonts is None:
    rFonts = OxmlElement('w:rFonts')
    rPr.append(rFonts)
rFonts.set(qn('w:ascii'), 'Times New Roman')
rFonts.set(qn('w:hAnsi'), 'Times New Roman')
rFonts.set(qn('w:cs'), 'Times New Roman')
rFonts.set(qn('w:eastAsia'), 'Times New Roman')

# Интервал 1.5 по умолчанию
pf = style.paragraph_format
pf.line_spacing = 1.5
pf.space_before = Pt(0)
pf.space_after = Pt(0)
pf.first_line_indent = Cm(1.25)

# --- Стили заголовков ---
def set_heading_style(name, size, bold=True, indent_first=False):
    s = doc.styles[name]
    s.font.name = 'Times New Roman'
    s.font.size = Pt(size)
    s.font.bold = bold
    s.font.color.rgb = RGBColor(0,0,0)
    rPr2 = s.element.get_or_add_rPr()
    rF = rPr2.find(qn('w:rFonts'))
    if rF is None:
        rF = OxmlElement('w:rFonts'); rPr2.append(rF)
    rF.set(qn('w:ascii'),'Times New Roman')
    rF.set(qn('w:hAnsi'),'Times New Roman')
    rF.set(qn('w:cs'),'Times New Roman')
    rF.set(qn('w:eastAsia'),'Times New Roman')
    pf = s.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing = 1.5
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.first_line_indent = Cm(0) if not indent_first else Cm(1.25)
    pf.keep_with_next = True

set_heading_style('Heading 1', 14)
set_heading_style('Heading 2', 14)
set_heading_style('Heading 3', 14)
set_heading_style('Heading 4', 14)

# --- Номера страниц снизу по центру ---
def add_page_number(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld_begin = OxmlElement('w:fldChar'); fld_begin.set(qn('w:fldCharType'),'begin')
    instr = OxmlElement('w:instrText'); instr.text = 'PAGE'
    fld_sep = OxmlElement('w:fldChar'); fld_sep.set(qn('w:fldCharType'),'separate')
    fld_end = OxmlElement('w:fldChar'); fld_end.set(qn('w:fldCharType'),'end')
    run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_sep); run._r.append(fld_end)

add_page_number(section)

# --- Парсинг строк ---
# Регексы для определения уровня заголовка
RE_H1_NAMED = re.compile(r'^(ОГЛАВЛЕНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК ЛИТЕРАТУРЫ|ПРИЛОЖЕНИЕ\s+[А-Я]\.)')
RE_H1_GLAVA = re.compile(r'^ГЛАВА\s+\d+\.')
RE_H2 = re.compile(r'^(\d+)\.(\d+)\s+\S')   # 1.1
RE_H3 = re.compile(r'^(\d+)\.(\d+)\.(\d+)\s+\S')  # 1.1.1
RE_H4 = re.compile(r'^(\d+)\.(\d+)\.(\d+)\.(\d+)\s+\S')  # 1.1.1.1
RE_TABLE_HEAD = re.compile(r'^Таблица\s+\S')
RE_FIGURE = re.compile(r'^Рисунок\s+\S')
RE_BULLET = re.compile(r'^[●•]\s+(.*)$')
RE_PSEUDOBULLET = re.compile(r'^[–-]\s+(.+)$')  # "– текст" в начале

# Заголовок «МЕТОДЫ» во введении тоже считаем H2 (по факту)
SPECIAL_H2_HEADERS = {'МЕТОДЫ'}

# Подзаголовки внутри R-режимов и УПС: "R1 «Плацдарм»", "R2 «Атака»", "R3 «Удержание»",
# "Цель.", "Условия применения.", "Рычаги.", "Показатели.", "Ограничения." и пр.
# Не делаем их заголовками отдельных уровней — оставим жирным курсивом как inline.
# Простоты ради — обычные абзацы.

def is_table_row(s, prev_is_table):
    # Строка считается строкой таблицы, если содержит >=2 табуляции
    return s.count('\t') >= 2

def add_paragraph(text, heading_level=None, alignment=None, indent=True):
    if heading_level is not None:
        p = doc.add_paragraph(text, style=f'Heading {heading_level}')
    else:
        p = doc.add_paragraph(text)
        if not indent:
            p.paragraph_format.first_line_indent = Cm(0)
    if alignment is not None:
        p.alignment = alignment
    return p

def add_bullet(text):
    p = doc.add_paragraph(text, style='List Bullet')
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(1.25)
    return p

lines = src.split('\n')

# Найдём блок оглавления: от строки 'ОГЛАВЛЕНИЕ' до ВТОРОГО вхождения 'ВВЕДЕНИЕ'
# (первое — пункт оглавления, второе — реальный заголовок главы).
toc_start = -1; toc_end = -1
vved_count = 0
for idx, ln in enumerate(lines):
    s = ln.strip()
    if s == 'ОГЛАВЛЕНИЕ':
        toc_start = idx
    elif s == 'ВВЕДЕНИЕ' and toc_start >= 0:
        vved_count += 1
        if vved_count == 2:
            toc_end = idx  # exclusive: реальное ВВЕДЕНИЕ
            break
print(f"TOC range: lines {toc_start}..{toc_end} (excl)")

i = 0
N = len(lines)
in_toc = False
while i < N:
    line = lines[i].rstrip('\r')
    s = line.strip()
    if not s:
        i += 1
        continue

    # Внутри блока ОГЛАВЛЕНИЕ → ВВЕДЕНИЕ — всё (кроме самого заголовка ОГЛАВЛЕНИЕ)
    # делаем обычным абзацем без отступа красной строки
    if toc_start >= 0 and toc_end > toc_start and toc_start < i < toc_end:
        p = doc.add_paragraph(s)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.5
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        i += 1
        continue

    # Таблица: накапливаем строки с табами
    if '\t' in s and s.count('\t') >= 2 and not RE_H2.match(s) and not RE_H3.match(s):
        rows = []
        while i < N:
            ll = lines[i].rstrip('\r').strip()
            if not ll:
                i += 1
                break
            if '\t' in ll and ll.count('\t') >= 1:
                rows.append(ll.split('\t'))
                i += 1
            else:
                break
        if rows:
            ncols = max(len(r) for r in rows)
            for r in rows:
                while len(r) < ncols:
                    r.append('')
            tbl = doc.add_table(rows=len(rows), cols=ncols)
            tbl.style = 'Table Grid'
            for ri, r in enumerate(rows):
                for ci, val in enumerate(r):
                    cell = tbl.cell(ri, ci)
                    cell.text = ''
                    p = cell.paragraphs[0]
                    p.paragraph_format.first_line_indent = Cm(0)
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    run = p.add_run(val)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(11)
                    if ri == 0:
                        run.font.bold = True
                    rPr3 = run._element.get_or_add_rPr()
                    rF3 = rPr3.find(qn('w:rFonts'))
                    if rF3 is None:
                        rF3 = OxmlElement('w:rFonts'); rPr3.append(rF3)
                    rF3.set(qn('w:ascii'),'Times New Roman')
                    rF3.set(qn('w:hAnsi'),'Times New Roman')
                    rF3.set(qn('w:eastAsia'),'Times New Roman')
                    rF3.set(qn('w:cs'),'Times New Roman')
            # пустой абзац после таблицы
            doc.add_paragraph('').paragraph_format.first_line_indent = Cm(0)
        continue

    # Заголовки
    m_h4 = RE_H4.match(s)
    m_h3 = RE_H3.match(s) if not m_h4 else None
    m_h2 = RE_H2.match(s) if not (m_h4 or m_h3) else None
    if RE_H1_NAMED.match(s) or RE_H1_GLAVA.match(s):
        p = doc.add_paragraph(s, style='Heading 1')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        # Page break перед ГЛАВА, ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК, ПРИЛОЖЕНИЕ — кроме первого ОГЛАВЛЕНИЕ
        if s != 'ОГЛАВЛЕНИЕ':
            run = p.runs[0] if p.runs else None
            # Вставим page break перед заголовком
            from docx.oxml import OxmlElement
            br = OxmlElement('w:br'); br.set(qn('w:type'),'page')
            # вставим в начало
            pPr = p._p
            new_run = OxmlElement('w:r')
            new_run.append(br)
            pPr.insert(0, new_run)
        i += 1
        continue
    if s in SPECIAL_H2_HEADERS:
        add_paragraph(s, heading_level=2, indent=False)
        i += 1
        continue
    if m_h4:
        add_paragraph(s, heading_level=4, indent=False)
        i += 1
        continue
    if m_h3:
        add_paragraph(s, heading_level=3, indent=False)
        i += 1
        continue
    if m_h2:
        add_paragraph(s, heading_level=2, indent=False)
        i += 1
        continue

    # Рисунок / Таблица заголовок (центрируем)
    if RE_TABLE_HEAD.match(s) or RE_FIGURE.match(s):
        p = doc.add_paragraph(s)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        for run in p.runs:
            run.font.italic = True
        i += 1
        continue

    # Маркированный список
    mb = RE_BULLET.match(s)
    if mb:
        add_bullet(mb.group(1))
        i += 1
        continue

    # Обычный абзац
    p = doc.add_paragraph(s)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    i += 1

# Сохранить
out = Path('/workspace/vkr/VKR_final.docx')
doc.save(str(out))
print(f"Saved: {out}, size: {out.stat().st_size} bytes")
