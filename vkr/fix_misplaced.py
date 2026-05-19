#!/usr/bin/env python3
"""Срочно: вырезаю блоки KFU, SW, ФИНМОДЕЛЬ из ОГЛАВЛЕНИЯ
и переношу в правильные места в основном тексте.
"""
import re
from pathlib import Path

p = Path('/workspace/vkr/vkr_final.txt')
t = p.read_text(encoding='utf-8')
n0 = len(t)

# Граница оглавления = первая строка "ВВЕДЕНИЕ" с тегом раздела (она же 2-я после ОГЛАВЛЕНИЕ).
# Но в нашем файле первая строка "ВВЕДЕНИЕ" — пункт оглавления, а реальное "ВВЕДЕНИЕ" — это
# заголовок раздела. Найдём конец оглавления как ПОСЛЕДНЮЮ строку перед реальным ВВЕДЕНИЕ.

lines = t.split('\n')

# Ищем строку, где сразу после "ВВЕДЕНИЕ" идёт реальный абзац (не "ГЛАВА 1...").
# Признак: после строки "ВВЕДЕНИЕ" следует абзац начинающийся с "Кикшеринг — сервис" или подобного.
real_intro_idx = -1
for i, ln in enumerate(lines):
    if ln.strip() == 'ВВЕДЕНИЕ' and i+1 < len(lines):
        nxt = lines[i+1].strip()
        if nxt.startswith('Кикшеринг'):
            real_intro_idx = i
            break
print(f"Real intro at line {real_intro_idx}")

# Найдём начало оглавления
oglav_idx = -1
for i, ln in enumerate(lines):
    if ln.strip() == 'ОГЛАВЛЕНИЕ':
        oglav_idx = i
        break
print(f"ОГЛАВЛЕНИЕ at line {oglav_idx}")

# Оглавление = строки от oglav_idx до real_intro_idx (exclusive)
# Внутри оглавления есть ВСТАВЛЕННЫЕ ошибочно блоки. Нужно их найти и вырезать.

# Маркеры вставленных блоков:
# 1) KFU блок начинается со строки "Ключевые факторы успеха на зрелом рынке кикшеринга."
# 2) SW блок начинается со строки "Сводный вывод по внутренней среде: сильные и слабые стороны."
# 3) Финмодель начинается со строки "3.4 Финансовая модель сценария S3: затраты, эффекты, чистый эффект"
#    (ДВА таких заголовка подряд — это мой косяк, второй надо удалить, как и весь блок).

# Безопаснее работать со строками.
new_lines = []
i = 0
while i < len(lines):
    ln = lines[i]
    s = ln.strip()
    # Внутри оглавления (между oglav_idx и real_intro_idx)
    if oglav_idx < i < real_intro_idx:
        # KFU block
        if s.startswith('Ключевые факторы успеха на зрелом рынке кикшеринга.'):
            # пропускаем до и включая строку "Региональные гэпы МТС Юрент относительно лидера." + 1 абзац после неё
            while i < real_intro_idx:
                ll = lines[i].strip()
                if ll.startswith('Региональные гэпы МТС Юрент относительно лидера.'):
                    # после неё ещё одна строка (продолжение), но это однострочный абзац
                    i += 1
                    # пропускаем пустые
                    while i < real_intro_idx and not lines[i].strip():
                        i += 1
                    break
                i += 1
            continue
        # SW block
        if s.startswith('Сводный вывод по внутренней среде:'):
            # пропускаем до конца блока: до "2.3 Стратегический выбор" (но НЕ включая его)
            while i < real_intro_idx:
                ll = lines[i].strip()
                if ll.startswith('2.3 Стратегический выбор'):
                    break
                i += 1
            continue
        # Дубль заголовка финмодели + сама финмодель
        if s == '3.4 Финансовая модель сценария S3: затраты, эффекты, чистый эффект':
            # Это либо первое (легит), либо второе. Если строка ПЕРЕД ним — это первый "3.4 Финансовая модель..."
            # (легитимный пункт оглавления), то текущий — дубль начала вставленного блока.
            prev_real = None
            for j in range(i-1, -1, -1):
                if lines[j].strip():
                    prev_real = lines[j].strip()
                    break
            if prev_real == '3.4 Финансовая модель сценария S3: затраты, эффекты, чистый эффект':
                # это второй (дубль) — начинается ошибочно вставленная финмодель
                # пропускаем до "3.5 Риски сценария S3..."
                while i < real_intro_idx:
                    ll = lines[i].strip()
                    if ll.startswith('3.5 Риски сценария S3'):
                        break
                    i += 1
                continue
    new_lines.append(ln)
    i += 1

t2 = '\n'.join(new_lines)
n1 = len(t2)
print(f"After cleanup: {n1} chars (delta {n1-n0:+d})")

# Сохраним промежуточно
Path('/workspace/vkr/vkr_final.txt').write_text(t2, encoding='utf-8')
print("Cleaned TOC. Now need to RE-INSERT blocks into the right places.")
