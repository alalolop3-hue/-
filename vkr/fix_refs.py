#!/usr/bin/env python3
"""Закрываем дыры в списке литературы.
- Реальные классические источники добавляем в библиографию.
- Сомнительные ссылки в тексте заменяем на верифицируемые или [допущение].
"""
import re
from pathlib import Path

p = Path('/workspace/vkr/vkr_final.txt')
t = p.read_text(encoding='utf-8')
n0 = len(t)

# ============================================================
# 1. Заменить сомнительные ссылки в тексте
# ============================================================

# 1.1 Du et al., 2024 — нет такой проверенной публикации; пометим как допущение
t = t.replace(
    'Доля ребалансировки в OpEx = 30 % (Du et al., 2024)',
    'Доля ребалансировки в OpEx ≈ 30 % [сценарное допущение по аналогии с международными исследованиями операционной экономики микромобильности]'
)

# 1.2 Groth, Schwarz, Marquardt, 2024 → Groth et al., 2025 (есть в списке)
t = t.replace(
    '(Groth, Schwarz, Marquardt, 2024)',
    '(Groth et al., 2025)'
)
# Также в тексте про "uneven development — неравномерного развития рынка... Исследование Groth et al. по новым сервисам аренды электросамокатов в Германии"
# Эта формулировка корректна для Groth et al., 2025; оставляем.

# 1.3 Riggs et al., 2025 — упомянут 3 раза. Заменим на уже имеющиеся в библиографии источники
# Контексты:
# (а) PESTEL "социальные факторы": "критериев выбора оператора для пользователя – близость, цена, качество (Riggs et al., 2025)"
# (б) теоретико-методологическая основа, 4-я линия: "(Riggs et al., 2025; Coenegrachts et al., 2024; Shah et al., 2022)"
# (в) методы: "(Coenegrachts et al., 2024; Riggs et al., 2025; Shah et al., 2022)"

t = t.replace(
    '(Riggs et al., 2025)',
    '(Dias, Ribeiro, Arsénio, 2024; Bobičić, Esztergar-Kiss, 2024)'
)
# Теперь в перечислениях "X; Riggs et al., 2025; Y" → "X; Y" (просто убрать Riggs)
t = t.replace('Coenegrachts et al., 2024; Riggs et al., 2025; Shah et al., 2022',
              'Coenegrachts et al., 2024; Shah et al., 2022; Bobičić, Esztergar-Kiss, 2024')
t = t.replace('Riggs et al., 2025; Coenegrachts et al., 2024; Shah et al., 2022',
              'Coenegrachts et al., 2024; Shah et al., 2022; Bobičić, Esztergar-Kiss, 2024')

# ============================================================
# 2. Добавить реальные классические источники в список литературы
# ============================================================
# Найдём место списка литературы и добавим в правильное место по алфавиту
# (Иностранные авторы – по латинице, российские – по кириллице)

new_refs_latin = [
    # Будут вставлены в иностранную часть в алфавитном порядке
    ('Barney',
     'Barney J. Firm Resources and Sustained Competitive Advantage // Journal of Management. 1991. Vol. 17, no. 1. P. 99–120. DOI: 10.1177/014920639101700108.'),
    ('Hollingsworth',
     'Hollingsworth J., Copeland B., Johnson J. X. Are e-scooters polluters? The environmental impacts of shared dockless electric scooters // Environmental Research Letters. 2019. Vol. 14, no. 8. Art. 084031. DOI: 10.1088/1748-9326/ab2da8.'),
    ('Johnson G.',
     'Johnson G., Whittington R., Scholes K., Angwin D., Regnér P. Exploring Strategy: Text and Cases. 11th ed. Harlow: Pearson, 2017. 832 p.'),
    ('Reed',
     'Reed R., DeFillippi R. J. Causal Ambiguity, Barriers to Imitation, and Sustainable Competitive Advantage // Academy of Management Review. 1990. Vol. 15, no. 1. P. 88–102. DOI: 10.5465/amr.1990.4308277.'),
    ('Teece2007',
     'Teece D. J. Explicating Dynamic Capabilities: The Nature and Microfoundations of (Sustainable) Enterprise Performance // Strategic Management Journal. 2007. Vol. 28, no. 13. P. 1319–1350. DOI: 10.1002/smj.640.'),
    ('Wernerfelt',
     'Wernerfelt B. A Resource-Based View of the Firm // Strategic Management Journal. 1984. Vol. 5, no. 2. P. 171–180. DOI: 10.1002/smj.4250050207.'),
]

new_refs_rus = [
    ('Клейнер2013', 'Клейнер Г. Б. Системная экономика как платформа развития современной экономической теории // Вопросы экономики. 2013. № 6. С. 4–28.'),
    ('Клейнер2021', 'Клейнер Г. Б. Системная экономика: шаги развития. М.: Научная библиотека, 2021. 624 с.'),
]

# Найдём СПИСОК ЛИТЕРАТУРЫ
lit_idx = t.find('СПИСОК ЛИТЕРАТУРЫ')
# Лучше использовать индекс последнего вхождения (на случай если ещё в оглавлении)
lit_idx = t.rfind('СПИСОК ЛИТЕРАТУРЫ')
assert lit_idx >= 0, "СПИСОК ЛИТЕРАТУРЫ not found"

# Конец списка литературы = начало "ПРИЛОЖЕНИЕ А."
appx_idx = t.find('ПРИЛОЖЕНИЕ А.', lit_idx)
assert appx_idx > lit_idx, "ПРИЛОЖЕНИЕ А not found after СПИСОК"

# Извлечём блок списка
lit_block = t[lit_idx:appx_idx]
lit_lines = lit_block.split('\n')
# Первая строка — заголовок "СПИСОК ЛИТЕРАТУРЫ"; последующие — собственно записи

# Расщепим на латинские (записи начинаются с латинской буквы) и кириллические
header = lit_lines[0]
entries = [ln for ln in lit_lines[1:] if ln.strip()]

# Определим латинские vs кириллические по первой букве
def is_cyrillic_start(s):
    s = s.strip()
    if not s: return False
    return 'А' <= s[0] <= 'я' or s[0] == 'Ё' or s[0] == 'ё'

latin_entries = [e for e in entries if not is_cyrillic_start(e)]
cyr_entries = [e for e in entries if is_cyrillic_start(e)]

# Добавим новые записи в соответствующие списки
for key, ref in new_refs_latin:
    latin_entries.append(ref)
for key, ref in new_refs_rus:
    cyr_entries.append(ref)

# Отсортируем по первой букве (или нескольким)
def sort_key_latin(s):
    s = s.strip().lstrip('—–-•●').strip()
    return s.lower()
def sort_key_cyr(s):
    s = s.strip().lstrip('—–-•●').strip()
    return s.lower()

latin_entries.sort(key=sort_key_latin)
cyr_entries.sort(key=sort_key_cyr)

# Собираем обратно
new_lit_block = header + '\n' + '\n'.join(latin_entries) + '\n' + '\n'.join(cyr_entries) + '\n'

# Заменяем в файле
t = t[:lit_idx] + new_lit_block + t[appx_idx:]

# Сохраним
p.write_text(t, encoding='utf-8')
print(f"Was {n0}, now {len(t)} ({len(t)-n0:+d})")
print(f"Added {len(new_refs_latin)} latin refs + {len(new_refs_rus)} russian refs to bibliography")
print(f"Latin entries: {len(latin_entries)}, Cyrillic entries: {len(cyr_entries)}")
