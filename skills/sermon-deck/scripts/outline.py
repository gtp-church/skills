#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Разбор PDF-конспекта проповеди.

Пастор выделяет ключевые пункты жирным — в PDF это отдельный жирный шрифт.
Скрипт печатает структуру: жирные заголовки, абзацы между ними и все
встреченные ссылки на Писание.

    python3 outline.py конспект.pdf
"""
import re
import subprocess
import sys

BOOKS = (r'Быт|Исх|Лев|Чис|Втор|Нав|Суд|Руф|Цар|Пар|Езд|Неем|Есф|Иов|Пс|Прит|Еккл|Песн|'
         r'Ис|Иер|Плач|Иез|Дан|Ос|Иоил|Ам|Авд|Ион|Мих|Наум|Авв|Соф|Агг|Зах|Мал|'
         r'Мф|Мк|Лк|Ин|Деян|Рим|Кор|Гал|Еф|Флп|Кол|Фес|Тим|Тит|Флм|Евр|Иак|Петр|Иуд|Откр|'
         r'Матфея|Марка|Луки|Иоанна|Римлянам|Галатам|Ефесянам|Филиппийцам|Колоссянам|'
         r'Тимофею|Титу|Евреям|Иакова|Петра|Бытие|Исход|Иезекииля|Исаии|Псалом')
REF = re.compile(r'\(?\b(?:[1-3]\s*)?(?:' + BOOKS + r')[а-я]*\.?\s*\d+[:,]\s*\d+(?:\s*[-–]\s*\d+)?\)?')


def runs(pdf):
    """Возвращает [(жирный?, текст)] в порядке чтения."""
    xml = subprocess.run(['pdftohtml', '-xml', '-i', '-stdout', pdf],
                         capture_output=True, text=True, check=True).stdout
    specs = {m.group(1): m.group(2) for m in
             re.finditer(r'<fontspec id="(\d+)"[^>]*family="([^"]*)"', xml)}
    bold_ids = {i for i, fam in specs.items() if 'bold' in fam.lower()}
    out = []
    for m in re.finditer(r'<text[^>]*font="(\d+)">(.*?)</text>', xml, re.S):
        fid, body = m.groups()
        txt = re.sub(r'<[^>]+>', '', body)
        for a, b in (('&#160;', ' '), ('&amp;', '&'), ('&lt;', '<'), ('&gt;', '>'),
                     ('&#34;', '"'), ('&quot;', '"'), ('&apos;', "'")):
            txt = txt.replace(a, b)
        txt = txt.replace('​', '').replace('\xad', '').strip()
        if txt:
            out.append((fid in bold_ids or '<b>' in body, txt))
    return out


def main(pdf):
    rows = runs(pdf)
    print('=' * 70)
    print('ЖИРНЫМ — то, что пастор выделил как ключевое. Это и есть разделы.')
    print('=' * 70)
    buf = []
    # ссылки ищем по склеенному тексту: в PDF «(Иез.» и «18:20)» часто
    # попадают в разные строки и по отдельности не распознаются
    refs = [m.group(0) for m in REF.finditer(' '.join(t for _, t in rows))]
    for bold, txt in rows:
        if bold:
            if buf:
                print('    ' + ' '.join(buf)[:600] + ('…' if len(' '.join(buf)) > 600 else ''))
                buf = []
            print(f'\n### {txt}')
        else:
            buf.append(txt)
    if buf:
        print('    ' + ' '.join(buf)[:600])
    print('\n' + '=' * 70)
    print('ССЫЛКИ НА ПИСАНИЕ (каждой нужен отдельный слайд):')
    seen = []
    for r in refs:
        r = r.strip('()')
        if r not in seen:
            seen.append(r)
    for r in seen:
        print('   ', r)
    print('\nТекст стиха, которого нет в конспекте, сверяй по azbyka.ru —')
    print('не приводи Писание по памяти.')


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else sys.exit(__doc__))
