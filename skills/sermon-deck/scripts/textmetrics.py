# -*- coding: utf-8 -*-
"""Ширины строк для подбора кегля.

Метрики лежат в assets/font_metrics.json (доли em), поэтому скилл считает
размеры даже там, где фирменные шрифты не установлены. Если нужного шрифта
в JSON нет, читаем его из ~/Library/Fonts через fontTools.
"""
import json
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_METRICS_PATH = os.path.join(_HERE, '..', 'assets', 'font_metrics.json')
_SYSTEM_FONTS = {
    'Akrobat Bold': 'Akrobat-Bold.otf',
    'Heading Pro': 'HeadingPro-Regular.ttf',
    'Heading Pro ExtBd': 'HeadingPro-ExtraBold.ttf',
    'Intro ': 'Intro.otf',
}
_cache = {}


def _from_system(font):
    from fontTools.ttLib import TTFont          # импорт только если правда нужен
    path = os.path.expanduser('~/Library/Fonts/' + _SYSTEM_FONTS[font])
    f = TTFont(path, fontNumber=0)
    upm = f['head'].unitsPerEm
    hmtx = f['hmtx']
    return {chr(cp): hmtx[g][0] / upm for cp, g in f.getBestCmap().items() if g in hmtx.metrics}


def _widths(font):
    if font not in _cache:
        try:
            with open(_METRICS_PATH, encoding='utf-8') as fh:
                bundled = json.load(fh)
        except FileNotFoundError:
            bundled = {}
        if font in bundled:
            _cache[font] = bundled[font]
        elif font in _SYSTEM_FONTS:
            _cache[font] = _from_system(font)
        else:
            raise KeyError(f'нет метрик для шрифта {font!r}: добавь его в assets/font_metrics.json')
    return _cache[font]


def text_width_pt(text, font, size_pt):
    """Ширина одной строки в пунктах. Кернинг не учитываем — запас его перекрывает."""
    w = _widths(font)
    fallback = w.get('о', 0.5)
    return sum(w.get(ch, fallback) for ch in text) * size_pt


def wrap_lines(text, font, size_pt, avail_pt):
    """Жадный перенос по пробелам — так же, как это делает PowerPoint."""
    lines, cur = [], ''
    for word in text.split(' '):
        trial = word if not cur else cur + ' ' + word
        if text_width_pt(trial, font, size_pt) <= avail_pt or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines


def n_lines(text, font, size_pt, avail_pt):
    return len(wrap_lines(text, font, size_pt, avail_pt))


def longest_word_pt(text, font, size_pt):
    """Слово шире строки PowerPoint разрывает посередине — это надо ловить заранее."""
    return max(text_width_pt(w, font, size_pt) for w in text.split(' ') if w)
