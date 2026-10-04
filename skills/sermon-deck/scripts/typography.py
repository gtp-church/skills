"""Typography for projected Russian text; source notes stay untouched."""
import re

PREPOSITIONS = frozenset([
    'в', 'во', 'к', 'ко', 'с', 'со', 'у', 'о', 'об', 'обо', 'на', 'по', 'за',
    'из', 'от', 'до', 'для', 'при', 'без', 'над', 'под', 'про', 'перед', 'через',
    'между', 'ради', 'вне', 'вдоль', 'вместо', 'около', 'после', 'против', 'сквозь', 'среди',
    'пред', 'чрез', 'безо', 'изо', 'ото', 'надо', 'подо', 'предо',
])
_PATTERN = re.compile(
    r'(?<!\w)(' + '|'.join(sorted(PREPOSITIONS, key=len, reverse=True))
    + r')([ \t]+)(?=[«“"(]?[\w])', re.IGNORECASE)


def display_text(text):
    """Use е on screen and bind prepositions to the following word."""
    text = text.replace('ё', 'е').replace('Ё', 'Е')
    return _PATTERN.sub(lambda match: match.group(1) + '\u00a0', text)


def trailing_preposition(line):
    words = re.findall(r'[А-Яа-яЁё]+', line)
    return bool(words and words[-1].lower() in PREPOSITIONS)
