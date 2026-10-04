#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx"]
# ///
# -*- coding: utf-8 -*-
"""Проверка готовой презентации: вылезания текста, фон, логотип, заметки,
плюс рендер в PDF и PNG для визуального контроля.

    python3 check.py Проповедь.pptx [папка_для_рендера]

PDF полезен и сам по себе — как запасной вариант для показа.
"""
import os
import subprocess
import shutil
import sys
import xml.etree.ElementTree as ET

from pptx import Presentation
from pptx.util import Emu
from packagecheck import package_problems
from typography import trailing_preposition

def find_soffice():
    """LibreOffice from PATH, an explicit override, or its macOS app bundle."""
    override = os.environ.get('SOFFICE')
    if override:
        return override
    for command in ('soffice', 'libreoffice'):
        found = shutil.which(command)
        if found:
            return found
    mac_path = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    if os.path.isfile(mac_path):
        return mac_path
    raise FileNotFoundError('Install LibreOffice or set SOFFICE to its executable path')

LOGO_TOP_IN = 6.71          # выше этой линии должен заканчиваться текст


def audit(path):
    prs = Presentation(path)
    problems = package_problems(path)
    for i, slide in enumerate(prs.slides, 1):
        names = [sh.name for sh in slide.shapes]
        if not any(n.startswith('Rectangle 9') for n in names):
            problems.append(f'слайд {i}: нет чёрной подложки')
        if not any(sh.shape_type == 13 for sh in slide.shapes):
            problems.append(f'слайд {i}: нет логотипа')
        if slide.placeholders and len(slide.placeholders):
            problems.append(f'слайд {i}: остались пустые заполнители макета')
        if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame.text.strip():
            problems.append(f'слайд {i}: пустые заметки докладчика')
        for sh in slide.shapes:
            if not (sh.has_text_frame and sh.text_frame.text.strip()):
                continue
            if 'ё' in sh.text.lower():
                problems.append(f'слайд {i}, «{sh.name}»: на экране есть ё')
            for line in sh.text.splitlines()[:-1]:
                if trailing_preposition(line):
                    problems.append(f'слайд {i}, «{sh.name}»: висячий предлог перед явным переносом')
            bottom = Emu(sh.top).inches + Emu(sh.height).inches
            if bottom > 7.1:
                problems.append(f'слайд {i}, «{sh.name}»: текст уходит вниз до {bottom:.2f}"')
            if Emu(sh.top).inches < -0.01 or Emu(sh.left).inches < -0.01:
                problems.append(f'слайд {i}, «{sh.name}»: рамка за краем слайда')
    return prs, problems


def render(path, outdir):
    os.makedirs(outdir, exist_ok=True)
    subprocess.run([find_soffice(), '--headless', '--convert-to', 'pdf', '--outdir', outdir, path],
                   capture_output=True, check=True)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + '.pdf')
    subprocess.run(['pdftoppm', '-png', '-r', '60', pdf, os.path.join(outdir, 's')],
                   capture_output=True, check=True)
    return pdf


def audit_rendered_text(pdf):
    """Check real line endings after rendering, including automatic wraps."""
    result = subprocess.run(['pdftotext', '-bbox-layout', pdf, '-'],
                            text=True, capture_output=True, check=True)
    root = ET.fromstring(result.stdout)
    problems = []
    for i, page in enumerate(root.findall('.//{*}page'), 1):
        for line in page.findall('.//{*}line'):
            text = ' '.join(word.text or '' for word in line.findall('{*}word'))
            if trailing_preposition(text):
                problems.append(f'PDF, слайд {i}: висячий предлог в строке «{text}»')
    return problems


def main():
    path = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else 'render'
    package_errors = package_problems(path)
    if package_errors:
        print('ошибки пакета PPTX:\n  ' + '\n  '.join(package_errors), file=sys.stderr)
        return 1
    prs, problems = audit(path)
    print(f'{len(prs.slides)} слайдов')
    print('замечания:', '\n  ' + '\n  '.join(problems) if problems else 'нет')
    try:
        pdf = render(path, outdir)
        rendered_problems = audit_rendered_text(pdf)
        problems.extend(rendered_problems)
        if rendered_problems:
            print('переносы в PDF:\n  ' + '\n  '.join(rendered_problems))
        pages = len([f for f in os.listdir(outdir) if f.startswith('s-')])
        print(f'рендер: {pdf} ({pages} страниц)')
        print('Посмотри PNG в этой папке — особенно разделители и самые длинные')
        print('списки: расчёт кегля надёжен, но глазами всё равно проверь.')
    except (OSError, subprocess.SubprocessError) as e:
        print('рендер не удался:', e, file=sys.stderr)
        return 1
    return 1 if problems else 0


if __name__ == '__main__':
    sys.exit(main() if len(sys.argv) > 1 else __doc__)
