"""Read-only checks for PPTX package defects that can trigger PowerPoint Repair.

XML parsers accept an unbound prefix inside an xsi:type attribute value. Office
does not: that value is an XML QName and needs its own namespace declaration.
"""
import posixpath
import zipfile
from collections import Counter
from urllib.parse import unquote

from lxml import etree

XSI_TYPE = '{http://www.w3.org/2001/XMLSchema-instance}type'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def package_problems(path):
    """Return actionable errors without rewriting any part of the package."""
    problems = []
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            parts = set(names)
            duplicate = [name for name, count in Counter(names).items() if count > 1]
            problems.extend(f'дублирующаяся часть пакета: {name}' for name in duplicate)
            bad_crc = archive.testzip()
            if bad_crc:
                problems.append(f'ошибка CRC: {bad_crc}')
            for required in ('[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml'):
                if required not in parts:
                    problems.append(f'нет обязательной части пакета: {required}')
            roots = {}
            for name in names:
                if not name.endswith(('.xml', '.rels')):
                    continue
                try:
                    parser = etree.XMLParser(resolve_entities=False, no_network=True)
                    root = etree.fromstring(archive.read(name), parser)
                except etree.XMLSyntaxError as error:
                    problems.append(f'{name}: неверный XML: {error}')
                    continue
                roots[name] = root
                for node in root.iter():
                    value = node.get(XSI_TYPE)
                    if value and ':' in value:
                        prefix = value.split(':', 1)[0]
                        if prefix not in node.nsmap:
                            problems.append(
                                f'{name}: xsi:type="{value}" использует '
                                f'необъявленный префикс {prefix} '
                                f'в {etree.QName(node).localname}')
            relationships = {}
            for name, root in roots.items():
                if not name.endswith('.rels'):
                    continue
                owner = '' if name == '_rels/.rels' else posixpath.join(
                    posixpath.dirname(posixpath.dirname(name)),
                    posixpath.basename(name)[:-5])
                ids = set()
                for relation in root:
                    rid = relation.get('Id')
                    if rid in ids:
                        problems.append(f'{name}: повторный Id связи {rid}')
                    ids.add(rid)
                    if relation.get('TargetMode') == 'External':
                        continue
                    target = unquote(relation.get('Target', '')).split('#', 1)[0]
                    resolved = posixpath.normpath(posixpath.join(
                        posixpath.dirname(owner), target)).lstrip('/')
                    if resolved not in parts:
                        problems.append(f'{name}: связь {rid} ведёт в отсутствующую часть {resolved}')
                relationships[owner] = ids
            for name, root in roots.items():
                if name.endswith('.rels'):
                    continue
                for node in root.iter():
                    for key, value in node.attrib.items():
                        if key.startswith('{' + REL_NS + '}') and value not in relationships.get(name, set()):
                            problems.append(f'{name}: не разрешается связь {value}')
    except (OSError, zipfile.BadZipFile) as error:
        problems.append(f'не удалось прочитать пакет PPTX: {error}')
    return problems


def require_valid_package(path):
    problems = package_problems(path)
    if problems:
        raise ValueError(f'Некорректный пакет PPTX: {path}\n' + '\n'.join(problems))
