"""Build single-equation Flat OPC fragments for WPS Writer."""

from __future__ import annotations

import base64
from io import BytesIO
import re
from zipfile import ZipFile

from convert import CT_NS, DOCUMENT, NS, R, RELS, REL_NS, TYPES
from convert import Match, Package, Renderer, UnsupportedFormula, etree, q, scan

FLAT_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage'
DOCUMENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
OLE_TYPE = 'application/vnd.openxmlformats-officedocument.oleObject'
REL_TYPE = 'application/vnd.openxmlformats-package.relationships+xml'
MAIN_REL = NS['r'] + '/officeDocument'
# WPS uses control characters for paragraph, cell, object and field boundaries.
# A vertical tab is a soft line break and can belong to a display formula.
TEXT_GROUP = re.compile(r'[^\x00-\x08\x0c-\x1f\ufffc]+')
MAX_FORMULAS = 2000


def utf16_length(text: str) -> int:
    return len(text.encode('utf-16-le', errors='surrogatepass')) // 2


def scan_segments(segments: list[str]) -> dict:
    matches = []
    warning_count = 0
    for segment_index, text in enumerate(segments):
        offset = 0
        previous = 0
        for group in TEXT_GROUP.finditer(text):
            parsed, warnings = scan(group[0].replace('\v', '\n'))
            warning_count += len(warnings)
            for match in parsed:
                start = group.start() + match.start
                end = group.start() + match.end
                utf16_start = offset + utf16_length(text[previous:start])
                offset = utf16_start + utf16_length(text[start:end])
                previous = end
                matches.append(dict(segment=segment_index, start=utf16_start, end=offset,
                                    latex=match.latex, display=match.display))
                if len(matches) > MAX_FORMULAS:
                    raise ValueError(f'选区公式超过 {MAX_FORMULAS} 个，请分批转换。')
    return dict(matches=matches, warningCount=warning_count)


def empty_document() -> bytes:
    doc = etree.Element(q('w:document'), nsmap=NS)
    etree.SubElement(doc, q('w:body'))
    with BytesIO() as buffer:
        with ZipFile(buffer, 'w') as source:
            source.writestr(DOCUMENT, etree.tostring(doc))
            source.writestr(TYPES, f'<Types xmlns="{CT_NS}"/>')
        return buffer.getvalue()


EMPTY_DOCUMENT = empty_document()


def add_part(package, name, content_type, data):
    part = etree.SubElement(package, f'{{{FLAT_NS}}}part', {
        f'{{{FLAT_NS}}}name': '/' + name,
        f'{{{FLAT_NS}}}contentType': content_type,
    })
    if isinstance(data, bytes):
        etree.SubElement(part, f'{{{FLAT_NS}}}binaryData').text = base64.b64encode(data).decode('ascii')
    else:
        etree.SubElement(part, f'{{{FLAT_NS}}}xmlData').append(data)


def make_fragment(renderer: Renderer, latex: str, display: bool, size: float) -> str:
    with ZipFile(BytesIO(EMPTY_DOCUMENT)) as source:
        equation_package = Package(source)
    equation_package.default_size = size
    run = equation_package.make_equation(Match(0, len(latex), latex, display), etree.Element(R), renderer)
    paragraph = etree.SubElement(equation_package.doc.find(q('w:body')), q('w:p'))
    # WPS adds this marker when WordOpenXML is taken from part of a paragraph.
    properties = etree.SubElement(paragraph, q('w:pPr'))
    etree.SubElement(properties, q('w:doNotImportParagraphEnd'))
    paragraph.append(run)

    package = etree.Element(f'{{{FLAT_NS}}}package', nsmap={'pkg': FLAT_NS})
    root_rels = etree.Element(f'{{{REL_NS}}}Relationships', nsmap={None: REL_NS})
    etree.SubElement(root_rels, f'{{{REL_NS}}}Relationship', Id='rId1', Type=MAIN_REL, Target=DOCUMENT)
    add_part(package, '_rels/.rels', REL_TYPE, root_rels)
    add_part(package, DOCUMENT, DOCUMENT_TYPE, equation_package.doc)
    add_part(package, RELS, REL_TYPE, equation_package.rels)
    for name, data in equation_package.added.items():
        add_part(package, name, 'image/png' if name.endswith('.png') else OLE_TYPE, data)
    return etree.tostring(package, encoding='unicode')


def render_batch(items: list[dict]) -> list[dict]:
    if len(items) > MAX_FORMULAS:
        raise ValueError(f'单次最多转换 {MAX_FORMULAS} 个公式。')
    for item in items:
        if (not isinstance(item, dict) or not isinstance(item.get('latex'), str)
                or not item['latex'].strip() or not isinstance(item.get('display'), bool)):
            raise ValueError('公式请求格式有误。')
    results = []
    with Renderer() as renderer:
        for item in items:
            size = item.get('size')
            if type(size) not in (float, int) or not 6 <= size <= 72:
                size = 11
            try:
                xml = make_fragment(renderer, item['latex'], item['display'], size)
            except UnsupportedFormula as error:
                results.append(dict(error=str(error)))
            except (UnicodeError, etree.XMLSyntaxError):
                results.append(dict(error='公式包含无法写入 MathType 的字符'))
            else:
                results.append(dict(xml=xml))
    return results
