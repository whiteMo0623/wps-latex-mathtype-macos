"""Replace delimited LaTeX in a local DOCX with embedded MathType equations."""

from __future__ import annotations

import argparse
import base64
from copy import deepcopy
from dataclasses import dataclass
import itertools
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from zipfile import BadZipFile, ZipFile

HERE = Path(__file__).resolve().parent

from lxml import etree
from docx_equation.mathtype.ooxml import NS, REL_NS, CT_NS, IMAGE_REL, OLE_REL, make_object_run
from mtef_writer import UnsupportedFormula, compound_object


def q(name: str) -> str:
    prefix, local = name.split(':')
    return f'{{{NS[prefix]}}}{local}'


P, R, T, PPR, RPR, BR = map(q, ('w:p', 'w:r', 'w:t', 'w:pPr', 'w:rPr', 'w:br'))
SPACE = '{http://www.w3.org/XML/1998/namespace}space'
DOCUMENT = 'word/document.xml'
RELS = 'word/_rels/document.xml.rels'
TYPES = '[Content_Types].xml'
OPENERS = {r'\(': r'\)', r'\[': r'\]', '$': '$', '$$': '$$'}
DELIMITER = re.compile(r'\\[\[\]()]|\${1,2}')


@dataclass(frozen=True)
class Match:
    start: int
    end: int
    latex: str
    display: bool


def escaped(text: str, start: int) -> bool:
    previous = start
    while previous > 0 and text[previous - 1] == '\\':
        previous -= 1
    return (start - previous) % 2 == 1


def scan(text: str) -> tuple[list[Match], list[str]]:
    tokens = [m for m in DELIMITER.finditer(text) if not escaped(text, m.start())]
    matches, unmatched = [], []
    i = 0
    while i < len(tokens):
        opening = tokens[i]
        key = opening[0]
        i += 1
        if key not in OPENERS:
            unmatched.append(f'找不到开始定界符 {key}')
            continue
        if key == '$' and (not text[opening.end():] or text[opening.end()].isspace()):
            continue
        if i == len(tokens) or tokens[i][0] != OPENERS[key]:
            if key != '$' or (text[opening.end():] and not text[opening.end()].isdigit()):
                unmatched.append(f'找不到结束定界符 {OPENERS[key]}')
            continue
        closing = tokens[i]
        inner = text[opening.end():closing.start()]
        after = text[closing.end():closing.end() + 1]
        if key == '$' and (not inner or inner[-1].isspace() or after.isdigit() or '\n' in inner):
            if inner and not inner[0].isdigit():
                unmatched.append('无法确定单个美元符号之间的公式范围')
            continue
        i += 1
        if inner.strip():
            matches.append(Match(opening.start(), closing.end(), inner.strip(), key in (r'\[', '$$')))
        else:
            unmatched.append('公式定界符之间没有内容')
    return matches, unmatched


def text_run(run: etree._Element) -> bool:
    return run.tag == R and any(c.tag == T or c.tag == BR for c in run) and all(
        c.tag in (RPR, T) or (c.tag == BR and c.get(q('w:type')) in (None, 'textWrapping')) for c in run
    )


def node_text(node: etree._Element) -> str:
    if node.tag == T:
        return node.text or ''
    return '\n' if node.tag == BR else ''


def run_text(run: etree._Element) -> str:
    return ''.join(map(node_text, run))


def slice_run(run: etree._Element, start: int, end: int) -> etree._Element:
    if start == 0 and end == len(run_text(run)):
        return run
    copy = etree.Element(R, run.attrib, nsmap=run.nsmap)
    position = 0
    for child in run:
        if child.tag == RPR:
            copy.append(deepcopy(child))
            continue
        value = node_text(child)
        lower, upper = max(0, start - position), min(len(value), end - position)
        position += len(value)
        if lower >= upper:
            continue
        piece = deepcopy(child)
        if child.tag == T:
            piece.text = value[lower:upper]
            if piece.text != piece.text.strip():
                piece.set(SPACE, 'preserve')
        copy.append(piece)
    return copy


def spans(runs):
    end = 0
    for run in runs:
        content = run_text(run)
        yield end, end + len(content), run
        end += len(content)


def replace_group(paragraph, runs, changes):
    positions = list(spans(runs))
    offset = paragraph.index(runs[0])
    pieces = []
    cursor = 0
    end = positions[-1][1]
    for match, replacement in [*changes, (Match(end, end, '', False), None)]:
        for lower, upper, run in positions:
            start, stop = max(lower, cursor), min(upper, match.start)
            if start < stop:
                pieces.append(slice_run(run, start - lower, stop - lower))
        if replacement is not None:
            pieces.append(replacement)
        cursor = match.end
    for run in runs:
        paragraph.remove(run)
    for index, piece in enumerate(pieces, offset):
        paragraph.insert(index, piece)


def paragraph_groups(paragraph, field_depth: int):
    group, groups = [], []
    for element in paragraph:
        if text_run(element) and not field_depth:
            group.append(element)
            continue
        if group:
            groups.append(group)
            group = []
        for field in element.findall(q('w:fldChar')) if element.tag == R else ():
            kind = field.get(q('w:fldCharType'))
            if kind == 'begin':
                field_depth += 1
            if kind == 'end':
                field_depth = max(0, field_depth - 1)
    if group:
        groups.append(group)
    return groups, field_depth


def plain_paragraph(p) -> str | None:
    if p.tag != P or any(c.tag != PPR and not text_run(c) for c in p):
        return None
    if p.find(f'{PPR}/{q("w:sectPr")}') is not None:
        return None
    return ''.join(run_text(c) for c in p if c.tag == R)


def display_block(p):
    first = plain_paragraph(p)
    if first is None or not first.lstrip().startswith((r'\[', '$$')) or scan(first)[0]:
        return None
    paragraphs = [p]
    value = first
    for sibling in itertools.islice(p.itersiblings(), 80):
        part = plain_paragraph(sibling)
        if part is None or part.lstrip().startswith((r'\[', '$$')) and part.strip() != '$$':
            return None
        value += '\n' + part
        paragraphs.append(sibling)
        matches, _ = scan(value)
        if len(matches) == 1:
            match = matches[0]
            if not value[:match.start].strip() and not value[match.end:].strip():
                return paragraphs, match
            return None
    return None


class Renderer:
    def __enter__(self):
        bundled_node = HERE.parent / 'node/bin/node'
        node = str(bundled_node) if bundled_node.is_file() else shutil.which('node')
        if node is None:
            raise RuntimeError('找不到 Node.js。请使用完整发行包，或为源码开发环境安装 Node.js。')
        self.process = subprocess.Popen(
            [node, str(HERE / 'render-equations.cjs')],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True, encoding='utf-8', bufsize=1,
        )
        self.cache = {}
        return self

    def __exit__(self, *_):
        self.process.stdin.close()
        self.process.stdout.close()
        self.process.wait(timeout=5)

    def convert(self, match: Match, size: float):
        if len(match.latex) > 6000:
            raise UnsupportedFormula('公式源码过长')
        key = (match.latex, match.display, size)
        if key not in self.cache:
            request = json.dumps(dict(latex=match.latex, display=match.display, size=size))
            self.process.stdin.write(request + '\n')
            message = self.process.stdout.readline()
            if not message:
                raise RuntimeError('公式渲染程序退出')
            rendered = json.loads(message)
            if 'error' in rendered:
                raise UnsupportedFormula(rendered['error'])
            ole = compound_object(rendered['mathml'], match.display)
            preview = base64.b64decode(rendered['png'])
            if len(self.cache) >= 256:
                self.cache.clear()
            self.cache[key] = (ole, preview, rendered['widthPx'], rendered['heightPx'], rendered['depth'])
        return self.cache[key]


def parse(data: bytes):
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))


def serialize(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


class Package:
    def __init__(self, source: ZipFile):
        self.doc = parse(source.read(DOCUMENT))
        self.rels = parse(source.read(RELS)) if RELS in source.namelist() else etree.Element(f'{{{REL_NS}}}Relationships', nsmap={None: REL_NS})
        self.types = parse(source.read(TYPES))
        self.default_size = 11.0
        if 'word/styles.xml' in source.namelist():
            default = parse(source.read('word/styles.xml')).find(
                f'{q("w:docDefaults")}/{q("w:rPrDefault")}/{RPR}/{q("w:sz")}'
            )
            if default is not None:
                self.default_size = self.font_size(default, 11.0)
        self.existing = set(source.namelist())
        self.rel_ids = {r.get('Id') for r in self.rels}
        self.shape_ids = set(self.doc.xpath('//v:shape/@id|//o:OLEObject/@ShapeID', namespaces=NS))
        self.object_ids = set(self.doc.xpath('//o:OLEObject/@ObjectID', namespaces=NS))
        self.added = {}
        self.index = 0
        self.count = 0

    @staticmethod
    def font_size(node, fallback):
        value = node.get(q('w:val'), '') if node is not None else ''
        return int(value) / 2 if value.isdecimal() and 12 <= int(value) <= 144 else fallback

    def relationship(self, kind, name, rel_id):
        etree.SubElement(self.rels, f'{{{REL_NS}}}Relationship', Id=rel_id, Type=kind, Target=name)

    def make_equation(self, match, run, renderer):
        size = self.font_size(run.find(f'{RPR}/{q("w:sz")}'), self.default_size)
        ole, preview, width, height, depth = renderer.convert(match, size)
        while True:
            self.index += 1
            i = self.index
            image = f'media/mt-preview-{i}.png'
            data = f'embeddings/mt-equation-{i}.bin'
            img_rel, ole_rel = f'rIdMTPng{i}', f'rIdMTOle{i}'
            if (f'word/{image}' not in self.existing and f'word/{data}' not in self.existing
                and img_rel not in self.rel_ids and ole_rel not in self.rel_ids
                and f'_x0000_i{1024+i}' not in self.shape_ids and f'_{1618029000+i}' not in self.object_ids):
                break
        replacement = make_object_run(img_rel, ole_rel, width, height, index=i, preview_pt_per_px=0.2, max_width_pt=421)
        obj = replacement.find(q('w:object'))
        obj.set(q('w:dxaOrig'), str(width * 4))
        obj.set(q('w:dyaOrig'), str(height * 4))
        obj.find(q('v:shape')).set('style', f'width:{width/5:.1f}pt;height:{height/5:.1f}pt')
        if not match.display and depth > 0.25:
            props = etree.Element(RPR)
            etree.SubElement(props, q('w:position'), {q('w:val'): str(-round(2 * depth))})
            replacement.insert(0, props)
        self.relationship(IMAGE_REL, image, img_rel)
        self.relationship(OLE_REL, data, ole_rel)
        self.added[f'word/{image}'] = preview
        self.added[f'word/{data}'] = ole
        self.count += 1
        return replacement

    def write(self, source, target):
        changes = {DOCUMENT: serialize(self.doc), RELS: serialize(self.rels)}
        for part in self.added:
            if part.endswith('.bin'):
                etree.SubElement(self.types, f'{{{CT_NS}}}Override', PartName='/' + part,
                                 ContentType='application/vnd.openxmlformats-officedocument.oleObject')
        if not any(t.get('Extension', '').lower() == 'png' for t in self.types):
            etree.SubElement(self.types, f'{{{CT_NS}}}Default', Extension='png', ContentType='image/png')
        changes[TYPES] = serialize(self.types)
        with ZipFile(target, 'w', compression=8) as out:
            for info in source.infolist():
                content = changes.pop(info.filename) if info.filename in changes else source.read(info)
                out.writestr(info, content)
            for name, data in {**changes, **self.added}.items():
                out.writestr(name, data)


def process(source: ZipFile, renderer):
    package = Package(source)
    body = package.doc.find(q('w:body'))
    if body is None:
        raise ValueError('需要另存为标准 Word .docx')
    paragraphs = [p for p in body.iter(P) if p.getparent().tag in (q('w:body'), q('w:tc'), q('w:sdtContent'))]
    removed, issues = set(), []
    field_depth = 0

    def convert_match(match, first_run, position):
        try:
            return package.make_equation(match, first_run, renderer)
        except UnsupportedFormula as exc:
            preview = ' '.join(match.latex.split())[:160]
            issues.append(f'正文/表格第 {position} 段 · {exc}\n  {preview}')
            return None

    for position, paragraph in enumerate(paragraphs, 1):
        if paragraph in removed:
            continue
        block = display_block(paragraph) if not field_depth else None
        if block:
            batch, match = block
            removed.update(batch[1:])
            equation = convert_match(match, paragraph.find(R), position)
            if equation is not None:
                for run in paragraph.findall(R):
                    paragraph.remove(run)
                paragraph.append(equation)
                for old in batch[1:]:
                    old.getparent().remove(old)
            continue
        groups, field_depth = paragraph_groups(paragraph, field_depth)
        for group in groups:
            combined = ''.join(map(run_text, group))
            matches, warnings = scan(combined)
            issues.extend(f'正文/表格第 {position} 段 · {warning}' for warning in warnings)
            changes = []
            for match in matches:
                first_run = next(r for lower, upper, r in spans(group) if lower <= match.start < upper)
                equation = convert_match(match, first_run, position)
                if equation is not None:
                    changes.append((match, equation))
            if changes:
                replace_group(paragraph, group, changes)
    return package, issues


def unique_copy(path: Path):
    candidate = path.with_name(f'{path.stem}_MathType7.docx')
    for i in itertools.count(2):
        report = candidate.with_name(candidate.stem + '_未转换.txt')
        if not candidate.exists() and not report.exists():
            return candidate
        candidate = path.with_name(f'{path.stem}_MathType7_{i}.docx')


def convert_file(path: Path, renderer, output_dir: Path | None):
    if path.suffix.lower() != '.docx':
        raise ValueError('请在 WPS 中先另存为 .docx')
    target = unique_copy(path if output_dir is None else output_dir / path.name)
    with ZipFile(path) as source:
        if any('vbaproject' in n.lower() for n in source.namelist()):
            raise ValueError('请选择另存为无宏的 .docx 文件')
        package, issues = process(source, renderer)
        if package.count:
            with tempfile.NamedTemporaryFile(dir=target.parent, prefix='.mathtype-', delete=False) as tmp:
                temp_path = Path(tmp.name)
            try:
                package.write(source, temp_path)
                os.link(temp_path, target)
            finally:
                temp_path.unlink(missing_ok=True)
    print(f'{path.name} · 转换 {package.count} 个，提示 {len(issues)} 条')
    if package.count:
        print(f'  {target}')
    if issues:
        report = target.with_name(target.stem + '_未转换.txt')
        report.write_text('\n\n'.join(issues) + '\n', encoding='utf-8')
        print(f'  检查清单 {report}')
    return bool(issues)


def main():
    parser = argparse.ArgumentParser(description='将 .docx 正文与表格中的 LaTeX 文本原位转换为 MathType 7 OLE。')
    parser.add_argument('docx', nargs='+', type=Path)
    parser.add_argument('--output-dir', type=Path)
    args = parser.parse_args()
    if args.output_dir is not None and not args.output_dir.is_dir():
        parser.error('--output-dir 必须为现有目录')
    errors = False
    with Renderer() as renderer:
        for path in args.docx:
            try:
                errors |= convert_file(path.expanduser().absolute(), renderer, args.output_dir)
            except (OSError, BadZipFile, KeyError, etree.XMLSyntaxError, ValueError) as exc:
                print(f'{path.name} · 转换失败 · {exc}', file=sys.stderr)
                errors = True
    return int(errors)


if __name__ == '__main__':
    sys.exit(main())
