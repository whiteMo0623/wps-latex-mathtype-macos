"""Encode a supported presentation MathML expression as a MathType OLE object."""

from __future__ import annotations

from functools import lru_cache
import io
import re
import struct
import unicodedata
from uuid import UUID

from lxml import etree
import olefile
from docx_equation.mathtype import mtef
from docx_equation.mathtype.ole import build_mathtype_ole_object


CLASS_ID = UUID('0002CE03-0000-0000-C000-000000000046')
END = b'\0'
NO_LINE = b'\x01\x01'
FENCES = {'(': 1, ')': 1, '{': 2, '}': 2, '[': 3, ']': 3, '|': 4,
          '‖': 5, '∥': 5, '⌊': 6, '⌋': 6, '⌈': 7, '⌉': 7, '⟨': 0, '⟩': 0}
OPERATORS = {'∫': (15, 1), '∬': (15, 2), '∭': (15, 3), '∮': (15, 5),
             '∑': (16, 0), '∏': (17, 0), '∐': (18, 0), '⋃': (19, 0), '⋂': (20, 0)}
GREEK = dict(zip('ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'))
GREEK.update(zip('αβχδεφγηιϕκλμνοπθρστυϖωξψζ', 'abcdefghijklmnopqrstuvwxyz'))
SYMBOL_CODES = {**GREEK, 'ϵ': 'e', 'ϱ': 'r', '∞': 0xa5, '±': 0xb1, '≤': 0xa3, '≥': 0xb3,
                '≠': 0xb9, '≈': 0xbb, '∈': 0xce, '∉': 0xcf, '∂': 0xb6, '∇': 0xd1,
                '→': 0xae, '←': 0xac, '↔': 0xab, '∫': 0xf2, '∑': 0xe5, '∏': 0xd5,
                '∩': 0xc7, '∪': 0xc8, '⊂': 0xcc, '⊆': 0xcd, '·': 0xd7, '′': 0xa2}
STYLES = {'normal': 1, 'italic': 3, 'bold': 7, 'bold-italic': 9, 'monospace': 10}
FANCY_NAMES = {'double-struck': 'DOUBLE-STRUCK', 'script': 'SCRIPT', 'fraktur': 'FRAKTUR'}
BLACKBOARD_GLYPHS = {'ℝ': 0xa1, 'ℤ': 0xa2, 'ℂ': 0xa3, 'ℚ': 0xa4, 'ℕ': 0xa5}
SHORT_ACCENTS = {'^': 9, '~': 8, '˜': 8, '¯': 17, '‾': 17, '˙': 2, '¨': 3}


class UnsupportedFormula(ValueError):
    pass


def template(kind: int, variation: int, contents: bytes) -> bytes:
    return mtef._template(kind, variation, 0) + contents + END


def line(contents: bytes | None) -> bytes:
    return NO_LINE if contents is None else b'\x01\x00' + contents + END


def tag(node: etree._Element) -> str:
    return etree.QName(node).localname


def text_of(node: etree._Element) -> str:
    return ''.join(node.itertext()).strip()


def char_record(char: str, style: int, flags: int = 0) -> bytes:
    if char in ('\u2061', '\u2062', '\u2063', '\u2064'):
        return b''
    if char in BLACKBOARD_GLYPHS:
        # MathType needs both MTCode and the position in its MT Extra font.
        return bytes([2, flags | 4, 139]) + struct.pack('<H', ord(char)) + bytes([BLACKBOARD_GLYPHS[char]])
    if char in mtef._OPERATOR_CODE:
        code, font_code = mtef._OPERATOR_CODE[char]
        return bytes([2, 4, 134]) + struct.pack('<H', code) + bytes([font_code])
    if ord(char) > 0xffff or unicodedata.category(char) in ('Cc', 'Cs', 'Co'):
        raise UnsupportedFormula(f'MathType 字符表未覆盖 {char!r}')
    if style in (4, 5, 6) and char in SYMBOL_CODES:
        mapping = SYMBOL_CODES[char]
        code = ord(mapping) if isinstance(mapping, str) else mapping
        return bytes([2, flags | 4, 128 + style]) + struct.pack('<H', ord(char)) + bytes([code])
    return bytes([2, flags, 128 + style]) + struct.pack('<H', ord(char))


@lru_cache(maxsize=128)
def fancy_char(char: str, variant: str) -> str:
    role = 'CAPITAL' if char.isupper() else 'SMALL'
    style = FANCY_NAMES[variant]
    for name in (f'{style} {role} {char.upper()}', f'MATHEMATICAL {style} {role} {char.upper()}'):
        try:
            return unicodedata.lookup(name)
        except KeyError:
            continue
    raise UnsupportedFormula(f'不支持 {variant} 字体的 {char!r}')


def token(node: etree._Element, variant: str | None) -> bytes:
    value = ''.join(node.itertext()).replace('\u00a0', ' ')
    kind = tag(node)
    if variant in FANCY_NAMES:
        value = ''.join(fancy_char(c, variant) for c in value)
        if variant == 'double-struck' and any(c not in BLACKBOARD_GLYPHS for c in value):
            raise UnsupportedFormula(f'MathType 字符表未覆盖双线体 {value!r}')
        variant = 'normal'
    if variant is not None and variant not in STYLES:
        raise UnsupportedFormula(f'不支持 MathML 字体 {variant}')
    result = bytearray()
    for index, char in enumerate(value):
        if variant:
            style = STYLES[variant]
        elif char in GREEK or char in ('ϵ', 'ϱ'):
            style = 5 if char.isupper() else 4
        elif kind == 'mn':
            style = 8
        elif kind == 'mo' and not char.isalpha():
            style = 6
        elif kind == 'mi' and len(value) == 1:
            style = 3
        else:
            style = 2 if kind in ('mi', 'mo') else 1
        flags = 2 if style == 2 and index == 0 and len(value) > 1 else 0
        result.extend(char_record(char, style, flags))
    return bytes(result)


def encode(node: etree._Element, variant: str | None = None, level: int = 0) -> bytes:
    variant = node.get('mathvariant', variant)
    if any(node.get(x) for x in ('mathcolor', 'mathbackground')):
        raise UnsupportedFormula('暂不转换带颜色的公式')
    kind = tag(node)
    children = list(node)
    full = bytes([10 + min(level, 2)])
    sub = bytes([10 + min(level + 1, 2)])
    if kind in ('mi', 'mn', 'mo', 'mtext'):
        return token(node, variant)
    if kind == 'mspace':
        width = node.get('width', '0em')
        m = re.fullmatch(r'(-?\d*\.?\d+)em', width)
        if m and float(m[1]) > 0:
            # MTEF distinguishes thin and wider spacing from a text space.
            return char_record(' ', 1) * max(1, min(4, round(float(m[1]) * 4)))
        return b''
    if kind in ('math', 'mrow', 'mstyle', 'mtd', 'TeXAtom'):
        if kind == 'mrow' and len(children) > 1:
            left, right = children[0], children[-1]
            if (left.get('data-mjx-texclass'), right.get('data-mjx-texclass')) == ('OPEN', 'CLOSE'):
                a, b = text_of(left), text_of(right)
                selector = FENCES.get(a or b)
                if selector is None or (b and FENCES.get(b) != selector):
                    raise UnsupportedFormula(f'不支持伸缩括号 {a}{b}')
                if a == '⟨':
                    a = '〈'
                if b == '⟩':
                    b = '〉'
                body = b''.join(encode(c, variant, level) for c in children[1:-1])
                fence_chars = b''.join(char_record(c, 22) for c in a + b)
                return template(selector, bool(a) + 2 * bool(b), line(body) + fence_chars)
        return b''.join(encode(c, variant, level) for c in children)
    if kind == 'mfrac' and len(children) == 2:
        if node.get('linethickness') not in (None, '1'):
            raise UnsupportedFormula('暂不转换无横线分式或二项式模板')
        a, b = [line(encode(c, variant, level)) for c in children]
        return template(11, 0, a + full + b)
    if kind in ('msqrt', 'mroot'):
        body = children if kind == 'msqrt' else children[:1]
        index = None if kind == 'msqrt' else encode(children[1], variant, level + 1)
        radicand = b''.join(encode(c, variant, level) for c in body)
        return template(10, int(index is not None), line(radicand) + sub + line(index)) + full
    if kind in ('msub', 'msup', 'msubsup', 'munder', 'mover', 'munderover'):
        base, *scripts = children
        base_text = text_of(base)
        if kind == 'mover':
            accents = {'¯': 13, '‾': 13, '^': 33, '~': 32, '˜': 32, '→': 31}
            accent_node = scripts[0]
            accent = text_of(accent_node)
            base_record = encode(base, variant, level)
            if (accent in SHORT_ACCENTS and len(base_text) == 1 and base_record[:1] == b'\x02'
                    and len(base_record) == (6 if base_record[1] & 4 else 5)
                    and (accent_node.get('stretchy') == 'false' or accent in ('˙', '¨'))):
                # A short accent belongs to the character, including when a script follows it.
                options = base_record[1] | 1
                return bytes([2, options]) + base_record[2:] + bytes([6, 0, SHORT_ACCENTS[accent], 0])
            if accent in accents:
                return template(accents[accent], 2 if accent == '→' else 0, line(base_record))
            raise UnsupportedFormula(f'不支持的顶部标记 {accent!r}')
        if kind == 'munder' and text_of(scripts[0]) in ('_', '‾', '\u0332'):
            return template(12, 0, line(encode(base, variant, level)))
        lower = encode(scripts[0], variant, level + 1) if kind != 'msup' else None
        upper = encode(scripts[-1], variant, level + 1) if kind in ('msup', 'msubsup', 'munderover') else None
        options = (0x10 if lower is not None else 0) | (0x20 if upper is not None else 0)
        if base_text in OPERATORS:
            selector, options2 = OPERATORS[base_text]
            # Count and loop are in the low nibble. Limits use bits 4 and 5.
            variation = options2 | options | (0x40 if kind in ('munder', 'munderover') else 0)
            operator = char_record(base_text, 6)
            return template(selector, variation, line(b'') + sub + line(lower) + line(upper) + b'\x0d' + operator) + full
        base_record = encode(base, variant, level)
        if kind in ('munder', 'munderover'):
            return template(23, options, line(base_record) + sub + line(lower) + line(upper)) + full
        selector = {'msub': 27, 'msup': 28, 'msubsup': 29}[kind]
        return base_record + template(selector, 0, sub + line(lower) + line(upper)) + full
    if kind == 'mtable':
        if any(tag(row) != 'mtr' for row in children) or node.get('rowlines') or node.get('columnlines'):
            raise UnsupportedFormula('暂不转换带标号或分隔线的矩阵')
        rows = [[encode(c, variant, level) for c in row] for row in children]
        cols = max(map(len, rows), default=0)
        if not 0 < len(rows) < 128 or not 0 < cols < 128:
            raise UnsupportedFormula('矩阵尺寸超出支持范围')
        aligns = node.get('columnalign', 'center')
        if aligns == 'right left':
            return b'\x04\x00\x04\x03' + b''.join(line(b''.join(row)) for row in rows) + END
        if len(set(aligns.split())) != 1:
            raise UnsupportedFormula('暂不转换带混合列对齐的矩阵')
        # Match native MATRIX records. Matrix alignment codes differ from PILE codes.
        align = {'left': 0, 'center': 1, 'right': 2}[aligns.split()[0]]
        cells = b''.join(line(cell) for row in rows for cell in row + [b''] * (cols - len(row)))
        return bytes([5, 0, 1, align, 1, len(rows), cols]) + bytes((len(rows) + 4)//4 + (cols + 4)//4) + cells + END
    raise UnsupportedFormula(f'暂不转换公式结构 {kind}')


def compound_object(mathml: str, display: bool) -> bytes:
    root = etree.fromstring(mathml.encode(), etree.XMLParser(resolve_entities=False, no_network=True))
    body = encode(root)
    if not body:
        raise UnsupportedFormula('公式为空')
    preamble = bytearray(mtef._preamble('DSMT4'))
    preamble[11] = int(not display)
    # Text, functions, variables and vector styles share the primary font.
    preamble = preamble.replace(b'\x11\x05Times New Roman\x00', b'\x11\x04MT Extra\x00', 1)
    preamble = preamble.replace(b'Arial\x00', b'Times New Roman\x00', 1)
    # Configure user styles 1 and 2 for bold italic and monospace.
    preamble[-7:] = b'\x01\x03\x03\x00\x04\x00\x01\x00'
    native = bytes(preamble) + b'\x0a' + line(body) + END
    if len(native) + 28 >= 4096:
        raise UnsupportedFormula('此公式超过当前 OLE 对象的长度上限')

    cfb = bytearray(build_mathtype_ole_object(native))
    directory = 512 * (1 + struct.unpack_from('<I', cfb, 48)[0])
    cfb[directory + 80:directory + 96] = CLASS_ID.bytes_le
    # The compound directory uses length-first name ordering and a red-black tree.
    # Stream order from the upstream packager is CompObj, Ole, ObjInfo, Equation Native.
    struct.pack_into('<I', cfb, directory + 76, 1)
    nodes = [(1, 1, 2, 3), (2, 1, -1, -1), (3, 1, -1, 4), (4, 0, -1, -1)]
    for number, color, left, right in nodes:
        entry = directory + 128 * number
        cfb[entry + 67] = color
        struct.pack_into('<ii', cfb, entry + 68, left, right)
    result = bytes(cfb)
    with olefile.OleFileIO(io.BytesIO(result), raise_defects=olefile.DEFECT_INCORRECT) as ole:
        if ole.root.clsid != str(CLASS_ID).upper() or ole.openstream('Equation Native').read()[28:] != native:
            raise ValueError('MathType 对象写入后读取失败')
    return result
