"""Recover LaTeX from a MathType OLE object's MTEF v5 payload.

Record layout follows the archived Design Science MTEF v5 reference at
https://rtf2latex2e.sourceforge.net/MTEF5.html and the native MathType 7 samples.
"""

from __future__ import annotations

import base64
from dataclasses import dataclass
from io import BytesIO
import re
import struct
import unicodedata

import olefile

from fragments import MAX_FORMULAS


MAX_OLE_BYTES = 512 * 1024
NATIVE_HEADER = struct.Struct('<HIHI4I')


class UnsupportedEquation(ValueError):
    pass


@dataclass(frozen=True)
class Character:
    code: int
    face: int
    function_start: bool
    accents: tuple[int, ...]


@dataclass(frozen=True)
class Slot:
    items: tuple[Record, ...]


@dataclass(frozen=True)
class Template:
    selector: int
    variation: int
    items: tuple[Record, ...]


@dataclass(frozen=True)
class Matrix:
    columns: int
    justify: int
    rows: tuple[tuple[Slot, ...], ...]


@dataclass(frozen=True)
class Pile:
    justify: int
    lines: tuple[Slot, ...]


Record = Character | Slot | Template | Matrix | Pile


def read_native(ole_data: bytes) -> bytes:
    if len(ole_data) > MAX_OLE_BYTES:
        raise UnsupportedEquation('公式对象过大')
    with olefile.OleFileIO(BytesIO(ole_data), raise_defects=olefile.DEFECT_FATAL) as ole:
        # WPS writes a nonzero file-header CLSID on some saved copies.
        if any(reason != 'incorrect CLSID in OLE header' for _, reason in ole.parsing_issues):
            raise UnsupportedEquation('OLE 容器结构异常')
        if not ole.exists('Equation Native'):
            raise UnsupportedEquation('缺少 Equation Native 流')
        data = ole.openstream('Equation Native').read()
    if len(data) < NATIVE_HEADER.size:
        raise UnsupportedEquation('Equation Native 数据不完整')
    header_size, _, _, length, *_ = NATIVE_HEADER.unpack_from(data)
    if header_size != NATIVE_HEADER.size or not 0 < length <= len(data) - header_size:
        raise UnsupportedEquation('Equation Native 的长度无效')
    return data[header_size:header_size + length]


class MTEFReader:
    # MathType's v5 records have byte tags. Definition and size records carry
    # formatting; the recursive object lists contain the mathematical structure.
    def __init__(self, data: bytes):
        self.data = data
        self.position = 0
        self.fonts: list[tuple[int, str]] = [(0, '')]
        self.styles: dict[int, tuple[int, int]] = {}

    def take(self, size: int) -> bytes:
        end = self.position + size
        if end > len(self.data):
            raise UnsupportedEquation('MTEF 数据在公式结束前截断')
        result = self.data[self.position:end]
        self.position = end
        return result

    def byte(self) -> int:
        return self.take(1)[0]

    def word(self) -> int:
        return int.from_bytes(self.take(2), 'little')

    def unsigned(self) -> int:
        value = self.byte()
        return self.word() if value == 255 else value

    def c_string(self) -> bytes:
        end = self.data.find(b'\0', self.position)
        if end < 0:
            raise UnsupportedEquation('MTEF 字符串缺少结束标记')
        return self.take(end - self.position + 1)[:-1]

    def options(self, allowed: int) -> int:
        options = self.byte()
        if options & ~allowed:
            raise UnsupportedEquation('存在不支持的 MTEF 记录选项')
        if options & 8 and self.take(2) == b'\x80\x80':
            self.take(4)
        return options

    def dimensions(self):
        count = self.byte()
        nibble_position = self.position * 2
        for _ in range(count):
            nibble_position += 1  # units
            while True:
                byte_index, lower = divmod(nibble_position, 2)
                if byte_index >= len(self.data):
                    raise UnsupportedEquation('MTEF 尺寸表不完整')
                value = self.data[byte_index]
                nibble_position += 1
                if (value & 15 if lower else value >> 4) == 15:
                    break
        self.position = (nibble_position + 1) // 2

    def read_equation(self) -> tuple[bool, tuple[Record, ...]]:
        if self.byte() != 5:
            raise UnsupportedEquation('当前支持 MTEF v5 公式')
        self.take(4)  # platform, product, and the generating MathType version
        self.c_string()
        display = not bool(self.byte() & 1)
        items = self.records(0)
        if not items or any(self.data[self.position:]):
            raise UnsupportedEquation('MTEF 公式为空或含有未解析的数据')
        return display, items

    def character(self) -> Character:
        opts = self.options(0x3f)
        face_byte = self.byte()
        face = self.word() - 32768 if face_byte == 255 else face_byte - 128
        code = None if opts & 0x20 else self.word()
        if opts & 0x14 == 0x14:
            raise UnsupportedEquation('字符包含冲突的字体编码')
        font_position = self.byte() if opts & 4 else self.word() if opts & 0x10 else None
        if code is None:
            font, _ = self.styles.get(face, (0, 0))
            if font >= len(self.fonts) or self.fonts[font][0] != 1 or font_position is None:
                raise UnsupportedEquation('存在缺少可识别 MTCode 的字符')
            code = font_position
        accents = []
        if opts & 1:
            while (tag := self.byte()) != 0:
                if tag != 6:
                    raise UnsupportedEquation('字符装饰记录不完整')
                self.options(8)
                accents.append(self.byte())
        return Character(code, face, bool(opts & 2), tuple(accents))

    def records(self, depth: int) -> tuple[Record, ...]:
        if depth > 64:
            raise UnsupportedEquation('公式嵌套层数过多')
        items = []
        while (tag := self.byte()) != 0:
            if tag == 1:
                opts = self.options(0x0f)
                if opts & 4:
                    self.word()
                if opts & 2:
                    self.ruler()
                items.append(Slot(()) if opts & 1 else Slot(self.records(depth + 1)))
            elif tag == 2:
                items.append(self.character())
            elif tag == 3:
                self.options(8)
                selector, variation = self.take(2)
                if variation & 0x80:
                    variation = (variation & 0x7f) | (self.byte() << 8)
                self.byte()  # template-specific alignment and sizing
                items.append(Template(selector, variation, self.records(depth + 1)))
            elif tag == 4:
                opts = self.options(0x0a)
                justify, _ = self.take(2)
                if opts & 2:
                    self.ruler()
                items.append(Pile(justify, require_slots(self.records(depth + 1))))
            elif tag == 5:
                self.options(8)
                _, justify, _, rows, columns = self.take(5)
                if not rows or not columns or rows * columns > 4096:
                    raise UnsupportedEquation('矩阵尺寸无效或过大')
                partitions = self.take((rows + 4) // 4 + (columns + 4) // 4)
                if any(partitions):
                    raise UnsupportedEquation('暂不转换带矩阵分隔线的公式')
                cells = require_slots(self.records(depth + 1))
                if len(cells) != rows * columns:
                    raise UnsupportedEquation('矩阵单元格数与矩阵尺寸不一致')
                items.append(Matrix(columns, justify, tuple(
                    cells[i:i + columns] for i in range(0, len(cells), columns)
                )))
            elif tag == 100 and depth == 0 and not items:
                self.take(self.unsigned())  # reserved extension before the equation
            else:
                self.definition(tag)
        return tuple(items)

    def ruler(self):
        if self.byte() != 7:
            raise UnsupportedEquation('MTEF 标尺记录不完整')
        self.take(self.byte() * 3)

    def definition(self, tag: int):
        if tag == 8:
            self.unsigned()
            self.byte()
        elif tag == 9:
            self.take({100: 3, 101: 2}.get(self.byte(), 1))
        elif 10 <= tag <= 14:
            pass
        elif tag == 15:
            self.unsigned()  # color
        elif tag == 16:
            flags = self.byte()
            self.take(8 if flags & 1 else 6)
            if flags & 4:
                self.c_string()
        elif tag == 17:
            encoding = self.unsigned()
            self.fonts.append((encoding, self.c_string().decode('ascii', errors='replace')))
        elif tag == 18:
            self.byte()
            self.dimensions()
            self.dimensions()
            for face in range(1, self.byte() + 1):
                font = self.unsigned()
                if font:
                    self.styles[face] = (font, self.byte())
        elif tag == 19:
            self.c_string()
        elif tag == 102:
            self.take(self.unsigned())  # MathType comment
        else:
            raise UnsupportedEquation(f'暂不支持 MTEF 记录类型 {tag}')


def require_slots(items: tuple[Record, ...]) -> tuple[Slot, ...]:
    if any(not isinstance(item, Slot) for item in items):
        raise UnsupportedEquation('公式槽位包含未识别的记录')
    return items


GREEK = dict(zip(
    'αβγδεζηθικλμνξοπρστυφχψωϵϑϕϖϱςϰΓΔΘΛΞΠΣΥΦΨΩ',
    r'\alpha \beta \gamma \delta \varepsilon \zeta \eta \theta \iota \kappa \lambda \mu \nu \xi o \pi \rho \sigma \tau \upsilon \phi \chi \psi \omega \epsilon \vartheta \varphi \varpi \varrho \varsigma \varkappa \Gamma \Delta \Theta \Lambda \Xi \Pi \Sigma \Upsilon \Phi \Psi \Omega'.split(),
    strict=True,
))
SYMBOLS = {
    **GREEK,
    '−': '-', '±': r'\pm', '∓': r'\mp', '×': r'\times', '÷': r'\div',
    '·': r'\cdot', '⋅': r'\cdot', '∗': r'\ast', '∘': r'\circ', '•': r'\bullet',
    '≤': r'\le', '≥': r'\ge', '≠': r'\ne', '≈': r'\approx', '≃': r'\simeq',
    '≅': r'\cong', '≡': r'\equiv', '∼': r'\sim', '∝': r'\propto',
    '≪': r'\ll', '≫': r'\gg', '⊥': r'\perp', '∥': r'\parallel', '∣': r'\mid',
    '∈': r'\in', '∉': r'\notin', '∋': r'\ni', '∅': r'\varnothing',
    '⊂': r'\subset', '⊃': r'\supset', '⊆': r'\subseteq', '⊇': r'\supseteq',
    '∪': r'\cup', '∩': r'\cap', '∖': r'\setminus', '⊕': r'\oplus', '⊗': r'\otimes',
    '∀': r'\forall', '∃': r'\exists', '∄': r'\nexists', '¬': r'\neg',
    '∧': r'\land', '∨': r'\lor', '⊢': r'\vdash', '⊨': r'\models',
    '→': r'\to', '←': r'\leftarrow', '↔': r'\leftrightarrow', '↦': r'\mapsto',
    '↑': r'\uparrow', '↓': r'\downarrow', '⇒': r'\Rightarrow',
    '⇐': r'\Leftarrow', '⇔': r'\Leftrightarrow', '⟶': r'\longrightarrow',
    '⟵': r'\longleftarrow', '↗': r'\nearrow', '↘': r'\searrow',
    '∞': r'\infty', '∂': r'\partial', '∇': r'\nabla',
    '∫': r'\int', '∬': r'\iint', '∭': r'\iiint', '∮': r'\oint',
    '∑': r'\sum', '∏': r'\prod', '∐': r'\coprod', '⋃': r'\bigcup', '⋂': r'\bigcap',
    '…': r'\ldots', '⋯': r'\cdots', '⋮': r'\vdots', '⋱': r'\ddots',
    '⌊': r'\lfloor', '⌋': r'\rfloor', '⌈': r'\lceil', '⌉': r'\rceil',
    '⟨': r'\langle', '⟩': r'\rangle', '〈': r'\langle', '〉': r'\rangle', '‖': r'\Vert',
    'ℝ': r'\mathbb{R}', 'ℤ': r'\mathbb{Z}', 'ℂ': r'\mathbb{C}',
    'ℚ': r'\mathbb{Q}', 'ℕ': r'\mathbb{N}', 'ℙ': r'\mathbb{P}', 'ℍ': r'\mathbb{H}',
    'ℓ': r'\ell', 'ℏ': r'\hbar', 'ℵ': r'\aleph', '℘': r'\wp',
    '∠': r'\angle', '△': r'\triangle', '∴': r'\therefore', '∵': r'\because',
    '′': r'\prime', '″': r'\prime\prime', '°': r'{}^\circ',
    '\u00a0': r'\ ', '\u2002': r'\enspace', '\u2003': r'\quad',
    '\u2004': r'\;', '\u2005': r'\:', '\u2009': r'\,',
    '\uef01': '', '\uef02': r'\,', '\uef03': r'\:', '\uef04': r'\;',
    '\uef05': r'\quad', '\uef06': r'\qquad', '\uef07': r'\,', '\uef08': r'\enspace',
    '\ued10': r'\mathrm{d}', '\ued11': r'\mathrm{e}', '\ued12': r'\mathrm{i}',
    '\ued13': r'\mathrm{j}', '\ued16': r'\mathrm{D}',
    '~': r'\sim', '\\': r'\backslash', ' ': r'\,',
}
ESCAPES = {c: '\\' + c for c in '#$%&_{}'}
FUNCTIONS = set('arccos arcsin arctan arg cos cosh cot coth csc deg det dim exp gcd hom inf ker '
                'lg lim liminf limsup ln log max min Pr sec sin sinh sup tan tanh'.split())
FENCES = (
    (r'\langle', r'\rangle'), ('(', ')'), (r'\{', r'\}'), ('[', ']'),
    ('|', '|'), (r'\lVert', r'\rVert'), (r'\lfloor', r'\rfloor'), (r'\lceil', r'\rceil'),
)
BIG_OPERATORS = {
    16: r'\sum', 17: r'\prod', 18: r'\coprod', 19: r'\bigcup', 20: r'\bigcap',
}
INTEGRALS = {1: r'\int', 2: r'\iint', 3: r'\iiint', 5: r'\oint'}
ACCENTS = {
    2: 'dot', 3: 'ddot', 4: 'dddot', 8: 'tilde', 9: 'hat', 11: 'vec',
    12: 'overleftarrow', 13: 'overleftrightarrow', 17: 'bar', 24: 'ddddot',
    29: 'underline', 33: 'underrightarrow', 34: 'underleftarrow', 35: 'underleftrightarrow',
}
PRIMES = {5: 1, 6: 2, 18: 3}
CONTROL_WORD_END = re.compile(r'\\[A-Za-z]+$')


def function_letter(code: int) -> bool:
    return 65 <= code <= 90 or 97 <= code <= 122


def join_latex(pieces) -> str:
    parts = []
    for piece in pieces:
        if not piece:
            continue
        if parts and piece[0].isascii() and piece[0].isalpha() and CONTROL_WORD_END.search(parts[-1]):
            parts.append(' ')
        parts.append(piece)
    return ''.join(parts)


def command(name: str, *arguments: str) -> str:
    return '\\' + name + ''.join('{' + value + '}' for value in arguments)


def math_char(code: int) -> str:
    char = chr(code)
    if char in SYMBOLS:
        return SYMBOLS[char]
    if char in ESCAPES:
        return ESCAPES[char]
    if char.isascii() and char.isprintable() and char != '^':
        return char
    name = unicodedata.name(char, '')
    fancy = re.fullmatch(r'(?:MATHEMATICAL )?(DOUBLE-STRUCK|FRAKTUR|SCRIPT) (CAPITAL|SMALL) ([A-Z])', name)
    if fancy:
        style, case, letter = fancy.groups()
        return command({'DOUBLE-STRUCK': 'mathbb', 'FRAKTUR': 'mathfrak', 'SCRIPT': 'mathcal'}[style],
                       letter if case == 'CAPITAL' else letter.lower())
    raise UnsupportedEquation(f'暂不支持 MTCode U+{code:04X}')


def text_latex(value: str) -> str:
    escapes = {**ESCAPES, '\\': r'\textbackslash{}', '^': r'\textasciicircum{}', '~': r'\textasciitilde{}'}
    if any(not ch.isprintable() or unicodedata.category(ch)[0] == 'C' for ch in value):
        raise UnsupportedEquation('文本中存在不支持的字符')
    return ''.join(escapes.get(ch, ch) for ch in value)


class LaTeXWriter:
    def __init__(self, reader: MTEFReader):
        self.reader = reader
        # The bundled writer uses different matrix alignment values from the native samples.
        self.app_matrix_encoding = any(name == 'MT Extra Tiger' for _, name in reader.fonts) and (
            reader.styles.get(9) == (1, 3)
        )

    def sequence(self, items: tuple[Record, ...]) -> str:
        parts = []
        index = 0
        scripted = False
        while index < len(items):
            item = items[index]
            if isinstance(item, Template) and item.selector in (27, 28, 29):
                lower, upper = self.arguments(item, 2)
                if item.variation & ~1 or (item.selector == 27 and upper) or (item.selector == 28 and lower):
                    raise UnsupportedEquation('上下标模板包含未识别的槽位')
                if item.variation & 1 or not parts:
                    parts.append('{}')
                if scripted:
                    parts[-1] = '{' + parts[-1] + '}'
                parts[-1] += self.scripts(lower, upper)
                scripted = True
            elif (isinstance(item, Character) and item.face in (1, 2, 12) and not item.accents
                  and (item.face != 2 or function_letter(item.code))):
                face = item.face
                chars = [chr(item.code)]
                while index + 1 < len(items):
                    next_item = items[index + 1]
                    if (not isinstance(next_item, Character) or next_item.face != face
                            or next_item.accents or next_item.function_start):
                        break
                    if face == 2 and not function_letter(next_item.code):
                        break
                    index += 1
                    chars.append(chr(next_item.code))
                value = ''.join(chars)
                if face == 2:
                    parts.append('\\' + value if value in FUNCTIONS else command('operatorname', value))
                else:
                    parts.append(self.text(value))
                scripted = False
            else:
                parts.append(self.render(item))
                scripted = isinstance(item, Character) and any(accent in PRIMES for accent in item.accents)
            index += 1
        return join_latex(parts)

    @staticmethod
    def scripts(lower: str, upper: str) -> str:
        return ('_{' + lower + '}' if lower else '') + ('^{' + upper + '}' if upper else '')

    @staticmethod
    def text(value: str) -> str:
        if not value.strip(' '):
            return r'\,' * len(value)
        plain = value.strip(' ')
        if plain.isascii() and plain.isalnum():
            return (r'\,' * (len(value) - len(value.lstrip(' '))) + command('mathrm', plain)
                    + r'\,' * (len(value) - len(value.rstrip(' '))))
        return command('text', text_latex(value))

    def character(self, item: Character) -> str:
        result = self.text(chr(item.code)) if item.face in (1, 12) else math_char(item.code)
        font, style = self.reader.styles.get(item.face, (0, 0))
        family = self.reader.fonts[font][1].lower() if 0 < font < len(self.reader.fonts) else ''
        if result and item.face in (7, 9, 10):
            if 'courier' in family or 'monaco' in family or 'mono' in family:
                result = command('mathtt', result)
            elif style & 1 or item.face == 7:
                macro = 'boldsymbol' if style & 2 or not result.isalnum() else 'mathbf'
                result = command(macro, result)
            elif item.face in (9, 10) and result.isascii() and result.isalnum() and not style & 2:
                result = command('mathrm', result)
        for accent in item.accents:
            if accent in PRIMES:
                result += '^{' + r'\prime' * PRIMES[accent] + '}'
            elif accent == 10:
                result = r'\not ' + result
            elif accent in ACCENTS:
                result = command(ACCENTS[accent], result)
            else:
                raise UnsupportedEquation(f'暂不支持字符装饰 {accent}')
        return result

    def arguments(self, template: Template, count: int, trailing_characters: int = 0) -> tuple[str, ...]:
        slots = require_slots(template.items[:count])
        tail = template.items[count:]
        if (len(slots) != count or len(tail) != trailing_characters
                or any(not isinstance(item, Character) for item in tail)):
            raise UnsupportedEquation(f'MathType 模板 {template.selector} 的槽位数不符合预期')
        return tuple(self.sequence(slot.items) for slot in slots)

    def render(self, item: Record) -> str:
        if isinstance(item, Character):
            return self.character(item)
        if isinstance(item, Slot):
            return self.sequence(item.items)
        if isinstance(item, Matrix):
            alignment = (
                {0: 'l', 1: 'c', 2: 'r'} if self.app_matrix_encoding else {0: 'l', 1: 'l', 2: 'c', 3: 'r'}
            ).get(item.justify, 'c')
            body = r' \\ '.join(' & '.join(self.sequence(cell.items) for cell in row) for row in item.rows)
            return (r'\begin{matrix}' + body + r'\end{matrix}' if alignment == 'c'
                    else r'\begin{array}{' + alignment * item.columns + '}' + body + r'\end{array}')
        if isinstance(item, Pile):
            lines = [self.sequence(line.items) for line in item.lines]
            if len(lines) <= 1:
                return ''.join(lines)
            return r'\begin{gathered}' + r' \\ '.join(lines) + r'\end{gathered}'
        return self.template(item)

    def template(self, item: Template) -> str:
        selector, variant = item.selector, item.variation
        if selector < len(FENCES):
            if not 0 < variant <= 3:
                raise UnsupportedEquation('伸缩括号模板的左右定界符不明确')
            left, right = FENCES[selector]
            body, = self.arguments(item, 1, variant.bit_count())
            return join_latex((r'\left' + (left if variant & 1 else '.'), body,
                               r'\right' + (right if variant & 2 else '.')))
        if selector == 9:
            body, = self.arguments(item, 1, 2)
            left, right = (chr(character.code) for character in item.items[-2:])
            if left not in '()[]' or right not in '()[]':
                raise UnsupportedEquation('暂不支持此区间括号')
            return r'\left' + left + body + r'\right' + right
        if selector == 10 and variant in (0, 1):
            parts = self.arguments(item, len(item.items))
            if len(parts) == 1 and not variant:
                return command('sqrt', parts[0])
            if len(parts) == 2 and bool(parts[1]) == bool(variant):
                body, degree = parts
                return r'\sqrt' + ('[' + degree + ']' if degree else '') + '{' + body + '}'
        if selector == 11 and not variant & ~7:
            numerator, denominator = self.arguments(item, 2)
            return command('tfrac' if variant & 1 else 'frac', numerator, denominator)
        if selector in (12, 13) and variant in (0, 1):
            body, = self.arguments(item, 1)
            macro = 'underline' if selector == 12 else 'overline'
            return command(macro, command(macro, body) if variant else body)
        if selector == 14 and not variant & 3:
            top, bottom = self.arguments(item, 2, 1)
            arrow = item.items[-1].code
            if arrow in (0x2190, 0x2192):
                macro = r'\xleftarrow' if arrow == 0x2190 else r'\xrightarrow'
                return macro + ('[' + bottom + ']' if bottom else '') + '{' + top + '}'
            if arrow == 0x2194:
                relation = command('overset', top, r'\leftrightarrow') if top else r'\leftrightarrow'
                return command('underset', bottom, relation) if bottom else relation
        if selector == 15 or selector in BIG_OPERATORS or selector in (21, 22):
            body, lower, upper = self.arguments(item, 3, 1)
            if selector == 15:
                operator = INTEGRALS.get(variant & 0x0f)
            else:
                operator = BIG_OPERATORS.get(selector) or self.character(item.items[-1])
            if operator and not variant & ~0x017f:
                if variant & 0x40 and (lower or upper):
                    operator += r'\limits'
                return join_latex((operator + self.scripts(lower, upper), body))
        if selector == 23:
            base, lower, upper = self.arguments(item, 3)
            if base.startswith('\\') and base[1:] in FUNCTIONS:
                return base + self.scripts(lower, upper)
            if upper:
                base = command('overset', upper, base)
            return command('underset', lower, base) if lower else base
        if selector == 24 and variant in (0, 1):
            body, label = self.arguments(item, 2, 1)
            macro = 'overbrace' if variant else 'underbrace'
            return command(macro, body) + self.scripts(label if not variant else '', label if variant else '')
        if selector == 31 and variant in (1, 2, 3):
            body, = self.arguments(item, 1)
            return command({1: 'overleftarrow', 2: 'overrightarrow', 3: 'overleftrightarrow'}[variant], body)
        if selector in (32, 33) and not variant:
            body, = self.arguments(item, 1)
            return command('widetilde' if selector == 32 else 'widehat', body)
        if selector == 37 and variant == 0x1e:
            body, = self.arguments(item, 1)
            return command('boxed', body)
        raise UnsupportedEquation(f'暂不支持 MathType 模板 {selector}（类型 {variant}）')


def equation_to_latex(ole_data: bytes) -> dict:
    reader = MTEFReader(read_native(ole_data))
    display, items = reader.read_equation()
    writer = LaTeXWriter(reader)
    if all(isinstance(item, Slot) for item in items) and len(items) > 1:
        latex = writer.render(Pile(2, items))
    else:
        latex = writer.sequence(items)
    if not latex or len(latex) > 24000:
        raise UnsupportedEquation('公式为空或 LaTeX 过长')
    return {'latex': latex, 'display': display}


def decode_batch(objects: list[str]) -> list[dict]:
    if len(objects) > MAX_FORMULAS:
        raise ValueError(f'单次最多转换 {MAX_FORMULAS} 个公式。')
    results = []
    for encoded in objects:
        try:
            if len(encoded) > ((MAX_OLE_BYTES + 2) // 3) * 4:
                raise UnsupportedEquation('公式对象过大')
            raw = base64.b64decode(encoded, validate=True)
            results.append(equation_to_latex(raw))
        except UnsupportedEquation as error:
            results.append({'error': str(error)})
        except (OSError, IndexError, ValueError):
            results.append({'error': '该对象的 OLE 或 MTEF 数据无效'})
    return results
