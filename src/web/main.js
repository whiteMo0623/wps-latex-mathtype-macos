'use strict';

const MAIN_STORY = 1;
const NORMAL_SELECTION = 2;
const INLINE_SHAPE_SELECTION = 7;
const EMBEDDED_OLE = 1;
const MATHTYPE_PROG_ID = 'Equation.DSMT4';
const MATHTYPE_CLASS = /^Equation\.DSMT\d+$/i;
const FLAT_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const OLE_NS = 'urn:schemas-microsoft-com:office:office';
const DOCUMENT_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
let conversionInProgress = false;

function getSelection(app, reverse = false) {
  const doc = app.ActiveDocument;
  if (!doc) {
    throw new Error('请打开一个 WPS 文字文档。');
  }
  const selection = app.Selection;
  const range = selection.Range;
  if (range.Start === range.End) {
    throw new Error(reverse ? '请选中 MathType 公式，或按 Command+A 全选。'
      : '请选中含 LaTeX 的文字。转换正文时可按 Command+A 全选。');
  }
  const type = selection.Type;
  if (selection.StoryType !== MAIN_STORY
      || (type !== NORMAL_SELECTION && !(reverse && type === INLINE_SHAPE_SELECTION))) {
    throw new Error('请选择正文或表格中的连续文字。');
  }
  if (doc.Path && !/\.docx$/i.test(doc.Name)) {
    throw new Error('请在 .docx 文档中使用此命令，以便保存 MathType 对象。');
  }
  assertEditable(doc);
  if (!reverse && typeof range.InsertXML !== 'function') {
    throw new Error('当前 WPS 版本缺少原位插入公式对象的接口。');
  }
  return {
    doc, fullName: doc.FullName, start: range.Start, end: range.End, text: range.Text, type,
  };
}

function assertEditable(doc) {
  if (doc.ReadOnly || doc.ProtectionType !== -1) {
    throw new Error('当前文档受到编辑保护。');
  }
  if (doc.TrackRevisions) {
    throw new Error('请关闭修订后再转换公式。');
  }
}

function assertSameSelection(app, snapshot) {
  const active = app.ActiveDocument;
  const selection = app.Selection;
  const range = selection.Range;
  // WPS creates a new JS wrapper on each lookup of the same document.
  if (!active || active.FullName !== snapshot.fullName || snapshot.doc.FullName !== snapshot.fullName
      || selection.StoryType !== MAIN_STORY || selection.Type !== snapshot.type
      || range.Start !== snapshot.start || range.End !== snapshot.end || range.Text !== snapshot.text) {
    throw new Error('转换期间文档或选区已变更，请重新选择后转换。');
  }
  assertEditable(snapshot.doc);
}

function positionedText(text, length) {
  if (text.length === length) return text;
  // WPS Mac emits a cell/row marker after CR without allocating a position for it.
  const withoutCellMarkers = text.replace(/\r\x07/g, '\r');
  return withoutCellMarkers.length === length ? withoutCellMarkers : null;
}

function textSpans({ doc, start, end }) {
  const selection = doc.Range(start, end);
  const shapes = selection.InlineShapes;
  const limits = [];
  for (let i = 1, count = shapes.Count; i <= count; i++) {
    const shape = shapes.Item(i).Range;
    limits.push([Math.max(start, shape.Start), Math.min(end, shape.End)]);
  }
  limits.sort((left, right) => left[0] - right[0]);
  const spans = [];
  let unmapped = 0;
  function addSpan(from, to) {
    if (from >= to) return;
    const range = doc.Range(from, to);
    const text = positionedText(range.Text, to - from);
    if (text !== null) {
      spans.push({ start: from, text });
      return;
    }
    // A field or another editor structure can change offsets in just one paragraph.
    const paragraphs = range.Paragraphs;
    for (let i = 1, count = paragraphs.Count; i <= count; i++) {
      const paragraph = paragraphs.Item(i).Range;
      const lower = Math.max(from, paragraph.Start);
      const upper = Math.min(to, paragraph.End);
      if (lower >= upper) continue;
      const part = positionedText(doc.Range(lower, upper).Text, upper - lower);
      if (part !== null) {
        spans.push({ start: lower, text: part });
      } else {
        unmapped++;
      }
    }
  }
  let cursor = start;
  for (const [left, right] of limits) {
    addSpan(cursor, left);
    cursor = Math.max(cursor, right);
  }
  addSpan(cursor, end);
  return { spans, unmapped };
}

async function requestJSON(path, body, token) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 180000);
  try {
    const response = await fetch(path, body === undefined ? { signal: controller.signal } : {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-MathType-Session': token },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || '本地公式服务返回错误。');
    return result;
  } catch (error) {
    if (error instanceof TypeError || error.name === 'AbortError') {
      throw new Error('无法连接本地公式服务或转换超时。请检查加载项服务后重试。');
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

function planReplacements(doc, spans, matches) {
  const edits = [];
  let skipped = 0;
  for (const match of matches) {
    const span = spans[match.segment];
    if (!span || !Number.isInteger(match.start) || !Number.isInteger(match.end)
        || match.start < 0 || match.start >= match.end || match.end > span.text.length) {
      throw new Error('公式位置校验失败，请重新选择后转换。');
    }
    const start = span.start + match.start;
    const end = span.start + match.end;
    const source = span.text.slice(match.start, match.end);
    const range = doc.Range(start, end);
    if (range.Text !== source || range.Fields.Count || range.Revisions.Count || range.InlineShapes.Count
        || range.OMaths.Count || range.Hyperlinks.Count || range.ContentControls.Count) {
      skipped++;
      continue;
    }
    try {
      rangeBookmarks(range);
    } catch (error) {
      skipped++;
      continue;
    }
    edits.push({ start, end, source, options: {
      latex: match.latex, display: match.display, size: doc.Range(start, start + 1).Font.Size,
    } });
  }
  return { edits, skipped };
}

function insertEquation(doc, edit, xml) {
  replaceAndVerify(doc, edit, range => range.Text === edit.source, range => range.InsertXML(xml), replacement => {
    const shapes = replacement.InlineShapes;
    if (replacement.Text !== '\x01' || shapes.Count !== 1) {
      throw new Error('WPS 未在原位置插入单个行内对象。');
    }
    const shape = shapes.Item(1);
    if (shape.Type !== EMBEDDED_OLE || shape.OLEFormat.ProgID !== MATHTYPE_PROG_ID) {
      throw new Error('WPS 未保留 MathType 对象类型。');
    }
  });
}

function restoreFailedReplacement(doc, unchanged, attempts = 1) {
  while (!unchanged() && attempts > 0) {
    attempts--;
    if (!doc.Undo(1)) break;
  }
  if (!unchanged()) {
    throw new Error('WPS 无法恢复本次替换。请检查当前公式及其所在段落。');
  }
}

async function convertSelection(app) {
  const snapshot = getSelection(app);
  const { spans, unmapped } = textSpans(snapshot);
  const { token } = await requestJSON('/api/session');
  const { matches, warningCount } = await requestJSON('/api/scan', {
    segments: spans.map(span => span.text),
  }, token);
  assertSameSelection(app, snapshot);
  const { edits, skipped } = planReplacements(snapshot.doc, spans, matches);
  const notes = [];
  if (unmapped) notes.push(unmapped + ' 个范围的字符位置无法校对，已跳过。');
  if (warningCount) notes.push(warningCount + ' 处定界符需要检查。');
  if (skipped) notes.push(skipped + ' 个公式位于带特殊结构的范围，已保留。');
  if (!edits.length) {
    return [matches.length ? '没有可安全替换的公式。' : '选区中没有找到完整的 LaTeX 公式。', ...notes].join('\n');
  }
  const { results } = await requestJSON('/api/render', { items: edits.map(edit => edit.options) }, token);
  assertSameSelection(app, snapshot);
  if (!Array.isArray(results) || results.length !== edits.length) {
    throw new Error('本地服务返回的公式数不一致，请重试。');
  }
  const failures = results.filter(result => typeof result.xml !== 'string');
  if (failures.length) {
    const reason = String(failures[0].error || '解析失败').slice(0, 180);
    notes.push(failures.length + ' 个公式生成失败，已保留 LaTeX。首个错误为 ' + reason + '。');
  }
  let converted = 0;
  // Earlier coordinates still refer to the original text after the later ranges change.
  for (let index = edits.length - 1; index >= 0; index--) {
    const xml = results[index].xml;
    if (typeof xml !== 'string') continue;
    try {
      insertEquation(snapshot.doc, edits[index], xml);
      converted++;
    } catch (error) {
      return ['已转换 ' + converted + ' 个 MathType 对象，本次转换已停止。', String(error.message || error), ...notes].join('\n');
    }
  }
  return ['已转换 ' + converted + ' 个 MathType 对象。', ...notes].join('\n');
}

function mathTypeShape(shape) {
  return shape.Type === EMBEDDED_OLE && MATHTYPE_CLASS.test(shape.OLEFormat.ProgID);
}

function equationPayload(range) {
  const xml = new DOMParser().parseFromString(range.WordOpenXML, 'application/xml');
  if (xml.documentElement.namespaceURI !== FLAT_NS || xml.documentElement.localName !== 'package') {
    throw new Error('WPS 未提供该对象的 Open XML 数据。');
  }
  const parts = new Map();
  for (const part of xml.documentElement.getElementsByTagNameNS(FLAT_NS, 'part')) {
    parts.set(part.getAttributeNS(FLAT_NS, 'name'), part);
  }
  const documentPart = parts.get('/word/document.xml');
  const objects = documentPart && documentPart.getElementsByTagNameNS(OLE_NS, 'OLEObject');
  if (!objects || objects.length !== 1 || objects[0].getAttribute('Type') !== 'Embed'
      || !MATHTYPE_CLASS.test(objects[0].getAttribute('ProgID'))) {
    throw new Error('选中的单个 MathType 对象无法从文档片段定位。');
  }
  const id = objects[0].getAttributeNS(DOCUMENT_REL_NS, 'id');
  const relPart = parts.get('/word/_rels/document.xml.rels');
  const relationship = relPart && Array.from(relPart.getElementsByTagNameNS(PACKAGE_REL_NS, 'Relationship'))
    .find(entry => entry.getAttribute('Id') === id);
  if (!id || !relationship || relationship.getAttribute('Type') !== DOCUMENT_REL_NS + '/oleObject'
      || relationship.getAttribute('TargetMode') === 'External') {
    throw new Error('找不到该 MathType 对象的嵌入数据。');
  }
  const base = new URL('http://document.local/word/document.xml');
  const target = new URL(relationship.getAttribute('Target'), base);
  const part = target.origin === base.origin && parts.get(target.pathname);
  const binary = part && part.getElementsByTagNameNS(FLAT_NS, 'binaryData')[0];
  const payload = binary && binary.textContent.replace(/\s/g, '');
  if (!payload) throw new Error('WPS 返回的 MathType 对象数据为空。');
  return payload;
}

function selectionObjects({ doc, start, end }) {
  const shapes = doc.Range(start, end).InlineShapes;
  const edits = [];
  const warnings = [];
  for (let index = 1, count = shapes.Count; index <= count; index++) {
    const shape = shapes.Item(index);
    if (!mathTypeShape(shape)) continue;
    const range = shape.Range;
    if (range.Start < start || range.End > end || range.Text !== '\x01') {
      warnings.push('请完整选中该公式对象。');
      continue;
    }
    if (range.Revisions.Count || range.Hyperlinks.Count || range.ContentControls.Count) {
      warnings.push('带有修订、链接或内容控件的公式已保留。');
      continue;
    }
    try {
      rangeBookmarks(range);
      edits.push({ start: range.Start, end: range.End, payload: equationPayload(range) });
    } catch (error) {
      warnings.push(String(error.message || error));
    }
  }
  edits.sort((left, right) => left.start - right.start);
  return { edits, warnings };
}

function rangeBookmarks(range) {
  const marks = range.Bookmarks;
  const result = [];
  for (let i = 1, count = marks.Count; i <= count; i++) {
    const mark = marks.Item(i);
    const bounds = mark.Range;
    const inside = position => position > range.Start && position < range.End;
    if (inside(bounds.Start) || inside(bounds.End)) {
      throw new Error('书签端点位于公式内部，已保留该公式。');
    }
    result.push({ name: mark.Name, start: bounds.Start, end: bounds.End });
  }
  return result;
}

function bookmarkAt(doc, { name, start, end }) {
  if (!doc.Bookmarks.Exists(name)) return false;
  const range = doc.Bookmarks.Item(name).Range;
  return range.Start === start && range.End === end;
}

function replaceAndVerify(doc, { start, end }, matches, apply, verify) {
  const original = doc.Range(start, end);
  if (!matches(original)) {
    throw new Error('公式内容已变更，请重新选择后转换。');
  }
  const bookmarks = rangeBookmarks(original);
  const previousEnd = doc.Content.End;
  let undoSteps = 1;
  try {
    apply(original);
    const replacementEnd = end + doc.Content.End - previousEnd;
    if (replacementEnd <= start) throw new Error('WPS 未返回有效的替换范围。');
    verify(doc.Range(start, replacementEnd));
    const delta = replacementEnd - end;
    const mapPosition = position => position < end ? position : position + delta;
    for (const mark of bookmarks) {
      const moved = { name: mark.name, start: mapPosition(mark.start), end: mapPosition(mark.end) };
      if (!bookmarkAt(doc, moved)) {
        doc.Bookmarks.Add(moved.name, doc.Range(moved.start, moved.end));
        undoSteps++;
        if (!bookmarkAt(doc, moved)) throw new Error('WPS 未能保留公式书签 ' + moved.name + '。');
      }
    }
  } catch (error) {
    restoreFailedReplacement(doc, () => doc.Content.End === previousEnd && matches(doc.Range(start, end))
      && bookmarks.every(mark => bookmarkAt(doc, mark)), undoSteps);
    throw error;
  }
}

function insertLatexText(doc, edit, { latex, display }) {
  const { start, end, payload } = edit;
  const matches = range => range.Text === '\x01' && range.InlineShapes.Count === 1
    && mathTypeShape(range.InlineShapes.Item(1)) && equationPayload(range) === payload;
  const left = start > 0 ? doc.Range(start - 1, start).Text : '';
  const right = end < doc.Content.End ? doc.Range(end, end + 1).Text : '';
  let source = '$' + latex + '$';
  if (display) {
    source = '\\[' + latex + '\\]';
  } else if (left.endsWith('$') || /^(?:\$|\d)/.test(right)) {
    // Avoid adjacent dollar delimiters and the currency heuristic in the forward scanner.
    source = '\\(' + latex + '\\)';
  }
  replaceAndVerify(doc, edit, matches, range => { range.Text = source; }, replacement => {
    if (replacement.End !== start + source.length || replacement.Text !== source || replacement.InlineShapes.Count) {
      throw new Error('WPS 未在原位置完整写入 LaTeX。');
    }
  });
}

async function decodeSelection(app) {
  const snapshot = getSelection(app, true);
  const { edits, warnings } = selectionObjects(snapshot);
  if (!edits.length) {
    return warnings.length ? '没有可转换的 MathType 公式。\n' + warnings[0]
      : '选区中没有找到 MathType 嵌入对象。';
  }
  const { token } = await requestJSON('/api/session');
  const { results } = await requestJSON('/api/decode', { objects: edits.map(edit => edit.payload) }, token);
  assertSameSelection(app, snapshot);
  if (!Array.isArray(results) || results.length !== edits.length) {
    throw new Error('本地服务返回的公式数不一致，请重试。');
  }
  let converted = 0;
  for (let index = edits.length - 1; index >= 0; index--) {
    const result = results[index];
    if (typeof result.latex !== 'string') {
      warnings.push(String(result.error || '解析失败'));
      continue;
    }
    try {
      insertLatexText(snapshot.doc, edits[index], result);
      converted++;
    } catch (error) {
      return '已将 ' + converted + ' 个 MathType 公式转换为 LaTeX，本次转换已停止。\n' + String(error.message || error);
    }
  }
  const output = '已将 ' + converted + ' 个 MathType 公式转换为 LaTeX。';
  if (!warnings.length) return output;
  return output + '\n保留 ' + warnings.length + ' 个公式。首个原因为 ' + warnings[0].slice(0, 180);
}

function runConversion(conversion) {
  if (conversionInProgress) return true;
  conversionInProgress = true;
  conversion(window.Application)
    .then(message => window.alert(message))
    .catch(error => window.alert(String(error.message || error)))
    .finally(() => { conversionInProgress = false; });
  return true;
}

function OnMathTypeConvert() {
  return runConversion(convertSelection);
}

function OnMathTypeDecode() {
  return runConversion(decodeSelection);
}
