'use strict';

const {execFileSync} = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const readline = require('node:readline');

const {mathjax} = require('mathjax-full/js/mathjax.js');
const {TeX} = require('mathjax-full/js/input/tex.js');
const {SVG} = require('mathjax-full/js/output/svg.js');
const {liteAdaptor} = require('mathjax-full/js/adaptors/liteAdaptor.js');
const {RegisterHTMLHandler} = require('mathjax-full/js/handlers/html.js');
const {SerializedMmlVisitor} = require('mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js');
const {STATE} = require('mathjax-full/js/core/MathItem.js');
require('mathjax-full/js/input/tex/ams/AmsConfiguration.js');
require('mathjax-full/js/input/tex/boldsymbol/BoldsymbolConfiguration.js');

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const input = new TeX({
  packages: ['base', 'ams', 'boldsymbol'],
  tags: 'none',
  maxBuffer: 6000,
  maxMacros: 1000,
  formatError(_jax, err) { throw err; },
});
const document = mathjax.document('', {InputJax: input, OutputJax: new SVG({fontCache: 'none'})});
const visitor = new SerializedMmlVisitor();

function render({latex, display, size}, workspace) {
  input.reset();
  const tree = document.convert(latex, {display, end: STATE.CONVERT});
  const mathml = visitor.visitTree(tree);
  input.reset();
  const rendered = document.convert(latex, {display});
  const svg = adaptor.tags(rendered, 'svg')[0];
  const [, top, width, height] = adaptor.getAttribute(svg, 'viewBox').trim().split(/\s+/).map(Number);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw new Error('公式没有可显示的内容');
  }
  // MathJax viewBox uses 1000 units per em. Keep the same scale in both axes.
  const scale = Math.min(size / 1000, 420 / width);
  const pixelsPerPoint = 5;
  const widthPx = Math.max(1, Math.ceil(width * scale * pixelsPerPoint));
  const heightPx = Math.max(1, Math.ceil(height * scale * pixelsPerPoint));
  if (heightPx > 3500) throw new Error('公式预览过高');
  adaptor.setAttribute(svg, 'width', String(widthPx));
  adaptor.setAttribute(svg, 'height', String(heightPx));
  adaptor.removeAttribute(svg, 'style');
  const source = path.join(workspace, 'preview.svg');
  const preview = path.join(workspace, 'preview.png');
  fs.writeFileSync(source, adaptor.outerHTML(svg), 'utf8');
  // SVG support comes from macOS ImageIO through sips.
  execFileSync('/usr/bin/sips', ['-s', 'format', 'png', source, '--out', preview], {
    stdio: 'ignore',
    timeout: 5000,
  });
  const png = fs.readFileSync(preview, 'base64');
  return {mathml, png, widthPx, heightPx, depth: Math.max(0, (top + height) * scale)};
}

(async () => {
  const workspace = fs.mkdtempSync(path.join(os.tmpdir(), 'wps-math-'));
  try {
    for await (const line of readline.createInterface({input: process.stdin, crlfDelay: Infinity})) {
      try {
        const result = render(JSON.parse(line), workspace);
        process.stdout.write(JSON.stringify(result) + '\n');
      } catch (err) {
        process.stdout.write(JSON.stringify({error: String(err.message || err)}) + '\n');
      }
    }
  } finally {
    fs.rmSync(workspace, {recursive: true, force: true});
  }
})();
