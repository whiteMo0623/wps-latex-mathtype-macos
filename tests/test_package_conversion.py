from base64 import b64decode
from io import BytesIO
import json
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
from threading import Thread
import unittest
from urllib.request import ProxyHandler, Request, build_opener
from zipfile import ZipFile

from docx import Document
from lxml import etree
from PIL import Image


SRC = Path(os.environ.get('MATHTYPE_SRC', Path(__file__).resolve().parents[1] / 'src'))
sys.path.insert(0, str(SRC))

from bridge import Bridge
from convert import DOCUMENT, NS, q
from fragments import FLAT_NS


class PackagedConversionTests(unittest.TestCase):
    def test_selection_to_ole_and_back_over_http(self):
        opener = build_opener(ProxyHandler({}))
        with Bridge(0) as server:
            worker = Thread(target=server.serve_forever)
            worker.start()
            try:
                self.check_round_trip(opener, server.origin)
            finally:
                server.shutdown()
                worker.join()

    def check_round_trip(self, opener, origin):
        with opener.open(origin + '/api/session', timeout=3) as response:
            token = json.load(response)['token']

        def post(route, data):
            request = Request(origin + '/api/' + route, json.dumps(data).encode('utf-8'), {
                'Content-Type': 'application/json',
                'X-MathType-Session': token,
                'Origin': origin,
            })
            with opener.open(request, timeout=10) as response:
                return json.load(response)

        scanned = post('scan', {'segments': [r'😀 $x+1$，\[\frac{a+b}{c}\]']})
        matches = scanned['matches']
        self.assertEqual([match['display'] for match in matches], [False, True])
        self.assertEqual(matches[0]['start'], 3)
        generated = post('render', {'items': matches})['results']
        objects = []
        for result in generated:
            package = etree.fromstring(result['xml'].encode('utf-8'))
            parts = {part.get(f'{{{FLAT_NS}}}name'): part for part in package}
            ole = package.find('.//' + q('o:OLEObject'))
            self.assertIsNotNone(ole)
            self.assertEqual(ole.get('ProgID'), 'Equation.DSMT4')
            for name, part in parts.items():
                binary = part.findtext(f'{{{FLAT_NS}}}binaryData')
                if name.endswith('.bin'):
                    objects.append(binary)
                elif name.endswith('.png'):
                    with Image.open(BytesIO(b64decode(binary))) as preview:
                        self.assertEqual(preview.format, 'PNG')
                        self.assertIsNotNone(preview.convert('RGBA').getchannel('A').getbbox())
        self.assertEqual(post('decode', {'objects': objects})['results'], [
            {'latex': 'x+1', 'display': False},
            {'latex': r'\frac{a+b}{c}', 'display': True},
        ])

    def test_file_conversion_writes_a_copy(self):
        with TemporaryDirectory() as folder:
            source = Path(folder) / 'Local example.docx'
            document = Document()
            paragraph = document.add_paragraph()
            paragraph.add_run('保留文字 ').bold = True
            paragraph.add_run('$x^2$')
            document.add_table(rows=1, cols=1).cell(0, 0).text = r'\[\frac{1}{2}\]'
            document.save(source)
            original = source.read_bytes()
            subprocess.run(
                [sys.executable, '-BEs', str(SRC / 'convert.py'), str(source)],
                check=True, capture_output=True, timeout=15,
            )
            self.assertEqual(source.read_bytes(), original)
            with ZipFile(source.with_name('Local example_MathType7.docx')) as modified:
                root = etree.fromstring(modified.read(DOCUMENT))
            self.assertIn('保留文字 ', root.xpath('.//w:r[w:rPr/w:b]/w:t/text()', namespaces=NS))
            self.assertEqual(root.xpath('.//o:OLEObject/@ProgID', namespaces=NS),
                             ['Equation.DSMT4', 'Equation.DSMT4'])
