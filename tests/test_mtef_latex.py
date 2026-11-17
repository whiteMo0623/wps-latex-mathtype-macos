from base64 import b64encode
import os
from pathlib import Path
import sys
import unittest


SRC = Path(os.environ.get('MATHTYPE_SRC', Path(__file__).resolve().parents[1] / 'src'))
sys.path.insert(0, str(SRC))

from convert import Match, Renderer
from mtef_latex import decode_batch, equation_to_latex
from mtef_writer import compound_object


class MathTypeToLatexTests(unittest.TestCase):
    def test_round_trip_preserves_accents_and_matrix_cells(self):
        cases = [
            r'\hat{\theta}_{k}',
            r'\mathbf{X}\in\mathbb{R}^{n\times d}',
            r'A=\left[\begin{matrix}1 & -2 \\ 0 & 3\end{matrix}\right]',
        ]
        with Renderer() as renderer:
            native = [renderer.convert(Match(0, len(latex), latex, False), 11)[0] for latex in cases]
        formulas = [equation_to_latex(item) for item in native]
        self.assertEqual([item['latex'] for item in formulas], cases)
        self.assertFalse(any(item['display'] for item in formulas))

    def test_display_math_survives_the_wps_compound_header(self):
        mathml = '<math><mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mi>c</mi></mfrac></math>'
        wps_ole = bytearray(compound_object(mathml, display=True))
        wps_ole[8:24] = bytes.fromhex('03ce020000000000c000000000000046')
        self.assertEqual(equation_to_latex(bytes(wps_ole)), {
            'latex': r'\frac{a+b}{c}', 'display': True,
        })

    def test_invalid_object_does_not_block_the_rest_of_the_batch(self):
        good = b64encode(compound_object('<math><mi>x</mi></math>', display=False)).decode('ascii')
        bad = b64encode(b'not an OLE object').decode('ascii')
        results = decode_batch([bad, good])
        self.assertIn('error', results[0])
        self.assertEqual(results[1], {'latex': 'x', 'display': False})


if __name__ == '__main__':
    unittest.main()
