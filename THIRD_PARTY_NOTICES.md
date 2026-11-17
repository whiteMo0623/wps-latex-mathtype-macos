# Third-party software

The project code is licensed under MIT. The release builds include the components identified below under their respective terms. Paths are relative to the app's `Contents/Resources` directory.

| Component | License and notices in the app | Upstream |
| --- | --- | --- |
| CPython 3.12.16 | CPython and bundled-library notices in `python/` | [CPython source](https://github.com/python/cpython/tree/v3.12.16), [Python Build Standalone 20261110 source and binaries](https://github.com/astral-sh/python-build-standalone/releases/tag/20261110) |
| Node.js 22.23.5 | MIT and bundled-library notices in `node/LICENSE` | [Node.js 22.23.5 source and binaries](https://nodejs.org/dist/v22.23.5/) |
| MathJax 3.2.2 | Apache-2.0 in `node_modules/mathjax-full/LICENSE` | [MathJax source](https://github.com/mathjax/MathJax-src/tree/3.2.2) |
| docx-equation 0.3.0 | MIT in its wheel metadata | [docx-equation](https://pypi.org/project/docx-equation/0.3.0/) |
| lxml 6.1.3 | BSD and bundled-library notices in its wheel metadata | [lxml](https://pypi.org/project/lxml/6.1.3/) |
| olefile 0.47 | BSD in its wheel metadata | [olefile](https://pypi.org/project/olefile/0.47/) |
| Pillow 13.0.0 | MIT-CMU and bundled-library notices in its wheel metadata | [Pillow](https://pypi.org/project/Pillow/13.0.0/) |
| python-docx 1.2.0 | MIT in its wheel metadata | [python-docx](https://pypi.org/project/python-docx/1.2.0/) |
| typing-extensions 4.16.0 | PSF-2.0 in its wheel metadata | [typing-extensions](https://pypi.org/project/typing-extensions/4.16.0/) |

Python wheel metadata and license files are under `python/lib/python3.12/site-packages/*.dist-info`. The exact upstream binaries and Python wheel hashes are pinned in `scripts/runtimes.json` and `requirements.lock`. The JavaScript dependency and integrity information is in `package-lock.json`.

WPS Office is a product of Kingsoft. MathType is a product of WIRIS. The app uses the local installation of WPS and generates interoperable equation data. It contains no WPS Office or MathType editor binaries.
