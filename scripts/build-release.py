"""Build and verify a self-contained macOS release ZIP."""

import argparse
from hashlib import sha256
import json
import os
from pathlib import Path
import platform
import plistlib
import shutil
import subprocess
from tempfile import TemporaryDirectory
from urllib.parse import unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]
BUILD = ROOT / 'build'
CACHE = BUILD / 'cache'
APP_NAME = 'WPS LaTeX MathType'
EXECUTABLE = 'WPSLaTeXMathType'
MIN_MACOS = '13.0'


def run(*args, **options):
    subprocess.run(args, check=True, **options)


def digest(path):
    hash_value = sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            hash_value.update(block)
    return hash_value.hexdigest()


def download(spec):
    url, expected = spec['url'], spec['sha256']
    name = Path(unquote(urlsplit(url).path)).name
    file = CACHE / name
    if file.is_file() and digest(file) == expected:
        return file
    print(f'Downloading {name}', flush=True)
    with TemporaryDirectory(dir=CACHE) as folder:
        downloaded = Path(folder) / name
        run('/usr/bin/curl', '--fail', '--silent', '--show-error', '--location',
            '--connect-timeout', '20', '--retry', '2', '--max-time', '300',
            '--output', downloaded, url)
        if digest(downloaded) != expected:
            raise RuntimeError(f'SHA-256 mismatch: {name}')
        downloaded.replace(file)
    return file


def bundle_dependencies(workspace, resources, runtimes):
    run('/usr/bin/tar', '-xf', download(runtimes['python']), '-C', resources)
    python = resources / 'python/bin/python3'
    run(python, '-BI', '-m', 'pip', '--isolated', '--disable-pip-version-check', 'install',
        '--index-url', 'https://pypi.org/simple', '--only-binary', ':all:',
        '--require-hashes', '--no-compile', '-r', ROOT / 'requirements.lock')
    run(python, '-BI', '-m', 'pip', '--isolated', 'check')
    # pip stamps this unused console script with the absolute build directory.
    (python.parent / 'docx-equation').unlink(missing_ok=True)

    node_dist = workspace / 'node-dist'
    node_dist.mkdir()
    run('/usr/bin/tar', '-xf', download(runtimes['node']), '-C', node_dist, '--strip-components', '1')
    npm = node_dist / 'lib/node_modules/npm/bin/npm-cli.js'
    npm_workspace = workspace / 'npm'
    npm_workspace.mkdir()
    for name in ('package.json', 'package-lock.json'):
        shutil.copyfile(ROOT / name, npm_workspace / name)
    run(node_dist / 'bin/node', npm, 'ci', '--omit=dev', '--ignore-scripts',
        '--no-audit', '--no-fund', '--registry=https://registry.npmjs.org', cwd=npm_workspace)

    # render-equations.cjs uses MathJax's Node modules for base, AMS and
    # boldsymbol. The browser bundles and optional extensions stay in the build tree.
    installed_mathjax = npm_workspace / 'node_modules/mathjax-full'
    mathjax = resources / 'node_modules/mathjax-full'
    mathjax.mkdir(parents=True)
    shutil.copytree(installed_mathjax / 'js', mathjax / 'js')
    for name in ('LICENSE', 'package.json'):
        shutil.copyfile(installed_mathjax / name, mathjax / name)

    node = resources / 'node'
    (node / 'bin').mkdir(parents=True)
    shutil.copy2(node_dist / 'bin/node', node / 'bin/node')
    shutil.copyfile(node_dist / 'LICENSE', node / 'LICENSE')


def bundle_app(app, version, arch):
    contents = app / 'Contents'
    macos = contents / 'MacOS'
    macos.mkdir()
    info = {
        'CFBundleDevelopmentRegion': 'zh_CN',
        'CFBundleDisplayName': APP_NAME,
        'CFBundleExecutable': EXECUTABLE,
        'CFBundleIdentifier': 'org.wpslatexmathtype.macos',
        'CFBundleName': APP_NAME,
        'CFBundlePackageType': 'APPL',
        'CFBundleShortVersionString': version,
        'CFBundleVersion': version,
        'LSMinimumSystemVersion': MIN_MACOS,
        'NSHighResolutionCapable': True,
    }
    (contents / 'Info.plist').write_bytes(plistlib.dumps(info))
    run('/usr/bin/xcrun', '--sdk', 'macosx', 'swiftc', '-O',
        '-target', f'{arch}-apple-macos{MIN_MACOS}', ROOT / 'macos/Main.swift',
        '-o', macos / EXECUTABLE)
    run('/usr/bin/codesign', '--force', '--sign', '-', '--timestamp=none', app)


def verify_zip(workspace, archive, arch):
    portable = workspace / 'Extracted release'
    portable.mkdir()
    run('/usr/bin/ditto', '-x', '-k', archive, portable)
    app = portable / f'{APP_NAME}.app'
    run('/usr/bin/codesign', '--verify', '--deep', '--strict', app)
    for executable in (app / 'Contents/MacOS' / EXECUTABLE,
                       app / 'Contents/Resources/node/bin/node'):
        run('/usr/bin/lipo', '-verify_arch', arch, executable)
    resources = app / 'Contents/Resources'
    python = resources / 'python/bin/python3'
    isolated_home = workspace / 'Empty home'
    isolated_home.mkdir()
    environment = {
        'HOME': str(isolated_home),
        'PATH': '/usr/bin:/bin:/usr/sbin:/sbin',
        'MATHTYPE_SRC': str(resources / 'src'),
    }
    run(python, '-BEs', '-c',
        f'import platform; assert platform.machine() == {arch!r}, platform.machine()',
        env=environment)
    run(python, '-BEs', resources / 'src/bridge.py', '--check', env=environment)
    run(python, '-BEs', '-m', 'unittest', 'discover', '-s', ROOT / 'tests', '-v', env=environment)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--arch', choices=('arm64', 'x86_64'), default=platform.machine())
    args = parser.parse_args()
    if platform.system() != 'Darwin':
        parser.error('Release builds require macOS and Xcode Command Line Tools.')
    runtimes = json.loads((ROOT / 'scripts/runtimes.json').read_text())[args.arch]
    version = json.loads((ROOT / 'package.json').read_text())['version']
    CACHE.mkdir(parents=True, exist_ok=True)
    dist = ROOT / 'dist'
    dist.mkdir(exist_ok=True)
    filename = f'wps-latex-mathtype-macos-v{version}-{args.arch}.zip'

    with TemporaryDirectory(prefix=f'{args.arch}-', dir=BUILD) as directory:
        workspace = Path(directory)
        app = workspace / f'{APP_NAME}.app'
        resources = app / 'Contents/Resources'
        resources.mkdir(parents=True)
        bundle_dependencies(workspace, resources, runtimes)
        shutil.copytree(ROOT / 'src', resources / 'src',
                        ignore=shutil.ignore_patterns('__pycache__', '*.pyc', '.DS_Store'))
        for name in ('LICENSE', 'THIRD_PARTY_NOTICES.md', 'package-lock.json', 'requirements.lock'):
            shutil.copyfile(ROOT / name, resources / name)
        bundle_app(app, version, args.arch)
        archive = workspace / filename
        run('/usr/bin/ditto', '-c', '-k', '--norsrc', '--keepParent', app, archive)
        verify_zip(workspace, archive, args.arch)
        checksum = digest(archive)
        archive.replace(dist / filename)
        (dist / (filename + '.sha256')).write_text(f'{checksum}  {filename}\n', encoding='ascii')
    print(dist / filename)


if __name__ == '__main__':
    main()
