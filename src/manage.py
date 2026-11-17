"""Install and remove the per-user WPS add-in and its local launch agent."""

from __future__ import annotations

import argparse
import fcntl
import json
import os
from pathlib import Path
import plistlib
import socket
import subprocess
import sys
import tempfile
import time
from urllib.request import ProxyHandler, build_opener
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent
LIBRARY = Path.home() / 'Library'
WPS_DATA = LIBRARY / 'Containers/com.kingsoft.wpsoffice.mac/Data'
MANIFEST = WPS_DATA / '.kingsoft/wps/jsaddons/publish.xml'
LABEL = 'local.latex-to-mathtype.wps'
ADDIN_NAME = 'LaTeXToMathTypeSelection'
PLIST = LIBRARY / 'LaunchAgents' / (LABEL + '.plist')
LOG = LIBRARY / 'Logs/LaTeXToMathType-WPS.log'
DOMAIN = f'gui/{os.getuid()}'
SERVICE = DOMAIN + '/' + LABEL
PORT = 49265
URL = f'http://127.0.0.1:{PORT}/'


def read_optional(path: Path) -> bytes | None:
    try:
        return path.read_bytes()
    except FileNotFoundError:
        return None


def write_atomic(path: Path, data: bytes):
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=path.parent, delete=False) as file:
        temporary = Path(file.name)
        try:
            file.write(data)
            file.close()
            temporary.replace(path)
        finally:
            temporary.unlink(missing_ok=True)


def check_manifest_write(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    # POSIX mode bits can allow a write that macOS app-data protection denies.
    with tempfile.NamedTemporaryFile(dir=path.parent, prefix='.mathtype-permission-'):
        pass


def updated_manifest(data: bytes | None, install: bool) -> bytes | None:
    options = dict(name=ADDIN_NAME, type='wps', url=URL, debug='', enable='enable_dev', install='null')
    if data is None:
        if not install:
            return None
        root = ET.Element('jsplugins')
    else:
        parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True))
        root = ET.fromstring(data, parser=parser)
        if root.tag != 'jsplugins':
            raise ValueError('WPS publish.xml 的根元素不是 jsplugins。')
    entries = [child for child in root if child.tag == 'jspluginonline' and child.get('name') == ADDIN_NAME]
    if (install and len(entries) == 1 and entries[0].attrib == options) or (not install and not entries):
        return None
    for entry in entries:
        root.remove(entry)
    if install:
        ET.SubElement(root, 'jspluginonline', options)
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def launchctl(*args, check=True):
    result = subprocess.run(('/bin/launchctl', *args), text=True, capture_output=True)
    if check and result.returncode:
        raise RuntimeError(f'launchctl {args[0]} 失败。{result.stderr.strip()}')
    return result.returncode == 0


def wait_for_bridge():
    opener = build_opener(ProxyHandler({}))
    cutoff = time.monotonic() + 6
    while time.monotonic() < cutoff:
        try:
            with opener.open(URL + 'api/status', timeout=0.5) as response:
                if json.load(response).get('application') == LABEL:
                    return
        except (OSError, ValueError):
            pass
        time.sleep(0.15)
    raise RuntimeError(f'本地服务启动失败。请查看 {LOG}')


def report_installation():
    print('安装完成。请保存工作并退出 WPS，再重新打开任意 .docx 文档。')
    print(f'首次加载时，WPS 会请求信任 {ADDIN_NAME}（{URL}）。')


def install():
    if not WPS_DATA.is_dir():
        raise RuntimeError('请先打开一次 WPS Office for Mac，再运行安装程序。')
    needs_manifest = updated_manifest(read_optional(MANIFEST), install=True) is not None
    if needs_manifest:
        check_manifest_write(MANIFEST)
    subprocess.run((sys.executable, '-BEs', str(ROOT / 'bridge.py'), '--check'), check=True)
    was_running = launchctl('print', SERVICE, check=False)
    previous_plist = read_optional(PLIST)
    if not was_running:
        with socket.socket() as probe:
            probe.settimeout(0.5)
            if probe.connect_ex(('127.0.0.1', PORT)) == 0:
                raise RuntimeError(f'端口 {PORT} 正被占用。请退出占用端口的程序后重试。')
    LOG.parent.mkdir(parents=True, exist_ok=True)
    config = {
        'Label': LABEL,
        'ProgramArguments': [sys.executable, '-BEs', str(ROOT / 'bridge.py')],
        'RunAtLoad': True,
        'StandardOutPath': str(LOG),
        'StandardErrorPath': str(LOG),
    }
    next_plist = plistlib.dumps(config)
    if was_running:
        launchctl('bootout', SERVICE)
    try:
        write_atomic(PLIST, next_plist)
        launchctl('bootstrap', DOMAIN, str(PLIST))
        wait_for_bridge()
        new_manifest = updated_manifest(read_optional(MANIFEST), install=True)
        if new_manifest is not None:
            write_atomic(MANIFEST, new_manifest)
    except (OSError, RuntimeError, ValueError, ET.ParseError):
        launchctl('bootout', SERVICE, check=False)
        if previous_plist is None:
            PLIST.unlink(missing_ok=True)
        else:
            write_atomic(PLIST, previous_plist)
            if was_running:
                launchctl('bootstrap', DOMAIN, str(PLIST), check=False)
        raise
    report_installation()


def uninstall():
    previous_manifest = read_optional(MANIFEST)
    new_manifest = updated_manifest(previous_manifest, install=False)
    if new_manifest is not None:
        write_atomic(MANIFEST, new_manifest)
    stopped = False
    try:
        if launchctl('print', SERVICE, check=False):
            launchctl('bootout', SERVICE)
            stopped = True
        PLIST.unlink(missing_ok=True)
    except (OSError, RuntimeError):
        if new_manifest is not None:
            write_atomic(MANIFEST, previous_manifest)
        if stopped and PLIST.is_file():
            launchctl('bootstrap', DOMAIN, str(PLIST), check=False)
        raise
    print('加载项已卸载。重启 WPS 后功能区按钮消失。')


def run_with_folder_access(action):
    try:
        action()
    except PermissionError as error:
        if not error.filename or not Path(error.filename).is_relative_to(WPS_DATA):
            raise
        from wps_folder_picker import open_wps_folder

        folder = MANIFEST.parent
        try:
            while not folder.is_dir() and folder != WPS_DATA:
                folder = folder.parent
        except PermissionError:
            folder = MANIFEST.parent
        print('macOS 限制了对此 WPS 目录的访问。请在弹出的窗口中选择已定位的文件夹。', flush=True)
        with open_wps_folder(folder):
            action()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('install', 'uninstall'))
    args = parser.parse_args()
    try:
        lock_path = Path(tempfile.gettempdir()) / (LABEL + '.lock')
        with lock_path.open('a') as lock:
            fcntl.flock(lock, fcntl.LOCK_EX)
            run_with_folder_access(install if args.action == 'install' else uninstall)
    except (OSError, RuntimeError, ValueError, ET.ParseError, subprocess.CalledProcessError) as error:
        parser.exit(1, f'{error}\n')


if __name__ == '__main__':
    main()
