"""Loopback-only service for the WPS Writer add-in."""

from __future__ import annotations

import argparse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
import secrets
import sys
from threading import BoundedSemaphore

from fragments import render_batch, scan_segments
from mtef_latex import decode_batch

PORT = 49265
BUNDLE_ID = 'local.latex-to-mathtype.wps'
WEB_ROOT = Path(__file__).resolve().parent / 'web'
STATIC = {
    '/': ('index.html', 'text/html; charset=utf-8'),
    '/index.html': ('index.html', 'text/html; charset=utf-8'),
    '/ribbon.xml': ('ribbon.xml', 'application/xml; charset=utf-8'),
    '/main.js': ('main.js', 'text/javascript; charset=utf-8'),
}
MAX_REQUEST_BYTES = 16 * 1024 * 1024
RENDER_SLOT = BoundedSemaphore(1)


class Bridge(ThreadingHTTPServer):
    request_queue_size = 8

    def __init__(self, port: int):
        super().__init__(('127.0.0.1', port), Handler)
        self.host = f'127.0.0.1:{self.server_port}'
        self.origin = 'http://' + self.host
        self.token = secrets.token_urlsafe(32)


class Handler(BaseHTTPRequestHandler):
    server: Bridge

    def setup(self):
        self.request.settimeout(20)
        super().setup()

    def log_message(self, *_):
        pass

    def allowed_origin(self):
        return (self.headers.get('Host') == self.server.host
                and self.headers.get('Origin') in (None, self.server.origin)
                and self.headers.get('Sec-Fetch-Site') != 'cross-site')

    def respond(self, status: int, content_type: str, data: bytes):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Cross-Origin-Resource-Policy', 'same-origin')
        self.end_headers()
        self.wfile.write(data)

    def respond_json(self, status: int, payload: dict):
        self.respond(status, 'application/json; charset=utf-8', json.dumps(payload).encode('ascii'))

    def do_GET(self):
        if not self.allowed_origin():
            self.respond_json(403, {'error': '请求来源不受支持。'})
        elif self.path == '/api/status':
            self.respond_json(200, {'application': BUNDLE_ID})
        elif self.path == '/api/session':
            self.respond_json(200, {'token': self.server.token})
        elif self.path in STATIC:
            filename, content_type = STATIC[self.path]
            self.respond(200, content_type, (WEB_ROOT / filename).read_bytes())
        else:
            self.respond_json(404, {'error': '未找到资源。'})

    def do_POST(self):
        path = self.path
        token = self.headers.get('X-MathType-Session', '')
        if not self.allowed_origin() or not token.isascii() or not secrets.compare_digest(token, self.server.token):
            self.respond_json(403, {'error': '请从 WPS 加载项重新发起转换。'})
            return
        if path not in ('/api/scan', '/api/render', '/api/decode'):
            self.respond_json(404, {'error': '未找到资源。'})
            return
        if self.headers.get_content_type() != 'application/json':
            self.respond_json(415, {'error': '需要 JSON 请求。'})
            return
        try:
            length = int(self.headers.get('Content-Length', '0'))
        except ValueError:
            length = 0
        if not 0 < length <= MAX_REQUEST_BYTES:
            self.respond_json(413, {'error': '选区过大或请求为空，请缩小选区再试。'})
            return
        try:
            payload = json.loads(self.rfile.read(length))
            if not isinstance(payload, dict):
                raise ValueError('请求数据格式有误。')
            if path == '/api/scan':
                segments = payload.get('segments')
                if not isinstance(segments, list) or not all(isinstance(item, str) for item in segments):
                    raise ValueError('选区数据格式有误。')
                result = scan_segments(segments)
            elif path == '/api/decode':
                objects = payload.get('objects')
                if not isinstance(objects, list) or not objects or not all(isinstance(item, str) for item in objects):
                    raise ValueError('MathType 对象请求格式有误。')
                result = {'results': decode_batch(objects)}
            else:
                items = payload.get('items')
                if not isinstance(items, list) or not items:
                    raise ValueError('公式请求为空。')
                if not RENDER_SLOT.acquire(blocking=False):
                    self.respond_json(409, {'error': '正在转换另一批公式，请稍后再试。'})
                    return
                try:
                    result = {'results': render_batch(items)}
                finally:
                    RENDER_SLOT.release()
        except (UnicodeError, json.JSONDecodeError):
            self.respond_json(400, {'error': '请求数据格式有误。'})
        except ValueError as error:
            self.respond_json(400, {'error': str(error)})
        except Exception as error:
            print(f'本地公式生成失败 ({type(error).__name__})', file=sys.stderr, flush=True)
            self.respond_json(500, {'error': '本地公式生成失败。请重新启动加载项服务。'})
        else:
            self.respond_json(200, result)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=PORT)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    if args.check:
        result, = render_batch([{'latex': r'E=mc^2', 'display': False, 'size': 11}])
        if 'xml' not in result:
            raise RuntimeError(f"公式引擎自检失败。{result['error']}")
        return
    with Bridge(args.port) as server:
        print(f'WPS 公式加载项服务 {server.origin}', flush=True)
        server.serve_forever()


if __name__ == '__main__':
    main()
