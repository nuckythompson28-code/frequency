# -*- coding: utf-8 -*-
"""
정적 파일 서빙 + graveyard.json 읽기/쓰기
모든 기기에서 http://NAS-IP:8585 으로 접속 가능
"""
from http.server import HTTPServer, SimpleHTTPRequestHandler
import json, os, sys

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8585
BASE = os.path.dirname(os.path.abspath(__file__))
GRAVE_FILE = os.path.join(BASE, 'graveyard.json')


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE, **kwargs)

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_GET(self):
        # graveyard.json 없으면 빈 객체 반환
        if self.path == '/graveyard.json' and not os.path.exists(GRAVE_FILE):
            self.send_response(200)
            self._cors()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(b'{}')
            return
        super().do_GET()

    def do_POST(self):
        if self.path == '/graveyard.json':
            length = int(self.headers.get('Content-Length', 0))
            if length > 1024 * 100:  # 100KB 제한
                self.send_response(413)
                self.end_headers()
                self.wfile.write(b'Payload too large')
                return
            body = self.rfile.read(length)
            try:
                data = json.loads(body)
                with open(GRAVE_FILE, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self.send_response(200)
                self._cors()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(b'{"ok":true}')
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode())
        else:
            self.send_response(404)
            self.end_headers()

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def log_message(self, fmt, *args):
        # 정적 파일 요청은 조용히, POST만 출력
        if args and str(args[0]).startswith('POST'):
            print(f'[저장] {args}')


if __name__ == '__main__':
    os.chdir(BASE)
    print(f'=== Soltri 선생산 서버 ===')
    print(f'  로컬:  http://localhost:{PORT}')
    print(f'  네트워크: http://<NAS-IP>:{PORT}')
    print(f'  무덤 파일: {GRAVE_FILE}')
    print(f'  Ctrl+C 로 종료\n')
    HTTPServer(('0.0.0.0', PORT), Handler).serve_forever()
