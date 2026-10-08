# 開発用サーバー: 静的配信 + 書き出し結果の保存 (POST /dev/save?name=xxx.xlsx → dev/out/)
import http.server
import sys
import urllib.parse
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "dev" / "out"


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=str(ROOT), **kw)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        if u.path != "/dev/save":
            self.send_error(404)
            return
        name = Path(urllib.parse.parse_qs(u.query).get("name", ["out.xlsx"])[0]).name
        data = self.rfile.read(int(self.headers["Content-Length"]))
        OUT.mkdir(exist_ok=True)
        (OUT / name).write_bytes(data)
        self.send_response(200)
        self.end_headers()
        self.wfile.write(str(OUT / name).encode())


port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
http.server.ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
