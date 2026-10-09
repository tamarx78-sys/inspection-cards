# 開発用サーバー: 静的配信 + 書き出し結果の保存 (POST /dev/save?name=xxx.xlsx → dev/out/)
# 使い方: python serve.py [port] [--ext <フォルダ>] [--out <フォルダ>]
#   --ext を付けると、そのフォルダを /ext/ で読み取り専用に配信する
#   (社外秘の点検表をリポジトリのフォルダにコピーせずに試すため)
import http.server
import sys
import urllib.parse
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "dev" / "out"
args = sys.argv[1:]
EXT = None
if "--ext" in args:
    i = args.index("--ext")
    EXT = Path(args[i + 1]).resolve()
    del args[i:i + 2]
# --out: 書き出し結果の保存先 (社外秘のデータをリポジトリのフォルダに置かないため)
if "--out" in args:
    i = args.index("--out")
    OUT = Path(args[i + 1]).resolve()
    del args[i:i + 2]


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=str(ROOT), **kw)

    def translate_path(self, path):
        p = urllib.parse.unquote(urllib.parse.urlparse(path).path)
        if EXT and p.startswith("/ext/"):
            target = (EXT / p[len("/ext/"):]).resolve()
            return str(target) if target.is_relative_to(EXT) else str(EXT / "__forbidden__")
        return super().translate_path(path)

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
        OUT.mkdir(parents=True, exist_ok=True)
        (OUT / name).write_bytes(data)
        self.send_response(200)
        self.end_headers()
        self.wfile.write(str(OUT / name).encode())


port = int(args[0]) if args else 8765
http.server.ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
