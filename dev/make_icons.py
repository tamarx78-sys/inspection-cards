# アプリアイコン生成 (青地に白いチェック付きクリップボード)
import sys
from pathlib import Path
from PIL import Image, ImageDraw

out = Path(sys.argv[1])

def icon(size):
    s = size / 512
    im = Image.new("RGB", (size, size), "#1f6feb")
    d = ImageDraw.Draw(im)
    d.rounded_rectangle([136 * s, 96 * s, 376 * s, 432 * s], radius=28 * s, fill="white")
    d.rounded_rectangle([196 * s, 72 * s, 316 * s, 124 * s], radius=16 * s, fill="#cfe0ff")
    for i, y in enumerate((190, 270, 350)):
        d.line([(172 * s, y * s), (204 * s, (y + 26) * s), (250 * s, (y - 22) * s)], fill="#1a7f37", width=max(2, int(18 * s)), joint="curve")
        d.rounded_rectangle([268 * s, (y - 2) * s, 344 * s, (y + 14) * s], radius=8 * s, fill="#9aa5b3")
    return im

icon(192).save(out / "icon-192.png")
icon(512).save(out / "icon-512.png")
icon(180).save(out / "apple-touch-icon.png")
print("ok")
