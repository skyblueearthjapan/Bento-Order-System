"""Build src/icons.html: menu icons as base64 data URIs (trimmed, square, small)."""
import base64, io, sys
from PIL import Image

src_dir, out_path = sys.argv[1], sys.argv[2]
SIZE = 96
parts = []
for key in ['bento', 'okazu', 'don', 'men', 'toku']:
    im = Image.open(f'{src_dir}/{key}.png').convert('RGBA')
    bb = im.getchannel('A').point(lambda a: 255 if a > 20 else 0).getbbox()
    if bb:
        im = im.crop(bb)
    w, h = im.size
    side = int(max(w, h) * 1.04)
    sq = Image.new('RGBA', (side, side), (0, 0, 0, 0))
    sq.alpha_composite(im, ((side - w) // 2, (side - h) // 2))
    sq = sq.resize((SIZE, SIZE), Image.LANCZOS)
    buf = io.BytesIO()
    sq.save(buf, 'PNG', optimize=True)
    b64 = base64.b64encode(buf.getvalue()).decode('ascii')
    print(key, len(buf.getvalue()), 'bytes')
    parts.append(f"    {key}: 'data:image/png;base64,{b64}'")

html = ("<script>\n"
        "  // メニューのアイコン（assets/menu-icons/ の画像を 96px に縮めて埋め込んだもの。tools で再生成）\n"
        "  var MENU_ICONS = {\n" + ",\n".join(parts) + "\n  };\n"
        "</script>\n")
open(out_path, 'w', encoding='utf-8', newline='\n').write(html)
print('written', out_path, len(html))
