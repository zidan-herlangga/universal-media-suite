import os

from PIL import Image, ImageDraw, ImageFont

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "assets")
FONT_DIR = os.environ.get("WINDIR", r"C:\Windows") + r"\Fonts"
os.makedirs(OUT, exist_ok=True)


def font(name, size):
    path = os.path.join(FONT_DIR, name)
    if os.path.exists(path):
        return ImageFont.truetype(path, size)
    return ImageFont.load_default()


URL = "https://zidan-herlangga.github.io/universal-media-suite"


def draw_gradient(draw, w, h, top=(13, 17, 23), bottom=(22, 27, 34)):
    for y in range(h):
        t = y / h
        r = int(top[0] + (bottom[0] - top[0]) * t)
        g = int(top[1] + (bottom[1] - top[1]) * t)
        b = int(top[2] + (bottom[2] - top[2]) * t)
        draw.line([(0, y), (w, y)], fill=(r, g, b))


def dot_grid(draw, w, h, color=(33, 38, 45), step=48):
    for x in range(60, w, step):
        for y in range(60, h, step):
            draw.ellipse([x - 2, y - 2, x + 2, y + 2], fill=color)


def glow(draw, cx, cy, r0, color=(88, 166, 255), steps=8):
    for i in range(steps, 0, -1):
        rr = r0 + (steps - i) * 6
        alpha = int(255 * (i / steps) * 0.25)
        overlay = Image.new("RGBA", (rr * 2, rr * 2), (0, 0, 0, 0))
        o = ImageDraw.Draw(overlay)
        o.ellipse([0, 0, rr * 2 - 1, rr * 2 - 1], fill=color + (alpha,))
        draw._image.paste(overlay, (cx - rr, cy - rr), overlay)


def make_og():
    w, h = 1200, 630
    img = Image.new("RGB", (w, h), "#0d1117")
    draw = ImageDraw.Draw(img)
    draw_gradient(draw, w, h)
    dot_grid(draw, w, h)

    glow(draw, 96, 96, 22)

    f_mono = font("consola.ttf", 21)
    f_badge = font("seguisb.ttf", 22)
    f_title = font("seguisb.ttf", 62)
    f_sub = font("segoeui.ttf", 27)
    f_url = font("consola.ttf", 20)
    f_os = font("seguisb.ttf", 21)

    accent = "#58a6ff"
    white = "#ffffff"
    muted = "#8b949e"
    faint = "#6e7681"

    draw.rounded_rectangle([80, 84, 90, 460], radius=5, fill=accent)

    draw.ellipse([92, 88, 104, 100], fill=accent)
    draw.text((124, 84), "ums-suite", font=f_mono, fill=white)
    draw.text((124, 116), URL, font=f_url, fill=muted)

    badge = "RILIS RESMI v2.0.0"
    bb = draw.textbbox((0, 0), badge, font=f_badge)
    bx, by, bx2, by2 = bb
    bw, bh = bx2 - bx, by2 - by
    pad = 12
    draw.rounded_rectangle([96, 168, 96 + bw + pad * 2, 168 + bh + pad * 2], radius=18, outline=accent, width=2)
    draw.text((96 + pad, 168 + pad - 2), badge, font=f_badge, fill=accent)

    draw.text((94, 240), "Universal Media", font=f_title, fill=white)
    draw.text((94, 322), "& Document Suite", font=f_title, fill=white)

    draw.text((96, 432), "Konversi foto RAW, video, audio, dokumen & downloader media — 100% lokal dan gratis.",
              font=f_sub, fill=muted)

    draw.line([(96, 556), (w - 96, 556)], fill=(33, 38, 45), width=2)
    draw.text((96, 570), "Windows   ·   Linux   ·   macOS", font=f_os, fill=faint)
    ow = draw.textlength(URL, font=f_url)
    draw.text((w - 96 - ow, 572), "100% offline", font=f_url, fill=faint)

    img.save(os.path.join(OUT, "og-image.png"))


def bolt_points(s):
    return [
        (0.63 * s, 0.13 * s),
        (0.30 * s, 0.53 * s),
        (0.49 * s, 0.53 * s),
        (0.40 * s, 0.87 * s),
        (0.70 * s, 0.45 * s),
        (0.51 * s, 0.45 * s),
    ]


def make_favicon_png(size, path):
    s = size
    im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    r = int(s * 0.22)
    d.rounded_rectangle([0, 0, s - 1, s - 1], radius=r, fill="#0d1117")
    lw = max(1, int(s * 0.05))
    d.rounded_rectangle([lw, lw, s - 1 - lw, s - 1 - lw], radius=r - lw, outline="#30363d", width=lw)
    d.polygon(bolt_points(s), fill="#58a6ff")
    im.save(path)


def make_ico():
    pngs = []
    for s in (16, 32, 48, 64):
        im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        r = int(s * 0.22)
        d.rounded_rectangle([0, 0, s - 1, s - 1], radius=r, fill="#0d1117")
        lw = max(1, int(s * 0.05))
        d.rounded_rectangle([lw, lw, s - 1 - lw, s - 1 - lw], radius=r - lw, outline="#30363d", width=lw)
        d.polygon(bolt_points(s), fill="#58a6ff")
        pngs.append(im)
    pngs[0].save(
        os.path.join(OUT, "favicon.ico"),
        format="ICO",
        sizes=[(s, s) for s in (16, 32, 48, 64)],
        append_images=pngs[1:],
    )


SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
  <rect width="32" height="32" rx="7" fill="#0d1117"/>
  <rect x="1" y="1" width="30" height="30" rx="6" fill="none" stroke="#30363d" stroke-width="2"/>
  <path d="M19.6 5.2 9.6 17.6h5.2l-2.4 9.2 10-12.4h-5.2l2.4-9.2Z" fill="#58a6ff"/>
</svg>"""


def write_svg():
    with open(os.path.join(OUT, "favicon.svg"), "w", encoding="utf-8") as f:
        f.write(SVG)


make_og()
make_ico()
make_favicon_png(180, os.path.join(OUT, "favicon-180.png"))
write_svg()
print("Generated:", sorted(os.listdir(OUT)))