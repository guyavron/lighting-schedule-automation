"""
Simple product "renders" drawn in code, so the demo needs no real product photos
(manufacturer photos are copyrighted). One function per fixture type.
"""
from PIL import Image, ImageDraw, ImageFilter

W, H = 480, 400
FINISH = {"לבן": (236, 236, 232), "לבן מט": (228, 228, 224), "שחור": (40, 40, 42),
          "פליז מוברש": (176, 142, 86), "אפור גרפיט": (78, 80, 84), "אפור כהה": (92, 94, 98),
          "נירוסטה": (170, 174, 178), "אלומיניום": (196, 199, 202)}


def _shade(rgb, k):
    return tuple(max(0, min(255, int(c * k))) for c in rgb)


def _glow(img, box, strength=170):
    layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).ellipse(box, fill=(255, 214, 150, strength))
    layer = layer.filter(ImageFilter.GaussianBlur(18))
    img.alpha_composite(layer)


def _cyl(d, x0, y0, x1, y1, rgb, steps=24):
    """vertical cylinder with left-to-right shading"""
    w = x1 - x0
    for i in range(steps):
        k = 0.65 + 0.55 * (1 - abs((i + 0.5) / steps - 0.38) * 1.6)
        d.rectangle([x0 + w * i / steps, y0, x0 + w * (i + 1) / steps, y1], fill=_shade(rgb, k))


def draw(kind, finish, out_path):
    rgb = FINISH.get(finish, (200, 200, 200))
    img = Image.new("RGBA", (W, H), (255, 255, 255, 255))
    d = ImageDraw.Draw(img)
    cx = W // 2
    if kind in ("downlight", "trimless"):
        d.ellipse([cx - 150, 120, cx + 150, 220], fill=(225, 225, 222), outline=(200, 200, 198), width=2)
        rim = 0 if kind == "trimless" else 14
        d.ellipse([cx - 95 - rim, 140 - rim // 2, cx + 95 + rim, 200 + rim // 2], fill=_shade(rgb, 0.95))
        d.ellipse([cx - 80, 148, cx + 80, 192], fill=_shade(rgb, 0.55))
        d.ellipse([cx - 40, 160, cx + 40, 180], fill=(255, 236, 196))
        _glow(img, [cx - 60, 150, cx + 60, 190])
        d = ImageDraw.Draw(img)
        _cyl(d, cx - 70, 200, cx + 70, 300, _shade(rgb, 0.9))
        d.ellipse([cx - 70, 290, cx + 70, 312], fill=_shade(rgb, 0.7))
    elif kind in ("trackspot", "wallspot"):
        if kind == "trackspot":
            d.rectangle([40, 60, W - 40, 84], fill=_shade(rgb, 0.8))
            d.rectangle([cx - 14, 84, cx + 14, 130], fill=_shade(rgb, 0.9))
        else:
            d.rectangle([cx - 60, 50, cx + 60, 90], fill=_shade(rgb, 0.9))
            d.rectangle([cx - 10, 90, cx + 10, 130], fill=_shade(rgb, 0.8))
        body = [(cx - 50, 130), (cx + 50, 130), (cx + 95, 300), (cx - 5, 300)]
        d.polygon(body, fill=_shade(rgb, 0.95))
        d.polygon([(cx - 50, 130), (cx - 20, 130), (cx + 25, 300), (cx - 5, 300)], fill=_shade(rgb, 1.15))
        d.ellipse([cx - 8, 285, cx + 98, 318], fill=(255, 236, 196))
        _glow(img, [cx - 10, 280, cx + 100, 330])
    elif kind == "track":
        for i in range(3):
            y = 120 + i * 70
            d.rectangle([30, y, W - 30, y + 34], fill=_shade(rgb, 0.8 + i * 0.08))
            d.rectangle([30, y + 12, W - 30, y + 22], fill=_shade(rgb, 0.45))
    elif kind in ("profile", "floorprofile"):
        d.polygon([(40, 230), (W - 120, 120), (W - 40, 150), (120, 260)], fill=_shade(rgb, 0.9))
        d.polygon([(70, 236), (W - 110, 134), (W - 70, 148), (110, 250)], fill=(255, 240, 210))
        _glow(img, [80, 140, W - 80, 260], 120)
        d = ImageDraw.Draw(img)
        d.polygon([(40, 230), (120, 260), (120, 290), (40, 260)], fill=_shade(rgb, 0.6))
    elif kind == "pendant":
        d.line([cx, 0, cx, 140], fill=(90, 90, 90), width=3)
        d.pieslice([cx - 140, 120, cx + 140, 400], 180, 360, fill=_shade(rgb, 1.0))
        d.ellipse([cx - 140, 250, cx + 140, 290], fill=(255, 238, 200))
        _glow(img, [cx - 120, 250, cx + 120, 330])
    elif kind == "ring":
        for x in (cx - 120, cx + 120):
            d.line([x, 0, cx + (x - cx) * 0.9, 190], fill=(90, 90, 90), width=2)
        d.ellipse([cx - 180, 170, cx + 180, 270], outline=_shade(rgb, 1.0), width=22)
        d.ellipse([cx - 166, 190, cx + 166, 254], outline=(255, 232, 190), width=6)
    elif kind in ("updown", "sconce"):
        d.rectangle([cx - 110, 40, cx + 110, 360], fill=(242, 242, 240))
        if kind == "updown":
            _glow(img, [cx - 50, 20, cx + 50, 150]); _glow(img, [cx - 50, 250, cx + 50, 380])
            d = ImageDraw.Draw(img)
            _cyl(d, cx - 45, 140, cx + 45, 260, rgb)
        else:
            _glow(img, [cx - 80, 60, cx + 80, 340], 120)
            d = ImageDraw.Draw(img)
            d.ellipse([cx - 70, 110, cx + 70, 290], fill=_shade(rgb, 1.0), outline=_shade(rgb, 0.8), width=3)
    elif kind == "step":
        d.rectangle([60, 60, W - 60, 340], fill=(222, 220, 214))
        d.rounded_rectangle([cx - 90, 160, cx + 90, 240], 14, fill=_shade(rgb, 1.0))
        d.rectangle([cx - 70, 205, cx + 70, 225], fill=(255, 236, 196))
        _glow(img, [cx - 90, 210, cx + 90, 300], 130)
    elif kind == "inground":
        d.ellipse([cx - 170, 170, cx + 170, 290], fill=(150, 146, 138))
        d.ellipse([cx - 100, 195, cx + 100, 265], fill=_shade(rgb, 1.0))
        d.ellipse([cx - 70, 207, cx + 70, 253], fill=(255, 236, 196))
        _glow(img, [cx - 60, 120, cx + 60, 250], 150)
    elif kind == "floorlamp":
        d.line([cx - 80, 360, cx - 80, 120], fill=_shade(rgb, 1.0), width=8)
        d.arc([cx - 80, 40, cx + 160, 280], 180, 300, fill=_shade(rgb, 1.0), width=8)
        d.pieslice([cx + 60, 110, cx + 200, 230], 180, 360, fill=_shade(rgb, 1.0))
        d.ellipse([cx - 130, 350, cx - 30, 375], fill=_shade(rgb, 0.8))
        _glow(img, [cx + 70, 160, cx + 190, 260], 140)
    elif kind == "grazer":
        d.rectangle([40, 40, W - 40, 360], fill=(236, 234, 230))
        _glow(img, [70, 60, W - 70, 300], 150)
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([60, 290, W - 60, 322], 6, fill=_shade(rgb, 1.0))
        d.rectangle([70, 292, W - 70, 298], fill=(255, 236, 196))
    elif kind == "bollard":
        _cyl(d, cx - 45, 80, cx + 45, 360, rgb)
        d.rectangle([cx - 45, 110, cx + 45, 150], fill=(255, 236, 196))
        _glow(img, [cx - 90, 90, cx + 90, 180], 120)
        d = ImageDraw.Draw(img)
        d.ellipse([cx - 60, 350, cx + 60, 372], fill=(180, 180, 180))
    img = img.convert("RGB")
    # trim the empty margin (like a cut-out product photo), keep a small border
    from PIL import ImageChops
    bbox = ImageChops.difference(img, Image.new("RGB", img.size, (255, 255, 255))).point(
        lambda v: 255 if v > 12 else 0).getbbox()
    if bbox:
        pad = 14
        img = img.crop((max(0, bbox[0] - pad), max(0, bbox[1] - pad),
                        min(W, bbox[2] + pad), min(H, bbox[3] + pad)))
    img.save(out_path, "JPEG", quality=88)
    return out_path
