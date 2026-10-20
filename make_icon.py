"""app.ico 생성 (개발용, Pillow 필요). 실행: python make_icon.py"""
from PIL import Image, ImageDraw, ImageFont


def draw(size):
    s = size / 256
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    # 문서 모양
    d.rounded_rectangle([40 * s, 16 * s, 216 * s, 240 * s], radius=22 * s, fill=(31, 56, 100, 255))
    d.rounded_rectangle([58 * s, 34 * s, 198 * s, 222 * s], radius=12 * s, fill=(255, 255, 255, 255))
    try:
        font = ImageFont.truetype("malgunbd.ttf", int(64 * s))
    except OSError:
        font = ImageFont.load_default()
    d.text((128 * s, 92 * s), "MD", font=font, fill=(31, 56, 100, 255), anchor="mm")
    # 화살표
    d.polygon([(96 * s, 128 * s), (160 * s, 128 * s), (128 * s, 160 * s)], fill=(46, 117, 182, 255))
    # 줄
    for y in (178, 200):
        d.rounded_rectangle([80 * s, (y - 5) * s, 176 * s, (y + 5) * s], radius=5 * s,
                            fill=(46, 117, 182, 255) if y == 178 else (191, 191, 191, 255))
    return im


if __name__ == "__main__":
    big = draw(256)
    big.save("app.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    big.save("app_preview.png")
