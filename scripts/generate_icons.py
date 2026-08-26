from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
ASSETS.mkdir(exist_ok=True)

size = 1024
image = Image.new("RGBA", (size, size), "#17324d")
draw = ImageDraw.Draw(image)

for index in range(4):
    left = 115 + index * 205
    top = 160
    right = left + 155
    bottom = 790
    draw.rounded_rectangle((left, top, right, bottom), radius=26, fill="#ffffff")
    draw.rounded_rectangle((left + 24, top + 42, right - 24, top + 64), radius=8, fill="#4b91c9")
    for row in range(5):
        y = top + 115 + row * 85
        draw.rounded_rectangle((left + 24, y, right - 24, y + 16), radius=7, fill="#a9bfd2")

chain_y = 855
draw.rounded_rectangle((270, chain_y, 520, chain_y + 70), radius=35, outline="#55c2a3", width=30)
draw.rounded_rectangle((500, chain_y, 750, chain_y + 70), radius=35, outline="#55c2a3", width=30)

png_path = ASSETS / "MultiDocSync.png"
ico_path = ASSETS / "MultiDocSync.ico"
image.save(png_path)
image.save(ico_path, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print(png_path)
print(ico_path)
