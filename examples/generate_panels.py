"""Generate deterministic drawing fixtures; no experimental data."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

def generate(folder):
    folder.mkdir(parents=True, exist_ok=True)
    try:
        font = ImageFont.truetype("arial.ttf", 26)
    except OSError:
        font = ImageFont.load_default()
    for index, label in enumerate(("a", "b", "c", "d")):
        image = Image.new("RGB", (800, 650), "white")
        draw = ImageDraw.Draw(image)
        draw.rectangle((150, 100, 720, 510), outline="black", width=4)
        draw.text((65, 70), "(" + label + ")", fill="black", font=font)
        draw.text((390, 580), "X axis", fill="black", font=font)
        draw.text((35, 280), "Y axis", fill="black", font=font)
        points = [(x, 400 - (x - 180) // 3 + index * 10) for x in range(180, 681, 5)]
        draw.line(points, fill=(40, 90, 150), width=4)
        for x in range(150, 721, 114):
            draw.line((x, 510, x, 519), fill="black", width=3)
            draw.text((x - 5, 535), str((x-150)//114), fill="black", font=font)
        image.save(folder / ("panel-" + label + ".png"))

if __name__ == "__main__":
    generate(Path(__file__).resolve().parent / "panels")
