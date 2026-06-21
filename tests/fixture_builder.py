from __future__ import annotations

from pathlib import Path

from PIL import Image, PngImagePlugin
from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches


class DeckBuilder:
    def __init__(self) -> None:
        self.presentation = Presentation()
        self.presentation.slides.add_slide(self.presentation.slide_layouts[6])

    @property
    def slide(self):
        return self.presentation.slides[0]

    def add_text(self, name: str, text: str = "Hello", left: float = 1, top: float = 1) -> None:
        box = self.slide.shapes.add_textbox(Inches(left), Inches(top), Inches(2), Inches(0.5))
        box.name = name
        box.text = text

    def add_table(self, name: str, left: float = 1, top: float = 2) -> None:
        table_shape = self.slide.shapes.add_table(1, 1, Inches(left), Inches(top), Inches(2), Inches(0.5))
        table_shape.name = name
        table_shape.table.cell(0, 0).text = "value"

    def add_picture(self, name: str | None, image_path: Path, left: float = 1, top: float = 3) -> None:
        picture = self.slide.shapes.add_picture(str(image_path), Inches(left), Inches(top), Inches(1), Inches(1))
        if name is not None:
            picture.name = name

    def add_unknown(self, name: str, left: float = 4, top: float = 1) -> None:
        shape = self.slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(left), Inches(top), Inches(1), Inches(1))
        shape.name = name

    def add_slide(self) -> None:
        self.presentation.slides.add_slide(self.presentation.slide_layouts[6])

    def save(self, path: Path) -> None:
        self.presentation.save(path)


def make_png(
    path: Path,
    color: tuple[int, int, int] = (255, 0, 0),
    size: tuple[int, int] = (10, 10),
    changed_pixels: dict[tuple[int, int], tuple[int, int, int]] | None = None,
    metadata: dict[str, str] | None = None,
) -> Path:
    image = Image.new("RGB", size, color)
    for xy, pixel in (changed_pixels or {}).items():
        image.putpixel(xy, pixel)
    pnginfo = None
    if metadata:
        pnginfo = PngImagePlugin.PngInfo()
        for key, value in metadata.items():
            pnginfo.add_text(key, value)
    image.save(path, pnginfo=pnginfo)
    return path
