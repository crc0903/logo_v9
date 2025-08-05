import streamlit as st
from PIL import Image, ImageChops
import os
import io
import math
from pptx import Presentation
from pptx.util import Inches

PRELOADED_LOGO_DIR = "preloaded_logos"

def trim_whitespace(image):
    bg = Image.new(image.mode, image.size, (255, 255, 255, 0))
    diff = ImageChops.difference(image, bg)
    bbox = diff.getbbox()
    if bbox:
        return image.crop(bbox)
    return image

def create_logo_slide(prs, logos, canvas_width_in, canvas_height_in, logos_per_row):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    canvas_width_px = int(canvas_width_in * 96)
    canvas_height_px = int(canvas_height_in * 96)

    logo_count = len(logos)
    cols = logos_per_row if logos_per_row else max(1, round((logo_count / 1.5) ** 0.5 * (canvas_width_in / canvas_height_in) ** 0.3))
    rows = math.ceil(logo_count / cols)

    spacing_ratio = 0.92  # shrink cell size slightly to add spacing between logos
    cell_width = (canvas_width_px / cols) * spacing_ratio
    cell_height = (canvas_height_px / rows) * spacing_ratio

    left_margin = Inches((10 - canvas_width_in) / 2)
    top_margin = Inches((7.5 - canvas_height_in) / 2)

    for idx, logo in enumerate(logos):
        col = idx % cols
        row = idx // cols

        trimmed = trim_whitespace(logo)
        img_w, img_h = trimmed.size

        scale = min(cell_width / img_w, cell_height / img_h, 1.0)
        final_w = img_w * scale
        final_h = img_h * scale
