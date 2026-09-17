"""
Image Certificate Engine.
Processes raster certificate templates (.png, .jpg, .jpeg),
draws high-quality anti-aliased text, and exports directly to PDF.
"""

import os
from typing import Dict, Any
from PIL import Image, ImageDraw, ImageFont

def generate_pdf_from_image(
    template_path: str,
    output_pdf_path: str,
    student: Dict[str, Any]
):
    """
    Overlays student Name and Reg No on an image template and saves as PDF.
    """
    img = Image.open(template_path).convert("RGB")
    width, height = img.size
    draw = ImageDraw.Draw(img)

    name = student["name"].strip()
    reg_no = student["reg_no"].strip()

    # Dynamic font sizing based on image resolution
    base_font_size = int(height * 0.04) # ~4% of image height
    if len(name) > 30:
        base_font_size = int(base_font_size * 0.8)

    try:
        font_name = ImageFont.truetype("arialbd.ttf", base_font_size)
        font_reg = ImageFont.truetype("arial.ttf", int(base_font_size * 0.6))
    except Exception:
        font_name = ImageFont.load_default()
        font_reg = ImageFont.load_default()

    # Position in center vertically around 45% - 50%
    y_center_name = int(height * 0.44)
    y_center_reg = int(height * 0.50)

    # Center Name
    bbox_name = draw.textbbox((0, 0), name, font=font_name)
    w_name = bbox_name[2] - bbox_name[0]
    x_name = (width - w_name) // 2
    draw.text((x_name, y_center_name), name, font=font_name, fill=(20, 20, 20))

    # Center Reg No
    reg_text = f"(Reg No: {reg_no})"
    bbox_reg = draw.textbbox((0, 0), reg_text, font=font_reg)
    w_reg = bbox_reg[2] - bbox_reg[0]
    x_reg = (width - w_reg) // 2
    draw.text((x_reg, y_center_reg), reg_text, font=font_reg, fill=(70, 70, 70))

    img.save(output_pdf_path, "PDF", resolution=100.0)
