#!/usr/bin/env python3
"""
create_icon.py — Generates a simple placeholder icon.ico for Sales Analyzer.
Requires: Pillow (pip install Pillow)
"""

from PIL import Image, ImageDraw, ImageFont
import os

SIZES = [16, 32,48, 64, 128, 256]
BG_COLOR = (26, 26, 46)       # #1a1a2e
ACCENT    = (233, 69, 96)      # #e94560
TEXT_COL  = (234, 234, 234)    # #eaeaea


def make_frame(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # Rounded rect background
    r = size // 6
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=BG_COLOR)

    # Bar chart icon
    bar_count = 3
    pad = size * 0.15
    bar_w = (size - pad * 2) / (bar_count * 1.5 - 0.5)
    gap = bar_w * 0.5
    heights = [0.55, 0.75, 0.45]
    base_y = size - pad

    for i, (h_ratio) in enumerate(heights):
        bx = pad + i * (bar_w + gap)
        bh = (size - pad * 2) * h_ratio
        by = base_y - bh
        color = ACCENT if i == 1 else tuple(int(c * 0.7) for c in ACCENT)
        d.rounded_rectangle(
            [bx, by, bx + bar_w, base_y],
            radius=max(1, int(bar_w * 0.15)),
            fill=color,
        )

    return img


def main():
    frames = [make_frame(s) for s in SIZES]
    frames[0].save(
        "icon.ico",
        format="ICO",
        sizes=[(s, s) for s in SIZES],
        append_images=frames[1:],
    )
    print("✓ icon.ico creata")


if __name__ == "__main__":
    main()
