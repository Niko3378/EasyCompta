#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génère EasyCompta.ico — icône professionnelle multi-résolution.
Palette : bleu foncé #1F4E79, bleu vif #2E75B6, blanc, dorée #C9A84C.
"""
from PIL import Image, ImageDraw, ImageFont
import os, math

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "installer", "EasyCompta.ico")

C_FOND    = (31,  78, 121, 255)   # #1F4E79
C_ACCENT  = (46, 117, 182, 255)   # #2E75B6
C_OR      = (201, 168,  76, 255)  # #C9A84C
C_BLANC   = (255, 255, 255, 255)
C_BLANC50 = (255, 255, 255, 128)


def arrondi(draw, xy, radius, fill):
    """Dessine un rectangle arrondi (compatible Pillow < 8)."""
    x0, y0, x1, y1 = xy
    r = min(radius, (x1 - x0) // 2, (y1 - y0) // 2)
    if r <= 0 or x1 <= x0 or y1 <= y0:
        if x1 > x0 and y1 > y0:
            draw.rectangle([x0, y0, x1, y1], fill=fill)
        return
    draw.rectangle([x0+r, y0, x1-r, y1], fill=fill)
    draw.rectangle([x0, y0+r, x1, y1-r], fill=fill)
    draw.ellipse([x0, y0, x0+2*r, y0+2*r], fill=fill)
    draw.ellipse([x1-2*r, y0, x1, y0+2*r], fill=fill)
    draw.ellipse([x0, y1-2*r, x0+2*r, y1], fill=fill)
    draw.ellipse([x1-2*r, y1-2*r, x1, y1], fill=fill)


def make_icon(size: int) -> Image.Image:
    s = size
    scale = s / 64          # référence interne 64 px
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    r = max(2, int(10 * scale))

    # ── Fond arrondi bleu foncé ───────────────────────────────────────────────
    arrondi(d, (0, 0, s-1, s-1), r, C_FOND)

    # ── Bande dorée en bas (accent) ──────────────────────────────────────────
    bande_h = max(2, int(7 * scale))
    arrondi(d, (0, s - bande_h, s-1, s-1), r, C_OR)
    d.rectangle([0, s - bande_h, s-1, s - bande_h + r], fill=C_OR)

    # ── Lignes de "tableau" stylisées ────────────────────────────────────────
    if s >= 32:
        lw      = max(1, int(1.5 * scale))
        marg    = int(8 * scale)
        col_x   = int(marg + 14 * scale)
        top_y   = int(12 * scale)
        row_h   = int(8 * scale)
        n_rows  = 3

        # Ligne d'en-tête (bleu clair)
        d.rectangle([marg, top_y, s - marg, top_y + row_h - 1], fill=C_ACCENT)

        # Lignes de données (blanc transparent)
        for i in range(1, n_rows + 1):
            y = top_y + i * row_h
            if y + row_h - 1 < s - bande_h - 2:
                d.rectangle([marg, y, s - marg, y + row_h - 2],
                            fill=C_BLANC50)

        # Séparateur colonne
        d.rectangle([col_x, top_y, col_x + lw, top_y + n_rows * row_h],
                    fill=C_BLANC)

    # ── Lettre "E" centrée (petites tailles) ou "ECPT" (grandes) ───────────
    try:
        if s >= 128:
            font_size = int(22 * scale)
            try:
                font = ImageFont.truetype("arialbd.ttf", font_size)
            except Exception:
                font = ImageFont.truetype("arial.ttf", font_size)
            label = "EC"
        elif s >= 48:
            font_size = int(20 * scale)
            try:
                font = ImageFont.truetype("arialbd.ttf", font_size)
            except Exception:
                font = ImageFont.truetype("arial.ttf", font_size)
            label = "EC"
        else:
            font_size = max(8, int(18 * scale))
            try:
                font = ImageFont.truetype("arialbd.ttf", font_size)
            except Exception:
                font = ImageFont.load_default()
            label = "E"
    except Exception:
        font = ImageFont.load_default()
        label = "S"

    # Position texte centrée verticalement dans la zone bleue
    text_area_bottom = s - bande_h - 1
    try:
        bbox = d.textbbox((0, 0), label, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
    except Exception:
        tw, th = d.textsize(label, font=font)
    tx = (s - tw) // 2
    ty = (text_area_bottom - th) // 2

    # Ombre légère
    if s >= 32:
        d.text((tx+1, ty+1), label, font=font, fill=(0, 0, 0, 80))
    d.text((tx, ty), label, font=font, fill=C_BLANC)

    return img


def main():
    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = [make_icon(s) for s in sizes]

    # Pillow sauvegarde toutes les tailles dans un seul .ico
    images[0].save(OUT, format="ICO",
                   sizes=[(s, s) for s in sizes],
                   append_images=images[1:])
    print(f"Icone generee : {OUT}")
    for s, img in zip(sizes, images):
        print(f"  {s}x{s} px")


if __name__ == "__main__":
    main()
