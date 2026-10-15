"""Wandelt die SVG-Karten in PNGs um (für PowerPoint-Versionen ohne SVG-Unterstützung)."""
from pathlib import Path
import cairosvg

assets = Path(__file__).resolve().parent / 'assets'
for svg in sorted(assets.glob('*.svg')):
    cairosvg.svg2png(url=str(svg), write_to=str(svg.with_suffix('.png')), scale=2.5)
    print('PNG erstellt:', svg.with_suffix('.png').name)
