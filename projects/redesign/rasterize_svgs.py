"""Create PNG compatibility previews for embedded SVGs."""
from pathlib import Path
import cairosvg
root=Path(__file__).resolve().parent/'assets'
for folder in [root,root/'icons']:
    for svg in folder.glob('*.svg'):
        cairosvg.svg2png(url=str(svg),write_to=str(svg.with_suffix('.png')),scale=2.5)
print('SVG compatibility previews created')
