"""Erzeugt die beiden Karten (Übersicht und Gebäudeschäden) als SVG aus den Geodaten.
Vorher: pip install -r requirements.txt. Die Berechnung läuft in UTM 36N (EPSG:32636).
"""
from pathlib import Path
from html import escape
import json
import geopandas as gpd

ROOT = Path(__file__).resolve().parent
A = ROOT / 'assets'

INK = '#262626'
MUTED = '#595959'
BORDER = '#BFBFBF'
LAND = '#E7E6E6'
BLUE = '#2E75B6'
FONT = 'Arial'

# UNOSAT, Produkt 4165, Version 1, Bildstand 08.07.2025
COUNTS = {'PS0255': 40192, 'PS0260': 48181, 'PS0265': 18103, 'PS0270': 51219, 'PS0275': 35117}
NAMES = {'PS0255': 'Nordgaza', 'PS0260': 'Gaza', 'PS0265': 'Deir al-Balah', 'PS0270': 'Khan Yunis', 'PS0275': 'Rafah'}
CATEGORIES = [('Zerstört', 102067), ('Schwer beschädigt', 17421), ('Mäßig beschädigt', 41895), ('Mögliche Schäden', 31429)]
assert sum(COUNTS.values()) == sum(n for _, n in CATEGORIES) == 192812

G = gpd.read_file(A / 'gaza_governorates.geojson').to_crs(32636)
PL = gpd.read_file(A / 'gaza_places.geojson').to_crs(32636)
assert G.geometry.is_valid.all() and set(G.ADM2_PCODE) == set(COUNTS)
GEOM = G.geometry.union_all()
assert 360 < GEOM.area / 1e6 < 370


def num(n):
    return f'{n:,}'.replace(',', '.')


def text(x, y, t, size=18, fill=INK, weight='normal', anchor='start'):
    return (f'<text x="{x:.2f}" y="{y:.2f}" fill="{fill}" font-family="{FONT}" font-size="{size}" '
            f'font-weight="{weight}" text-anchor="{anchor}">{escape(str(t))}</text>')


def line(x1, y1, x2, y2, stroke=BORDER, w=1):
    return f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" stroke="{stroke}" stroke-width="{w}"/>'


def circle(x, y, r, fill=BLUE, stroke='none', sw=0):
    return f'<circle cx="{x:.2f}" cy="{y:.2f}" r="{r:.2f}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'


def write(name, w, h, parts):
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}">'
           f'<rect width="{w}" height="{h}" fill="#FFFFFF"/>' + ''.join(parts) + '</svg>')
    (A / name).write_text(svg)


def projection(x, y, w, h):
    x0, y0, x1, y1 = G.total_bounds
    scale = min(w / (x1 - x0), h / (y1 - y0))
    ox = x + (w - (x1 - x0) * scale) / 2
    oy = y + (h - (y1 - y0) * scale) / 2

    def pt(px, py):
        return ox + (px - x0) * scale, oy + (y1 - py) * scale
    return pt, scale


def path(geom, pt):
    parts = []
    geoms = [geom] if geom.geom_type == 'Polygon' else geom.geoms
    for p in geoms:
        for ring in [p.exterior, *p.interiors]:
            coords = [pt(x, y) for x, y in ring.coords]
            parts.append('M ' + ' L '.join(f'{x:.2f},{y:.2f}' for x, y in coords) + ' Z')
    return ' '.join(parts)


def overview():
    pt, scale = projection(95, 25, 340, 405)
    parts = [f'<path d="{path(GEOM.simplify(15), pt)}" fill="{LAND}" stroke="{INK}" stroke-width="1.8"/>']
    for geom in G.geometry:
        parts.append(f'<path d="{path(geom.simplify(15), pt)}" fill="none" stroke="#FFFFFF" stroke-width="1.8"/>')
    for osm, name in [('504266619', 'Gaza-Stadt'), ('299944872', 'Khan Yunis'), ('322022888', 'Rafah')]:
        q = PL[PL.osm_id == osm].iloc[0].geometry
        x, y = pt(q.x, q.y)
        assert GEOM.buffer(300).contains(q)
        parts += [circle(x, y, 5, BLUE, '#FFFFFF', 1.5), text(x + 12, y + 7, name, 21, INK, 'bold')]
    parts += [text(38, 174, 'Mittelmeer', 20, MUTED, 'normal'), text(434, 280, 'Israel', 20, MUTED),
              text(56, 432, 'Ägypten', 20, MUTED)]
    # Nordpfeil und Maßstab in Kartenkoordinaten
    parts += [text(533, 39, 'N', 15, INK, 'bold', 'middle'), line(533, 49, 533, 90, INK, 1.5),
              f'<path d="M 533 47 L 528 60 L 538 60 Z" fill="{INK}"/>']
    x, y = 440, 425
    length = 10000 * scale
    parts += [line(x, y, x + length, y, INK, 2), line(x, y - 4, x, y + 4, INK, 1.4),
              line(x + length, y - 4, x + length, y + 4, INK, 1.4),
              text(x, y + 22, '0', 13, MUTED), text(x + length, y + 22, '10 km', 13, MUTED, anchor='end')]
    write('gaza_overview.svg', 600, 460, parts)


def governorates():
    pt, scale = projection(12, 5, 280, 350)
    parts = []
    for _, row in G.iterrows():
        value = COUNTS[row.ADM2_PCODE]
        opacity = .18 + .72 * value / max(COUNTS.values())
        parts.append(f'<path d="{path(row.geometry.simplify(15), pt)}" fill="{BLUE}" fill-opacity="{opacity:.3f}" '
                     f'stroke="#FFFFFF" stroke-width="2.1"/>')
    for i, code in enumerate(['PS0255', 'PS0260', 'PS0265', 'PS0270', 'PS0275']):
        row = G[G.ADM2_PCODE == code].iloc[0]
        q = row.geometry.representative_point()
        x, y = pt(q.x, q.y)
        ly = 36 + i * 72
        parts += [circle(x, y, 3, INK),
                  f'<path d="M {x:.2f} {y:.2f} L 298 {ly} L 315 {ly}" fill="none" stroke="{MUTED}" stroke-width=".9"/>',
                  text(326, ly, NAMES[code], 18, INK), text(326, ly + 29, num(COUNTS[code]), 26, INK, 'bold')]
    write('damage_governorates.svg', 500, 380, parts)


overview()
governorates()
(A / 'damage_data.json').write_text(json.dumps({
    'source': 'UNOSAT product 4165, 8 July 2025, version 1',
    'counts_by_pcode': COUNTS, 'categories': CATEGORIES,
    'rounded_percent_affected': 78, 'total': 192812}, ensure_ascii=False, indent=2))
print('Karten erzeugt: gaza_overview.svg, damage_governorates.svg')
