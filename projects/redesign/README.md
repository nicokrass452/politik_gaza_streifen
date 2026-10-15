# Konfliktdossier Gazastreifen

Die Präsentation enthält 12 Folien mit deutschen Sprechernotizen und vier Folien pro Person. Die Namensplatzhalter lassen sich direkt in PowerPoint ersetzen.

## Grafiken und Quellen

Die Karten entstehen mit GeoPandas aus geografischen Grenzdaten und tatsächlichen Ortskoordinaten. Die Verarbeitung nutzt EPSG:32636 (UTM 36N). SVGs bleiben als Vektorgrafiken in der PowerPoint erhalten; PNGs dienen als Kompatibilitätsvorschau.

Die fünf Schadenswerte und vier Schadenskategorien stammen aus UNOSAT, Produkt 4165, Version 1, Bildstand 08.07.2025. Die Summe beträgt 192.812 betroffene Bauwerke. „78 %“ ist die gerundete Originalangabe und umfasst mögliche Schäden. Die Karte zeigt absolute Zahlen je Gouvernement, keine Schadensquote oder aktuelle Kontrollzonen.

Grenzen: ArcGIS GAZA_E, Layer 34, OCHA-P-Codes. Orte: OSM-Daten aus GAZA_E, Layer 32. Datenzugriff: 15.10.2026. Fachquellenzugriff: 14.10.2026. Die vollständigen Belege stehen in den Foliennotizen.

https://services3.arcgis.com/tuNLpt6Wfhd22qmO/ArcGIS/rest/services/GAZA_E/FeatureServer/34

https://services3.arcgis.com/tuNLpt6Wfhd22qmO/ArcGIS/rest/services/GAZA_E/FeatureServer/32

https://unosat.org/products/4165

Kleine Symbole: Lucide, ISC-Lizenz, Strichfarbe angepasst. Die Lizenz liegt bei den Symbolen. OSM-Orte: © OpenStreetMap-Mitwirkende.

## Neu erstellen

Die gespeicherten Geodaten und Grafiken sind enthalten. Zur Erstellung ist kein erneuter Download nötig.

```bash
pip install -r requirements.txt
python generate_graphics.py
python rasterize_svgs.py
python build_redesign.py
```

Die neue Datei erscheint unter `output/akuf_gazastreifen_neugestaltet.pptx`. Das Skript kontrolliert Folienzahl, Sprechernotizen, Bildproportionen und die Bildpositionen. Vor dem Vortrag die fertige Datei visuell prüfen.

Für die ursprüngliche Gestaltung wurden Nimbus Roman und Nimbus Sans verwendet. Falls sie fehlen, kann PowerPoint Ersatzschriften verwenden. Die Schriften sind nicht im Projekt enthalten.
