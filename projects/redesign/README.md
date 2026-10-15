# Der Krieg im Gazastreifen – Präsentation

12 Folien, drei Sprecher, jeweils vier Folien. Die Sprechernotizen liegen in
`assets/speaker_notes.json` und landen beim Bauen in den PowerPoint-Notizen.
Die Namensplatzhalter `[Name 1]` bis `[Name 3]` vor dem Vortrag in PowerPoint ersetzen.

## Neu bauen

```bash
pip install -r requirements.txt
python generate_graphics.py   # Karten als SVG aus den GeoJSON-Dateien
python rasterize_svgs.py      # PNG-Versionen der Karten
python build_redesign.py      # PowerPoint nach output/
```

Wer nur Texte ändern will, braucht nur den letzten Schritt; die Karten liegen fertig in `assets/`.

Schrift ist Calibri. Das Balkendiagramm und die Tabelle sind normale PowerPoint-Objekte
und lassen sich direkt bearbeiten.

## Karten und Zahlen

- Grenzen der Gouvernements: ArcGIS-Dienst GAZA_E, Layer 34 (OCHA-P-Codes)
- Städte: OpenStreetMap-Punkte aus GAZA_E, Layer 32
- Gebäudeschäden: UNOSAT, Produkt 4165, Version 1, Bildstand 8. Juli 2025.
  Summe 192.812 betroffene Gebäude, davon 102.067 zerstört. Die 78 % sind die
  gerundete UNOSAT-Angabe. Die Karte zeigt absolute Zahlen je Gouvernement.
- Bevölkerung: OCHA, Juli 2024 (ca. 2,1 Mio. auf 365 km²)

Geodaten abgerufen am 15.10.2026, alle anderen Quellen am 14.10.2026. Die vollständigen
Links stehen in den Foliennotizen und auf der Quellenfolie.

https://services3.arcgis.com/tuNLpt6Wfhd22qmO/ArcGIS/rest/services/GAZA_E/FeatureServer/34
https://services3.arcgis.com/tuNLpt6Wfhd22qmO/ArcGIS/rest/services/GAZA_E/FeatureServer/32
https://unosat.org/products/4165
