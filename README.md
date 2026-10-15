# Der Krieg im Gazastreifen – Konfliktanalyse nach AKUF

Gruppenpräsentation für Politik, Klasse 10 (Projekt Konfliktforschung). Drei Sprecher,
12 Folien, deutsche Sprechernotizen mit Quellen.

## Dateien

- `presentations/akuf_gazastreifen_neugestaltet.pptx` – die aktuelle Fassung, diese halten wir.
- `presentations/akuf_gazastreifen_praesentation_original.pptx` – erste Fassung, nur zum Vergleich.
- `previews/` – alle Folien als PNG, falls man schnell reinschauen will.
- `research/` – Quellenliste (`SOURCES.md`, `sources.csv`), Rohdaten (UNOSAT-Zahlen,
  Bevölkerung, Resolutionen), die drei UN-Resolutionen als PDF und die Sprechernotizen
  beider Fassungen unter `notes/`.
- `projects/redesign/` – Skripte und Geodaten, mit denen Karten und PowerPoint erzeugt werden.
- `projects/original/` – Skript und OCHA/UNOSAT-Karten der ersten Fassung.

Vor dem Vortrag die Platzhalter `[Name 1]`, `[Name 2]`, `[Name 3]` in PowerPoint ersetzen.

## Präsentation neu erzeugen

```bash
cd projects/redesign
pip install -r requirements.txt
python generate_graphics.py
python rasterize_svgs.py
python build_redesign.py
```

Die Datei liegt danach unter `projects/redesign/output/`. Alle Geodaten sind im Ordner
enthalten, es muss nichts heruntergeladen werden. Für die alte Fassung analog
`projects/original/build_presentation.py`.

## Stand der Daten

- Quellenrecherche: 14. Oktober 2026, Geodaten: 15. Oktober 2026
- UNOSAT-Schadensauswertung: Bildstand 8. Juli 2025 (Version 1)
- Bevölkerungszahl: OCHA, Juli 2024

Was die Zahlen genau bedeuten (und was nicht), steht in `research/DATA_NOTES.md`.
