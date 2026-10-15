# Complete Gaza / AKUF presentation project

This archive contains the original and redesigned German school presentations, all task-authored project code, all saved research assets and geographic data, the SVG graphics, full speaker notes, final slide previews and validation records.

## Start here

- `presentations/akuf_gazastreifen_neugestaltet.pptx`: latest cream, charcoal and terracotta design; use this for the presentation.
- `presentations/akuf_gazastreifen_praesentation_original.pptx`: the earlier slate-and-blue version, retained for comparison.
- Both finished decks have 12 slides, full German notes and four slides per speaker. Replace `[Name 1]`, `[Name 2]` and `[Name 3]` in PowerPoint before presenting.

## Contents

| Folder | Contents |
|---|---|
| `projects/original/` | Original Python builder, requirements, original OCHA and UNOSAT PDFs and PNGs |
| `projects/redesign/` | GeoPandas/SVG generator, PowerPoint builder, portable SVG rasterizer, all SVGs, previews, geographic inputs, metadata, icons and icon license |
| `research/` | Full source registry in Markdown/CSV/JSON, full slide text and speaking notes for both versions, statistical CSV/JSON inputs and source-document download records |
| `previews/` | All 12 final slide previews for each finished version |
| `validation/` | Saved geometry, image-fit, slide-count and package validation reports |
| `runtime_only/` | Task-authored JavaScript used with the original presentation runtime |
| `layout_skill_draft/` | The original user-defined layout skill draft, including image fitting |

Raw map PDFs are stored with `projects/original/assets/`; geographic files and graphic metadata are under `projects/redesign/assets/`. `research/SOURCES.md` maps sources to their stored files or prepared notes. The archive is a complete saved project, not an offline mirror of every website visited. Some web sources are represented by URLs and notes rather than complete article copies.

## Regenerate the redesigned deck

From `projects/redesign/`:

```bash
pip install -r requirements.txt
python generate_graphics.py
python rasterize_svgs.py
python build_redesign.py
```

The generated PowerPoint appears in that project's `output/` folder. All inputs are already included, so no data download is needed for this workflow. To rebuild using the included, already-rendered SVG/PNG assets, only the PowerPoint builder is needed.

Nimbus Roman and Nimbus Sans were used for the redesigned deck. They are not bundled. PowerPoint may substitute fonts on systems where they are unavailable; inspect layout after substitution. CairoSVG may also require a platform-specific Cairo installation.

## Regenerate the earlier version

From `projects/original/`:

```bash
pip install -r requirements.txt
python build_presentation.py
```

The map PDFs and image files are included. The original builder's download paths are only used when required map assets are missing.

## Research dates and data scope

- Subject research: 14 October 2026.
- Geographic and design research: 15 October 2026.
- UNOSAT damage imagery: 8 July 2025, version 1.
- Population estimate: OCHA, July 2024.

The historical figures retain their original dates. `research/DATA_NOTES.md` explains units, rounding and interpretation. Raw governorate and city geometry, source metadata, numerical inputs, prepared research notes and citations are included. The full building-level UNOSAT damage point dataset was not used and is not part of this archive.

The small icons retain their Lucide license. The considered fluter map and alternative boundary files are included as research assets, although they were not used in the final redesigned slides. Bibliographic sources remain subject to their original terms.

The optional runtime-specific JavaScript requires the original authoring runtime and original path setup. Generic installed packages, package caches, credentials, authentication helpers, duplicate temporary exports and nested copies of previous ZIP bundles are omitted. Project dependencies are listed in the requirements files.

## Integrity

`FILE_INDEX.csv` lists every payload file's size and SHA-256 checksum. `SHA256SUMS.txt` provides a checksum list. `ARCHIVE_SUMMARY.json` records file counts and archive coverage.
