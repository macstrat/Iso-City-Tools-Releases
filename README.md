# Sprite Pipeline App (Standalone)

This is a separate desktop tool for building model sprite packs from PNG exports.
It does not modify or run inside the Godot game.

## What this first pass does

- Load one or many source PNG files as a single pack.
- Calibrate each image independently (guides + fit mode).
- Drag guides directly on preview:
  - left base guide
  - center guide
  - right base guide
- Click-and-drag panning on preview and live FPS readout.
- High zoom range for pixel-level guide alignment.
- Drag-and-drop PNG files (or folders containing PNGs).
- Set full pack-level metadata:
  - id, name, set_id, category
  - tiles_x / tiles_y
  - variant_group / variant_label / group_label
  - manufacturer / link / instructions / notes
  - per-rotation offsets (0-3)
- Export one pack folder (or zip):
  - numbered sprites (`1`, `2`, `3`, ...)
  - `thumb`
  - one shared `metadata.json`
  - configurable image format: `webp` or `png`
  - WebP quality + lossless toggle
- Export metadata only (updates `metadata.json` without re-exporting images).

## Why it matches current runtime

The export keeps a bottom-center sprite anchor so it aligns with your existing
placement system (`Sprite2D.offset = (-width/2, -height)`).

## Install

From the repo root:

```powershell
py -m pip install -r tools/sprite_pipeline_app/requirements.txt
```

## Run

```powershell
py tools/sprite_pipeline_app/sprite_pipeline_app.py
```

## Notes

- Preset spans include:
  - `1x1` = `538`
  - `1x2` / `2x1` = `810`
  - `full_2x2` = `1080`
  - `2x3` / `3x2` = `1350`
- If `tkinterdnd2` is installed, OS drag/drop is enabled; otherwise use **Add PNGs**.
- Default export format is `webp` with quality `95` (lossy). This matches game support and keeps packs smaller.
- Padding is fixed to `32 px` in this build (matches your workflow).
