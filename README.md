# Youtube Tutorial
[![IMAGE ALT TEXT HERE](https://img.youtube.com/vi/XBx6i75ykos/0.jpg)](https://www.youtube.com/watch?v=XBx6i75ykos)
# Iso City Tools

Sprite pack creator tools for **MILS / Iso City**.

This repository is for **tools and releases** — not the full game source code.

## What is in this repo?

There are two different things here:

1. **Sprite Maker (Python script)**
   - The Python app is used to build sprite packs from PNG exports.
   - It helps you align/calibrate sprites and export pack metadata.
   - This is for creators/modders who want to make content.

2. **MILS Software Releases (ready to run)**
   - The **Releases** section contains the actual MILS application builds.
   - If you just want to use MILS, download from Releases.
   - No Python setup is required to run the released app.

---

## Quick Start

### I want to use MILS (normal users)
- Go to **Releases**
- Download the latest build for your platform
- Run it

### I want to create sprite packs (tool users)
- Install Python dependencies
- Watch the tutorial above
- Run the sprite pipeline tool
- Export packs for MILS

---

## Sprite Pipeline App (Python)

The sprite maker supports:

- Loading one or many PNGs into a pack
- Per-image calibration and guide placement
- Drag-and-drop PNG/folder import
- Metadata editing (`id`, `name`, `set_id`, category, variants, links, notes, etc.)
- Exporting pack folders (or zip) with:
  - numbered sprites (`1`, `2`, `3`, ...)
  - `thumb`
  - shared `metadata.json`
- Metadata-only export mode

### Install

py -m pip install -r tools/sprite_pipeline_app/requirements.txt



