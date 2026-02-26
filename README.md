# Welcome to M.I.L.S
## Mac's Interlocking Layout System
![Screenshot of application window](https://i.imgur.com/dn3Ji0H.jpeg)

---

## Quick Start

### I want to use MILS (normal users)
- Go to **Releases**
- Download the latest build for your platform
- Run it

### I want to create sprite packs (tool users)
- Install Python dependencies
- Watch the tutorial tutorial
- Run the sprite pipeline tool
- Export packs for MILS by placing generated zip file in the /models folder

---

### Quick Tutorial (Main App)

### 1) Start and load models
- Open the app.
- Click **Options** (top-left) and choose **Rescan Models** if your packs are new or updated.
- If no model is selected, **left click** empty map space to open the building picker.

### 2) Place your first building
- In the picker, choose a model.
- Move your mouse over the grid to preview placement.
- **Left click** to place.
- If **Place multiple** is disabled, placement mode ends after one placement.
- Press **Esc** to cancel placement mode anytime.

### 3) Move around the map
- **Middle mouse drag** to pan.
- **Mouse wheel up/down** to zoom in/out.
- **[** and **]** rotate the **view** (map camera angle), not the model itself.

### 4) Rotate and edit buildings
- While placing a model:
  - **Q / E** rotates the placement.
- For already placed models:
  - **Right click** a model to open context menu:
    - Rotate
    - Delete
    - Replace

### 5) Open model info
- With no selected model, **left click** an existing building to open metadata/info.

### 6) Save, load, and export
- **Options → Export Map** to save your layout as JSON.
- **Options → Import Map** to load a saved layout.
- **Options → Export Image...** for PNG/JPG/WEBP output.

### 7) Useful options
- **Center View on Map**
- **Reset View Rotation**
- **More Options...** for Key Bindings, Zoom Settings, Model Master List, and extra toggles.




## Sprite Creation Tutorial
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

py -m pip install -r requirements.txt

## Future Plans

- **Update the UI**: ...pretty self-explanitory 
- **8x8 Grid Support**: Add support for an `8x8` base grid for smaller addons like sidewalks and other micro-details, while maintaining compatibility with the current `16x16` workflow.
- **Elevated Models**: Add multi-height placement support for elevated structures such as train lines, tunnels, bridges, and layered city builds.
- **Moving Vehicles and Characters**: Introduce animated movement for cars, trains, and minifigs to make cities feel more dynamic and alive.
- **Top-Down View Mode**: Add a dedicated top-down camera mode to improve planning, alignment, and large-layout editing.
- **Streetview-Style Mode**: Add an immersive street-level exploration mode inspired by Google Street View for navigating finished city scenes.
- **Integrated Pack Distribution System**: Add a system that new packs can be added in app from a centralized source instead of having to constantly download from the releases page







