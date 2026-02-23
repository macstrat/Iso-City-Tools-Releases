import json
import math
import os
import re
import subprocess
import sys
import tempfile
import time
import threading
import zipfile
import importlib.util
from collections import deque
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def _ensure_runtime_dependencies() -> None:
    required = [
        ("Pillow", "PIL"),
        ("tkinterdnd2", "tkinterdnd2"),
    ]
    missing = [(pip_name, module_name) for pip_name, module_name in required if importlib.util.find_spec(module_name) is None]
    if not missing:
        return

    root = tk.Tk()
    root.withdraw()
    pip_names = [pip_name for pip_name, _module_name in missing]
    human = ", ".join(pip_names)
    should_install = messagebox.askyesno(
        "Missing Dependencies",
        (
            f"Missing required libraries: {human}\n\n"
            "Install now with pip?"
        ),
        parent=root,
    )
    if not should_install:
        root.destroy()
        raise SystemExit(f"Cannot launch without required dependencies: {human}")

    cmd = [sys.executable, "-m", "pip", "install"] + pip_names
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        details = (result.stderr or result.stdout or "").strip()
        if len(details) > 1800:
            details = details[-1800:]
        messagebox.showerror(
            "Install Failed",
            f"Failed to install dependencies ({human}).\n\n{details}",
            parent=root,
        )
        root.destroy()
        raise SystemExit("Dependency installation failed")

    messagebox.showinfo(
        "Dependencies Installed",
        f"Installed: {human}\n\nLaunching app.",
        parent=root,
    )
    root.destroy()
    importlib.invalidate_caches()


_ensure_runtime_dependencies()

from PIL import Image, ImageOps, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD

BaseTk = TkinterDnD.Tk


DEFAULT_CATEGORY = "Uncategorized"
DEFAULT_TARGET_SPAN_2X2 = 1080.0
FIXED_PADDING_PX = 32
PREVIEW_MIN_ZOOM = 0.08
PREVIEW_MAX_ZOOM_FALLBACK = 64.0
EXPORT_ALPHA_TRIM_THRESHOLD = 12
EXPORT_BOUNDS_ALPHA_THRESHOLD = 32
AUTO_SIDE_TRANSPARENCY_PX = 15
AUTO_EDGE_OPAQUE_RUN_PX = 10
EDGE_ALIGN_ALPHA_THRESHOLD = EXPORT_ALPHA_TRIM_THRESHOLD

# Keep in sync with the legacy metadata utility.
CATEGORY_OPTIONS = [
    "Food & Beverage",
    "Retail",
    "Hospitality",
    "Civic Services",
    "Infrastructure",
    "Recreation",
    "Entertainment",
    "Landmarks",
    "Streetscape",
    "Civic Memorial",
    "Automotive Services",
    "Residential",
    "Maintenance",
    "Uncategorized",
]

THEME_OPTIONS = [
    "Asian",
    "Cyberpunk",
    "European",
    "Modern",
    "Victorian",
    "Industrial",
    "Steampunk",
    "Medieval/Fantasy",
    "Sci-Fi",
    "Nordic/Scandinavian",
    "Mediterranean",
    "Colonial/Main Street",
    "Post-Apocalyptic",
]

_CATEGORY_ALIAS_MAP = {
    "food_beverage": "Food & Beverage",
    "food and beverage": "Food & Beverage",
    "food & beverage": "Food & Beverage",
    "retail": "Retail",
    "hospitality": "Hospitality",
    "civic_services": "Civic Services",
    "civic services": "Civic Services",
    "infrastructure": "Infrastructure",
    "recreation": "Recreation",
    "entertainment": "Entertainment",
    "landmarks": "Landmarks",
    "streetscape": "Streetscape",
    "civic_memorial": "Civic Memorial",
    "civic memorial": "Civic Memorial",
    "automotive_services": "Automotive Services",
    "automotive services": "Automotive Services",
    "residential": "Residential",
    "maintenance": "Maintenance",
    "uncategorized": "Uncategorized",
    "other": "Uncategorized",
}

FIT_MODE_PRESETS = {
    "1x1": 538.0,
    "1x2": 810.0,
    "2x1": 810.0,
    "full_2x2": 1080.0,
    "2x3": 1350.0,
    "3x2": 1350.0,
}


def _normalize_id(value: str) -> str:
    value = value.strip().lower()
    value = re.sub(r"[^a-z0-9\s_-]+", "", value)
    value = re.sub(r"\s+", "_", value)
    value = re.sub(r"_+", "_", value)
    value = re.sub(r"^-+|_+$", "", value)
    return value or "model"


def _safe_int(value: str, default: int) -> int:
    try:
        return int(value)
    except Exception:
        return default


def _safe_float(value: str, default: float) -> float:
    try:
        return float(value)
    except Exception:
        return default


def _round_half_up(value: float) -> int:
    # Deterministic rounding (avoids Python's banker's rounding at .5).
    if value >= 0:
        return int(math.floor(value + 0.5))
    return int(math.ceil(value - 0.5))


def _normalize_category(value: str) -> str:
    raw = value.strip()
    if raw == "":
        return DEFAULT_CATEGORY
    key = raw.lower().replace("-", " ").replace("_", " ")
    key = re.sub(r"\s+", " ", key).strip()
    if key in _CATEGORY_ALIAS_MAP:
        return _CATEGORY_ALIAS_MAP[key]
    for category in CATEGORY_OPTIONS:
        if key == category.lower():
            return category
    return raw


@dataclass
class SpriteImageItem:
    source_path: str
    guide_left: float
    guide_center: float
    guide_right: float
    baseline_y: float
    fit_mode: str = "full_2x2"
    target_span_px: float = DEFAULT_TARGET_SPAN_2X2
    offset_x: float = 0.0
    offset_y: float = 0.0
    width: int = 0
    height: int = 0
    _source_rgba: Optional[Image.Image] = None

    @classmethod
    def from_path(cls, source_path: str) -> "SpriteImageItem":
        with Image.open(source_path) as img:
            w, h = img.size
        center = w * 0.5
        span = min(DEFAULT_TARGET_SPAN_2X2, float(w))
        left = max(0.0, center - span * 0.5)
        right = min(float(w), center + span * 0.5)
        return cls(
            source_path=source_path,
            guide_left=left,
            guide_center=center,
            guide_right=right,
            baseline_y=float(h),
            width=w,
            height=h,
        )

    def source_rgba(self) -> Image.Image:
        if self._source_rgba is None:
            with Image.open(self.source_path) as img:
                self._source_rgba = img.convert("RGBA")
            self.width, self.height = self._source_rgba.size
        return self._source_rgba

    def measured_span(self) -> float:
        return max(1.0, self.guide_right - self.guide_left)

    def effective_target_span(self) -> float:
        preset = FIT_MODE_PRESETS.get(self.fit_mode, None)
        if preset is not None:
            return preset
        return max(1.0, self.target_span_px)

    def scale_factor(self) -> float:
        return self.effective_target_span() / self.measured_span()

    def label(self) -> str:
        return Path(self.source_path).name


@dataclass
class PackMetadata:
    id: str = "new_model"
    name: str = "New Model"
    set_id: str = ""
    category: str = DEFAULT_CATEGORY
    theme: str = ""
    tiles_x: int = 2
    tiles_y: int = 2
    variant_group: str = ""
    variant_label: str = ""
    group_label: str = ""
    manufacturer: str = ""
    link: str = ""
    instructions: str = ""
    notes: str = ""
    offsets: dict[str, list[float]] = field(
        default_factory=lambda: {
            "0": [0.0, 0.0],
            "1": [0.0, 0.0],
            "2": [0.0, 0.0],
            "3": [0.0, 0.0],
        }
    )

    def to_dict(self) -> dict:
        return {
            "id": _normalize_id(self.id),
            "name": self.name.strip() or "New Model",
            "set_id": self.set_id.strip(),
            "category": _normalize_category(self.category),
            "theme": self.theme.strip(),
            "tiles_x": max(1, int(self.tiles_x)),
            "tiles_y": max(1, int(self.tiles_y)),
            "variant_group": self.variant_group.strip(),
            "variant_label": self.variant_label.strip(),
            "group_label": self.group_label.strip(),
            "manufacturer": self.manufacturer.strip(),
            "link": self.link.strip(),
            "instructions": self.instructions.strip(),
            "notes": self.notes.strip(),
            "offsets": self.offsets,
        }


class SpritePipelineApp(BaseTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Sprite Pipeline (Standalone)")
        self.geometry("1560x1050")
        self.minsize(960, 680)

        self.items: list[SpriteImageItem] = []
        self.active_idx: Optional[int] = None
        self.pack_meta = PackMetadata()

        self.preview_zoom: float = 0.55
        self.pan_x: float = 0.0
        self.pan_y: float = 0.0
        self.drag_mode: Optional[str] = None
        self.preview_photo: Optional[ImageTk.PhotoImage] = None
        self._render_pending = False
        self._fps_times: deque[float] = deque(maxlen=90)
        self._fps_text = tk.StringVar(value="FPS: --")
        self._preview_lock = threading.Lock()
        self._preview_event = threading.Event()
        self._preview_shutdown = False
        self._preview_job_id = 0
        self._preview_next_job: Optional[dict] = None
        self._preview_latest_result: Optional[dict] = None
        self._preview_applied_job_id = 0
        self._scene_ox: float = 0.0
        self._scene_oy: float = 0.0
        self._scene_disp_w: int = 0
        self._scene_disp_h: int = 0
        self._canvas_image_id: Optional[int] = None
        self._canvas_rect_id: Optional[int] = None
        self._canvas_left_id: Optional[int] = None
        self._canvas_center_id: Optional[int] = None
        self._canvas_right_id: Optional[int] = None
        self._zoom_render_after_id: Optional[str] = None
        self._pan_render_after_id: Optional[str] = None
        self._last_canvas_mouse: tuple[float, float] = (0.0, 0.0)
        self._suppress_image_apply = False

        self.status_var = tk.StringVar(value="Ready.")
        self.export_format_var = tk.StringVar(value="webp")
        self.webp_quality_var = tk.StringVar(value="95")
        self.webp_lossless_var = tk.BooleanVar(value=False)
        self.zip_name_var = tk.StringVar(value="sprite_pack.zip")
        self._last_auto_zip_name = self.zip_name_var.get().strip() or "sprite_pack.zip"

        self.image_vars: dict[str, tk.StringVar] = {
            "fit_mode": tk.StringVar(value="full_2x2"),
            "target_span": tk.StringVar(value=f"{DEFAULT_TARGET_SPAN_2X2:.2f}"),
            "guide_left": tk.StringVar(value="0"),
            "guide_center": tk.StringVar(value="0"),
            "guide_right": tk.StringVar(value="0"),
            "baseline_y": tk.StringVar(value="0"),
            "offset_x": tk.StringVar(value="0"),
            "offset_y": tk.StringVar(value="0"),
        }
        self.meta_vars: dict[str, tk.StringVar] = {
            "id": tk.StringVar(value=self.pack_meta.id),
            "name": tk.StringVar(value=self.pack_meta.name),
            "set_id": tk.StringVar(value=self.pack_meta.set_id),
            "category": tk.StringVar(value=self.pack_meta.category),
            "theme": tk.StringVar(value=self.pack_meta.theme),
            "tiles_x": tk.StringVar(value=str(self.pack_meta.tiles_x)),
            "tiles_y": tk.StringVar(value=str(self.pack_meta.tiles_y)),
            "variant_group": tk.StringVar(value=self.pack_meta.variant_group),
            "variant_label": tk.StringVar(value=self.pack_meta.variant_label),
            "group_label": tk.StringVar(value=self.pack_meta.group_label),
            "manufacturer": tk.StringVar(value=self.pack_meta.manufacturer),
            "link": tk.StringVar(value=self.pack_meta.link),
            "instructions": tk.StringVar(value=self.pack_meta.instructions),
            "notes": tk.StringVar(value=self.pack_meta.notes),
        }
        self.offset_vars: dict[str, tk.StringVar] = {}
        for idx in range(4):
            self.offset_vars[f"offset_{idx}_x"] = tk.StringVar(value="0")
            self.offset_vars[f"offset_{idx}_y"] = tk.StringVar(value="0")

        self.bulk_root_var = tk.StringVar(value="")
        self.bulk_status_var = tk.StringVar(value="Choose a folder and scan for metadata.json files.")
        self.bulk_apply_status_var = tk.StringVar(value="")
        self.bulk_entries: list[dict] = []
        self.bulk_entry_by_iid: dict[str, dict] = {}
        self.bulk_field_vars: dict[str, tk.StringVar] = {}
        self.bulk_apply_vars: dict[str, tk.BooleanVar] = {}
        self.bulk_mode_var = tk.StringVar(value="single")
        self.bulk_sort_column = "id"
        self.bulk_sort_reverse = False
        self.bulk_tree_column_labels: dict[str, str] = {}
        self.bulk_single_vars: dict[str, tk.StringVar] = {}
        self.bulk_single_inputs: dict[str, tk.Widget] = {}
        self.bulk_single_status_var = tk.StringVar(value="Select one row to edit a single metadata file.")

        self._build_ui()
        self._sync_zip_name_to_id(force=True)
        self._setup_dnd()
        self._refresh_image_list()
        self._preview_thread = threading.Thread(target=self._preview_worker_loop, daemon=True)
        self._preview_thread.start()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(16, self._poll_preview_results)
        self._request_render()

    def _build_ui(self) -> None:
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True, padx=10, pady=10)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        notebook = ttk.Notebook(root)
        notebook.grid(row=0, column=0, sticky="nsew")
        sprite_editor_tab = ttk.Frame(notebook)
        bulk_metadata_tab = ttk.Frame(notebook)
        notebook.add(sprite_editor_tab, text="Sprite Editor")
        notebook.add(bulk_metadata_tab, text="Bulk Metadata Editor")

        self._build_sprite_editor_tab(sprite_editor_tab)
        self._build_bulk_metadata_tab(bulk_metadata_tab)

    def _build_sprite_editor_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(2, weight=0)
        parent.rowconfigure(0, weight=1)

        left = ttk.Frame(parent)
        left.grid(row=0, column=0, sticky="nsw")
        center = ttk.Frame(parent)
        center.grid(row=0, column=1, sticky="nsew", padx=(10, 10))
        right = ttk.Frame(parent)
        right.grid(row=0, column=2, sticky="nsew")
        center.columnconfigure(0, weight=1)
        center.rowconfigure(0, weight=1)

        # Left panel: image sequence and export.
        header_row = ttk.Frame(left)
        header_row.pack(fill="x")
        ttk.Label(header_row, text="Pack Images (rotation/order)", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(header_row, text="New File / Reset", command=self._reset_sprite_editor).pack(side="right")
        self.listbox = tk.Listbox(left, width=42, height=24)
        self.listbox.pack(fill="both", expand=True, pady=(6, 6))
        self.listbox.bind("<<ListboxSelect>>", self._on_list_select)

        btn_row = ttk.Frame(left)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Add PNGs", command=self._add_images).pack(side="left")
        ttk.Button(btn_row, text="Remove", command=self._remove_selected).pack(side="left", padx=5)
        ttk.Button(btn_row, text="Up", command=self._move_up).pack(side="left")
        ttk.Button(btn_row, text="Down", command=self._move_down).pack(side="left", padx=5)

        ttk.Separator(left, orient="horizontal").pack(fill="x", pady=10)
        ttk.Label(left, text="Export", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        export = ttk.Frame(left)
        export.pack(fill="x", pady=(6, 0))
        export.columnconfigure(1, weight=1)
        ttk.Label(export, text="Format").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Combobox(
            export,
            values=["webp", "png"],
            textvariable=self.export_format_var,
            state="readonly",
            width=12,
        ).grid(row=0, column=1, sticky="w")
        ttk.Label(export, text="WebP Quality").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(export, textvariable=self.webp_quality_var, width=10).grid(row=1, column=1, sticky="w")
        ttk.Checkbutton(export, text="WebP Lossless", variable=self.webp_lossless_var).grid(
            row=2, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(export, text="Zip Name").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(export, textvariable=self.zip_name_var, width=24).grid(row=3, column=1, sticky="ew")
        ttk.Button(left, text="Export Pack Folder", command=self._export_folder).pack(fill="x", pady=(10, 4))
        ttk.Button(left, text="Export Pack Zip", command=self._export_zip).pack(fill="x")
        ttk.Button(left, text="Export Metadata Only", command=self._export_metadata_only).pack(fill="x", pady=(6, 0))

        # Center panel: preview.
        self.canvas = tk.Canvas(center, bg="#1c1c1c", highlightthickness=1, highlightbackground="#4a4a4a")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.bind("<Configure>", lambda _e: self._request_render())
        self.canvas.bind("<ButtonPress-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel_zoom)
        self.canvas.bind("<Motion>", self._on_canvas_motion)

        hud = ttk.Frame(center)
        hud.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        hud.columnconfigure(2, weight=1)
        ttk.Label(hud, text="Zoom").grid(row=0, column=0, sticky="w")
        self.zoom_var = tk.DoubleVar(value=self.preview_zoom * 100.0)
        ttk.Scale(hud, from_=8, to=6400, variable=self.zoom_var, command=self._on_zoom_slider).grid(
            row=0, column=1, sticky="ew", padx=(6, 12)
        )
        ttk.Label(hud, textvariable=self._fps_text).grid(row=0, column=2, sticky="e")
        ttk.Label(
            hud,
            text="Drag guides directly. Drag empty space to pan.",
            foreground="#808080",
        ).grid(row=1, column=0, columnspan=3, sticky="w")

        # Right panel: image calibration + full pack metadata.
        right.columnconfigure(1, weight=1)
        ttk.Label(right, text="Selected Image Calibration", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )
        self.metrics_label = ttk.Label(right, text="", foreground="#8a8a8a")
        self.metrics_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 8))

        self._add_image_field(
            right,
            "Footprint Preset",
            "fit_mode",
            2,
            combobox=[
                "1x1",
                "1x2",
                "2x1",
                "full_2x2",
                "2x3",
                "3x2",
            ],
        )
        self._add_image_field(right, "Target Span", "target_span", 3)
        self._add_image_field(right, "Guide Left", "guide_left", 4)
        self._add_image_field(right, "Guide Center", "guide_center", 5)
        self._add_image_field(right, "Guide Right", "guide_right", 6)
        self._add_image_field(right, "Baseline Y", "baseline_y", 7)
        self._add_image_field(right, "Offset X", "offset_x", 8)
        self._add_image_field(right, "Offset Y", "offset_y", 9)
        auto_row = ttk.Frame(right)
        auto_row.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(4, 2))
        ttk.Button(auto_row, text="Auto Align Current", command=self._auto_align_current).pack(side="left")
        ttk.Button(auto_row, text="Auto Align All", command=self._auto_align_all).pack(side="left", padx=(6, 0))

        ttk.Separator(right, orient="horizontal").grid(row=11, column=0, columnspan=2, sticky="ew", pady=10)
        ttk.Label(right, text="Pack Metadata (shared)", font=("Segoe UI", 10, "bold")).grid(
            row=12, column=0, columnspan=2, sticky="w"
        )
        row_idx = 13
        for key, label in [
            ("id", "ID"),
            ("name", "Name"),
            ("set_id", "Set ID"),
            ("category", "Category"),
            ("theme", "Theme"),
            ("tiles_x", "Tiles X"),
            ("tiles_y", "Tiles Y"),
            ("variant_group", "Variant Group"),
            ("variant_label", "Variant Label"),
            ("group_label", "Group Label"),
            ("manufacturer", "Manufacturer"),
            ("link", "Link"),
            ("instructions", "Instructions"),
            ("notes", "Notes"),
        ]:
            ttk.Label(right, text=label).grid(row=row_idx, column=0, sticky="w", pady=2)
            if key in ("category", "theme"):
                choices = CATEGORY_OPTIONS if key == "category" else THEME_OPTIONS
                widget = ttk.Combobox(right, values=choices, textvariable=self.meta_vars[key], state="normal")
            else:
                widget = ttk.Entry(right, textvariable=self.meta_vars[key])
            widget.grid(row=row_idx, column=1, sticky="ew", pady=2)
            widget.bind("<FocusOut>", lambda _e: self._apply_pack_metadata_fields())
            widget.bind("<Return>", lambda _e: self._apply_pack_metadata_fields())
            row_idx += 1

        offsets_frame = ttk.LabelFrame(right, text="Per-Rotation Offsets")
        offsets_frame.grid(row=row_idx, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        for idx in range(4):
            r = ttk.Frame(offsets_frame)
            r.pack(fill="x", pady=1)
            ttk.Label(r, text=f"Offset {idx}", width=10).pack(side="left")
            ttk.Label(r, text="X").pack(side="left")
            ex = ttk.Entry(r, textvariable=self.offset_vars[f"offset_{idx}_x"], width=8)
            ex.pack(side="left", padx=(2, 8))
            ttk.Label(r, text="Y").pack(side="left")
            ey = ttk.Entry(r, textvariable=self.offset_vars[f"offset_{idx}_y"], width=8)
            ey.pack(side="left", padx=(2, 0))
            ex.bind("<FocusOut>", lambda _e: self._apply_pack_metadata_fields())
            ey.bind("<FocusOut>", lambda _e: self._apply_pack_metadata_fields())
            ex.bind("<Return>", lambda _e: self._apply_pack_metadata_fields())
            ey.bind("<Return>", lambda _e: self._apply_pack_metadata_fields())
        self.offsets_legend = tk.Canvas(
            offsets_frame,
            width=220,
            height=120,
            bg="white",
            highlightthickness=1,
            highlightbackground="#bfbfbf",
        )
        self.offsets_legend.pack(pady=(6, 4))
        self._draw_offsets_legend()

        ttk.Label(right, textvariable=self.status_var).grid(row=row_idx + 1, column=0, columnspan=2, sticky="w", pady=(8, 0))

    def _build_bulk_metadata_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        controls = ttk.Frame(parent)
        controls.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        controls.columnconfigure(1, weight=1)
        ttk.Label(controls, text="Root Folder").grid(row=0, column=0, sticky="w")
        ttk.Entry(controls, textvariable=self.bulk_root_var).grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Button(controls, text="Browse", command=self._bulk_choose_root).grid(row=0, column=2, padx=(0, 6))
        ttk.Button(controls, text="Scan", command=self._bulk_scan_root).grid(row=0, column=3)
        ttk.Label(controls, textvariable=self.bulk_status_var, foreground="#707070").grid(
            row=1, column=0, columnspan=4, sticky="w", pady=(4, 0)
        )

        body = ttk.Frame(parent)
        body.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        list_frame = ttk.LabelFrame(body, text="Discovered Metadata")
        list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.bulk_tree = ttk.Treeview(
            list_frame,
            columns=("id", "name", "category", "theme", "variant_options", "manufacturer", "source", "location"),
            show="headings",
            selectmode="extended",
        )
        self.bulk_tree_column_labels = {
            "id": "ID",
            "name": "Name",
            "category": "Category",
            "theme": "Theme",
            "variant_options": "Variant Options",
            "manufacturer": "Manufacturer",
            "source": "Source",
            "location": "Location",
        }
        for col, label in self.bulk_tree_column_labels.items():
            self.bulk_tree.heading(col, text=label, command=lambda c=col: self._bulk_sort_by_column(c))
        self.bulk_tree.column("id", width=170, anchor="w", stretch=False)
        self.bulk_tree.column("name", width=190, anchor="w", stretch=False)
        self.bulk_tree.column("category", width=130, anchor="w")
        self.bulk_tree.column("theme", width=130, anchor="w")
        self.bulk_tree.column("variant_options", width=170, anchor="w")
        self.bulk_tree.column("manufacturer", width=150, anchor="w")
        self.bulk_tree.column("source", width=70, anchor="center")
        self.bulk_tree.column("location", width=380, anchor="w")
        self.bulk_tree.grid(row=0, column=0, sticky="nsew")
        self._bulk_update_tree_heading_labels()
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.bulk_tree.yview)
        list_scroll.grid(row=0, column=1, sticky="ns")
        self.bulk_tree.configure(yscrollcommand=list_scroll.set)

        self.bulk_tree.bind("<<TreeviewSelect>>", self._bulk_on_tree_select)

        right_panel = ttk.Frame(body)
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(0, weight=0)
        right_panel.rowconfigure(1, weight=0)
        right_panel.rowconfigure(2, weight=1)

        mode_row = ttk.Frame(right_panel)
        mode_row.grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(mode_row, text="Edit Mode").pack(side="left", padx=(0, 8))
        ttk.Radiobutton(
            mode_row,
            text="Single",
            value="single",
            variable=self.bulk_mode_var,
            command=self._bulk_update_mode_visibility,
        ).pack(side="left")
        ttk.Radiobutton(
            mode_row,
            text="Bulk",
            value="bulk",
            variable=self.bulk_mode_var,
            command=self._bulk_update_mode_visibility,
        ).pack(side="left", padx=(8, 0))

        single_frame = ttk.LabelFrame(right_panel, text="Selected Metadata (Single File)")
        single_frame.grid(row=1, column=0, sticky="ew")
        single_frame.columnconfigure(1, weight=1)
        self.bulk_single_frame = single_frame

        single_fields = [
            ("id", "ID"),
            ("name", "Name"),
            ("set_id", "Set ID"),
            ("category", "Category"),
            ("theme", "Theme"),
            ("tiles_x", "Tiles X"),
            ("tiles_y", "Tiles Y"),
            ("variant_group", "Variant Group"),
            ("variant_label", "Variant Label"),
            ("group_label", "Group Label"),
            ("manufacturer", "Manufacturer"),
            ("link", "Link"),
            ("instructions", "Instructions"),
            ("notes", "Notes"),
        ]

        for row_idx, (key, label) in enumerate(single_fields):
            self.bulk_single_vars[key] = tk.StringVar(value="")
            ttk.Label(single_frame, text=label).grid(row=row_idx, column=0, sticky="w", pady=2)
            if key in ("category", "theme"):
                choices = CATEGORY_OPTIONS if key == "category" else THEME_OPTIONS
                widget = ttk.Combobox(single_frame, values=choices, textvariable=self.bulk_single_vars[key], state="normal")
            else:
                widget = ttk.Entry(single_frame, textvariable=self.bulk_single_vars[key])
            widget.grid(row=row_idx, column=1, sticky="ew", pady=2)
            self.bulk_single_inputs[key] = widget

        single_offsets_frame = ttk.LabelFrame(single_frame, text="Per-Rotation Offsets")
        single_offsets_frame.grid(row=len(single_fields), column=0, columnspan=2, sticky="ew", pady=(6, 0))
        for idx in range(4):
            self.bulk_single_vars[f"offset_{idx}_x"] = tk.StringVar(value="0")
            self.bulk_single_vars[f"offset_{idx}_y"] = tk.StringVar(value="0")
            row = ttk.Frame(single_offsets_frame)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=f"Offset {idx}", width=10).pack(side="left")
            ttk.Label(row, text="X").pack(side="left")
            ex = ttk.Entry(row, textvariable=self.bulk_single_vars[f"offset_{idx}_x"], width=8)
            ex.pack(side="left", padx=(2, 8))
            ttk.Label(row, text="Y").pack(side="left")
            ey = ttk.Entry(row, textvariable=self.bulk_single_vars[f"offset_{idx}_y"], width=8)
            ey.pack(side="left", padx=(2, 0))
            self.bulk_single_inputs[f"offset_{idx}_x"] = ex
            self.bulk_single_inputs[f"offset_{idx}_y"] = ey

        self.bulk_single_offsets_legend = tk.Canvas(
            single_frame,
            width=220,
            height=120,
            bg="white",
            highlightthickness=1,
            highlightbackground="#bfbfbf",
        )
        self.bulk_single_offsets_legend.grid(row=len(single_fields) + 1, column=0, columnspan=2, sticky="w", pady=(6, 2))
        self._draw_offsets_legend(self.bulk_single_offsets_legend)

        single_actions = ttk.Frame(single_frame)
        single_actions.grid(row=len(single_fields) + 2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.bulk_single_save_btn = ttk.Button(single_actions, text="Save Selected", command=self._bulk_save_selected_single)
        self.bulk_single_save_btn.pack(side="left")
        self.bulk_single_reload_btn = ttk.Button(single_actions, text="Reload Selected", command=self._bulk_reload_selected_single)
        self.bulk_single_reload_btn.pack(side="left", padx=(6, 0))
        ttk.Label(single_frame, textvariable=self.bulk_single_status_var, foreground="#707070").grid(
            row=len(single_fields) + 3, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

        edit_frame = ttk.LabelFrame(right_panel, text="Bulk Update Fields")
        edit_frame.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        edit_frame.columnconfigure(2, weight=1)
        self.bulk_edit_frame = edit_frame

        bulk_fields = [
            ("id", "ID"),
            ("name", "Name"),
            ("set_id", "Set ID"),
            ("category", "Category"),
            ("theme", "Theme"),
            ("tiles_x", "Tiles X"),
            ("tiles_y", "Tiles Y"),
            ("variant_group", "Variant Group"),
            ("variant_label", "Variant Label"),
            ("group_label", "Group Label"),
            ("manufacturer", "Manufacturer"),
            ("link", "Link"),
            ("instructions", "Instructions"),
            ("notes", "Notes"),
        ]

        for row_idx, (key, label) in enumerate(bulk_fields):
            self.bulk_apply_vars[key] = tk.BooleanVar(value=False)
            self.bulk_field_vars[key] = tk.StringVar(value="")
            cb = ttk.Checkbutton(edit_frame, variable=self.bulk_apply_vars[key])
            cb.grid(row=row_idx, column=0, sticky="w", padx=(0, 4))
            ttk.Label(edit_frame, text=label).grid(row=row_idx, column=1, sticky="w", pady=2)
            if key in ("category", "theme"):
                choices = CATEGORY_OPTIONS if key == "category" else THEME_OPTIONS
                w = ttk.Combobox(edit_frame, values=choices, textvariable=self.bulk_field_vars[key], state="normal")
            else:
                w = ttk.Entry(edit_frame, textvariable=self.bulk_field_vars[key])
            w.grid(row=row_idx, column=2, sticky="ew", pady=2)
            if key in ("id", "name"):
                self.bulk_apply_vars[key].set(False)
                cb.configure(state="disabled")
                w.configure(state="disabled")

        action_row = ttk.Frame(edit_frame)
        action_row.grid(row=len(bulk_fields), column=0, columnspan=3, sticky="ew", pady=(10, 0))
        ttk.Button(action_row, text="Apply To Selected", command=self._bulk_apply_to_selected).pack(side="left")
        ttk.Button(action_row, text="Clear Fields", command=self._bulk_clear_fields).pack(side="left", padx=(6, 0))
        ttk.Label(edit_frame, textvariable=self.bulk_apply_status_var, foreground="#707070").grid(
            row=len(bulk_fields) + 1, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )
        self._bulk_set_single_editor_enabled(False)
        self._bulk_update_mode_visibility()

    def _add_image_field(self, parent, label: str, key: str, row: int, combobox: Optional[list[str]] = None) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=2)
        if combobox is None:
            w = ttk.Entry(parent, textvariable=self.image_vars[key])
        else:
            w = ttk.Combobox(parent, values=combobox, textvariable=self.image_vars[key], state="readonly")
        w.grid(row=row, column=1, sticky="ew", pady=2)
        w.bind("<Return>", lambda _e: self._apply_image_fields())
        if combobox is not None:
            w.bind("<<ComboboxSelected>>", lambda _e: self._apply_image_fields())

    def _sync_zip_name_to_id(self, force: bool = False) -> None:
        id_value = _normalize_id(self.meta_vars["id"].get() or self.pack_meta.id or "sprite_pack")
        auto_name = f"{id_value}.zip"
        current = self.zip_name_var.get().strip()
        if force or current == "" or current == self._last_auto_zip_name:
            self.zip_name_var.set(auto_name)
            self._last_auto_zip_name = auto_name
            return
        # If user entered a custom name, keep it.
        if current.lower() == auto_name.lower():
            self._last_auto_zip_name = current

    def _reset_sprite_editor(self) -> None:
        self.items.clear()
        self.active_idx = None
        self.pack_meta = PackMetadata()
        self.preview_zoom = 0.55
        self.pan_x = 0.0
        self.pan_y = 0.0
        self.zoom_var.set(self.preview_zoom * 100.0)
        self.export_format_var.set("webp")
        self.webp_quality_var.set("95")
        self.webp_lossless_var.set(False)

        self.meta_vars["id"].set(self.pack_meta.id)
        self.meta_vars["name"].set(self.pack_meta.name)
        self.meta_vars["set_id"].set(self.pack_meta.set_id)
        self.meta_vars["category"].set(self.pack_meta.category)
        self.meta_vars["theme"].set(self.pack_meta.theme)
        self.meta_vars["tiles_x"].set(str(self.pack_meta.tiles_x))
        self.meta_vars["tiles_y"].set(str(self.pack_meta.tiles_y))
        self.meta_vars["variant_group"].set(self.pack_meta.variant_group)
        self.meta_vars["variant_label"].set(self.pack_meta.variant_label)
        self.meta_vars["group_label"].set(self.pack_meta.group_label)
        self.meta_vars["manufacturer"].set(self.pack_meta.manufacturer)
        self.meta_vars["link"].set(self.pack_meta.link)
        self.meta_vars["instructions"].set(self.pack_meta.instructions)
        self.meta_vars["notes"].set(self.pack_meta.notes)
        for idx in range(4):
            self.offset_vars[f"offset_{idx}_x"].set("0")
            self.offset_vars[f"offset_{idx}_y"].set("0")

        self._sync_zip_name_to_id(force=True)
        self._refresh_image_list()
        self._load_active_item_fields()
        self.status_var.set("Reset editor for a new model.")
        self._request_render()

    def _bulk_set_single_editor_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for widget in self.bulk_single_inputs.values():
            widget.configure(state=state)
        self.bulk_single_save_btn.configure(state=state)
        self.bulk_single_reload_btn.configure(state=state)

    def _bulk_update_mode_visibility(self) -> None:
        mode = self.bulk_mode_var.get().strip().lower()
        if mode == "bulk":
            self.bulk_single_frame.grid_remove()
            self.bulk_edit_frame.grid()
        else:
            self.bulk_edit_frame.grid_remove()
            self.bulk_single_frame.grid()
            self._bulk_on_tree_select()

    def _bulk_update_tree_heading_labels(self) -> None:
        arrow = "▼" if self.bulk_sort_reverse else "▲"
        for col, label in self.bulk_tree_column_labels.items():
            text = f"{label} {arrow}" if col == self.bulk_sort_column else label
            self.bulk_tree.heading(col, text=text, command=lambda c=col: self._bulk_sort_by_column(c))

    def _bulk_sort_by_column(self, column: str) -> None:
        if not self.bulk_entries:
            self.bulk_sort_column = column
            self.bulk_sort_reverse = False
            self._bulk_update_tree_heading_labels()
            return
        if self.bulk_sort_column == column:
            self.bulk_sort_reverse = not self.bulk_sort_reverse
        else:
            self.bulk_sort_column = column
            self.bulk_sort_reverse = False

        selected_entries: list[dict] = []
        for iid in self.bulk_tree.selection():
            entry = self.bulk_entry_by_iid.get(iid)
            if entry is not None:
                selected_entries.append(entry)

        self.bulk_entries.sort(
            key=lambda e: str(e.get(column, "")).strip().lower(),
            reverse=self.bulk_sort_reverse,
        )
        self._refresh_bulk_tree()
        self._bulk_update_tree_heading_labels()

        reselect_iids: list[str] = []
        for idx, entry in enumerate(self.bulk_entries):
            if entry in selected_entries:
                reselect_iids.append(str(idx))
        if reselect_iids:
            self.bulk_tree.selection_set(tuple(reselect_iids))
            self.bulk_tree.focus(reselect_iids[0])
            self.bulk_tree.see(reselect_iids[0])
            self._bulk_on_tree_select()

    def _bulk_clear_single_editor_fields(self) -> None:
        for var in self.bulk_single_vars.values():
            var.set("")

    def _bulk_selected_entry(self) -> Optional[dict]:
        selected = self.bulk_tree.selection()
        if len(selected) != 1:
            return None
        return self.bulk_entry_by_iid.get(selected[0], None)

    def _bulk_on_tree_select(self, _event=None) -> None:
        selected = self.bulk_tree.selection()
        if len(selected) == 0:
            self._bulk_clear_single_editor_fields()
            self._bulk_set_single_editor_enabled(False)
            self.bulk_single_status_var.set("Select one row to edit a single metadata file.")
            return
        if len(selected) > 1:
            self._bulk_clear_single_editor_fields()
            self._bulk_set_single_editor_enabled(False)
            self.bulk_single_status_var.set("Single-file editor is disabled while multiple rows are selected.")
            return
        entry = self.bulk_entry_by_iid.get(selected[0], None)
        if entry is None:
            self._bulk_clear_single_editor_fields()
            self._bulk_set_single_editor_enabled(False)
            self.bulk_single_status_var.set("Unable to resolve selected row.")
            return
        meta = self._bulk_read_entry_metadata(entry)
        if meta is None:
            self._bulk_clear_single_editor_fields()
            self._bulk_set_single_editor_enabled(False)
            self.bulk_single_status_var.set("Failed to load metadata from selected row.")
            return
        for key, var in self.bulk_single_vars.items():
            if key in ("tiles_x", "tiles_y"):
                var.set(str(meta.get(key, "")))
            elif key.startswith("offset_"):
                m = re.fullmatch(r"offset_(\d+)_(x|y)", key)
                if m is None:
                    var.set("0")
                    continue
                idx = m.group(1)
                axis = 0 if m.group(2) == "x" else 1
                offsets = meta.get("offsets", {})
                value = 0.0
                if isinstance(offsets, dict):
                    pair = offsets.get(idx, [0.0, 0.0])
                    if isinstance(pair, list) and len(pair) > axis:
                        value = _safe_float(str(pair[axis]), 0.0)
                var.set(f"{value:.2f}")
            else:
                var.set(str(meta.get(key, "")).strip())
        self._bulk_set_single_editor_enabled(True)
        self.bulk_single_status_var.set(f"Editing: {entry.get('location', '')}")

    def _bulk_read_entry_metadata(self, entry: dict) -> Optional[dict]:
        source = entry.get("source", "")
        if source == "folder":
            path = Path(str(entry.get("metadata_path", "")))
            return self._load_json_file(path)
        if source == "zip":
            zip_path = Path(str(entry.get("zip_path", "")))
            zip_chain = list(entry.get("zip_chain", []))
            zip_entry_path = str(entry.get("zip_entry_path", ""))
            return self._read_json_from_zip_path(zip_path, zip_chain, zip_entry_path)
        return None

    def _read_json_from_zip_path(self, zip_path: Path, zip_chain: list[str], zip_entry_path: str) -> Optional[dict]:
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                payload = self._read_bytes_from_zip_chain(zf, zip_chain, zip_entry_path)
            return self._parse_json_bytes(payload)
        except Exception:
            return None

    def _read_bytes_from_zip_chain(self, zip_file: zipfile.ZipFile, zip_chain: list[str], leaf_path: str) -> bytes:
        if not zip_chain:
            return zip_file.read(leaf_path)
        nested_payload = zip_file.read(zip_chain[0])
        with zipfile.ZipFile(BytesIO(nested_payload), "r") as nested_zip:
            return self._read_bytes_from_zip_chain(nested_zip, zip_chain[1:], leaf_path)

    def _bulk_build_single_updates(self, current_meta: dict) -> dict:
        updates: dict[str, object] = {}
        for key, var in self.bulk_single_vars.items():
            if key.startswith("offset_"):
                continue
            raw = var.get()
            if key == "id":
                normalized = _normalize_id(raw)
                if normalized:
                    updates[key] = normalized
                elif "id" in current_meta:
                    updates[key] = current_meta["id"]
            elif key == "name":
                name = raw.strip()
                if name != "":
                    updates[key] = name
                elif "name" in current_meta:
                    updates[key] = current_meta["name"]
            elif key == "category":
                updates[key] = _normalize_category(raw)
            elif key in ("tiles_x", "tiles_y"):
                default_int = int(current_meta.get(key, 2)) if str(current_meta.get(key, "")).strip() != "" else 2
                updates[key] = max(1, _safe_int(raw, default_int))
            else:
                updates[key] = raw.strip()
        existing_offsets = current_meta.get("offsets", {})
        if not isinstance(existing_offsets, dict):
            existing_offsets = {}
        offsets: dict[str, list[float]] = {}
        for idx in range(4):
            default_pair = existing_offsets.get(str(idx), [0.0, 0.0])
            if not isinstance(default_pair, list) or len(default_pair) < 2:
                default_pair = [0.0, 0.0]
            default_x = _safe_float(str(default_pair[0]), 0.0)
            default_y = _safe_float(str(default_pair[1]), 0.0)
            raw_x = self.bulk_single_vars[f"offset_{idx}_x"].get()
            raw_y = self.bulk_single_vars[f"offset_{idx}_y"].get()
            offsets[str(idx)] = [
                _safe_float(raw_x, default_x),
                _safe_float(raw_y, default_y),
            ]
        updates["offsets"] = offsets
        return updates

    def _bulk_reselect_entry(self, entry: dict) -> None:
        try:
            idx = self.bulk_entries.index(entry)
        except ValueError:
            self._bulk_on_tree_select()
            return
        iid = str(idx)
        self.bulk_tree.selection_set(iid)
        self.bulk_tree.focus(iid)
        self.bulk_tree.see(iid)
        self._bulk_on_tree_select()

    def _bulk_save_selected_single(self) -> None:
        entry = self._bulk_selected_entry()
        if entry is None:
            messagebox.showinfo("Single metadata edit", "Select exactly one row first.")
            return
        current_meta = self._bulk_read_entry_metadata(entry)
        if current_meta is None:
            messagebox.showerror("Single metadata edit", "Failed to read metadata for selected row.")
            return
        new_data = self._apply_updates_to_metadata(current_meta, self._bulk_build_single_updates(current_meta))
        source = entry.get("source", "")
        if source == "folder":
            path = Path(str(entry.get("metadata_path", "")))
            with path.open("w", encoding="utf-8") as f:
                json.dump(new_data, f, indent=2, ensure_ascii=True)
        elif source == "zip":
            zip_path = Path(str(entry.get("zip_path", "")))
            zip_chain = list(entry.get("zip_chain", []))
            zip_entry_path = str(entry.get("zip_entry_path", ""))
            self._rewrite_zip_metadata_full(zip_path, zip_chain, zip_entry_path, new_data)
        else:
            messagebox.showerror("Single metadata edit", f"Unknown source type: {source}")
            return
        self._bulk_update_entry_summary(entry, new_data)
        self._refresh_bulk_tree()
        self._bulk_reselect_entry(entry)
        self.bulk_single_status_var.set(f"Saved: {entry.get('location', '')}")

    def _bulk_reload_selected_single(self) -> None:
        entry = self._bulk_selected_entry()
        if entry is None:
            messagebox.showinfo("Single metadata edit", "Select exactly one row first.")
            return
        self._bulk_on_tree_select()

    def _bulk_choose_root(self) -> None:
        selected = filedialog.askdirectory(title="Choose root folder to scan")
        if not selected:
            return
        self.bulk_root_var.set(selected)

    def _bulk_scan_root(self) -> None:
        root_raw = self.bulk_root_var.get().strip()
        if not root_raw:
            root_raw = filedialog.askdirectory(title="Choose root folder to scan")
            if not root_raw:
                return
            self.bulk_root_var.set(root_raw)
        root_path = Path(root_raw)
        if not root_path.exists() or not root_path.is_dir():
            messagebox.showerror("Bulk metadata scan", f"Folder does not exist:\n{root_path}")
            return
        entries = self._collect_bulk_metadata_entries(root_path)
        self.bulk_entries = entries
        self._refresh_bulk_tree()
        self.bulk_status_var.set(f"Found {len(entries)} metadata.json file(s).")
        self.bulk_apply_status_var.set("")

    def _collect_bulk_metadata_entries(self, root_path: Path) -> list[dict]:
        entries: list[dict] = []
        for dirpath, _dirnames, filenames in os.walk(root_path):
            base = Path(dirpath)
            if "metadata.json" in filenames:
                metadata_path = base / "metadata.json"
                data = self._load_json_file(metadata_path)
                if data is not None:
                    entries.append(
                        {
                            "source": "folder",
                            "metadata_path": str(metadata_path),
                            "zip_path": "",
                            "zip_entry_path": "",
                            "id": str(data.get("id", "")).strip(),
                            "name": str(data.get("name", "")).strip(),
                            "category": str(data.get("category", "")).strip(),
                            "theme": str(data.get("theme", "")).strip(),
                            "variant_options": self._extract_variant_options_text(data),
                            "manufacturer": str(data.get("manufacturer", "")).strip(),
                            "location": str(metadata_path.parent.relative_to(root_path)),
                        }
                    )
            for filename in filenames:
                if not filename.lower().endswith(".zip"):
                    continue
                zip_path = base / filename
                entries.extend(self._collect_zip_metadata_entries(zip_path, root_path))
        entries.sort(key=lambda e: (e["id"].lower(), e["name"].lower(), e["location"].lower()))
        return entries

    def _collect_zip_metadata_entries(self, zip_path: Path, root_path: Path) -> list[dict]:
        out: list[dict] = []
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                self._collect_zip_metadata_entries_recursive(
                    zip_file=zf,
                    out=out,
                    zip_path=zip_path,
                    root_path=root_path,
                    zip_chain=[],
                )
        except Exception:
            pass
        return out

    def _collect_zip_metadata_entries_recursive(
        self,
        zip_file: zipfile.ZipFile,
        out: list[dict],
        zip_path: Path,
        root_path: Path,
        zip_chain: list[str],
    ) -> None:
        for member in zip_file.namelist():
            member_norm = member.replace("\\", "/")
            lower_name = member_norm.lower()
            if lower_name.endswith("/") or member_norm == "":
                continue
            if lower_name == "metadata.json" or lower_name.endswith("/metadata.json"):
                try:
                    payload = zip_file.read(member)
                    parsed = self._parse_json_bytes(payload)
                    if parsed is None:
                        continue
                except Exception:
                    continue
                location_parts = [str(zip_path.relative_to(root_path))] + zip_chain + [member_norm]
                out.append(
                    {
                        "source": "zip",
                        "metadata_path": "",
                        "zip_path": str(zip_path),
                        "zip_chain": list(zip_chain),
                        "zip_entry_path": member_norm,
                        "id": str(parsed.get("id", "")).strip(),
                        "name": str(parsed.get("name", "")).strip(),
                        "category": str(parsed.get("category", "")).strip(),
                        "theme": str(parsed.get("theme", "")).strip(),
                        "variant_options": self._extract_variant_options_text(parsed),
                        "manufacturer": str(parsed.get("manufacturer", "")).strip(),
                        "location": "::".join(location_parts),
                    }
                )
                continue
            if not lower_name.endswith(".zip"):
                continue
            try:
                nested_payload = zip_file.read(member)
                with zipfile.ZipFile(BytesIO(nested_payload), "r") as nested_zip:
                    self._collect_zip_metadata_entries_recursive(
                        zip_file=nested_zip,
                        out=out,
                        zip_path=zip_path,
                        root_path=root_path,
                        zip_chain=zip_chain + [member_norm],
                    )
            except Exception:
                continue

    def _parse_json_bytes(self, payload: bytes) -> Optional[dict]:
        try:
            parsed = json.loads(payload.decode("utf-8"))
            if isinstance(parsed, dict):
                return parsed
        except Exception:
            return None
        return None

    def _load_json_file(self, path: Path) -> Optional[dict]:
        try:
            with path.open("r", encoding="utf-8") as f:
                parsed = json.load(f)
            if isinstance(parsed, dict):
                return parsed
        except Exception:
            return None
        return None

    def _extract_variant_options_text(self, data: dict) -> str:
        variant_options = data.get("variant_options", None)
        if isinstance(variant_options, list):
            return ", ".join(str(item).strip() for item in variant_options if str(item).strip())
        if isinstance(variant_options, str):
            return variant_options.strip()
        parts = [
            str(data.get("variant_group", "")).strip(),
            str(data.get("variant_label", "")).strip(),
            str(data.get("group_label", "")).strip(),
        ]
        parts = [part for part in parts if part]
        return " / ".join(parts)

    def _bulk_update_entry_summary(self, entry: dict, data: dict) -> None:
        entry["id"] = str(data.get("id", "")).strip()
        entry["name"] = str(data.get("name", "")).strip()
        entry["category"] = str(data.get("category", "")).strip()
        entry["theme"] = str(data.get("theme", "")).strip()
        entry["variant_options"] = self._extract_variant_options_text(data)
        entry["manufacturer"] = str(data.get("manufacturer", "")).strip()

    def _refresh_bulk_tree(self) -> None:
        for iid in self.bulk_tree.get_children():
            self.bulk_tree.delete(iid)
        self.bulk_entry_by_iid.clear()
        for idx, entry in enumerate(self.bulk_entries):
            iid = str(idx)
            self.bulk_entry_by_iid[iid] = entry
            self.bulk_tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    entry.get("id", ""),
                    entry.get("name", ""),
                    entry.get("category", ""),
                    entry.get("theme", ""),
                    entry.get("variant_options", ""),
                    entry.get("manufacturer", ""),
                    entry.get("source", ""),
                    entry.get("location", ""),
                ),
            )
        self._bulk_on_tree_select()

    def _bulk_clear_fields(self) -> None:
        for var in self.bulk_apply_vars.values():
            var.set(False)
        for var in self.bulk_field_vars.values():
            var.set("")
        self.bulk_apply_status_var.set("")

    def _bulk_build_update_payload(self) -> dict[str, object]:
        updates: dict[str, object] = {}
        for key, enabled_var in self.bulk_apply_vars.items():
            if not enabled_var.get():
                continue
            raw = self.bulk_field_vars[key].get()
            if key == "id":
                updates[key] = _normalize_id(raw)
            elif key == "category":
                updates[key] = _normalize_category(raw)
            elif key in ("tiles_x", "tiles_y"):
                updates[key] = max(1, _safe_int(raw, 2))
            else:
                updates[key] = raw.strip()
        return updates

    def _bulk_apply_to_selected(self) -> None:
        selected = self.bulk_tree.selection()
        if not selected:
            messagebox.showinfo("Bulk metadata", "Select one or more rows first.")
            return
        updates = self._bulk_build_update_payload()
        if not updates:
            messagebox.showinfo("Bulk metadata", "Select at least one field checkbox to apply.")
            return

        updated_count = 0
        errors: list[str] = []
        for iid in selected:
            entry = self.bulk_entry_by_iid.get(iid)
            if entry is None:
                continue
            try:
                self._bulk_apply_single_entry(entry, updates)
                updated_count += 1
            except Exception as exc:
                location = entry.get("location", "unknown")
                errors.append(f"{location}: {exc}")

        self._refresh_bulk_tree()
        self.bulk_apply_status_var.set(f"Updated {updated_count}/{len(selected)} selected metadata file(s).")
        if errors:
            messagebox.showwarning("Bulk metadata warnings", "\n".join(errors[:10]))

    def _bulk_apply_single_entry(self, entry: dict, updates: dict[str, object]) -> None:
        source = entry.get("source", "")
        if source == "folder":
            path = Path(str(entry.get("metadata_path", "")))
            data = self._load_json_file(path)
            if data is None:
                raise ValueError("Unable to read metadata.json")
            new_data = self._apply_updates_to_metadata(data, updates)
            with path.open("w", encoding="utf-8") as f:
                json.dump(new_data, f, indent=2, ensure_ascii=True)
            self._bulk_update_entry_summary(entry, new_data)
            return
        if source == "zip":
            zip_path = Path(str(entry.get("zip_path", "")))
            zip_chain = list(entry.get("zip_chain", []))
            zip_entry_path = str(entry.get("zip_entry_path", ""))
            new_data = self._rewrite_zip_metadata(zip_path, zip_chain, zip_entry_path, updates)
            self._bulk_update_entry_summary(entry, new_data)
            return
        raise ValueError(f"Unknown source type: {source}")

    def _apply_updates_to_metadata(self, data: dict, updates: dict[str, object]) -> dict:
        updated = dict(data)
        for key, value in updates.items():
            updated[key] = value
        return updated

    def _rewrite_zip_metadata(
        self, zip_path: Path, zip_chain: list[str], zip_entry_path: str, updates: dict[str, object]
    ) -> dict:
        tmp_fd, tmp_name = tempfile.mkstemp(prefix="sprite-meta-", suffix=".zip", dir=str(zip_path.parent))
        os.close(tmp_fd)
        tmp_path = Path(tmp_name)
        updated_meta: Optional[dict] = None
        try:
            with zipfile.ZipFile(zip_path, "r") as src_zip:
                with zipfile.ZipFile(tmp_path, "w") as dst_zip:
                    for info in src_zip.infolist():
                        payload = src_zip.read(info.filename)
                        current_name = info.filename.replace("\\", "/")
                        if zip_chain:
                            if current_name == zip_chain[0]:
                                payload, updated_meta = self._rewrite_nested_zip_payload(
                                    payload,
                                    zip_chain[1:],
                                    zip_entry_path,
                                    updates,
                                )
                        elif current_name == zip_entry_path:
                            parsed = self._parse_json_bytes(payload)
                            if parsed is None:
                                raise ValueError("metadata.json inside zip is not an object")
                            updated_meta = self._apply_updates_to_metadata(parsed, updates)
                            payload = json.dumps(updated_meta, indent=2, ensure_ascii=True).encode("utf-8")
                        dst_zip.writestr(info, payload)
            if updated_meta is None:
                raise ValueError("metadata.json entry not found in zip")
            os.replace(tmp_path, zip_path)
            return updated_meta
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except Exception:
                    pass

    def _rewrite_zip_metadata_full(
        self, zip_path: Path, zip_chain: list[str], zip_entry_path: str, new_meta: dict
    ) -> None:
        tmp_fd, tmp_name = tempfile.mkstemp(prefix="sprite-meta-", suffix=".zip", dir=str(zip_path.parent))
        os.close(tmp_fd)
        tmp_path = Path(tmp_name)
        replaced = False
        try:
            with zipfile.ZipFile(zip_path, "r") as src_zip:
                with zipfile.ZipFile(tmp_path, "w") as dst_zip:
                    for info in src_zip.infolist():
                        payload = src_zip.read(info.filename)
                        current_name = info.filename.replace("\\", "/")
                        if zip_chain:
                            if current_name == zip_chain[0]:
                                payload, nested_replaced = self._rewrite_nested_zip_payload_full(
                                    payload,
                                    zip_chain[1:],
                                    zip_entry_path,
                                    new_meta,
                                )
                                if nested_replaced:
                                    replaced = True
                        elif current_name == zip_entry_path:
                            replaced = True
                            payload = json.dumps(new_meta, indent=2, ensure_ascii=True).encode("utf-8")
                        dst_zip.writestr(info, payload)
            if not replaced:
                raise ValueError("metadata.json entry not found in zip")
            os.replace(tmp_path, zip_path)
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except Exception:
                    pass

    def _rewrite_nested_zip_payload(
        self,
        zip_payload: bytes,
        zip_chain: list[str],
        zip_entry_path: str,
        updates: dict[str, object],
    ) -> tuple[bytes, Optional[dict]]:
        updated_meta: Optional[dict] = None
        src_buffer = BytesIO(zip_payload)
        out_buffer = BytesIO()
        with zipfile.ZipFile(src_buffer, "r") as src_zip:
            with zipfile.ZipFile(out_buffer, "w") as dst_zip:
                for info in src_zip.infolist():
                    payload = src_zip.read(info.filename)
                    current_name = info.filename.replace("\\", "/")
                    if zip_chain:
                        if current_name == zip_chain[0]:
                            payload, updated_meta = self._rewrite_nested_zip_payload(
                                payload,
                                zip_chain[1:],
                                zip_entry_path,
                                updates,
                            )
                    elif current_name == zip_entry_path:
                        parsed = self._parse_json_bytes(payload)
                        if parsed is None:
                            raise ValueError("metadata.json inside nested zip is not an object")
                        updated_meta = self._apply_updates_to_metadata(parsed, updates)
                        payload = json.dumps(updated_meta, indent=2, ensure_ascii=True).encode("utf-8")
                    dst_zip.writestr(info, payload)
        if updated_meta is None:
            return zip_payload, None
        return out_buffer.getvalue(), updated_meta

    def _rewrite_nested_zip_payload_full(
        self,
        zip_payload: bytes,
        zip_chain: list[str],
        zip_entry_path: str,
        new_meta: dict,
    ) -> tuple[bytes, bool]:
        replaced = False
        src_buffer = BytesIO(zip_payload)
        out_buffer = BytesIO()
        with zipfile.ZipFile(src_buffer, "r") as src_zip:
            with zipfile.ZipFile(out_buffer, "w") as dst_zip:
                for info in src_zip.infolist():
                    payload = src_zip.read(info.filename)
                    current_name = info.filename.replace("\\", "/")
                    if zip_chain:
                        if current_name == zip_chain[0]:
                            payload, nested_replaced = self._rewrite_nested_zip_payload_full(
                                payload,
                                zip_chain[1:],
                                zip_entry_path,
                                new_meta,
                            )
                            if nested_replaced:
                                replaced = True
                    elif current_name == zip_entry_path:
                        replaced = True
                        payload = json.dumps(new_meta, indent=2, ensure_ascii=True).encode("utf-8")
                    dst_zip.writestr(info, payload)
        if not replaced:
            return zip_payload, False
        return out_buffer.getvalue(), True

    def _draw_offsets_legend(self, canvas: Optional[tk.Canvas] = None) -> None:
        c = canvas if canvas is not None else self.offsets_legend
        c.delete("all")
        src_w = 1089.338
        src_h = 516.427
        dst_w = int(c.cget("width"))
        dst_h = int(c.cget("height"))
        margin = 4.0
        scale = min((dst_w - margin * 2.0) / src_w, (dst_h - margin * 2.0) / src_h)
        draw_w = src_w * scale
        draw_h = src_h * scale
        offset_x = (dst_w - draw_w) * 0.5
        offset_y = (dst_h - draw_h) * 0.5

        def p(x: float, y: float) -> tuple[float, float]:
            return (offset_x + x * scale, offset_y + y * scale)

        def rect(x: float, y: float, w: float, h: float, color: str) -> None:
            a = p(x, y)
            b = p(x + w, y + h)
            c.create_rectangle(a[0], a[1], b[0], b[1], fill=color, outline="")

        poly = [
            p(1084.669, 258.213),
            p(544.669, 514.213),
            p(4.669, 258.213),
            p(544.669, 2.213),
        ]
        c.create_polygon(
            poly[0][0], poly[0][1],
            poly[1][0], poly[1][1],
            poly[2][0], poly[2][1],
            poly[3][0], poly[3][1],
            outline="#231f20",
            fill="",
            width=2,
        )

        rect(904.076, 244.781, 88.31, 26.864, "#ED1C24")
        rect(67.076, 244.781, 88.31, 26.864, "#ED1C24")

        x_path = (
            "M365.865,366.891l-16.144-27.977c-6.574-10.702-10.702-17.642-14.646-24.965h-0.375"
            "c-3.57,7.323-7.132,14.08-13.706,25.149l-15.211,27.793H287.01l38.678-64.026l-37.173-62.527"
            "h18.957l16.71,29.674c4.701,8.255,8.263,14.646,11.642,21.403h0.566c3.57-7.506,6.757-13.331,"
            "11.451-21.403l17.275-29.674h18.774l-38.487,61.595l39.427,64.958H365.865z"
        )
        y_path = (
            "M730.689,366.891V313.2l-39.993-72.862h18.59l17.826,34.933c4.892,9.57,8.638,17.275,"
            "12.582,26.096h0.382c3.562-8.271,7.889-16.526,12.765-26.096l18.208-34.933h18.59L747.2,313.001v53.89"
            "H730.689z"
        )
        self._draw_svg_path_fill(c, x_path, p, "#231f20")
        self._draw_svg_path_fill(c, y_path, p, "#231f20")

        plus_points = [
            p(602.166, 406.897), p(562.16, 406.897), p(562.16, 366.891), p(527.178, 366.891),
            p(527.178, 406.897), p(487.172, 406.897), p(487.172, 441.879), p(527.178, 441.879),
            p(527.178, 481.885), p(562.16, 481.885), p(562.16, 441.879), p(602.166, 441.879),
        ]
        flat_plus: list[float] = []
        for px, py in plus_points:
            flat_plus.extend([px, py])
        c.create_polygon(*flat_plus, fill="#00A651", outline="")

    def _draw_svg_path_fill(self, canvas: tk.Canvas, path_d: str, transform, fill_color: str, curve_steps: int = 10) -> None:
        points = self._svg_path_to_points(path_d, curve_steps)
        if len(points) < 3:
            return
        flat: list[float] = []
        for x, y in points:
            tx, ty = transform(x, y)
            flat.extend([tx, ty])
        canvas.create_polygon(*flat, fill=fill_color, outline="")

    def _svg_path_to_points(self, path_d: str, curve_steps: int = 10) -> list[tuple[float, float]]:
        tokens = re.findall(r"[A-Za-z]|-?\d*\.?\d+", path_d)
        i = 0
        cmd = ""
        cx = 0.0
        cy = 0.0
        sx = 0.0
        sy = 0.0
        out: list[tuple[float, float]] = []

        def is_num(tok: str) -> bool:
            return bool(re.fullmatch(r"-?\d*\.?\d+", tok))

        def cubic(p0, p1, p2, p3, steps: int) -> list[tuple[float, float]]:
            pts: list[tuple[float, float]] = []
            for step in range(1, steps + 1):
                t = step / float(steps)
                mt = 1.0 - t
                x = mt * mt * mt * p0[0] + 3.0 * mt * mt * t * p1[0] + 3.0 * mt * t * t * p2[0] + t * t * t * p3[0]
                y = mt * mt * mt * p0[1] + 3.0 * mt * mt * t * p1[1] + 3.0 * mt * t * t * p2[1] + t * t * t * p3[1]
                pts.append((x, y))
            return pts

        while i < len(tokens):
            tok = tokens[i]
            if re.fullmatch(r"[A-Za-z]", tok):
                cmd = tok
                i += 1
            if cmd == "":
                break
            if cmd == "M":
                if i + 1 >= len(tokens):
                    break
                cx = float(tokens[i]); cy = float(tokens[i + 1]); i += 2
                sx, sy = cx, cy
                out.append((cx, cy))
                cmd = "L"
            elif cmd == "m":
                if i + 1 >= len(tokens):
                    break
                cx += float(tokens[i]); cy += float(tokens[i + 1]); i += 2
                sx, sy = cx, cy
                out.append((cx, cy))
                cmd = "l"
            elif cmd == "L":
                while i + 1 < len(tokens) and is_num(tokens[i]) and is_num(tokens[i + 1]):
                    cx = float(tokens[i]); cy = float(tokens[i + 1]); i += 2
                    out.append((cx, cy))
            elif cmd == "l":
                while i + 1 < len(tokens) and is_num(tokens[i]) and is_num(tokens[i + 1]):
                    cx += float(tokens[i]); cy += float(tokens[i + 1]); i += 2
                    out.append((cx, cy))
            elif cmd == "H":
                while i < len(tokens) and is_num(tokens[i]):
                    cx = float(tokens[i]); i += 1
                    out.append((cx, cy))
            elif cmd == "h":
                while i < len(tokens) and is_num(tokens[i]):
                    cx += float(tokens[i]); i += 1
                    out.append((cx, cy))
            elif cmd == "V":
                while i < len(tokens) and is_num(tokens[i]):
                    cy = float(tokens[i]); i += 1
                    out.append((cx, cy))
            elif cmd == "v":
                while i < len(tokens) and is_num(tokens[i]):
                    cy += float(tokens[i]); i += 1
                    out.append((cx, cy))
            elif cmd == "C":
                while i + 5 < len(tokens) and all(is_num(tokens[i + j]) for j in range(6)):
                    p0 = (cx, cy)
                    p1 = (float(tokens[i]), float(tokens[i + 1]))
                    p2 = (float(tokens[i + 2]), float(tokens[i + 3]))
                    p3 = (float(tokens[i + 4]), float(tokens[i + 5]))
                    i += 6
                    out.extend(cubic(p0, p1, p2, p3, curve_steps))
                    cx, cy = p3
            elif cmd == "c":
                while i + 5 < len(tokens) and all(is_num(tokens[i + j]) for j in range(6)):
                    p0 = (cx, cy)
                    p1 = (cx + float(tokens[i]), cy + float(tokens[i + 1]))
                    p2 = (cx + float(tokens[i + 2]), cy + float(tokens[i + 3]))
                    p3 = (cx + float(tokens[i + 4]), cy + float(tokens[i + 5]))
                    i += 6
                    out.extend(cubic(p0, p1, p2, p3, curve_steps))
                    cx, cy = p3
            elif cmd in ("Z", "z"):
                out.append((sx, sy))
            else:
                i += 1
        return out

    def _setup_dnd(self) -> None:
        if DND_FILES is None:
            self.status_var.set("Ready. Drag/drop optional (install tkinterdnd2).")
            return
        try:
            # Register multiple targets; some environments only dispatch on widgets.
            self.drop_target_register(DND_FILES)
            self.listbox.drop_target_register(DND_FILES)
            self.canvas.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
            self.listbox.dnd_bind("<<Drop>>", self._on_drop)
            self.canvas.dnd_bind("<<Drop>>", self._on_drop)
            self.status_var.set("Ready. Drag/drop enabled.")
        except Exception:
            self.status_var.set("Ready. Drag/drop failed to initialize; using file picker.")

    def _on_drop(self, event) -> None:  # pragma: no cover
        paths = self._expand_paths_from_input(list(self.tk.splitlist(event.data)))
        self._ingest_paths(paths)

    def _add_images(self) -> None:
        paths = filedialog.askopenfilenames(title="Select PNG images", filetypes=[("PNG files", "*.png")])
        self._ingest_paths(self._expand_paths_from_input(list(paths)))

    def _expand_paths_from_input(self, raw_paths: list[str]) -> list[str]:
        expanded: list[str] = []
        seen: set[str] = set()
        for raw in raw_paths:
            if raw is None:
                continue
            p = str(raw).strip().strip('"').strip("{}")
            if p == "":
                continue
            path_obj = Path(p)
            if path_obj.is_dir():
                candidates = sorted(path_obj.glob("*.png"))
            else:
                candidates = [path_obj]
            for candidate in candidates:
                candidate_str = str(candidate)
                if not candidate_str.lower().endswith(".png"):
                    continue
                key = os.path.normcase(candidate_str)
                if key in seen:
                    continue
                seen.add(key)
                expanded.append(candidate_str)
        return expanded

    def _ingest_paths(self, paths: list[str]) -> None:
        if not paths:
            return
        existing = {os.path.normcase(it.source_path) for it in self.items}
        added = 0
        for path in paths:
            if os.path.normcase(path) in existing:
                continue
            try:
                item = SpriteImageItem.from_path(path)
                self._auto_align_item_guides(item)
                self.items.append(item)
                existing.add(os.path.normcase(path))
                added += 1
            except Exception as exc:
                self.status_var.set(f"Skip {Path(path).name}: {exc}")
        if added > 0:
            self.active_idx = len(self.items) - 1
            self._apply_default_pack_naming(paths)
            self._refresh_image_list()
            self._load_active_item_fields()
            self.status_var.set(f"Added {added} image(s).")
            self._request_render()

    def _auto_align_current(self) -> None:
        item = self._active_item()
        if item is None:
            return
        self._auto_align_item_guides(item)
        self._load_active_item_fields()
        self._request_render()
        self.status_var.set("Auto-aligned current image guides.")

    def _auto_align_all(self) -> None:
        if not self.items:
            return
        for item in self.items:
            self._auto_align_item_guides(item)
        self._load_active_item_fields()
        self._request_render()
        self.status_var.set(f"Auto-aligned guides for {len(self.items)} image(s).")

    def _auto_align_item_guides(self, item: SpriteImageItem) -> None:
        img = item.source_rgba()
        alpha = img.getchannel("A")
        pix = alpha.load()
        w, h = alpha.size
        if w <= 0 or h <= 0:
            return
        threshold = EXPORT_ALPHA_TRIM_THRESHOLD

        bottom_y = -1
        for y in range(h - 1, -1, -1):
            found = False
            for x in range(w):
                if pix[x, y] >= threshold:
                    found = True
                    break
            if found:
                bottom_y = y
                break
        if bottom_y < 0:
            return

        bottom_xs = [x for x in range(w) if pix[x, bottom_y] >= threshold]
        if bottom_xs:
            left_bottom = min(bottom_xs)
            right_bottom = max(bottom_xs)
            center = (left_bottom + right_bottom + 1.0) * 0.5
        else:
            center = w * 0.5
            left_bottom = 0
            right_bottom = w - 1

        left_guide = float(left_bottom)
        right_guide = float(right_bottom)
        left_auto, right_auto = self._detect_plate_side_edges(alpha, threshold)
        if left_auto is not None:
            left_guide = left_auto
        if right_auto is not None:
            right_guide = right_auto

        if right_guide <= left_guide:
            left_guide = float(left_bottom)
            right_guide = float(max(left_bottom + 1, right_bottom))

        item.guide_center = center
        item.guide_left = left_guide
        item.guide_right = right_guide
        item.baseline_y = float(h)

    def _detect_plate_side_edges(self, alpha: Image.Image, threshold: int) -> tuple[Optional[float], Optional[float]]:
        pix = alpha.load()
        w, h = alpha.size
        if w <= 0 or h <= 0:
            return (None, None)

        bottom_y = -1
        for y in range(h - 1, -1, -1):
            has_opaque = False
            for x in range(w):
                if pix[x, y] >= threshold:
                    has_opaque = True
                    break
            if has_opaque:
                bottom_y = y
                break
        if bottom_y < 0:
            return (None, None)

        row_bounds: dict[int, tuple[int, int]] = {}
        for y in range(h):
            left = None
            right = None
            for x in range(w):
                if pix[x, y] >= threshold:
                    left = x
                    break
            if left is None:
                continue
            for x in range(w - 1, -1, -1):
                if pix[x, y] >= threshold:
                    right = x
                    break
            if right is None:
                continue
            row_bounds[y] = (left, right)

        def outside_transparent_run_down(x: int, start_y: int) -> int:
            if x < 0 or x >= w:
                return h - start_y
            run = 0
            y = start_y
            while y < h and pix[x, y] < threshold:
                run += 1
                y += 1
            return run

        def find_vertical_side_edge(is_left: bool) -> Optional[float]:
            min_y = AUTO_EDGE_OPAQUE_RUN_PX - 1
            for y in range(bottom_y, min_y - 1, -1):
                xs: list[int] = []
                ok = True
                for k in range(AUTO_EDGE_OPAQUE_RUN_PX):
                    yy = y - k
                    bounds = row_bounds.get(yy, None)
                    if bounds is None:
                        ok = False
                        break
                    xs.append(bounds[0] if is_left else bounds[1])
                if not ok:
                    continue
                if max(xs) - min(xs) > 1:
                    continue
                edge_x = int(round(sum(xs) / float(len(xs))))
                outside_x = edge_x - 1 if is_left else edge_x + 1
                if outside_transparent_run_down(outside_x, y) >= AUTO_SIDE_TRANSPARENCY_PX:
                    return float(edge_x)
            return None

        return (find_vertical_side_edge(True), find_vertical_side_edge(False))

    def _apply_default_pack_naming(self, incoming_paths: list[str]) -> None:
        if not incoming_paths:
            return
        first = Path(incoming_paths[0])
        file_stem = first.stem.strip()
        parent_name = ""
        try:
            common_parent = Path(os.path.commonpath([str(Path(p).parent) for p in incoming_paths]))
            parent_name = common_parent.name.strip()
        except Exception:
            parent_name = first.parent.name.strip() if first.parent else ""
        guess_name = parent_name or file_stem or "New Model"
        guess_id = _normalize_id(guess_name)

        id_is_default = self.pack_meta.id.strip() in ("", "new_model")
        name_is_default = self.pack_meta.name.strip().lower() in ("", "new model")

        if id_is_default:
            self.pack_meta.id = guess_id
            self.meta_vars["id"].set(self.pack_meta.id)
        if name_is_default:
            self.pack_meta.name = guess_name
            self.meta_vars["name"].set(self.pack_meta.name)
        self._sync_zip_name_to_id()

    def _refresh_image_list(self) -> None:
        self.listbox.delete(0, tk.END)
        for idx, item in enumerate(self.items):
            self.listbox.insert(tk.END, f"{idx + 1}. {item.label()}")
        if self.active_idx is not None and 0 <= self.active_idx < len(self.items):
            self.listbox.selection_set(self.active_idx)
            self.listbox.see(self.active_idx)

    def _on_list_select(self, _event=None) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        self._apply_image_fields_to_index(self.active_idx)
        self.active_idx = int(sel[0])
        self._load_active_item_fields()
        self._request_render()

    def _active_item(self) -> Optional[SpriteImageItem]:
        if self.active_idx is None or not (0 <= self.active_idx < len(self.items)):
            return None
        return self.items[self.active_idx]

    def _load_active_item_fields(self) -> None:
        item = self._active_item()
        if item is None:
            for v in self.image_vars.values():
                v.set("")
            self.metrics_label.config(text="")
            return
        self._suppress_image_apply = True
        try:
            self.image_vars["fit_mode"].set(item.fit_mode)
            self.image_vars["target_span"].set(f"{item.target_span_px:.2f}")
            self.image_vars["guide_left"].set(f"{item.guide_left:.2f}")
            self.image_vars["guide_center"].set(f"{item.guide_center:.2f}")
            self.image_vars["guide_right"].set(f"{item.guide_right:.2f}")
            self.image_vars["baseline_y"].set(f"{item.baseline_y:.2f}")
            self.image_vars["offset_x"].set(f"{item.offset_x:.2f}")
            self.image_vars["offset_y"].set(f"{item.offset_y:.2f}")
        finally:
            self._suppress_image_apply = False
        self._update_metrics(item)

    def _apply_image_fields(self) -> None:
        self._apply_image_fields_to_index(self.active_idx)

    def _tiles_from_preset(self, preset: str) -> Optional[tuple[int, int]]:
        name = (preset or "").strip().lower()
        if name == "full_2x2":
            return (2, 2)
        m = re.fullmatch(r"(\d+)x(\d+)", name)
        if m is None:
            return None
        return (max(1, int(m.group(1))), max(1, int(m.group(2))))

    def _apply_image_fields_to_index(self, idx: Optional[int]) -> None:
        if self._suppress_image_apply:
            return
        if idx is None or not (0 <= idx < len(self.items)):
            return
        item = self.items[idx]
        if item is None:
            return
        item.fit_mode = self.image_vars["fit_mode"].get().strip() or "full_2x2"
        item.target_span_px = max(1.0, _safe_float(self.image_vars["target_span"].get(), item.target_span_px))
        item.guide_left = min(max(0.0, _safe_float(self.image_vars["guide_left"].get(), item.guide_left)), float(item.width))
        item.guide_center = min(max(0.0, _safe_float(self.image_vars["guide_center"].get(), item.guide_center)), float(item.width))
        item.guide_right = min(max(item.guide_left + 1.0, _safe_float(self.image_vars["guide_right"].get(), item.guide_right)), float(item.width))
        item.baseline_y = min(max(0.0, _safe_float(self.image_vars["baseline_y"].get(), item.baseline_y)), float(item.height))
        item.offset_x = _safe_float(self.image_vars["offset_x"].get(), item.offset_x)
        item.offset_y = _safe_float(self.image_vars["offset_y"].get(), item.offset_y)
        if item.fit_mode in FIT_MODE_PRESETS and FIT_MODE_PRESETS[item.fit_mode] is not None:
            item.target_span_px = float(FIT_MODE_PRESETS[item.fit_mode])
        tiles = self._tiles_from_preset(item.fit_mode)
        if tiles is not None and idx == 0:
            self.pack_meta.tiles_x = tiles[0]
            self.pack_meta.tiles_y = tiles[1]
            self.meta_vars["tiles_x"].set(str(tiles[0]))
            self.meta_vars["tiles_y"].set(str(tiles[1]))
        self._load_active_item_fields()
        self._request_render()

    def _apply_pack_metadata_fields(self) -> None:
        self.pack_meta.id = _normalize_id(self.meta_vars["id"].get() or self.pack_meta.id)
        self.pack_meta.name = self.meta_vars["name"].get().strip() or self.pack_meta.name
        self.pack_meta.set_id = self.meta_vars["set_id"].get().strip()
        self.pack_meta.category = _normalize_category(self.meta_vars["category"].get())
        self.pack_meta.theme = self.meta_vars["theme"].get().strip()
        self.pack_meta.tiles_x = max(1, _safe_int(self.meta_vars["tiles_x"].get(), self.pack_meta.tiles_x))
        self.pack_meta.tiles_y = max(1, _safe_int(self.meta_vars["tiles_y"].get(), self.pack_meta.tiles_y))
        self.pack_meta.variant_group = self.meta_vars["variant_group"].get().strip()
        self.pack_meta.variant_label = self.meta_vars["variant_label"].get().strip()
        self.pack_meta.group_label = self.meta_vars["group_label"].get().strip()
        self.pack_meta.manufacturer = self.meta_vars["manufacturer"].get().strip()
        self.pack_meta.link = self.meta_vars["link"].get().strip()
        self.pack_meta.instructions = self.meta_vars["instructions"].get().strip()
        self.pack_meta.notes = self.meta_vars["notes"].get().strip()
        offsets: dict[str, list[float]] = {}
        for idx in range(4):
            ox = _safe_float(self.offset_vars[f"offset_{idx}_x"].get(), 0.0)
            oy = _safe_float(self.offset_vars[f"offset_{idx}_y"].get(), 0.0)
            offsets[str(idx)] = [ox, oy]
        self.pack_meta.offsets = offsets
        self.meta_vars["id"].set(self.pack_meta.id)
        self.meta_vars["category"].set(self.pack_meta.category)
        self._sync_zip_name_to_id()

    def _move_up(self) -> None:
        item = self._active_item()
        if item is None or self.active_idx in (None, 0):
            return
        idx = self.active_idx
        assert idx is not None
        self.items[idx - 1], self.items[idx] = self.items[idx], self.items[idx - 1]
        self.active_idx = idx - 1
        self._refresh_image_list()
        self._request_render()

    def _move_down(self) -> None:
        item = self._active_item()
        if item is None or self.active_idx is None or self.active_idx >= len(self.items) - 1:
            return
        idx = self.active_idx
        self.items[idx + 1], self.items[idx] = self.items[idx], self.items[idx + 1]
        self.active_idx = idx + 1
        self._refresh_image_list()
        self._request_render()

    def _remove_selected(self) -> None:
        if self.active_idx is None:
            return
        del self.items[self.active_idx]
        if not self.items:
            self.active_idx = None
        else:
            self.active_idx = min(self.active_idx, len(self.items) - 1)
        self._refresh_image_list()
        self._load_active_item_fields()
        self._request_render()

    def _on_zoom_slider(self, _event=None) -> None:
        old_zoom = self.preview_zoom
        new_zoom = self._clamp_zoom_for_active(float(self.zoom_var.get()) / 100.0)
        self._apply_zoom_at_canvas_point(new_zoom, self._last_canvas_mouse[0], self._last_canvas_mouse[1], old_zoom)
        self.preview_zoom = new_zoom
        self.zoom_var.set(self.preview_zoom * 100.0)
        self._schedule_zoom_render()

    def _on_mousewheel_zoom(self, event) -> None:
        self._last_canvas_mouse = (float(event.x), float(event.y))
        old_zoom = self.preview_zoom
        delta = 0.06 if event.delta > 0 else -0.06
        new_zoom = self._clamp_zoom_for_active(self.preview_zoom + delta)
        self._apply_zoom_at_canvas_point(new_zoom, float(event.x), float(event.y), old_zoom)
        self.preview_zoom = new_zoom
        self.zoom_var.set(self.preview_zoom * 100.0)
        self._schedule_zoom_render()

    def _apply_zoom_at_canvas_point(self, new_zoom: float, canvas_x: float, canvas_y: float, old_zoom: Optional[float] = None) -> None:
        item = self._active_item()
        if item is None:
            return
        if old_zoom is None:
            old_zoom = self.preview_zoom
        if old_zoom <= 0 or new_zoom <= 0:
            return

        cw = max(1.0, float(self.canvas.winfo_width()))
        ch = max(1.0, float(self.canvas.winfo_height()))
        old_disp_w = float(item.width) * old_zoom
        old_disp_h = float(item.height) * old_zoom
        old_ox = (cw - old_disp_w) * 0.5 + self.pan_x
        old_oy = (ch - old_disp_h) * 0.5 + self.pan_y

        # Keep the same image-space point under the cursor while zooming.
        ix = (canvas_x - old_ox) / old_zoom
        iy = (canvas_y - old_oy) / old_zoom

        new_disp_w = float(item.width) * new_zoom
        new_disp_h = float(item.height) * new_zoom
        new_pan_x = canvas_x - ix * new_zoom - (cw - new_disp_w) * 0.5
        new_pan_y = canvas_y - iy * new_zoom - (ch - new_disp_h) * 0.5
        self.pan_x = new_pan_x
        self.pan_y = new_pan_y

    def _schedule_zoom_render(self) -> None:
        # Coalesce rapid wheel/slider events so we render at controlled cadence.
        if self._zoom_render_after_id is not None:
            self.after_cancel(self._zoom_render_after_id)
        self._zoom_render_after_id = self.after(25, self._flush_zoom_render)

    def _flush_zoom_render(self) -> None:
        self._zoom_render_after_id = None
        self._request_render()

    def _schedule_pan_render(self) -> None:
        if self._pan_render_after_id is not None:
            return
        # Refresh viewport tile while panning without flooding render queue.
        self._pan_render_after_id = self.after(33, self._flush_pan_render)

    def _flush_pan_render(self) -> None:
        self._pan_render_after_id = None
        self._request_render()

    def _on_canvas_motion(self, event) -> None:
        self._last_canvas_mouse = (float(event.x), float(event.y))

    def _clamp_zoom_for_active(self, requested_zoom: float) -> float:
        return min(PREVIEW_MAX_ZOOM_FALLBACK, max(PREVIEW_MIN_ZOOM, requested_zoom))

    def _request_render(self) -> None:
        if self._render_pending:
            return
        self._render_pending = True
        self.after_idle(self._enqueue_preview_job)

    def _enqueue_preview_job(self) -> None:
        self._render_pending = False
        item = self._active_item()
        if item is None:
            self.canvas.delete("all")
            self._canvas_image_id = None
            self._canvas_rect_id = None
            self._canvas_left_id = None
            self._canvas_center_id = None
            self._canvas_right_id = None
            self.canvas.create_text(
                20,
                20,
                anchor="nw",
                fill="#9a9a9a",
                text="Add PNGs.\nOrder is used for export naming (1,2,3,4...).",
            )
            self._record_fps(0.001)
            return

        cw = max(1, self.canvas.winfo_width())
        ch = max(1, self.canvas.winfo_height())
        self._preview_job_id += 1
        job = {
            "job_id": self._preview_job_id,
            "source": item.source_rgba(),
            "img_w": item.width,
            "img_h": item.height,
            "zoom": self.preview_zoom,
            "pan_x": self.pan_x,
            "pan_y": self.pan_y,
            "canvas_w": cw,
            "canvas_h": ch,
            "guide_left": item.guide_left,
            "guide_center": item.guide_center,
            "guide_right": item.guide_right,
            "baseline_y": item.baseline_y,
        }
        with self._preview_lock:
            self._preview_next_job = job
        self._preview_event.set()

    def _preview_worker_loop(self) -> None:
        while not self._preview_shutdown:
            self._preview_event.wait()
            if self._preview_shutdown:
                break
            while True:
                with self._preview_lock:
                    job = self._preview_next_job
                    self._preview_next_job = None
                    if job is None:
                        self._preview_event.clear()
                        break
                start = time.perf_counter()
                zoom = float(job["zoom"])
                disp_w = max(1, int(round(int(job["img_w"]) * zoom)))
                disp_h = max(1, int(round(int(job["img_h"]) * zoom)))
                ox = (int(job["canvas_w"]) - disp_w) * 0.5 + float(job["pan_x"])
                oy = (int(job["canvas_h"]) - disp_h) * 0.5 + float(job["pan_y"])

                cw = int(job["canvas_w"])
                ch = int(job["canvas_h"])
                src_w = int(job["img_w"])
                src_h = int(job["img_h"])

                # Render only the visible source region for fast zoom/pan on large images.
                src_x0 = max(0.0, (-ox) / zoom)
                src_y0 = max(0.0, (-oy) / zoom)
                src_x1 = min(float(src_w), (cw - ox) / zoom)
                src_y1 = min(float(src_h), (ch - oy) / zoom)

                if src_x1 <= src_x0 or src_y1 <= src_y0:
                    preview = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
                    draw_x = 0.0
                    draw_y = 0.0
                else:
                    crop_l = int(math.floor(src_x0))
                    crop_t = int(math.floor(src_y0))
                    crop_r = int(math.ceil(src_x1))
                    crop_b = int(math.ceil(src_y1))
                    crop_l = max(0, min(src_w, crop_l))
                    crop_t = max(0, min(src_h, crop_t))
                    crop_r = max(crop_l + 1, min(src_w, crop_r))
                    crop_b = max(crop_t + 1, min(src_h, crop_b))
                    crop = job["source"].crop((crop_l, crop_t, crop_r, crop_b))
                    target_w = max(1, int(round((crop_r - crop_l) * zoom)))
                    target_h = max(1, int(round((crop_b - crop_t) * zoom)))
                    # Pixel-precise mode when zoomed in, smoother downsample when zoomed out.
                    if zoom >= 1.0:
                        resample = Image.Resampling.NEAREST
                    else:
                        resample = Image.Resampling.BILINEAR
                    preview = crop.resize((target_w, target_h), resample)
                    draw_x = ox + crop_l * zoom
                    draw_y = oy + crop_t * zoom

                result = {
                    "job_id": int(job["job_id"]),
                    "preview": preview,
                    "disp_w": disp_w,
                    "disp_h": disp_h,
                    "ox": ox,
                    "oy": oy,
                    "draw_x": draw_x,
                    "draw_y": draw_y,
                    "lx": ox + float(job["guide_left"]) * zoom,
                    "cx": ox + float(job["guide_center"]) * zoom,
                    "rx": ox + float(job["guide_right"]) * zoom,
                    "frame_dt": time.perf_counter() - start,
                }
                with self._preview_lock:
                    self._preview_latest_result = result

    def _poll_preview_results(self) -> None:
        result: Optional[dict] = None
        with self._preview_lock:
            if (
                self._preview_latest_result is not None
                and self._preview_latest_result["job_id"] > self._preview_applied_job_id
            ):
                result = self._preview_latest_result
        if result is not None:
            self._apply_preview_result(result)
        if not self._preview_shutdown:
            self.after(16, self._poll_preview_results)

    def _apply_preview_result(self, result: dict) -> None:
        self._preview_applied_job_id = int(result["job_id"])
        self.canvas.delete("all")
        self.preview_photo = ImageTk.PhotoImage(result["preview"])
        self._scene_ox = float(result["ox"])
        self._scene_oy = float(result["oy"])
        self._scene_disp_w = int(result["disp_w"])
        self._scene_disp_h = int(result["disp_h"])
        draw_x = float(result.get("draw_x", self._scene_ox))
        draw_y = float(result.get("draw_y", self._scene_oy))
        self._canvas_image_id = self.canvas.create_image(
            draw_x,
            draw_y,
            image=self.preview_photo,
            anchor="nw",
        )
        self._canvas_rect_id = self.canvas.create_rectangle(
            self._scene_ox,
            self._scene_oy,
            self._scene_ox + self._scene_disp_w,
            self._scene_oy + self._scene_disp_h,
            outline="#4f4f4f",
        )
        self._canvas_left_id = self.canvas.create_line(0, 0, 0, 0, fill="#32cd32", width=2)
        self._canvas_center_id = self.canvas.create_line(0, 0, 0, 0, fill="#ff3b30", width=2)
        self._canvas_right_id = self.canvas.create_line(0, 0, 0, 0, fill="#32cd32", width=2)
        self._update_overlay_positions()
        self._record_fps(float(result["frame_dt"]))
        item = self._active_item()
        if item is not None:
            self._update_metrics(item)

    def _update_overlay_positions(self) -> None:
        item = self._active_item()
        if item is None:
            return
        if self._canvas_left_id is None or self._canvas_center_id is None or self._canvas_right_id is None:
            return
        zoom = self.preview_zoom
        oy = self._scene_oy
        by = self._scene_oy + self._scene_disp_h
        lx = self._scene_ox + item.guide_left * zoom
        cx = self._scene_ox + item.guide_center * zoom
        rx = self._scene_ox + item.guide_right * zoom
        self.canvas.coords(self._canvas_left_id, lx, oy, lx, by)
        self.canvas.coords(self._canvas_center_id, cx, oy, cx, by)
        self.canvas.coords(self._canvas_right_id, rx, oy, rx, by)

    def _pan_scene(self, dx: float, dy: float) -> None:
        if dx == 0 and dy == 0:
            return
        if self._canvas_image_id is None:
            return
        self.pan_x += dx
        self.pan_y += dy
        self._scene_ox += dx
        self._scene_oy += dy
        self.canvas.move("all", dx, dy)

    def _record_fps(self, frame_dt: float) -> None:
        if frame_dt <= 0:
            return
        now = time.perf_counter()
        self._fps_times.append(now)
        if len(self._fps_times) >= 2:
            elapsed = self._fps_times[-1] - self._fps_times[0]
            fps = (len(self._fps_times) - 1) / elapsed if elapsed > 0 else 0.0
            self._fps_text.set(f"FPS: {fps:5.1f}   frame {frame_dt*1000:5.1f} ms")
        else:
            self._fps_text.set(f"FPS: --   frame {frame_dt*1000:5.1f} ms")

    def _canvas_to_image_xy(self, x: float, y: float) -> tuple[float, float]:
        item = self._active_item()
        if item is None:
            return (0.0, 0.0)
        zoom = self.preview_zoom
        disp_w = item.width * zoom
        disp_h = item.height * zoom
        cw = max(1, self.canvas.winfo_width())
        ch = max(1, self.canvas.winfo_height())
        ox = (cw - disp_w) * 0.5 + self.pan_x
        oy = (ch - disp_h) * 0.5 + self.pan_y
        return ((x - ox) / zoom, (y - oy) / zoom)

    def _on_canvas_press(self, event) -> None:
        item = self._active_item()
        if item is None:
            return
        ix, _iy = self._canvas_to_image_xy(event.x, event.y)
        hit = 10.0 / max(self.preview_zoom, 0.08)
        if abs(ix - item.guide_left) <= hit:
            self.drag_mode = "guide_left"
        elif abs(ix - item.guide_center) <= hit:
            self.drag_mode = "guide_center"
        elif abs(ix - item.guide_right) <= hit:
            self.drag_mode = "guide_right"
        else:
            self.drag_mode = "pan"
            self._pan_last = (event.x, event.y)

    def _on_canvas_drag(self, event) -> None:
        item = self._active_item()
        if item is None or self.drag_mode is None:
            return
        ix, _iy = self._canvas_to_image_xy(event.x, event.y)
        if self.drag_mode == "guide_left":
            item.guide_left = min(max(0.0, ix), item.guide_right - 1.0)
            self.image_vars["guide_left"].set(f"{item.guide_left:.2f}")
            self._update_overlay_positions()
        elif self.drag_mode == "guide_center":
            item.guide_center = min(max(0.0, ix), float(item.width))
            self.image_vars["guide_center"].set(f"{item.guide_center:.2f}")
            self._update_overlay_positions()
        elif self.drag_mode == "guide_right":
            item.guide_right = max(min(float(item.width), ix), item.guide_left + 1.0)
            self.image_vars["guide_right"].set(f"{item.guide_right:.2f}")
            self._update_overlay_positions()
        else:
            last_x, last_y = self._pan_last
            self._pan_scene(event.x - last_x, event.y - last_y)
            self._pan_last = (event.x, event.y)
            self._schedule_pan_render()

    def _on_canvas_release(self, _event) -> None:
        if self.drag_mode == "pan":
            self._request_render()
        self.drag_mode = None

    def _on_close(self) -> None:
        if self._zoom_render_after_id is not None:
            try:
                self.after_cancel(self._zoom_render_after_id)
            except Exception:
                pass
            self._zoom_render_after_id = None
        if self._pan_render_after_id is not None:
            try:
                self.after_cancel(self._pan_render_after_id)
            except Exception:
                pass
            self._pan_render_after_id = None
        self._preview_shutdown = True
        self._preview_event.set()
        self.destroy()

    def _update_metrics(self, item: SpriteImageItem) -> None:
        self.metrics_label.config(
            text=(
                f"Measured span: {item.measured_span():.2f}px   "
                f"Target span: {item.effective_target_span():.2f}px   "
                f"Scale: {item.scale_factor():.4f}x"
            )
        )

    def _image_ext(self) -> str:
        return "webp" if self.export_format_var.get().strip().lower() == "webp" else "png"

    def _save_encoded(self, image: Image.Image, path: Path) -> None:
        ext = path.suffix.lower()
        if ext == ".webp":
            save_kwargs = self._webp_save_kwargs()
            image.save(path, format="WEBP", **save_kwargs)
        else:
            image.save(path, format="PNG")

    def _encode_bytes(self, image: Image.Image, ext: str) -> bytes:
        b = BytesIO()
        if ext == "webp":
            save_kwargs = self._webp_save_kwargs()
            image.save(b, format="WEBP", **save_kwargs)
        else:
            image.save(b, format="PNG")
        return b.getvalue()

    def _webp_save_kwargs(self) -> dict:
        lossless = bool(self.webp_lossless_var.get())
        kwargs = {
            "method": 6,
            "lossless": lossless,
            # Keep RGB in transparent areas to avoid fringe/matte shifts.
            "exact": True,
            "alpha_quality": 100,
        }
        if not lossless:
            kwargs["quality"] = max(1, min(100, _safe_int(self.webp_quality_var.get(), 95)))
        return kwargs

    def _export_sprite_image(self, item: SpriteImageItem) -> Image.Image:
        src = item.source_rgba()
        target_span = max(1.0, item.effective_target_span())
        measured_span = max(1e-6, item.measured_span())
        target_span_i = int(round(target_span))

        # Strict guide-driven solve: choose integer width that makes the
        # scaled guide span land exactly on target integer pixels when possible.
        nominal_sw = max(1, _round_half_up(item.width * (target_span / measured_span)))
        best_sw = nominal_sw
        best_err = float("inf")
        for dsw in range(-128, 129):
            cand_sw = max(1, nominal_sw + dsw)
            cand_scale = float(cand_sw) / float(max(1, item.width))
            cand_span = measured_span * cand_scale
            cand_span_i = int(round(cand_span))
            err = abs(cand_span_i - target_span_i) * 1000.0 + abs(cand_span - target_span)
            if err < best_err:
                best_err = err
                best_sw = cand_sw
                if cand_span_i == target_span_i:
                    break

        sw = max(1, best_sw)
        scale_x = float(sw) / float(max(1, item.width))
        scale_y = scale_x
        sh = max(1, _round_half_up(item.height * scale_y))
        # Match Photoshop-style interpolation more closely for edge stability.
        scaled = src.resize((sw, sh), Image.Resampling.BICUBIC)

        # Match Photoshop action behavior by using actual sprite pixel bounds,
        # not the full canvas size, when computing output dimensions.
        # Ignore very faint antialias fringe when computing bounds so
        # bottom alignment matches Photoshop-like visual edges.
        alpha = scaled.getchannel("A")
        # Use a tighter threshold for bounds so tiny AA fringe does not inflate width.
        alpha_mask = alpha.point(lambda a: 255 if a >= EXPORT_BOUNDS_ALPHA_THRESHOLD else 0)
        alpha_bbox = alpha_mask.getbbox()
        if alpha_bbox is None:
            alpha_left = 0.0
            alpha_top = 0.0
            alpha_right = float(sw)
            alpha_bottom = float(sh)
        else:
            alpha_left = float(alpha_bbox[0])
            alpha_top = float(alpha_bbox[1])
            alpha_right = float(alpha_bbox[2])
            alpha_bottom = float(alpha_bbox[3])

        center_scaled = item.guide_center * scale_x
        baseline_scaled = item.baseline_y * scale_y

        dist_left = max(0.0, center_scaled - alpha_left)
        dist_right = max(0.0, alpha_right - center_scaled)
        half_w = max(dist_left, dist_right) + FIXED_PADDING_PX
        out_w = max(1, int(math.ceil(half_w * 2.0)))
        # Keep center anchor on an exact pixel to prevent 1px L/R wobble.
        if out_w % 2 != 0:
            out_w += 1

        # Vertical rule: 32px top padding, 0px bottom padding.
        # Visible sprite bounds must touch the output bottom edge.
        alpha_h = max(1.0, alpha_bottom - alpha_top)
        out_h = max(1, int(math.ceil(alpha_h + FIXED_PADDING_PX)))

        paste_x = _round_half_up(out_w * 0.5 - center_scaled)
        paste_y = _round_half_up(out_h - alpha_bottom)
        out = Image.new("RGBA", (out_w, out_h), (0, 0, 0, 0))
        out.alpha_composite(scaled, (paste_x, paste_y))
        return out

    def _build_pack_metadata_json(self) -> dict:
        self._apply_pack_metadata_fields()
        return self.pack_meta.to_dict()

    def _export_metadata_only(self) -> None:
        out_dir = filedialog.askdirectory(title="Choose folder to write metadata.json")
        if not out_dir:
            return
        target = Path(out_dir) / "metadata.json"
        try:
            with target.open("w", encoding="utf-8") as f:
                json.dump(self._build_pack_metadata_json(), f, indent=2, ensure_ascii=True)
        except Exception as exc:
            messagebox.showerror("Export metadata", f"Failed to write metadata.json:\n{exc}")
            return
        self.status_var.set(f"Metadata updated: {target}")

    def _export_folder(self) -> None:
        if not self.items:
            messagebox.showinfo("Export", "No images loaded.")
            return
        out_root = filedialog.askdirectory(title="Choose output root")
        if not out_root:
            return
        ext = self._image_ext()
        model_dir = Path(out_root) / _normalize_id(self.meta_vars["id"].get())
        model_dir.mkdir(parents=True, exist_ok=True)
        exported = 0
        errors: list[str] = []
        for idx, item in enumerate(self.items, start=1):
            try:
                out = self._export_sprite_image(item)
                self._save_encoded(out, model_dir / f"{idx}.{ext}")
                if idx == 1:
                    thumb = ImageOps.contain(out, (256, 256), Image.Resampling.LANCZOS)
                    self._save_encoded(thumb, model_dir / f"thumb.{ext}")
                exported += 1
            except Exception as exc:
                errors.append(f"{item.label()}: {exc}")

        with (model_dir / "metadata.json").open("w", encoding="utf-8") as f:
            json.dump(self._build_pack_metadata_json(), f, indent=2, ensure_ascii=True)
        self.status_var.set(f"Exported pack folder: {exported}/{len(self.items)} images.")
        if errors:
            messagebox.showwarning("Export warnings", "\n".join(errors[:10]))

    def _export_zip(self) -> None:
        if not self.items:
            messagebox.showinfo("Export", "No images loaded.")
            return
        out_zip = filedialog.asksaveasfilename(
            title="Save pack zip",
            filetypes=[("Zip files", "*.zip")],
            defaultextension=".zip",
            initialfile=self.zip_name_var.get().strip() or "sprite_pack.zip",
        )
        if not out_zip:
            return
        ext = self._image_ext()
        folder = _normalize_id(self.meta_vars["id"].get())
        exported = 0
        errors: list[str] = []
        with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for idx, item in enumerate(self.items, start=1):
                try:
                    out = self._export_sprite_image(item)
                    zf.writestr(f"{folder}/{idx}.{ext}", self._encode_bytes(out, ext))
                    if idx == 1:
                        thumb = ImageOps.contain(out, (256, 256), Image.Resampling.LANCZOS)
                        zf.writestr(f"{folder}/thumb.{ext}", self._encode_bytes(thumb, ext))
                    exported += 1
                except Exception as exc:
                    errors.append(f"{item.label()}: {exc}")
            zf.writestr(
                f"{folder}/metadata.json",
                json.dumps(self._build_pack_metadata_json(), indent=2, ensure_ascii=True).encode("utf-8"),
            )
        self.status_var.set(f"Exported pack zip: {exported}/{len(self.items)} images.")
        if errors:
            messagebox.showwarning("Export warnings", "\n".join(errors[:10]))


def main() -> int:
    app = SpritePipelineApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
