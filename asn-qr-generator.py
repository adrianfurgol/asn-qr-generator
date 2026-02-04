#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
ASN / Paperless QR Label Generator (PDF)

Workflow:
1) Select a sheet template (or Custom).
   - If a template is selected, layout is loaded and we jump directly to "Number of pages".
   - If Custom is selected, the user MUST enter all layout values (no defaults for layout).

Rules:
- Page size + margins define a fixed "sheetbox" (must NOT move).
- rows/cols define the grid; label size is computed automatically.
- Gaps: horizontal = left/right, vertical = up/down.
- QR is as large as possible:
    - fixed 0.5 mm padding top/bottom
    - width maximized while leaving space for text on the right
- Text uses the remaining space to the right of the QR.

Advanced submenu (optional):
- Debug frames
- X/Y offset (printer alignment)
- X/Y scale (drift correction), anchored at sheetbox top-left.
  Note: The RED sheetbox debug frame is NOT affected by offset/scale.

Requires:
  py -m pip install reportlab "qrcode[pil]" pillow
"""

import io
import sys
from dataclasses import dataclass
from typing import Optional, Tuple

import qrcode
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors


# --- Page presets (mm) ---
A4_MM = (210.0, 297.0)
LETTER_MM = (215.9, 279.4)  # 8.5" x 11"


# --- Templates ---
TEMPLATES = [
    {
        "key": "avery_l4731rev_25",
        "name": "Avery L4731REV-25 (A4)",
        "page_name": "A4",
        "page_w_mm": A4_MM[0],
        "page_h_mm": A4_MM[1],
        "margin_top_mm": 13.6,
        "margin_bottom_mm": 13.6,
        "margin_left_mm": 8.5,
        "margin_right_mm": 8.5,
        "rows": 27,
        "cols": 7,
        "gap_x_mm": 2.5,   # horizontal (left/right)
        "gap_y_mm": 0.0,   # vertical (up/down)
        "deadzone_left_mm": 1.0,
        "deadzone_right_mm": 0.0,
    }
]


@dataclass
class Config:
    # Page
    page_name: str
    page_w_mm: float
    page_h_mm: float

    # Output
    out_pdf: str

    # Margins (mm)
    margin_top_mm: float
    margin_bottom_mm: float
    margin_left_mm: float
    margin_right_mm: float

    # Grid
    rows: int
    cols: int
    pages: int

    # Gaps between labels (mm)
    gap_x_mm: float  # horizontal (left/right)
    gap_y_mm: float  # vertical (up/down)

    # Dead zones inside label (mm)
    deadzone_left_mm: float
    deadzone_right_mm: float

    # Code settings
    prefix: str
    start_number: int
    leading_zeros: int  # 0 = none

    # Advanced
    advanced_enabled: bool
    debug_label_frames: bool
    debug_sheetbox_frame: bool
    offset_x_mm: float
    offset_y_mm: float
    scale_x: float
    scale_y: float


# -----------------------------
# CLI helpers (English)
# -----------------------------
def ask_str(prompt: str, default: Optional[str] = None, required: bool = False) -> str:
    p = f"{prompt} [{default}]: " if default is not None else f"{prompt}: "
    while True:
        v = input(p).strip()
        if not v and default is not None and not required:
            return default
        if v:
            return v
        if required:
            print("This value is required.")
        else:
            print("Please enter a value.")


def ask_int(prompt: str, default: Optional[int] = None, min_value: Optional[int] = None, required: bool = False) -> int:
    p = f"{prompt} [{default}]: " if default is not None else f"{prompt}: "
    while True:
        raw = input(p).strip()
        if not raw:
            if default is not None and not required:
                val = default
            else:
                print("This value is required.")
                continue
        else:
            try:
                val = int(raw)
            except ValueError:
                print("Please enter a valid integer.")
                continue
        if min_value is not None and val < min_value:
            print(f"Value must be >= {min_value}.")
            continue
        return val


def ask_float(prompt: str, default: Optional[float] = None, min_value: Optional[float] = None, required: bool = False) -> float:
    p = f"{prompt} [{default}]: " if default is not None else f"{prompt}: "
    while True:
        raw = input(p).strip().replace(",", ".")
        if not raw:
            if default is not None and not required:
                val = float(default)
            else:
                print("This value is required.")
                continue
        else:
            try:
                val = float(raw)
            except ValueError:
                print("Please enter a valid number (e.g., 13.6).")
                continue
        if min_value is not None and val < min_value:
            print(f"Value must be >= {min_value}.")
            continue
        return val


def ask_yes_no(prompt: str, default: bool = True) -> bool:
    d = "y" if default else "n"
    raw = input(f"{prompt} (y/n) [{d}]: ").strip().lower()
    if not raw:
        return default
    return raw in ("y", "yes", "true", "1")


def ask_menu_choice(prompt: str, valid: Tuple[str, ...], default: Optional[str] = None, required: bool = False) -> str:
    while True:
        suffix = f" [default {default}]" if default is not None and not required else ""
        raw = input(f"{prompt}{suffix}: ").strip()
        if not raw:
            if default is not None and not required:
                return default
            if required:
                print("This value is required.")
                continue
        if raw in valid:
            return raw
        print(f"Please enter one of: {', '.join(valid)}")


# -----------------------------
# QR + formatting
# -----------------------------
def format_code(prefix: str, number: int, leading_zeros: int) -> Tuple[str, str]:
    if leading_zeros <= 0:
        num_part = str(number)
    else:
        num_part = f"{number:0{leading_zeros}d}"
    return f"{prefix}{num_part}", num_part


def make_qr_image(data: str, pixels: int = 1400) -> Image.Image:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=1,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    return img.resize((pixels, pixels), Image.NEAREST)


# -----------------------------
# Text fitting
# -----------------------------
def pick_font_size_to_fit(
    c: canvas.Canvas,
    text: str,
    font: str,
    max_w: float,
    max_size: float,
    min_size: float,
    step: float = 0.5,
) -> float:
    fs = max_size
    while fs >= min_size - 1e-9:
        if c.stringWidth(text, font, fs) <= max_w:
            return fs
        fs -= step
    return min_size


def draw_code_text(
    c: canvas.Canvas,
    x: float,
    y_center: float,
    w: float,
    h: float,
    full_code: str,
    prefix: str,
    num_part: str,
):
    # Keep text modest so the QR dominates.
    font = "Helvetica-Bold"

    max_fs = 9.0
    min_fs = 3.5

    fs = pick_font_size_to_fit(c, full_code, font, w, max_fs, min_fs, step=0.5)
    c.setFont(font, fs)
    if c.stringWidth(full_code, font, fs) <= w:
        c.drawString(x, y_center - (fs / 2.7), full_code)
        return

    # Fallback: 2 lines
    line1 = prefix
    line2 = num_part

    fs1 = pick_font_size_to_fit(c, line1, font, w, max_fs - 0.5, min_fs, step=0.5)
    fs2 = pick_font_size_to_fit(c, line2, font, w, max_fs, min_fs, step=0.5)
    fs_use = min(fs1, fs2)

    c.setFont(font, fs_use)
    line_gap = min(h * 0.32, fs_use * 1.2)
    c.drawString(x, y_center + (line_gap / 2.0) - (fs_use / 2.7), line1)
    c.drawString(x, y_center - (line_gap / 1.2) - (fs_use / 2.7), line2)


# -----------------------------
# Validation (layout math)
# -----------------------------
def validate_layout(cfg: Config) -> None:
    page_w_pt = cfg.page_w_mm * mm
    page_h_pt = cfg.page_h_mm * mm

    mt = cfg.margin_top_mm * mm
    mb = cfg.margin_bottom_mm * mm
    ml = cfg.margin_left_mm * mm
    mr = cfg.margin_right_mm * mm

    sheet_w = page_w_pt - ml - mr
    sheet_h = page_h_pt - mt - mb
    if sheet_w <= 0 or sheet_h <= 0:
        raise ValueError("Margins are too large: sheetbox would be <= 0.")

    gap_x = cfg.gap_x_mm * mm
    gap_y = cfg.gap_y_mm * mm

    if cfg.cols > 1 and sheet_w - (cfg.cols - 1) * gap_x <= 0:
        raise ValueError("Horizontal gap too large: no width left for labels.")
    if cfg.rows > 1 and sheet_h - (cfg.rows - 1) * gap_y <= 0:
        raise ValueError("Vertical gap too large: no height left for labels.")

    label_w = (sheet_w - (cfg.cols - 1) * gap_x) / cfg.cols
    label_h = (sheet_h - (cfg.rows - 1) * gap_y) / cfg.rows
    if label_w <= 0 or label_h <= 0:
        raise ValueError("Computed label size is <= 0 (check rows/cols/gaps/margins).")

    dz_l = cfg.deadzone_left_mm * mm
    dz_r = cfg.deadzone_right_mm * mm
    if dz_l + dz_r >= label_w:
        raise ValueError("Dead zones too large: content width would be <= 0.")


# -----------------------------
# PDF generation
# -----------------------------
def generate_pdf(cfg: Config) -> None:
    page_w_pt = cfg.page_w_mm * mm
    page_h_pt = cfg.page_h_mm * mm

    c = canvas.Canvas(cfg.out_pdf, pagesize=(page_w_pt, page_h_pt))

    mt = cfg.margin_top_mm * mm
    mb = cfg.margin_bottom_mm * mm
    ml = cfg.margin_left_mm * mm
    mr = cfg.margin_right_mm * mm

    # Fixed sheetbox (NOT affected by offset/scale)
    sheet_x = ml
    sheet_y = mb
    sheet_w = page_w_pt - ml - mr
    sheet_h = page_h_pt - mt - mb
    sheet_top = sheet_y + sheet_h

    gap_x = cfg.gap_x_mm * mm   # horizontal (X)
    gap_y = cfg.gap_y_mm * mm   # vertical (Y)

    label_w = (sheet_w - (cfg.cols - 1) * gap_x) / cfg.cols
    label_h = (sheet_h - (cfg.rows - 1) * gap_y) / cfg.rows

    # Advanced transforms
    off_x = (cfg.offset_x_mm * mm) if cfg.advanced_enabled else 0.0
    off_y = (cfg.offset_y_mm * mm) if cfg.advanced_enabled else 0.0
    scale_x = cfg.scale_x if cfg.advanced_enabled else 1.0
    scale_y = cfg.scale_y if cfg.advanced_enabled else 1.0

    # Scale-sensitive internal geometry (so drift correction scales EVERYTHING)
    dz_l = (cfg.deadzone_left_mm * mm) * scale_x
    dz_r = (cfg.deadzone_right_mm * mm) * scale_x

    # Constants (scaled)
    qr_vpad = (0.5 * mm) * scale_y
    qr_text_gap = (0.6 * mm) * scale_x
    text_right_pad = (0.6 * mm) * scale_x
    min_text_w = (6.0 * mm) * scale_x

    per_page = cfg.rows * cfg.cols
    num = cfg.start_number

    # Debug styles
    if cfg.advanced_enabled and (cfg.debug_label_frames or cfg.debug_sheetbox_frame):
        c.setLineWidth(0.25)
        c.setStrokeColor(colors.lightgrey)

    for _ in range(cfg.pages):
        # RED sheetbox frame
        if cfg.advanced_enabled and cfg.debug_sheetbox_frame:
            # same anchor logic as the labels: scale around sheetbox top-left, then offset
            sheet_x_t = sheet_x + off_x
            sheet_y_t = (sheet_top - sheet_h * scale_y) + off_y
            sheet_w_t = sheet_w * scale_x
            sheet_h_t = sheet_h * scale_y

            c.setStrokeColor(colors.red)
            c.setLineWidth(0.25)
            c.rect(sheet_x_t, sheet_y_t, sheet_w_t, sheet_h_t, stroke=1, fill=0)

            # restore debug style for subsequent label frames
            c.setStrokeColor(colors.lightgrey)
            c.setLineWidth(0.25)


        for i in range(per_page):
            r = i // cfg.cols
            col = i % cfg.cols

            # Nominal bottom-left anchored to sheetbox
            x_nom = sheet_x + col * (label_w + gap_x)
            y_nom = sheet_top - (r + 1) * label_h - r * gap_y

            # Apply scale anchored at sheetbox top-left, then offset
            x = sheet_x + (x_nom - sheet_x) * scale_x + off_x
            y = sheet_top - (sheet_top - y_nom) * scale_y + off_y

            w = label_w * scale_x
            h = label_h * scale_y

            # Debug label frames (affected by offset/scale on purpose)
            if cfg.advanced_enabled and cfg.debug_label_frames:
                c.rect(x, y, w, h, stroke=1, fill=0)

            # Content box (dead zones left/right)
            content_x = x + dz_l
            content_w = w - dz_l - dz_r
            if content_w <= 0:
                raise ValueError("Content width is <= 0 (dead zones / scale).")

            full_code, num_part = format_code(cfg.prefix, num, cfg.leading_zeros)

            # QR size: vertical max based on label height with fixed padding
            qr_h_max = h - 2 * qr_vpad
            if qr_h_max <= 0:
                raise ValueError("Label height too small for 0.5mm QR top/bottom padding.")

            # QR width max: must leave space for text
            qr_w_max = content_w - qr_text_gap - min_text_w
            if qr_w_max <= 0:
                raise ValueError("Label too narrow for QR + text (reduce columns/deadzones/gaps).")

            qr_size = min(qr_h_max, qr_w_max)

            # Place QR left-aligned, exact padding top/bottom
            qr_x = content_x
            qr_y = y + qr_vpad

            img = make_qr_image(full_code, pixels=1400)
            bio = io.BytesIO()
            img.save(bio, format="PNG")
            bio.seek(0)
            c.drawImage(ImageReader(bio), qr_x, qr_y, width=qr_size, height=qr_size, mask=None)

            # Text uses remaining area to the right
            text_x = qr_x + qr_size + qr_text_gap
            text_w = (content_x + content_w) - text_x - text_right_pad
            text_w = max(1.0, text_w)
            y_center = y + h / 2.0

            draw_code_text(c, text_x, y_center, text_w, h, full_code, cfg.prefix, num_part)

            num += 1

        c.showPage()

    c.save()


# -----------------------------
# Setup / menus / confirmation
# -----------------------------
def select_template() -> Optional[dict]:
    print("\nSelect a sheet template:")
    for idx, t in enumerate(TEMPLATES, start=1):
        print(f"  {idx} - {t['name']}")
    print(f"  {len(TEMPLATES) + 1} - Custom")

    valid = tuple(str(i) for i in range(1, len(TEMPLATES) + 2))
    choice = ask_menu_choice("Enter your choice", valid, default="1", required=False)

    if choice == str(len(TEMPLATES) + 1):
        return None
    return TEMPLATES[int(choice) - 1]


def collect_custom_layout() -> dict:
    print("\nCustom layout selected.")
    print("You must enter ALL layout values (no defaults).")

    # Page layout menu (no defaults)
    print("\nSelect page layout:")
    print("  1 - A4")
    print("  2 - Letter")
    print("  3 - Custom")
    choice = ask_menu_choice("Enter 1, 2, or 3", ("1", "2", "3"), default=None, required=True)

    if choice == "1":
        page_name = "A4"
        page_w_mm, page_h_mm = A4_MM
    elif choice == "2":
        page_name = "Letter"
        page_w_mm, page_h_mm = LETTER_MM
    else:
        page_name = "Custom"
        page_w_mm = ask_float("Custom page width (mm)", default=None, min_value=1.0, required=True)
        page_h_mm = ask_float("Custom page height (mm)", default=None, min_value=1.0, required=True)

    print("\nMargins (mm) (define a fixed sheetbox):")
    margin_top = ask_float("Top margin", default=None, min_value=0.0, required=True)
    margin_bottom = ask_float("Bottom margin", default=None, min_value=0.0, required=True)
    margin_left = ask_float("Left margin", default=None, min_value=0.0, required=True)
    margin_right = ask_float("Right margin", default=None, min_value=0.0, required=True)

    print("\nGrid:")
    rows = ask_int("Rows", default=None, min_value=1, required=True)
    cols = ask_int("Columns", default=None, min_value=1, required=True)

    print("\nLabel gaps (mm):")
    gap_x = ask_float("Horizontal gap (left/right)", default=None, min_value=0.0, required=True)
    gap_y = ask_float("Vertical gap (up/down)", default=None, min_value=0.0, required=True)

    print("\nQR box dead zone inside label (mm):")
    dz_l = ask_float("Dead zone LEFT", default=None, min_value=0.0, required=True)
    dz_r = ask_float("Dead zone RIGHT", default=None, min_value=0.0, required=True)

    return {
        "page_name": page_name,
        "page_w_mm": page_w_mm,
        "page_h_mm": page_h_mm,
        "margin_top_mm": margin_top,
        "margin_bottom_mm": margin_bottom,
        "margin_left_mm": margin_left,
        "margin_right_mm": margin_right,
        "rows": rows,
        "cols": cols,
        "gap_x_mm": gap_x,
        "gap_y_mm": gap_y,
        "deadzone_left_mm": dz_l,
        "deadzone_right_mm": dz_r,
    }


def build_config_interactive() -> Config:
    print("\nASN QR Label Generator")
    print("---------------------------------")

    chosen_template = select_template()

    out_pdf = ask_str("Output PDF filename", "asn_labels.pdf")

    if chosen_template is not None:
        # Load template and jump directly to "Number of pages"
        layout = chosen_template
        print(f"\nLoaded template: {layout['name']}")
    else:
        # Custom: no layout defaults, user must enter everything
        layout = collect_custom_layout()

    # Jump here for template, continue here for custom as well
    pages = ask_int("\nNumber of pages", 1, 1, required=False)

    print("\nCode settings:")
    prefix = ask_str("Prefix", "ASN")
    start_number = ask_int("Start number", 1, 0)
    leading_zeros = ask_int("Leading zeros (0 = none)", 5, 0)  # NEW DEFAULT = 5

    # Advanced submenu
    advanced = ask_yes_no("\nDo you wish to set advanced options?", default=False)

    debug_labels = False
    debug_sheet = False
    offset_x = 0.0
    offset_y = 0.0
    scale_x = 1.0
    scale_y = 1.0

    if advanced:
        print("\nAdvanced options:")
        debug_labels = ask_yes_no("Draw label frames (debug)", default=False)
        debug_sheet = ask_yes_no("Draw sheetbox frame (debug, red)", default=False)

        print("\nPrinter offset (mm):")
        offset_x = ask_float("Offset X (right + / left -)", 0.0, None)
        offset_y = ask_float("Offset Y (up + / down -)", 0.0, None)

        print("\nPrinter scale (drift correction):")
        scale_x = ask_float("Scale X (default 1.0)", 1.0, 0.1)
        scale_y = ask_float("Scale Y (default 1.0)", 1.0, 0.1)

    cfg = Config(
        page_name=layout["page_name"],
        page_w_mm=float(layout["page_w_mm"]),
        page_h_mm=float(layout["page_h_mm"]),
        out_pdf=out_pdf,

        margin_top_mm=float(layout["margin_top_mm"]),
        margin_bottom_mm=float(layout["margin_bottom_mm"]),
        margin_left_mm=float(layout["margin_left_mm"]),
        margin_right_mm=float(layout["margin_right_mm"]),

        rows=int(layout["rows"]),
        cols=int(layout["cols"]),
        pages=pages,

        gap_x_mm=float(layout["gap_x_mm"]),
        gap_y_mm=float(layout["gap_y_mm"]),

        deadzone_left_mm=float(layout["deadzone_left_mm"]),
        deadzone_right_mm=float(layout["deadzone_right_mm"]),

        prefix=prefix,
        start_number=start_number,
        leading_zeros=leading_zeros,

        advanced_enabled=advanced,
        debug_label_frames=debug_labels,
        debug_sheetbox_frame=debug_sheet,
        offset_x_mm=offset_x,
        offset_y_mm=offset_y,
        scale_x=scale_x,
        scale_y=scale_y,
    )

    # Validate layout math early (so user sees it before generating)
    validate_layout(cfg)
    return cfg


def print_summary(cfg: Config) -> None:
    print("\n---------------------------------")
    print("Please confirm these settings:")
    print("---------------------------------")
    print(f"Output file:            {cfg.out_pdf}")
    print(f"Page layout:            {cfg.page_name} ({cfg.page_w_mm:.1f} x {cfg.page_h_mm:.1f} mm)")
    print(f"Margins (mm):           top={cfg.margin_top_mm}, bottom={cfg.margin_bottom_mm}, left={cfg.margin_left_mm}, right={cfg.margin_right_mm}")
    print(f"Grid:                   rows={cfg.rows}, cols={cfg.cols}")
    print(f"Gaps (mm):              horizontal={cfg.gap_x_mm}, vertical={cfg.gap_y_mm}")
    print(f"Dead zones (mm):        left={cfg.deadzone_left_mm}, right={cfg.deadzone_right_mm}")
    print(f"Pages:                  {cfg.pages}")
    print(f"Prefix:                 {cfg.prefix}")
    print(f"Start number:           {cfg.start_number}")
    print(f"Leading zeros:          {cfg.leading_zeros}")

    if cfg.advanced_enabled:
        print("Advanced:               enabled")
        print(f"  Debug label frames:   {cfg.debug_label_frames}")
        print(f"  Debug sheetbox frame: {cfg.debug_sheetbox_frame}")
        print(f"  Offset (mm):          x={cfg.offset_x_mm}, y={cfg.offset_y_mm}")
        print(f"  Scale:                x={cfg.scale_x}, y={cfg.scale_y}")
    else:
        print("Advanced:               disabled")
    print("---------------------------------\n")


def main():
    while True:
        try:
            cfg = build_config_interactive()
            print_summary(cfg)

            if not ask_yes_no("Do you wish to create the file with the above listed settings?", default=True):
                print("\nRestarting...\n")
                continue

            generate_pdf(cfg)
            print(f"\nFile Created: {cfg.out_pdf}")
            break

        except KeyboardInterrupt:
            print("\nCancelled.")
            sys.exit(1)
        except Exception as e:
            print(f"\nERROR: {e}")
            print("Restarting...\n")
            continue


if __name__ == "__main__":
    main()
