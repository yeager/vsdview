"""Convert Visio files (.vsdx, .vsd, .vssx, .vss) to SVG.

Backend priority:
1. libvisio (vsd2xhtml) — lightweight, accurate
2. Built-in .vsdx XML parser — zero dependencies, .vsdx only

Author: Daniel Nylander <daniel@danielnylander.se>
"""

import base64
import gettext
import math
import mimetypes
import os
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

_ = gettext.gettext

# Visio XML namespaces
_NS = {
    "v": "http://schemas.microsoft.com/office/visio/2012/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
_VNS = _NS["v"]
_VTAG = f"{{{_VNS}}}"

# Supported file extensions
VISIO_EXTENSIONS = {".vsd", ".vsdx", ".vsdm"}
STENCIL_EXTENSIONS = {".vss", ".vssx", ".vssm"}
ALL_EXTENSIONS = VISIO_EXTENSIONS | STENCIL_EXTENSIONS

# Visio color index table (standard colors)
_VISIO_COLORS = {
    0: "#000000",  # Black
    1: "#FFFFFF",  # White
    2: "#FF0000",  # Red
    3: "#00FF00",  # Green
    4: "#0000FF",  # Blue
    5: "#FFFF00",  # Yellow
    6: "#FF00FF",  # Magenta
    7: "#00FFFF",  # Cyan
    8: "#800000",  # Dark Red
    9: "#008000",  # Dark Green
    10: "#000080", # Dark Blue
    11: "#808000", # Dark Yellow (Olive)
    12: "#800080", # Dark Magenta (Purple)
    13: "#008080", # Dark Cyan (Teal)
    14: "#C0C0C0", # Light Gray
    15: "#808080", # Dark Gray
    16: "#993366", # Rose
    17: "#333399", # Indigo
    18: "#333333", # Charcoal
    19: "#003300", # Forest
    20: "#003366", # Marine
    21: "#993300", # Brown
    22: "#993366", # Plum
    23: "#333399", # Navy
    24: "#E6E6E6", # Pale Gray
}

# Visio line patterns
_LINE_PATTERNS = {
    0: "none",           # No line
    1: "",               # Solid
    2: "4,3",            # Dash
    3: "1,3",            # Dot
    4: "4,3,1,3",        # Dash-dot
    5: "4,3,1,3,1,3",   # Dash-dot-dot
    6: "8,3",            # Long dash
    7: "1,1",            # Dense dot
    8: "8,3,1,3",        # Long dash-dot
    9: "8,3,1,3,1,3",   # Long dash-dot-dot
    10: "12,6",          # Extra-long dash
    16: "6,3,6,3",       # Dash-dash
}

# Inches to SVG pixels conversion
_INCH_TO_PX = 72.0

# Arrow size lookup (BeginArrowSize/EndArrowSize 0-6 -> scale factor)
_ARROW_SIZES = {0: 0.5, 1: 0.65, 2: 0.8, 3: 1.0, 4: 1.5, 5: 2.0, 6: 2.5}

# MIME types for embedded images
_IMAGE_MIMETYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".emf": "image/x-emf",
    ".wmf": "image/x-wmf",
    ".tiff": "image/tiff",
    ".tif": "image/tiff",
    ".svg": "image/svg+xml",
}

# Relationship namespace
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _lighten_color(hex_color: str, factor: float = 0.7) -> str:
    """Lighten a hex color by blending towards white.

    factor=0.0 returns original, factor=1.0 returns white.
    """
    hex_color = hex_color.strip().lstrip("#")
    if len(hex_color) != 6:
        return "#E8E8E8"
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
    except ValueError:
        return "#E8E8E8"
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"#{r:02X}{g:02X}{b:02X}"


def _is_black(color: str) -> bool:
    """Check if a color is black or near-black."""
    if not color:
        return False
    c = color.strip().upper()
    return c in ("#000000", "#000", "0")


def _hsl_to_rgb(h: int, s: int, l: int) -> str:
    """Convert Visio HSL (h=0-255, s=0-255, l=0-255) to #RRGGBB."""
    # Normalize to 0-1 range
    hf = (h / 255.0) * 360.0
    sf = s / 255.0
    lf = l / 255.0
    # HSL to RGB conversion
    if sf == 0:
        r = g = b = lf
    else:
        def hue2rgb(p, q, t):
            if t < 0: t += 1
            if t > 1: t -= 1
            if t < 1/6: return p + (q - p) * 6 * t
            if t < 1/2: return q
            if t < 2/3: return p + (q - p) * (2/3 - t) * 6
            return p
        q = lf * (1 + sf) if lf < 0.5 else lf + sf - lf * sf
        p = 2 * lf - q
        hn = hf / 360.0
        r = hue2rgb(p, q, hn + 1/3)
        g = hue2rgb(p, q, hn)
        b = hue2rgb(p, q, hn - 1/3)
    return f"#{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}"


def _resolve_color(val: str) -> str:
    """Convert a Visio color value to an SVG color string.

    Handles: color index, #RRGGBB, RGB(r,g,b), HSL(h,s,l), THEMEVAL(), etc.
    Returns empty string for unresolvable values (caller decides default).
    """
    if not val:
        return ""
    val = val.strip()

    # THEMEVAL or formula — return empty (use default, not black)
    if "THEMEVAL" in val or "THEME" in val or val.startswith("="):
        return ""

    # #RRGGBB or #RGB
    if val.startswith("#"):
        return val

    # RGB(r,g,b) function
    m = re.match(r"RGB\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", val, re.IGNORECASE)
    if m:
        r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"#{r:02X}{g:02X}{b:02X}"

    # HSL(h,s,l) function — Visio uses 0-255 range
    m = re.match(r"HSL\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", val, re.IGNORECASE)
    if m:
        return _hsl_to_rgb(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # Numeric index
    try:
        idx = int(val)
        return _VISIO_COLORS.get(idx, "")
    except ValueError:
        pass

    # Try float index
    try:
        idx = int(float(val))
        return _VISIO_COLORS.get(idx, "")
    except (ValueError, TypeError):
        pass

    return ""


def _get_dash_array(pattern: int, weight: float) -> str:
    """Get SVG stroke-dasharray for a Visio line pattern."""
    p = _LINE_PATTERNS.get(pattern, "")
    if not p or p == "none":
        return ""
    # Scale dash pattern by stroke weight
    scale = max(weight, 0.5)
    parts = [str(float(x) * scale) for x in p.split(",")]
    return ",".join(parts)


def _safe_float(val: str | None, default: float = 0.0) -> float:
    """Parse a float value, returning default on failure."""
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def _escape_xml(text: str) -> str:
    """Escape text for XML/SVG output."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


# ---------------------------------------------------------------------------
# Embedded image support
# ---------------------------------------------------------------------------

def _extract_media(zf: zipfile.ZipFile) -> dict[str, bytes]:
    """Extract all files from visio/media/ in the ZIP.

    Returns {filename: bytes} e.g. {"image1.png": b"..."}
    """
    media = {}
    for name in zf.namelist():
        if name.startswith("visio/media/"):
            fname = name.split("/")[-1]
            if fname:
                try:
                    media[fname] = zf.read(name)
                except (KeyError, zipfile.BadZipFile):
                    pass
    return media


def _parse_rels(zf: zipfile.ZipFile, page_file: str) -> dict[str, str]:
    """Parse relationship file for a page to map rId -> target path.

    For visio/pages/page1.xml, the rels file is
    visio/pages/_rels/page1.xml.rels
    """
    page_dir = os.path.dirname(page_file)
    page_basename = os.path.basename(page_file)
    rels_path = f"{page_dir}/_rels/{page_basename}.rels"

    rels = {}
    try:
        rels_xml = zf.read(rels_path)
        root = ET.fromstring(rels_xml)
        for rel in root.findall(f"{{{_RELS_NS}}}Relationship"):
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            if rid and target:
                rels[rid] = target
    except (KeyError, ET.ParseError):
        pass
    return rels


def _image_to_data_uri(data: bytes, filename: str) -> str:
    """Convert image bytes to a base64 data URI."""
    ext = os.path.splitext(filename)[1].lower()
    mime = _IMAGE_MIMETYPES.get(ext, "image/png")
    b64 = base64.b64encode(data).decode("ascii")
    return f"data:{mime};base64,{b64}"


def _save_image_file(data: bytes, filename: str, output_dir: str) -> str:
    """Save image to output directory and return relative filename.

    For large images, saves to file instead of embedding as data URI
    to avoid XML parser buffer limits in rsvg.
    """
    dest = os.path.join(output_dir, filename)
    with open(dest, "wb") as f:
        f.write(data)
    return filename  # Relative path for SVG href


def _parse_foreign_data(shape_elem: ET.Element) -> dict | None:
    """Parse ForeignData element from a shape.

    Returns {"type": "bitmap"|"metafile", "data": base64_str, "rel_id": rIdN}
    or None if no foreign data.
    """
    fd = shape_elem.find(f"{_VTAG}ForeignData")
    if fd is None:
        return None

    info = {
        "foreign_type": fd.get("ForeignType", ""),
        "compression": fd.get("CompressionType", ""),
        "data": None,
        "rel_id": None,
    }

    # Check for Rel element (can be in Visio namespace or r: namespace)
    rel_elem = fd.find(f"{_VTAG}Rel")
    if rel_elem is None:
        rel_elem = fd.find(f"{{{_NS['r']}}}Rel")
    if rel_elem is not None:
        # The r:id attribute may use full namespace
        info["rel_id"] = rel_elem.get(f"{{{_NS['r']}}}id", "")
        if not info["rel_id"]:
            info["rel_id"] = rel_elem.get("r:id", "")
        if not info["rel_id"]:
            for attr_name, attr_val in rel_elem.attrib.items():
                if attr_name.endswith("}id") or attr_name == "id":
                    info["rel_id"] = attr_val
                    break
    else:
        # Inline data
        text = fd.text
        if text and text.strip():
            info["data"] = text.strip()

    return info


# ---------------------------------------------------------------------------
# Arrow marker SVG generation
# ---------------------------------------------------------------------------

def _arrow_marker_defs(used_markers: set[str]) -> list[str]:
    """Generate SVG <defs> for arrow markers.

    used_markers: set of marker IDs like "arrow_end_3", "arrow_start_2"
    """
    if not used_markers:
        return []

    lines = ["<defs>"]
    for marker_id in sorted(used_markers):
        # Parse: arrow_{start|end}_{size}_{color}
        parts = marker_id.split("_", 3)
        direction = parts[1] if len(parts) > 1 else "end"
        size_idx = int(parts[2]) if len(parts) > 2 else 3
        color = f"#{parts[3]}" if len(parts) > 3 else "#333333"

        scale = _ARROW_SIZES.get(size_idx, 1.0)
        marker_w = 10 * scale
        marker_h = 7 * scale

        if direction == "start":
            # Reverse triangle for start
            lines.append(
                f'<marker id="{marker_id}" markerWidth="{marker_w:.1f}" '
                f'markerHeight="{marker_h:.1f}" refX="0" refY="{marker_h/2:.1f}" '
                f'orient="auto" markerUnits="strokeWidth">'
                f'<polygon points="{marker_w:.1f} 0, 0 {marker_h/2:.1f}, '
                f'{marker_w:.1f} {marker_h:.1f}" fill="{color}"/>'
                f'</marker>'
            )
        else:
            # Forward triangle for end
            lines.append(
                f'<marker id="{marker_id}" markerWidth="{marker_w:.1f}" '
                f'markerHeight="{marker_h:.1f}" refX="{marker_w:.1f}" '
                f'refY="{marker_h/2:.1f}" orient="auto" markerUnits="strokeWidth">'
                f'<polygon points="0 0, {marker_w:.1f} {marker_h/2:.1f}, '
                f'0 {marker_h:.1f}" fill="{color}"/>'
                f'</marker>'
            )

    lines.append("</defs>")
    return lines


# ---------------------------------------------------------------------------
# Master shape parsing
# ---------------------------------------------------------------------------

def _parse_master_shapes(zf: zipfile.ZipFile) -> dict[str, dict]:
    """Parse full shape data from master files.

    Returns {master_id: {shape_id: shape_dict, ...}, ...}
    Each shape_dict has: cells, geometry, text, char_formats, para_formats, sub_shapes
    """
    # First, read masters.xml to map Master ID to file
    master_map = {}  # Master ID -> master file number
    try:
        masters_xml = zf.read("visio/masters/masters.xml")
        root = ET.fromstring(masters_xml)
        for master_el in root.findall(f"{_VTAG}Master"):
            mid = master_el.get("ID", "")
            rel_el = master_el.find(f"{{{_NS['r']}}}Rel")
            if rel_el is None:
                rel_el = master_el.find(f"Rel")
            # Also try to find via Rel element with rId
            # The master file number usually matches the ID
            master_map[mid] = mid
    except (KeyError, ET.ParseError):
        pass

    masters = {}
    for name in zf.namelist():
        if not (name.startswith("visio/masters/master") and name.endswith(".xml")):
            continue
        if "masters.xml" in name:
            continue
        master_num = Path(name).stem.replace("master", "")
        try:
            root = ET.fromstring(zf.read(name))
        except (ET.ParseError, KeyError):
            continue

        shapes_data = {}
        for shape in root.iter(f"{_VTAG}Shape"):
            sd = _parse_single_shape(shape)
            shapes_data[sd["id"]] = sd

        if shapes_data:
            masters[master_num] = shapes_data

    return masters


def _parse_single_shape(shape_elem: ET.Element) -> dict:
    """Parse a single <Shape> element into a rich dict."""
    sd = {
        "id": shape_elem.get("ID", ""),
        "name": shape_elem.get("Name", ""),
        "name_u": shape_elem.get("NameU", ""),
        "type": shape_elem.get("Type", "Shape"),
        "master": shape_elem.get("Master", ""),
        "master_shape": shape_elem.get("MasterShape", ""),
        "cells": {},
        "geometry": [],
        "text": "",
        "text_parts": [],
        "char_formats": {},
        "para_formats": {},
        "sub_shapes": [],
        "controls": {},      # Row_N -> {X, Y, ...}
        "connections": {},    # IX -> {X, Y, ...}
        "foreign_data": None, # ForeignData info for embedded images
    }

    # Parse top-level cells
    for cell in shape_elem.findall(f"{_VTAG}Cell"):
        n = cell.get("N", "")
        v = cell.get("V", "")
        f = cell.get("F", "")
        sd["cells"][n] = {"V": v, "F": f}

    # Parse Section elements
    for section in shape_elem.findall(f"{_VTAG}Section"):
        sec_name = section.get("N", "")

        if sec_name == "Geometry":
            geo = _parse_geometry_section(section)
            if geo:
                sd["geometry"].append(geo)

        elif sec_name == "Controls":
            for row in section.findall(f"{_VTAG}Row"):
                row_ix = row.get("IX", "0")
                ctrl = {}
                for cell in row.findall(f"{_VTAG}Cell"):
                    ctrl[cell.get("N", "")] = cell.get("V", "")
                sd["controls"][f"Row_{row_ix}"] = ctrl

        elif sec_name == "Connection":
            for row in section.findall(f"{_VTAG}Row"):
                row_ix = row.get("IX", "0")
                conn = {}
                for cell in row.findall(f"{_VTAG}Cell"):
                    conn[cell.get("N", "")] = cell.get("V", "")
                sd["connections"][row_ix] = conn

        elif sec_name == "Character":
            for row in section.findall(f"{_VTAG}Row"):
                row_ix = row.get("IX", "0")
                fmt = {}
                for cell in row.findall(f"{_VTAG}Cell"):
                    fmt[cell.get("N", "")] = cell.get("V", "")
                sd["char_formats"][row_ix] = fmt

        elif sec_name == "Paragraph":
            for row in section.findall(f"{_VTAG}Row"):
                row_ix = row.get("IX", "0")
                fmt = {}
                for cell in row.findall(f"{_VTAG}Cell"):
                    fmt[cell.get("N", "")] = cell.get("V", "")
                sd["para_formats"][row_ix] = fmt

    # Also parse Geom sections that are direct children (alternative format)
    for geom_idx in range(20):  # Max 20 geometry sections
        geom_section = shape_elem.find(f"{_VTAG}Geom")
        if geom_section is not None and geom_section not in []:
            break

    # Parse text
    text_elem = shape_elem.find(f"{_VTAG}Text")
    if text_elem is not None:
        sd["text"] = "".join(text_elem.itertext()).strip()
        sd["text_parts"] = _parse_text_element(text_elem)

    # Parse sub-shapes (for groups)
    shapes_container = shape_elem.find(f"{_VTAG}Shapes")
    if shapes_container is not None:
        for sub_shape in shapes_container.findall(f"{_VTAG}Shape"):
            sd["sub_shapes"].append(_parse_single_shape(sub_shape))

    # Parse ForeignData (embedded images)
    fd_info = _parse_foreign_data(shape_elem)
    if fd_info:
        sd["foreign_data"] = fd_info

    return sd


def _parse_geometry_section(section: ET.Element) -> dict:
    """Parse a Geometry section into a list of geometry rows."""
    geo = {"rows": [], "no_fill": False, "no_line": False, "no_show": False}

    # Check section-level cells
    for cell in section.findall(f"{_VTAG}Cell"):
        n = cell.get("N", "")
        v = cell.get("V", "0")
        if n == "NoFill" and v == "1":
            geo["no_fill"] = True
        elif n == "NoLine" and v == "1":
            geo["no_line"] = True
        elif n == "NoShow" and v == "1":
            geo["no_show"] = True

    for row in section.findall(f"{_VTAG}Row"):
        row_type = row.get("T", "")
        row_data = {"type": row_type, "cells": {}}
        for cell in row.findall(f"{_VTAG}Cell"):
            n = cell.get("N", "")
            v = cell.get("V", "")
            f = cell.get("F", "")
            row_data["cells"][n] = {"V": v, "F": f}
        geo["rows"].append(row_data)

    return geo


def _parse_text_element(text_elem: ET.Element) -> list:
    """Parse a <Text> element into parts with formatting references."""
    parts = []
    current_cp = "0"
    current_pp = "0"

    # Process text content with inline elements
    if text_elem.text:
        parts.append({"text": text_elem.text, "cp": current_cp, "pp": current_pp})

    for child in text_elem:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "cp":
            current_cp = child.get("IX", "0")
        elif tag == "pp":
            current_pp = child.get("IX", "0")
        elif tag == "fld":
            # Field element — extract text
            field_text = "".join(child.itertext()).strip()
            if field_text:
                parts.append({"text": field_text, "cp": current_cp, "pp": current_pp})
        if child.tail:
            parts.append({"text": child.tail, "cp": current_cp, "pp": current_pp})

    return parts


# ---------------------------------------------------------------------------
# Geometry to SVG path conversion
# ---------------------------------------------------------------------------

def _geometry_to_path(geo: dict, w: float, h: float,
                      master_w: float = 0.0, master_h: float = 0.0) -> str:
    """Convert a parsed Geometry section to an SVG path 'd' attribute.

    Coordinates are in local shape space (inches), will be scaled to px.
    w, h are shape width/height in inches for relative coordinates.
    master_w, master_h: if geometry was inherited from a master, these are
    the master's original dimensions for coordinate scaling.
    """
    if geo.get("no_show"):
        return ""

    # Compute scale factors if geometry came from a master with different dims
    sx = w / master_w if master_w > 1e-6 and abs(master_w - w) > 1e-6 else 1.0
    sy = h / master_h if master_h > 1e-6 and abs(master_h - h) > 1e-6 else 1.0

    d_parts = []
    cx, cy = 0.0, 0.0  # Current point (inches)

    for row in geo["rows"]:
        rt = row["type"]
        cells = row["cells"]

        if rt == "MoveTo":
            x = _safe_float(cells.get("X", {}).get("V")) * sx
            y = _safe_float(cells.get("Y", {}).get("V")) * sy
            d_parts.append(f"M {x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}")
            cx, cy = x, y

        elif rt == "RelMoveTo":
            x = _safe_float(cells.get("X", {}).get("V"))
            y = _safe_float(cells.get("Y", {}).get("V"))
            ax, ay = x * w, y * h
            d_parts.append(f"M {ax * _INCH_TO_PX:.2f} {(h - ay) * _INCH_TO_PX:.2f}")
            cx, cy = ax, ay

        elif rt == "LineTo":
            x = _safe_float(cells.get("X", {}).get("V")) * sx
            y = _safe_float(cells.get("Y", {}).get("V")) * sy
            d_parts.append(f"L {x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}")
            cx, cy = x, y

        elif rt == "RelLineTo":
            x = _safe_float(cells.get("X", {}).get("V"))
            y = _safe_float(cells.get("Y", {}).get("V"))
            ax, ay = x * w, y * h
            d_parts.append(f"L {ax * _INCH_TO_PX:.2f} {(h - ay) * _INCH_TO_PX:.2f}")
            cx, cy = ax, ay

        elif rt == "ArcTo":
            x = _safe_float(cells.get("X", {}).get("V")) * sx
            y = _safe_float(cells.get("Y", {}).get("V")) * sy
            a = _safe_float(cells.get("A", {}).get("V")) * sy  # bulge scales with Y
            # A is the bulge/sagitta of the arc
            _append_arc(d_parts, cx, cy, x, y, a, h)
            cx, cy = x, y

        elif rt == "EllipticalArcTo":
            x = _safe_float(cells.get("X", {}).get("V")) * sx
            y = _safe_float(cells.get("Y", {}).get("V")) * sy
            a = _safe_float(cells.get("A", {}).get("V")) * sx  # control point X
            b = _safe_float(cells.get("B", {}).get("V")) * sy  # control point Y
            c = _safe_float(cells.get("C", {}).get("V"))  # major/minor ratio
            d_val = _safe_float(cells.get("D", {}).get("V"))  # angle
            _append_elliptical_arc(d_parts, cx, cy, x, y, a, b, c, d_val, h)
            cx, cy = x, y

        elif rt == "NURBSTo":
            x = _safe_float(cells.get("X", {}).get("V")) * sx
            y = _safe_float(cells.get("Y", {}).get("V")) * sy
            # Approximate NURBS as line for now (proper NURBS requires complex computation)
            d_parts.append(f"L {x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}")
            cx, cy = x, y

        elif rt == "RelCurveTo":
            x = _safe_float(cells.get("X", {}).get("V"))
            y = _safe_float(cells.get("Y", {}).get("V"))
            a = _safe_float(cells.get("A", {}).get("V"))
            b = _safe_float(cells.get("B", {}).get("V"))
            c = _safe_float(cells.get("C", {}).get("V"))
            dd = _safe_float(cells.get("D", {}).get("V"))
            # Cubic bezier with relative coordinates
            cp1x, cp1y = a * w, b * h
            cp2x, cp2y = c * w, dd * h
            ex, ey = x * w, y * h
            d_parts.append(
                f"C {cp1x * _INCH_TO_PX:.2f} {(h - cp1y) * _INCH_TO_PX:.2f} "
                f"{cp2x * _INCH_TO_PX:.2f} {(h - cp2y) * _INCH_TO_PX:.2f} "
                f"{ex * _INCH_TO_PX:.2f} {(h - ey) * _INCH_TO_PX:.2f}"
            )
            cx, cy = ex, ey

        elif rt == "Ellipse":
            # Full ellipse: center (X,Y), point on major axis (A,B), point on minor axis (C,D)
            ex = _safe_float(cells.get("X", {}).get("V")) * sx
            ey = _safe_float(cells.get("Y", {}).get("V")) * sy
            ea = _safe_float(cells.get("A", {}).get("V")) * sx
            eb = _safe_float(cells.get("B", {}).get("V")) * sy
            ec = _safe_float(cells.get("C", {}).get("V")) * sx
            ed = _safe_float(cells.get("D", {}).get("V")) * sy
            rx = math.sqrt((ea - ex) ** 2 + (eb - ey) ** 2)
            ry = math.sqrt((ec - ex) ** 2 + (ed - ey) ** 2)
            if rx < 0.001:
                rx = 0.001
            if ry < 0.001:
                ry = 0.001
            cpx = ex * _INCH_TO_PX
            cpy = (h - ey) * _INCH_TO_PX
            rpx = rx * _INCH_TO_PX
            rpy = ry * _INCH_TO_PX
            # SVG ellipse as two arcs
            d_parts.append(
                f"M {cpx - rpx:.2f} {cpy:.2f} "
                f"A {rpx:.2f} {rpy:.2f} 0 1 0 {cpx + rpx:.2f} {cpy:.2f} "
                f"A {rpx:.2f} {rpy:.2f} 0 1 0 {cpx - rpx:.2f} {cpy:.2f} Z"
            )

        elif rt == "PolylineTo":
            x = _safe_float(cells.get("X", {}).get("V")) * sx
            y = _safe_float(cells.get("Y", {}).get("V")) * sy
            # Try to parse the formula for intermediate points
            a_cell = cells.get("A", {})
            formula = a_cell.get("F", "")
            pts = _parse_polyline_formula(formula, w, h)
            if pts:
                for px_val, py_val in pts:
                    d_parts.append(f"L {px_val * _INCH_TO_PX:.2f} {(h - py_val) * _INCH_TO_PX:.2f}")
            d_parts.append(f"L {x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}")
            cx, cy = x, y

        elif rt == "SplineStart":
            x = _safe_float(cells.get("X", {}).get("V")) * sx
            y = _safe_float(cells.get("Y", {}).get("V")) * sy
            d_parts.append(f"M {x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}")
            cx, cy = x, y

        elif rt == "SplineKnot":
            x = _safe_float(cells.get("X", {}).get("V")) * sx
            y = _safe_float(cells.get("Y", {}).get("V")) * sy
            d_parts.append(f"L {x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}")
            cx, cy = x, y

        elif rt == "InfiniteLine":
            x = _safe_float(cells.get("X", {}).get("V")) * sx
            y = _safe_float(cells.get("Y", {}).get("V")) * sy
            a = _safe_float(cells.get("A", {}).get("V")) * sx
            b = _safe_float(cells.get("B", {}).get("V")) * sy
            d_parts.append(f"M {x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}")
            d_parts.append(f"L {a * _INCH_TO_PX:.2f} {(h - b) * _INCH_TO_PX:.2f}")
            cx, cy = a, b

    return " ".join(d_parts)


def _append_arc(d_parts: list, cx: float, cy: float, x: float, y: float,
                bulge: float, h: float):
    """Append an arc segment (ArcTo) using SVG arc command.

    bulge (A) is the sagitta — distance from the midpoint of the chord to the arc.
    If bulge is 0, it's a straight line.
    """
    if abs(bulge) < 1e-6:
        d_parts.append(f"L {x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}")
        return

    # Compute arc from chord and sagitta
    dx = x - cx
    dy = y - cy
    chord = math.sqrt(dx * dx + dy * dy)
    if chord < 1e-10:
        return

    # Radius from sagitta: r = (chord²/4 + sagitta²) / (2 * |sagitta|)
    sagitta = abs(bulge)
    radius = (chord * chord / 4 + sagitta * sagitta) / (2 * sagitta)
    radius_px = radius * _INCH_TO_PX

    # Determine sweep direction
    large_arc = 1 if sagitta > chord / 2 else 0
    sweep = 0 if bulge > 0 else 1

    d_parts.append(
        f"A {radius_px:.2f} {radius_px:.2f} 0 {large_arc} {sweep} "
        f"{x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}"
    )


def _append_elliptical_arc(d_parts: list, cx: float, cy: float,
                           x: float, y: float, a: float, b: float,
                           c: float, d_val: float, h: float):
    """Append an elliptical arc segment.

    (a,b) = control point, c = aspect ratio, d_val = rotation angle.
    For simplicity, approximate with SVG arc.
    """
    # Compute approximate radius from control point
    mid_x = (cx + x) / 2
    mid_y = (cy + y) / 2
    dist_to_control = math.sqrt((a - mid_x) ** 2 + (b - mid_y) ** 2)
    chord = math.sqrt((x - cx) ** 2 + (y - cy) ** 2)

    if chord < 1e-10:
        return

    sagitta = dist_to_control
    if sagitta < 1e-6:
        d_parts.append(f"L {x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}")
        return

    rx = (chord * chord / 4 + sagitta * sagitta) / (2 * sagitta)
    ry = rx * c if c > 0 else rx
    angle_deg = math.degrees(d_val) if d_val else 0

    rx_px = abs(rx * _INCH_TO_PX)
    ry_px = abs(ry * _INCH_TO_PX)
    if rx_px < 0.1:
        rx_px = 0.1
    if ry_px < 0.1:
        ry_px = 0.1

    # Determine arc direction from control point position relative to chord
    cross = (x - cx) * (b - cy) - (y - cy) * (a - cx)
    sweep = 1 if cross < 0 else 0
    large_arc = 0

    d_parts.append(
        f"A {rx_px:.2f} {ry_px:.2f} {angle_deg:.1f} {large_arc} {sweep} "
        f"{x * _INCH_TO_PX:.2f} {(h - y) * _INCH_TO_PX:.2f}"
    )


def _parse_polyline_formula(formula: str, w: float, h: float) -> list[tuple[float, float]]:
    """Parse a POLYLINE formula to extract points."""
    # Format: POLYLINE(0, 0, x1, y1, x2, y2, ...)
    pts = []
    m = re.match(r"POLYLINE\s*\((.*)\)", formula, re.IGNORECASE)
    if not m:
        return pts
    try:
        vals = [float(v.strip()) for v in m.group(1).split(",")]
        # Skip first two values (flags), then pairs
        for i in range(2, len(vals) - 1, 2):
            pts.append((vals[i], vals[i + 1]))
    except (ValueError, IndexError):
        pass
    return pts


# ---------------------------------------------------------------------------
# Shape merging (master inheritance)
# ---------------------------------------------------------------------------

def _merge_shape_with_master(shape: dict, masters: dict,
                              parent_master_id: str = "") -> dict:
    """Merge a shape with its master, local values override master values.

    For sub-shapes in groups, parent_master_id is the group's Master ID,
    and the sub-shape's master_shape references a shape within that master.
    """
    master_id = shape.get("master", "") or parent_master_id
    master_shape_id = shape.get("master_shape", "")

    if not master_id or master_id not in masters:
        return shape

    master_shapes = masters[master_id]

    # Find the right master shape
    master_sd = None
    if master_shape_id and master_shape_id in master_shapes:
        master_sd = master_shapes[master_shape_id]
    elif master_shapes:
        # Use first shape in master
        master_sd = next(iter(master_shapes.values()))

    if not master_sd:
        return shape

    # Merge cells: master provides defaults, local overrides
    merged_cells = dict(master_sd.get("cells", {}))
    merged_cells.update({k: v for k, v in shape["cells"].items() if v.get("V")})
    shape["cells"] = merged_cells

    # Merge geometry: use local if present, otherwise master
    if not shape["geometry"] and master_sd.get("geometry"):
        shape["geometry"] = master_sd["geometry"]
        # Store master's original dimensions for geometry coordinate scaling
        master_w_val = master_sd.get("cells", {}).get("Width", {}).get("V")
        master_h_val = master_sd.get("cells", {}).get("Height", {}).get("V")
        if master_w_val:
            shape["_master_w"] = _safe_float(master_w_val)
        if master_h_val:
            shape["_master_h"] = _safe_float(master_h_val)

    # Merge text: use local if present, otherwise master
    if not shape["text"] and master_sd.get("text"):
        txt = master_sd["text"]
        if txt not in ("Label", "Abc"):
            shape["text"] = txt
            if not shape["text_parts"] and master_sd.get("text_parts"):
                shape["text_parts"] = master_sd["text_parts"]

    # Merge character and paragraph formats
    if not shape["char_formats"] and master_sd.get("char_formats"):
        shape["char_formats"] = master_sd["char_formats"]
    if not shape["para_formats"] and master_sd.get("para_formats"):
        shape["para_formats"] = master_sd["para_formats"]

    # Merge controls and connections
    if not shape.get("controls") and master_sd.get("controls"):
        shape["controls"] = master_sd["controls"]
    if not shape.get("connections") and master_sd.get("connections"):
        shape["connections"] = master_sd["connections"]

    return shape


# ---------------------------------------------------------------------------
# Shape to SVG rendering
# ---------------------------------------------------------------------------

def _get_cell_val(shape: dict, name: str, default: str = "") -> str:
    """Get a cell value from a shape."""
    cell = shape.get("cells", {}).get(name, {})
    return cell.get("V", default)


def _get_cell_float(shape: dict, name: str, default: float = 0.0) -> float:
    """Get a cell value as float."""
    return _safe_float(_get_cell_val(shape, name), default)


def _compute_transform(shape: dict, page_h: float) -> str:
    """Compute SVG transform for a shape.

    Handles PinX/PinY positioning, LocPinX/LocPinY, rotation, and flipping.
    Returns SVG transform attribute value.
    """
    pin_x = _get_cell_float(shape, "PinX") * _INCH_TO_PX
    pin_y = (page_h - _get_cell_float(shape, "PinY")) * _INCH_TO_PX
    loc_pin_x = _get_cell_float(shape, "LocPinX") * _INCH_TO_PX
    loc_pin_y_raw = _get_cell_float(shape, "LocPinY")
    w = _get_cell_float(shape, "Width")
    h = _get_cell_float(shape, "Height")
    loc_pin_y = (h - loc_pin_y_raw) * _INCH_TO_PX  # Flip Y for local pin

    angle = _get_cell_float(shape, "Angle")
    flip_x = _get_cell_val(shape, "FlipX") == "1"
    flip_y = _get_cell_val(shape, "FlipY") == "1"

    parts = []

    # Translate so pin point is at correct page position
    tx = pin_x - loc_pin_x
    ty = pin_y - loc_pin_y

    parts.append(f"translate({tx:.2f},{ty:.2f})")

    # Apply rotation around local pin
    if abs(angle) > 1e-6:
        angle_deg = -math.degrees(angle)  # Visio angles are CCW, SVG CW
        parts.append(f"rotate({angle_deg:.2f},{loc_pin_x:.2f},{loc_pin_y:.2f})")

    # Apply flips around local pin
    if flip_x or flip_y:
        sx = -1 if flip_x else 1
        sy = -1 if flip_y else 1
        # Translate to origin, scale, translate back
        parts.append(f"translate({loc_pin_x:.2f},{loc_pin_y:.2f})")
        parts.append(f"scale({sx},{sy})")
        parts.append(f"translate({-loc_pin_x:.2f},{-loc_pin_y:.2f})")

    return " ".join(parts)


def _render_shape_svg(shape: dict, page_h: float, masters: dict,
                       parent_master_id: str = "",
                       _depth: int = 0,
                       media: dict | None = None,
                       page_rels: dict | None = None,
                       used_markers: set | None = None,
                       output_dir: str | None = None) -> list[str]:
    """Render a single shape as SVG elements. Returns list of SVG strings."""
    shape = _merge_shape_with_master(shape, masters, parent_master_id)
    if media is None:
        media = {}
    if page_rels is None:
        page_rels = {}
    if used_markers is None:
        used_markers = set()

    lines = []

    # Skip shapes that are invisible or purely connection/control metadata
    vis_val = _get_cell_val(shape, "Visible")
    if vis_val == "0":
        return lines

    # Skip shapes with only connection points and no geometry/text (connection markers)
    if (shape.get("connections") and not shape.get("geometry")
            and not shape.get("text") and not shape.get("sub_shapes")):
        return lines

    # Handle shape type
    shape_type = shape.get("type", "Shape")

    w_inch = _get_cell_float(shape, "Width")
    h_inch = _get_cell_float(shape, "Height")
    w_px = w_inch * _INCH_TO_PX
    h_px = h_inch * _INCH_TO_PX

    # --- Style ---
    line_weight = _get_cell_float(shape, "LineWeight", 0.01) * _INCH_TO_PX
    if line_weight < 0.25:
        line_weight = 0.75
    elif line_weight > 20:
        line_weight = 20

    line_color = _resolve_color(_get_cell_val(shape, "LineColor")) or "#333333"
    fill_foregnd = _resolve_color(_get_cell_val(shape, "FillForegnd"))
    fill_bkgnd = _resolve_color(_get_cell_val(shape, "FillBkgnd"))

    # GUARD(color_index) in Visio stencils are theme accent placeholders.
    # Replace magenta (#FF00FF, color 6) with a sensible default accent.
    _ff_formula = shape.get("cells", {}).get("FillForegnd", {}).get("F", "")
    if "GUARD" in _ff_formula and fill_foregnd == "#FF00FF":
        fill_foregnd = "#5B9BD5"  # Default blue accent
    _fb_formula = shape.get("cells", {}).get("FillBkgnd", {}).get("F", "")
    if "GUARD" in _fb_formula and fill_bkgnd == "#FF00FF":
        fill_bkgnd = "#5B9BD5"
    fill_pattern = _get_cell_val(shape, "FillPattern", "1")
    line_pattern = int(_safe_float(_get_cell_val(shape, "LinePattern", "1")))
    rounding = _get_cell_float(shape, "Rounding") * _INCH_TO_PX

    # Determine fill
    fill_pat_int = int(_safe_float(fill_pattern, 1))
    if fill_pat_int == 0:
        fill = "none"
    elif fill_pat_int == 1:
        # Solid fill
        fill = fill_foregnd or fill_bkgnd or "none"
    elif fill_pat_int >= 2:
        # Pattern/texture fill — we can't render actual patterns, so approximate.
        # Use FillBkgnd as the visible base if available and not black.
        # FillForegnd=0 (black) with pattern > 1 is typically a theme placeholder.
        if fill_bkgnd and not _is_black(fill_bkgnd):
            fill = fill_bkgnd
        elif fill_foregnd and not _is_black(fill_foregnd):
            fill = _lighten_color(fill_foregnd, 0.7)
        else:
            # FillForegnd is black/missing with complex pattern = theme placeholder
            fill = "none"
    else:
        fill = "none"

    # No line if pattern 0
    stroke = line_color if line_pattern != 0 else "none"
    stroke_width = line_weight

    dash_array = _get_dash_array(line_pattern, stroke_width)

    # Build style string
    style_parts = [
        f'fill="{fill}"',
        f'stroke="{stroke}"',
        f'stroke-width="{stroke_width:.2f}"',
    ]
    if dash_array:
        style_parts.append(f'stroke-dasharray="{dash_array}"')

    style_str = " ".join(style_parts)

    # --- Check for 1D shape (connector/line) ---
    begin_x = _get_cell_val(shape, "BeginX")
    begin_y = _get_cell_val(shape, "BeginY")
    end_x = _get_cell_val(shape, "EndX")
    end_y = _get_cell_val(shape, "EndY")
    is_1d = bool(begin_x and end_x)

    # --- Group shapes ---
    if shape_type == "Group" or shape.get("sub_shapes"):
        transform = _compute_transform(shape, page_h)
        group_master_id = shape.get("master", "") or parent_master_id
        # Group's local coordinate system uses its own Width x Height
        group_h = h_inch

        # Clip group contents to group bounds.
        # The clip path is in the group's local (pre-rotation) coordinate space,
        # so we apply it inside the transform — width/height are pre-rotation dims.
        clip_id = f"clip_{shape['id']}"
        lines.append(
            f'<defs><clipPath id="{clip_id}">'
            f'<rect x="0" y="0" width="{w_px:.2f}" height="{h_px:.2f}"/>'
            f'</clipPath></defs>'
        )
        lines.append(
            f'<g transform="{transform}" clip-path="url(#{clip_id})">'
        )
        for sub in shape.get("sub_shapes", []):
            lines.extend(_render_shape_svg(
                sub, group_h, masters, group_master_id, _depth + 1,
                media, page_rels, used_markers, output_dir))
        lines.append('</g>')
        # Render text for the group itself, or shape name as label
        if shape["text"]:
            _append_text_svg(lines, shape, page_h, w_px, h_px)
        elif _depth == 0:
            label = shape.get("name_u") or shape.get("name", "")
            if label and label not in ("Sheet", "Group"):
                _append_name_label(lines, shape, page_h, w_px, h_px, label)
        return lines

    # --- Compute transform ---
    transform = _compute_transform(shape, page_h)

    # --- Geometry rendering ---
    has_geometry = bool(shape["geometry"])

    if has_geometry:
        master_w = shape.get("_master_w", 0.0)
        master_h = shape.get("_master_h", 0.0)
        for geo in shape["geometry"]:
            path_d = _geometry_to_path(geo, w_inch, h_inch, master_w, master_h)
            if not path_d:
                continue

            geo_fill = fill
            geo_stroke = stroke
            if geo.get("no_fill"):
                geo_fill = "none"
            if geo.get("no_line"):
                geo_stroke = "none"

            geo_style = (
                f'fill="{geo_fill}" stroke="{geo_stroke}" '
                f'stroke-width="{stroke_width:.2f}"'
            )
            if dash_array:
                geo_style += f' stroke-dasharray="{dash_array}"'

            lines.append(
                f'<path d="{path_d}" {geo_style} '
                f'transform="{transform}"/>'
            )

    elif is_1d:
        # 1D shape — check for geometry (routed connectors) first
        bx = _safe_float(begin_x) * _INCH_TO_PX
        by = (page_h - _safe_float(begin_y)) * _INCH_TO_PX
        ex_px = _safe_float(end_x) * _INCH_TO_PX
        ey_px = (page_h - _safe_float(end_y)) * _INCH_TO_PX

        # Arrow markers
        begin_arrow = int(_safe_float(_get_cell_val(shape, "BeginArrow", "0")))
        end_arrow = int(_safe_float(_get_cell_val(shape, "EndArrow", "0")))
        begin_arrow_size = int(_safe_float(_get_cell_val(shape, "BeginArrowSize", "2")))
        end_arrow_size = int(_safe_float(_get_cell_val(shape, "EndArrowSize", "2")))
        marker_color = stroke.lstrip("#") if stroke != "none" else "333333"
        marker_attrs = ""
        if begin_arrow > 0:
            mid = f"arrow_start_{begin_arrow_size}_{marker_color}"
            used_markers.add(mid)
            marker_attrs += f' marker-start="url(#{mid})"'
        if end_arrow > 0:
            mid = f"arrow_end_{end_arrow_size}_{marker_color}"
            used_markers.add(mid)
            marker_attrs += f' marker-end="url(#{mid})"'

        if has_geometry:
            # Routed connector — render geometry path instead of straight line
            master_w = shape.get("_master_w", 0.0)
            master_h = shape.get("_master_h", 0.0)
            for geo in shape["geometry"]:
                path_d = _geometry_to_path(geo, w_inch, h_inch, master_w, master_h)
                if not path_d:
                    continue
                geo_stroke = stroke
                if geo.get("no_line"):
                    geo_stroke = "none"
                lines.append(
                    f'<path d="{path_d}" fill="none" stroke="{geo_stroke}" '
                    f'stroke-width="{stroke_width:.2f}"'
                    + (f' stroke-dasharray="{dash_array}"' if dash_array else '')
                    + marker_attrs
                    + f' transform="{transform}"/>'
                )
        else:
            # Simple straight line
            lines.append(
                f'<line x1="{bx:.2f}" y1="{by:.2f}" x2="{ex_px:.2f}" y2="{ey_px:.2f}" '
                f'stroke="{stroke}" stroke-width="{stroke_width:.2f}"'
                + (f' stroke-dasharray="{dash_array}"' if dash_array else '')
                + marker_attrs
                + '/>'
            )

    else:
        # No geometry, no 1D — fall back to rectangle only if shape has
        # explicit fill/stroke (not for empty group sub-shapes)
        if w_px > 0 and h_px > 0 and (fill != "none" or shape.get("text")):
            rx_attr = f' rx="{rounding:.2f}"' if rounding > 0 else ""
            lines.append(
                f'<rect x="0" y="0" width="{w_px:.2f}" height="{h_px:.2f}" '
                f'{style_str}{rx_attr} transform="{transform}"/>'
            )

    # --- Embedded image rendering ---
    fd = shape.get("foreign_data")
    if fd and media:
        img_href = None
        if fd.get("rel_id") and fd["rel_id"] in page_rels:
            target = page_rels[fd["rel_id"]]
            img_name = target.split("/")[-1]
            if img_name in media:
                if output_dir:
                    img_href = _save_image_file(media[img_name], img_name, output_dir)
                else:
                    img_href = _image_to_data_uri(media[img_name], img_name)
        elif fd.get("data"):
            ext_map = {"PNG": ".png", "JPEG": ".jpeg", "BMP": ".bmp",
                       "GIF": ".gif", "TIFF": ".tiff"}
            comp = fd.get("compression", "PNG").upper()
            fake_ext = ext_map.get(comp, ".png")
            try:
                raw = base64.b64decode(fd["data"])
                fname = f"inline_{shape['id']}{fake_ext}"
                if output_dir:
                    img_href = _save_image_file(raw, fname, output_dir)
                else:
                    img_href = _image_to_data_uri(raw, fname)
            except Exception:
                pass

        if img_href:
            img_w = _get_cell_float(shape, "ImgWidth") or w_inch
            img_h = _get_cell_float(shape, "ImgHeight") or h_inch
            img_off_x = _get_cell_float(shape, "ImgOffsetX")
            img_off_y = _get_cell_float(shape, "ImgOffsetY")
            img_w_px = img_w * _INCH_TO_PX
            img_h_px = img_h * _INCH_TO_PX
            img_x_px = img_off_x * _INCH_TO_PX
            img_y_px = img_off_y * _INCH_TO_PX
            lines.append(
                f'<image x="{img_x_px:.2f}" y="{img_y_px:.2f}" '
                f'width="{img_w_px:.2f}" height="{img_h_px:.2f}" '
                f'href="{img_href}" '
                f'preserveAspectRatio="xMidYMid meet" '
                f'transform="{transform}"/>'
            )

    # --- Text rendering ---
    if shape["text"]:
        _append_text_svg(lines, shape, page_h, w_px, h_px)

    # No fallback rectangle for shapes inside groups (sub-shapes)
    # is handled by skipping the else branch when geometry/1D absent
    # and the shape has no meaningful content

    return lines


def _append_text_svg(lines: list, shape: dict, page_h: float,
                     w_px: float, h_px: float):
    """Append SVG text elements for a shape's text."""
    text = shape["text"]
    if not text:
        return

    # Text position
    pin_x = _get_cell_float(shape, "PinX") * _INCH_TO_PX
    pin_y = (page_h - _get_cell_float(shape, "PinY")) * _INCH_TO_PX

    # Text block offset
    txt_pin_x = _get_cell_float(shape, "TxtPinX")
    txt_pin_y = _get_cell_float(shape, "TxtPinY")
    txt_loc_pin_x = _get_cell_float(shape, "TxtLocPinX")
    txt_loc_pin_y = _get_cell_float(shape, "TxtLocPinY")

    tx = pin_x
    ty = pin_y

    # Get text formatting
    char_fmt = shape.get("char_formats", {}).get("0", {})
    font_size = _safe_float(char_fmt.get("Size"), 0.1111) * _INCH_TO_PX  # ~8pt default
    if font_size < 6:
        font_size = 8
    elif font_size > 72:
        font_size = 72

    text_color = _resolve_color(char_fmt.get("Color", "")) or "#000000"
    style_bits = int(_safe_float(char_fmt.get("Style", "0")))
    is_bold = bool(style_bits & 1)
    is_italic = bool(style_bits & 2)
    is_underline = bool(style_bits & 4)

    # Paragraph alignment
    para_fmt = shape.get("para_formats", {}).get("0", {})
    halign = int(_safe_float(para_fmt.get("HorzAlign", "1")))
    anchor_map = {0: "start", 1: "middle", 2: "end"}
    text_anchor = anchor_map.get(halign, "middle")

    # Font weight/style
    fw = ' font-weight="bold"' if is_bold else ""
    fs = ' font-style="italic"' if is_italic else ""
    td = ' text-decoration="underline"' if is_underline else ""

    # Text wrapping: parse TxtWidth for max text width
    txt_width = _get_cell_float(shape, "TxtWidth")
    txt_width_px = txt_width * _INCH_TO_PX if txt_width > 0 else w_px

    # Split text into lines, then wrap long lines
    text_lines = text.split("\n")

    # Simple word-wrap: estimate chars per line from font size
    if txt_width_px > 0 and font_size > 0:
        avg_char_w = font_size * 0.55  # Approximate average char width
        max_chars = max(4, int(txt_width_px / avg_char_w))
        wrapped_lines = []
        for tline in text_lines:
            if len(tline) <= max_chars:
                wrapped_lines.append(tline)
            else:
                words = tline.split()
                current = ""
                for word in words:
                    if current and len(current) + 1 + len(word) > max_chars:
                        wrapped_lines.append(current)
                        current = word
                    else:
                        current = current + " " + word if current else word
                if current:
                    wrapped_lines.append(current)
        text_lines = wrapped_lines

    if len(text_lines) == 1:
        escaped = _escape_xml(text_lines[0])
        lines.append(
            f'<text x="{tx:.2f}" y="{ty:.2f}" '
            f'text-anchor="{text_anchor}" dominant-baseline="central" '
            f'font-family="sans-serif" font-size="{font_size:.1f}" '
            f'fill="{text_color}"{fw}{fs}{td}>'
            f'{escaped}</text>'
        )
    else:
        # Multi-line text
        total_height = len(text_lines) * font_size * 1.2
        start_y = ty - total_height / 2 + font_size * 0.6
        for j, tline in enumerate(text_lines):
            if not tline.strip():
                continue
            escaped = _escape_xml(tline)
            ly = start_y + j * font_size * 1.2
            lines.append(
                f'<text x="{tx:.2f}" y="{ly:.2f}" '
                f'text-anchor="{text_anchor}" '
                f'font-family="sans-serif" font-size="{font_size:.1f}" '
                f'fill="{text_color}"{fw}{fs}{td}>'
                f'{escaped}</text>'
            )


def _append_name_label(lines: list, shape: dict, page_h: float,
                       w_px: float, h_px: float, label: str):
    """Append a shape name as a text label below the shape."""
    pin_x = _get_cell_float(shape, "PinX") * _INCH_TO_PX
    pin_y = (page_h - _get_cell_float(shape, "PinY")) * _INCH_TO_PX
    h_inch = _get_cell_float(shape, "Height")

    # Position label below the shape
    tx = pin_x
    ty = pin_y + h_inch * _INCH_TO_PX / 2 + 14

    # Scale font size relative to shape width (min 10px, max 48px)
    w_inch = _get_cell_float(shape, "Width")
    font_sz = max(10, min(48, w_inch * _INCH_TO_PX * 0.06))

    escaped = _escape_xml(label)
    lines.append(
        f'<text x="{tx:.2f}" y="{ty:.2f}" '
        f'text-anchor="middle" dominant-baseline="central" '
        f'font-family="sans-serif" font-size="{font_sz:.1f}" '
        f'fill="#333333">'
        f'{escaped}</text>'
    )


# ---------------------------------------------------------------------------
# Page dimension parsing
# ---------------------------------------------------------------------------

def _parse_page_dimensions(page_xml: bytes) -> tuple[float, float]:
    """Extract page width and height from a page XML.

    Returns (width_inches, height_inches).
    """
    try:
        root = ET.fromstring(page_xml)
    except ET.ParseError:
        return (8.5, 11.0)

    page_w = 8.5
    page_h = 11.0

    # Look for PageSheet
    page_sheet = root.find(f"{_VTAG}PageSheet")
    if page_sheet is not None:
        for cell in page_sheet.findall(f"{_VTAG}Cell"):
            n = cell.get("N", "")
            v = cell.get("V", "")
            if n == "PageWidth":
                page_w = _safe_float(v, 8.5)
            elif n == "PageHeight":
                page_h = _safe_float(v, 11.0)

    return page_w, page_h


def _parse_all_page_dimensions(zf: zipfile.ZipFile) -> list[tuple[float, float]]:
    """Parse page dimensions from pages.xml (the index file).

    Returns list of (width_inches, height_inches) per page.
    Falls back to individual page XML parsing.
    """
    dims = []
    try:
        pages_xml = zf.read("visio/pages/pages.xml")
        root = ET.fromstring(pages_xml)
        for page in root.findall(f"{_VTAG}Page"):
            pw, ph = 8.5, 11.0
            page_sheet = page.find(f"{_VTAG}PageSheet")
            if page_sheet is not None:
                for cell in page_sheet.findall(f"{_VTAG}Cell"):
                    n = cell.get("N", "")
                    v = cell.get("V", "")
                    if n == "PageWidth":
                        pw = _safe_float(v, 8.5)
                    elif n == "PageHeight":
                        ph = _safe_float(v, 11.0)
            dims.append((pw, ph))
    except (KeyError, ET.ParseError):
        pass
    return dims


# ---------------------------------------------------------------------------
# Main parser and SVG generation
# ---------------------------------------------------------------------------

def _parse_connects(page_xml_root: ET.Element) -> list[dict]:
    """Parse <Connect> elements from a page XML root."""
    connects = []
    connects_el = page_xml_root.find(f"{_VTAG}Connects")
    if connects_el is None:
        return connects
    for c in connects_el.findall(f"{_VTAG}Connect"):
        connects.append({
            "from_sheet": c.get("FromSheet", ""),
            "from_cell": c.get("FromCell", ""),
            "to_sheet": c.get("ToSheet", ""),
            "to_cell": c.get("ToCell", ""),
        })
    return connects


def _build_shape_index(shapes: list[dict]) -> dict[str, dict]:
    """Build a flat index of shape ID -> shape dict, including sub-shapes."""
    idx = {}
    for s in shapes:
        idx[s["id"]] = s
        for sub in s.get("sub_shapes", []):
            idx[sub["id"]] = sub
            # Also index deeper sub-shapes
            for subsub in sub.get("sub_shapes", []):
                idx[subsub["id"]] = subsub
    return idx


def _resolve_connection_point(shape: dict, cell_ref: str, page_h: float,
                               shape_index: dict) -> tuple[float, float] | None:
    """Resolve a connection cell reference to page coordinates (px).

    cell_ref like 'Controls.Row_1' or 'Connections.X1'
    """
    pin_x = _get_cell_float(shape, "PinX")
    pin_y = _get_cell_float(shape, "PinY")
    loc_pin_x = _get_cell_float(shape, "LocPinX")
    loc_pin_y = _get_cell_float(shape, "LocPinY")

    if cell_ref.startswith("Controls."):
        row_key = cell_ref.split(".", 1)[1]  # e.g. "Row_1"
        ctrl = shape.get("controls", {}).get(row_key)
        if ctrl:
            lx = _safe_float(ctrl.get("X"))
            ly = _safe_float(ctrl.get("Y"))
            # Local to page
            px = (pin_x - loc_pin_x + lx) * _INCH_TO_PX
            py = (page_h - (pin_y - loc_pin_y + ly)) * _INCH_TO_PX
            return (px, py)

    elif cell_ref.startswith("Connections."):
        # Parse "X1" -> row IX=0, "X2" -> IX=1, etc
        suffix = cell_ref.split(".", 1)[1]  # e.g. "X1"
        m = re.match(r"X(\d+)", suffix)
        if m:
            row_ix = str(int(m.group(1)) - 1)  # X1 -> IX=0
            conn = shape.get("connections", {}).get(row_ix)
            if conn:
                lx = _safe_float(conn.get("X"))
                ly = _safe_float(conn.get("Y"))
                px = (pin_x - loc_pin_x + lx) * _INCH_TO_PX
                py = (page_h - (pin_y - loc_pin_y + ly)) * _INCH_TO_PX
                return (px, py)

    return None


def _render_connections_svg(connects: list[dict], shape_index: dict,
                            page_h: float, masters: dict) -> list[str]:
    """Render connection lines as SVG elements."""
    lines = []
    for conn in connects:
        from_shape = shape_index.get(conn["from_sheet"])
        to_shape = shape_index.get(conn["to_sheet"])
        if not from_shape or not to_shape:
            continue

        # Merge with masters for connections/controls data
        from_shape = _merge_shape_with_master(
            from_shape, masters, from_shape.get("master", ""))
        to_shape = _merge_shape_with_master(
            to_shape, masters, to_shape.get("master", ""))

        from_pt = _resolve_connection_point(
            from_shape, conn["from_cell"], page_h, shape_index)
        to_pt = _resolve_connection_point(
            to_shape, conn["to_cell"], page_h, shape_index)

        if from_pt and to_pt:
            lines.append(
                f'<line x1="{from_pt[0]:.2f}" y1="{from_pt[1]:.2f}" '
                f'x2="{to_pt[0]:.2f}" y2="{to_pt[1]:.2f}" '
                f'stroke="#666666" stroke-width="0.75"/>'
            )
        elif to_pt:
            # No from-point (no Controls) — draw vertical drop from bus Y
            # Use bus shape's PinY as the connection Y
            bus_y = (page_h - _get_cell_float(from_shape, "PinY")) * _INCH_TO_PX
            lines.append(
                f'<line x1="{to_pt[0]:.2f}" y1="{bus_y:.2f}" '
                f'x2="{to_pt[0]:.2f}" y2="{to_pt[1]:.2f}" '
                f'stroke="#666666" stroke-width="0.75"/>'
            )

    return lines


def _parse_vsdx_shapes(page_xml: bytes, master_texts: dict | None = None,
                       masters: dict | None = None) -> list[dict]:
    """Parse shapes from a Visio page XML into rich shape dicts.

    Args:
        page_xml: Raw XML bytes of a page file.
        master_texts: Legacy param (ignored, kept for API compat).
        masters: Full master shapes dict from _parse_master_shapes.
    """
    shapes = []
    try:
        root = ET.fromstring(page_xml)
    except ET.ParseError:
        return shapes

    # Find all top-level shapes (direct children of Shapes element)
    shapes_container = root.find(f"{_VTAG}Shapes")
    if shapes_container is None:
        return shapes

    for shape_elem in shapes_container.findall(f"{_VTAG}Shape"):
        sd = _parse_single_shape(shape_elem)
        shapes.append(sd)

    return shapes


def _shapes_to_svg(shapes: list[dict], page_w: float, page_h: float,
                   masters: dict | None = None,
                   connects: list[dict] | None = None,
                   media: dict | None = None,
                   page_rels: dict | None = None,
                   bg_shapes: list[dict] | None = None,
                   bg_connects: list[dict] | None = None,
                   output_dir: str | None = None) -> str:
    """Generate SVG string from parsed shapes."""
    ET.register_namespace("", "http://www.w3.org/2000/svg")
    ET.register_namespace("xlink", "http://www.w3.org/1999/xlink")

    page_w_px = page_w * _INCH_TO_PX
    page_h_px = page_h * _INCH_TO_PX

    svg_lines = [
        f'<?xml version="1.0" encoding="UTF-8"?>',
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'width="{page_w_px:.0f}" height="{page_h_px:.0f}" '
        f'viewBox="0 0 {page_w_px:.0f} {page_h_px:.0f}">',
        f'<rect width="100%" height="100%" fill="white"/>',
    ]

    if masters is None:
        masters = {}
    if media is None:
        media = {}
    if page_rels is None:
        page_rels = {}

    used_markers: set[str] = set()

    # Render background page shapes first (behind foreground)
    if bg_shapes:
        svg_lines.append('<!-- Background page -->')
        for s in bg_shapes:
            svg_elements = _render_shape_svg(
                s, page_h, masters, media=media,
                page_rels=page_rels, used_markers=used_markers,
                output_dir=output_dir)
            svg_lines.extend(svg_elements)
        if bg_connects:
            bg_index = _build_shape_index(bg_shapes)
            svg_lines.extend(_render_connections_svg(
                bg_connects, bg_index, page_h, masters))

    # Render foreground shapes
    for s in shapes:
        svg_elements = _render_shape_svg(
            s, page_h, masters, media=media,
            page_rels=page_rels, used_markers=used_markers,
            output_dir=output_dir)
        svg_lines.extend(svg_elements)

    # Render connections
    if connects:
        shape_index = _build_shape_index(shapes)
        conn_lines = _render_connections_svg(connects, shape_index, page_h, masters)
        svg_lines.extend(conn_lines)

    svg_lines.append("</svg>")

    # Insert marker defs after the opening <svg> tag if needed
    if used_markers:
        marker_lines = _arrow_marker_defs(used_markers)
        # Insert after the background rect (index 3)
        for j, ml in enumerate(marker_lines):
            svg_lines.insert(3 + j, ml)

    return "\n".join(svg_lines)


# ---------------------------------------------------------------------------
# Legacy API compatibility
# ---------------------------------------------------------------------------

def _parse_master_texts(zf: zipfile.ZipFile) -> dict[str, dict[str, str]]:
    """Parse text from master shapes. Returns {master_id: {shape_id: text}}.

    Kept for API compatibility. Internally we use _parse_master_shapes now.
    """
    masters = {}
    for name in zf.namelist():
        if name.startswith("visio/masters/master") and name.endswith(".xml") and "masters.xml" not in name:
            master_num = Path(name).stem.replace("master", "")
            try:
                root = ET.fromstring(zf.read(name))
            except (ET.ParseError, KeyError):
                continue
            shape_texts = {}
            for shape in root.iter(f"{_VTAG}Shape"):
                shape_id = shape.get("ID", "")
                text_elem = shape.find(f"{_VTAG}Text")
                if text_elem is not None:
                    text = "".join(text_elem.itertext()).strip()
                    if text:
                        shape_texts[shape_id] = text
            if shape_texts:
                masters[master_num] = shape_texts
    return masters


# ---------------------------------------------------------------------------
# Public API functions
# ---------------------------------------------------------------------------

def find_vsd2xhtml() -> str | None:
    """Find vsd2xhtml from libvisio."""
    for name in ("vsd2xhtml", "vsd2raw"):
        path = shutil.which(name)
        if path:
            return path
    return None


def find_vss2xhtml() -> str | None:
    """Find vss2xhtml from libvisio."""
    path = shutil.which("vss2xhtml")
    if path:
        return path
    return None


def _convert_with_libvisio(input_path: str, output_dir: str, page: int | None = None) -> list[str]:
    """Convert using vsd2xhtml (libvisio)."""
    ext = Path(input_path).suffix.lower()

    if ext in STENCIL_EXTENSIONS:
        tool = find_vss2xhtml()
        if not tool:
            tool = find_vsd2xhtml()
    else:
        tool = find_vsd2xhtml()

    if not tool:
        return []

    basename = Path(input_path).stem

    cmd = [tool]
    if page is not None and "vsd2xhtml" in tool:
        cmd.extend(["--page", str(page)])
    cmd.append(input_path)

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=60,
    )

    if result.returncode != 0:
        return []

    xhtml_content = result.stdout
    if not xhtml_content.strip():
        return []

    svg_files = []
    try:
        ET.register_namespace("", "http://www.w3.org/2000/svg")
        ET.register_namespace("xlink", "http://www.w3.org/1999/xlink")

        root = ET.fromstring(xhtml_content)
        ns = {"svg": "http://www.w3.org/2000/svg", "xhtml": "http://www.w3.org/1999/xhtml"}
        svgs = root.findall(".//svg:svg", ns)
        if not svgs:
            svgs = root.findall(".//{http://www.w3.org/2000/svg}svg")

        for i, svg_elem in enumerate(svgs):
            svg_path = os.path.join(output_dir, f"{basename}_page{i + 1}.svg")
            svg_str = ET.tostring(svg_elem, encoding="unicode", xml_declaration=True)
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(svg_str)
            svg_files.append(svg_path)
    except ET.ParseError:
        svg_path = os.path.join(output_dir, f"{basename}.svg")
        with open(svg_path, "w", encoding="utf-8") as f:
            f.write(xhtml_content)
        svg_files.append(svg_path)

    return svg_files


def _parse_vsdx_page_names(zf: zipfile.ZipFile) -> list[str]:
    """Parse page names from pages.xml inside a .vsdx/.vssx ZIP."""
    names = []
    try:
        pages_xml = zf.read("visio/pages/pages.xml")
        root = ET.fromstring(pages_xml)
        for page in root.findall(f"{_VTAG}Page"):
            name = page.get("Name", "")
            names.append(name)
    except (KeyError, ET.ParseError):
        pass
    return names


def _get_page_files(zf: zipfile.ZipFile) -> list[str]:
    """Get sorted list of page XML files from a ZIP."""
    page_files = sorted(
        n for n in zf.namelist()
        if n.startswith("visio/pages/page") and n.endswith(".xml")
        and not n.endswith("pages.xml")
    )
    if not page_files:
        page_files = sorted(
            n for n in zf.namelist()
            if "page" in n.lower() and n.endswith(".xml")
            and "pages.xml" not in n.lower()
            and "_rels" not in n
        )
    return page_files


def _parse_background_pages(zf: zipfile.ZipFile) -> dict[int, int]:
    """Parse pages.xml to find background page references.

    Returns {page_index: background_page_index} (0-based).
    """
    bg_map = {}
    try:
        pages_xml = zf.read("visio/pages/pages.xml")
        root = ET.fromstring(pages_xml)
        pages = root.findall(f"{_VTAG}Page")

        # Build page ID -> index map
        page_id_to_idx = {}
        for i, page in enumerate(pages):
            pid = page.get("ID", "")
            if pid:
                page_id_to_idx[pid] = i

        # Find BackPage references
        for i, page in enumerate(pages):
            page_sheet = page.find(f"{_VTAG}PageSheet")
            if page_sheet is None:
                continue
            for cell in page_sheet.findall(f"{_VTAG}Cell"):
                if cell.get("N") == "BackPage":
                    back_id = cell.get("V", "")
                    if back_id and back_id in page_id_to_idx:
                        bg_map[i] = page_id_to_idx[back_id]
    except (KeyError, ET.ParseError):
        pass
    return bg_map


def _vsdx_to_svg(input_path: str, output_dir: str) -> list[str]:
    """Parse .vsdx (ZIP+XML) and generate SVG directly."""
    if not zipfile.is_zipfile(input_path):
        return []

    basename = Path(input_path).stem
    svg_files = []

    with zipfile.ZipFile(input_path, "r") as zf:
        masters = _parse_master_shapes(zf)
        media = _extract_media(zf)
        page_files = _get_page_files(zf)
        all_dims = _parse_all_page_dimensions(zf)
        bg_map = _parse_background_pages(zf)

        # Pre-parse all pages for background composition
        page_cache: dict[int, tuple] = {}  # idx -> (shapes, connects, page_rels)

        for i, page_file in enumerate(page_files):
            try:
                page_xml = zf.read(page_file)
            except (KeyError, zipfile.BadZipFile):
                continue

            shapes = _parse_vsdx_shapes(page_xml, masters=masters)
            try:
                page_root = ET.fromstring(page_xml)
                connects = _parse_connects(page_root)
            except ET.ParseError:
                connects = []

            page_rels = _parse_rels(zf, page_file)
            page_cache[i] = (shapes, connects, page_rels)

        for i, page_file in enumerate(page_files):
            if i not in page_cache:
                continue

            shapes, connects, page_rels = page_cache[i]
            if not shapes:
                continue

            if i < len(all_dims):
                page_w, page_h = all_dims[i]
            else:
                try:
                    page_xml = zf.read(page_file)
                    page_w, page_h = _parse_page_dimensions(page_xml)
                except (KeyError, zipfile.BadZipFile):
                    page_w, page_h = 8.5, 11.0

            # Background page composition
            bg_shapes = None
            bg_connects = None
            if i in bg_map:
                bg_idx = bg_map[i]
                if bg_idx in page_cache:
                    bg_shapes, bg_connects, _ = page_cache[bg_idx]

            svg_content = _shapes_to_svg(
                shapes, page_w, page_h, masters, connects,
                media, page_rels, bg_shapes, bg_connects, output_dir)
            svg_path = os.path.join(output_dir, f"{basename}_page{i + 1}.svg")
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(svg_content)
            svg_files.append(svg_path)

    return svg_files


def get_page_info(input_path: str) -> list[dict]:
    """Get page names and shape data from a Visio file.

    Returns list of dicts: [{"name": "Page-1", "shapes": [...], "index": 0}, ...]
    """
    ext = Path(input_path).suffix.lower()
    pages = []

    if ext in (".vsdx", ".vssx", ".vssm"):
        if not zipfile.is_zipfile(input_path):
            return pages

        with zipfile.ZipFile(input_path, "r") as zf:
            masters = _parse_master_shapes(zf)
            page_names = _parse_vsdx_page_names(zf)
            page_files = _get_page_files(zf)
            all_dims = _parse_all_page_dimensions(zf)

            for i, page_file in enumerate(page_files):
                try:
                    page_xml = zf.read(page_file)
                except (KeyError, zipfile.BadZipFile):
                    continue

                shapes = _parse_vsdx_shapes(page_xml, masters=masters)
                name = page_names[i] if i < len(page_names) else f"Page {i + 1}"
                page_w, page_h = all_dims[i] if i < len(all_dims) else _parse_page_dimensions(page_xml)
                pages.append({"name": name, "shapes": shapes, "index": i,
                              "page_w": page_w, "page_h": page_h})

    return pages


def extract_all_text(input_path: str) -> str:
    """Extract all text from a Visio file."""
    ext = Path(input_path).suffix.lower()

    if ext in (".vsdx", ".vssx", ".vssm"):
        pages = get_page_info(input_path)
        text_lines = []
        for page in pages:
            text_lines.append(f"--- {page['name']} ---")
            for shape in page["shapes"]:
                if shape.get("text"):
                    text_lines.append(shape["text"])
            text_lines.append("")
        return "\n".join(text_lines)

    if ext in (".vsd", ".vss"):
        tool = find_vsd2xhtml()
        if tool:
            try:
                result = subprocess.run(
                    [tool, input_path],
                    capture_output=True, text=True, timeout=60,
                )
                if result.returncode == 0:
                    try:
                        root = ET.fromstring(result.stdout)
                        texts = []
                        for elem in root.iter():
                            if elem.text and elem.text.strip():
                                texts.append(elem.text.strip())
                            if elem.tail and elem.tail.strip():
                                texts.append(elem.tail.strip())
                        return "\n".join(texts)
                    except ET.ParseError:
                        return result.stdout
            except Exception:
                pass

    return ""


def convert_vsd_to_svg(input_path: str, output_dir: str | None = None) -> list[str]:
    """Convert a Visio file to SVG pages.

    Returns a list of SVG file paths (one per page).
    Uses libvisio if available, otherwise built-in .vsdx parser.
    """
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="vsdview_")

    input_path = os.path.abspath(input_path)
    ext = Path(input_path).suffix.lower()

    if ext not in ALL_EXTENSIONS:
        raise RuntimeError(_("Unsupported file format: %s") % ext)

    # For .vsdx files, prefer built-in parser (handles images, arrows,
    # background pages). Fall back to libvisio only for .vsd/.vss.
    if ext in (".vsdx", ".vssx", ".vssm"):
        svg_files = _vsdx_to_svg(input_path, output_dir)
        if svg_files:
            return svg_files

    svg_files = _convert_with_libvisio(input_path, output_dir)
    if svg_files:
        return svg_files

    if ext in (".vsd", ".vss"):
        raise RuntimeError(
            _("Cannot open %s files without libvisio. Install it:\n"
              "  Ubuntu/Debian: sudo apt install libvisio-tools\n"
              "  Fedora: sudo dnf install libvisio-tools\n"
              "  macOS: brew install libvisio") % ext
        )

    raise RuntimeError(
        _("Could not parse the Visio file. "
          "The file may be corrupt or unsupported.\n"
          "For best results, install libvisio-tools.")
    )


def convert_vsd_page_to_svg(input_path: str, page_index: int, output_dir: str) -> str | None:
    """Convert a specific page of a Visio file to SVG. Returns SVG path or None."""
    ext = Path(input_path).suffix.lower()
    basename = Path(input_path).stem

    if ext in (".vsd", ".vss"):
        svg_files = _convert_with_libvisio(input_path, output_dir, page=page_index + 1)
        if svg_files:
            return svg_files[0]
        return None

    if ext in (".vsdx", ".vssx", ".vssm") and zipfile.is_zipfile(input_path):
        with zipfile.ZipFile(input_path, "r") as zf:
            masters = _parse_master_shapes(zf)
            media = _extract_media(zf)
            page_files = _get_page_files(zf)
            all_dims = _parse_all_page_dimensions(zf)
            bg_map = _parse_background_pages(zf)
            if page_index >= len(page_files):
                return None

            page_file = page_files[page_index]
            page_xml = zf.read(page_file)
            if page_index < len(all_dims):
                page_w, page_h = all_dims[page_index]
            else:
                page_w, page_h = _parse_page_dimensions(page_xml)
            shapes = _parse_vsdx_shapes(page_xml, masters=masters)
            if not shapes:
                return None

            try:
                page_root = ET.fromstring(page_xml)
                connects = _parse_connects(page_root)
            except ET.ParseError:
                connects = []

            page_rels = _parse_rels(zf, page_file)

            # Background page
            bg_shapes = None
            bg_connects = None
            if page_index in bg_map:
                bg_idx = bg_map[page_index]
                if bg_idx < len(page_files):
                    try:
                        bg_xml = zf.read(page_files[bg_idx])
                        bg_shapes = _parse_vsdx_shapes(bg_xml, masters=masters)
                        bg_root = ET.fromstring(bg_xml)
                        bg_connects = _parse_connects(bg_root)
                    except (KeyError, ET.ParseError):
                        pass

            svg_content = _shapes_to_svg(
                shapes, page_w, page_h, masters, connects,
                media, page_rels, bg_shapes, bg_connects, output_dir)
            svg_path = os.path.join(output_dir, f"{basename}_page{page_index + 1}.svg")
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(svg_content)
            return svg_path

    return None


def export_to_png(svg_path: str, output_path: str, width: int = 1920) -> str:
    """Export an SVG to PNG using rsvg-convert or cairosvg."""
    rsvg = shutil.which("rsvg-convert")
    if rsvg:
        subprocess.run(
            [rsvg, "-w", str(width), "-o", output_path, svg_path],
            check=True,
            timeout=30,
        )
        return output_path

    try:
        import cairosvg
        cairosvg.svg2png(url=svg_path, write_to=output_path, output_width=width)
        return output_path
    except ImportError:
        pass

    raise RuntimeError(
        _("Neither rsvg-convert nor cairosvg found. Install one:\n"
          "  Ubuntu/Debian: sudo apt install librsvg2-bin\n"
          "  pip install cairosvg")
    )


def export_to_pdf(svg_path: str, output_path: str) -> str:
    """Export an SVG to PDF using cairosvg."""
    try:
        import cairosvg
        cairosvg.svg2pdf(url=svg_path, write_to=output_path)
        return output_path
    except ImportError:
        raise RuntimeError(
            _("cairosvg is required for PDF export. Install it:\n"
              "  pip install cairosvg")
        )
