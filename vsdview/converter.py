"""Convert Visio files (.vsdx, .vsd) to SVG.

Backend priority:
1. libvisio (vsd2xhtml) — lightweight, accurate
2. Built-in .vsdx XML parser — zero dependencies, .vsdx only
"""

import gettext
import os
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


def find_vsd2xhtml() -> str | None:
    """Find vsd2xhtml from libvisio."""
    for name in ("vsd2xhtml", "vsd2raw"):
        path = shutil.which(name)
        if path:
            return path
    return None


def _convert_with_libvisio(input_path: str, output_dir: str) -> list[str]:
    """Convert using vsd2xhtml (libvisio)."""
    vsd2xhtml = find_vsd2xhtml()
    if not vsd2xhtml:
        return []

    basename = Path(input_path).stem
    output_file = os.path.join(output_dir, f"{basename}.xhtml")

    result = subprocess.run(
        [vsd2xhtml, input_path],
        capture_output=True,
        text=True,
        timeout=60,
    )

    if result.returncode != 0:
        return []

    # vsd2xhtml outputs XHTML with embedded SVG to stdout
    xhtml_content = result.stdout
    if not xhtml_content.strip():
        return []

    # Extract SVG elements from XHTML
    svg_files = []
    try:
        root = ET.fromstring(xhtml_content)
        ns = {"svg": "http://www.w3.org/2000/svg", "xhtml": "http://www.w3.org/1999/xhtml"}
        svgs = root.findall(".//svg:svg", ns)
        if not svgs:
            # Try without namespace
            svgs = root.findall(".//{http://www.w3.org/2000/svg}svg")

        for i, svg_elem in enumerate(svgs):
            svg_path = os.path.join(output_dir, f"{basename}_page{i + 1}.svg")
            svg_str = ET.tostring(svg_elem, encoding="unicode", xml_declaration=True)
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(svg_str)
            svg_files.append(svg_path)
    except ET.ParseError:
        # If XHTML parsing fails, save raw output as single SVG
        svg_path = os.path.join(output_dir, f"{basename}.svg")
        with open(svg_path, "w", encoding="utf-8") as f:
            f.write(xhtml_content)
        svg_files.append(svg_path)

    return svg_files


def _parse_vsdx_shapes(page_xml: bytes) -> list[dict]:
    """Parse shapes from a Visio page XML."""
    shapes = []
    try:
        root = ET.fromstring(page_xml)
    except ET.ParseError:
        return shapes

    for shape in root.iter(f"{{{_NS['v']}}}Shape"):
        s = {"type": "shape", "x": 0, "y": 0, "w": 72, "h": 36, "text": ""}

        # Get geometry
        for cell in shape.findall(f"{{{_NS['v']}}}Cell"):
            name = cell.get("N", "")
            val = cell.get("V", "0")
            try:
                fval = float(val)
            except (ValueError, TypeError):
                fval = 0

            # Visio uses inches, convert to points (72 dpi)
            if name == "PinX":
                s["x"] = fval * 72
            elif name == "PinY":
                s["y"] = fval * 72
            elif name == "Width":
                s["w"] = fval * 72
            elif name == "Height":
                s["h"] = fval * 72

        # Get text
        text_elem = shape.find(f"{{{_NS['v']}}}Text")
        if text_elem is not None:
            s["text"] = "".join(text_elem.itertext()).strip()

        shapes.append(s)

    return shapes


def _vsdx_to_svg(input_path: str, output_dir: str) -> list[str]:
    """Parse .vsdx (ZIP+XML) and generate SVG directly."""
    if not zipfile.is_zipfile(input_path):
        return []

    basename = Path(input_path).stem
    svg_files = []

    with zipfile.ZipFile(input_path, "r") as zf:
        # Find page files
        page_files = sorted(
            n for n in zf.namelist()
            if n.startswith("visio/pages/page") and n.endswith(".xml")
            and not n.endswith("pages.xml")
        )

        if not page_files:
            # Try alternate paths
            page_files = sorted(
                n for n in zf.namelist()
                if "page" in n.lower() and n.endswith(".xml")
                and "pages.xml" not in n.lower()
                and "_rels" not in n
            )

        for i, page_file in enumerate(page_files):
            try:
                page_xml = zf.read(page_file)
            except (KeyError, zipfile.BadZipFile):
                continue

            shapes = _parse_vsdx_shapes(page_xml)
            if not shapes:
                continue

            # Calculate page bounds
            max_x = max((s["x"] + s["w"] / 2) for s in shapes) if shapes else 800
            max_y = max((s["y"] + s["h"] / 2) for s in shapes) if shapes else 600
            page_w = max(max_x + 50, 400)
            page_h = max(max_y + 50, 300)

            # Generate SVG
            svg_lines = [
                f'<?xml version="1.0" encoding="UTF-8"?>',
                f'<svg xmlns="http://www.w3.org/2000/svg" '
                f'width="{page_w:.0f}" height="{page_h:.0f}" '
                f'viewBox="0 0 {page_w:.0f} {page_h:.0f}">',
                f'<rect width="100%" height="100%" fill="white"/>',
            ]

            for s in shapes:
                # Flip Y axis (Visio Y goes up, SVG Y goes down)
                sx = s["x"] - s["w"] / 2
                sy = page_h - s["y"] - s["h"] / 2
                sw = s["w"]
                sh = s["h"]

                # Draw rectangle
                svg_lines.append(
                    f'<rect x="{sx:.1f}" y="{sy:.1f}" '
                    f'width="{sw:.1f}" height="{sh:.1f}" '
                    f'fill="#e8f0fe" stroke="#4285f4" stroke-width="1.5" rx="4"/>'
                )

                # Draw text
                if s["text"]:
                    tx = s["x"]
                    ty = page_h - s["y"]
                    text_escaped = (
                        s["text"]
                        .replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                    )
                    svg_lines.append(
                        f'<text x="{tx:.1f}" y="{ty:.1f}" '
                        f'text-anchor="middle" dominant-baseline="central" '
                        f'font-family="sans-serif" font-size="11" fill="#333">'
                        f'{text_escaped}</text>'
                    )

            svg_lines.append("</svg>")

            svg_path = os.path.join(output_dir, f"{basename}_page{i + 1}.svg")
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write("\n".join(svg_lines))
            svg_files.append(svg_path)

    return svg_files


def convert_vsd_to_svg(input_path: str, output_dir: str | None = None) -> list[str]:
    """Convert a Visio file to SVG pages.

    Returns a list of SVG file paths (one per page).
    Uses libvisio if available, otherwise built-in .vsdx parser.
    """
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="vsdview_")

    input_path = os.path.abspath(input_path)

    # Try libvisio first (handles both .vsd and .vsdx)
    svg_files = _convert_with_libvisio(input_path, output_dir)
    if svg_files:
        return svg_files

    # Fall back to built-in .vsdx parser
    ext = Path(input_path).suffix.lower()
    if ext == ".vsdx":
        svg_files = _vsdx_to_svg(input_path, output_dir)
        if svg_files:
            return svg_files

    if ext == ".vsd":
        raise RuntimeError(
            _("Cannot open .vsd files without libvisio. Install it:\n"
              "  Ubuntu/Debian: sudo apt install libvisio-tools\n"
              "  Fedora: sudo dnf install libvisio-tools\n"
              "  macOS: brew install libvisio")
        )

    raise RuntimeError(
        _("Could not parse the Visio file. "
          "The file may be corrupt or unsupported.\n"
          "For best results, install libvisio-tools.")
    )


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
