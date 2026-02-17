"""Convert Visio files (.vsdx, .vsd, .vssx, .vss) to SVG.

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

# Supported file extensions
VISIO_EXTENSIONS = {".vsd", ".vsdx", ".vsdm"}
STENCIL_EXTENSIONS = {".vss", ".vssx", ".vssm"}
ALL_EXTENSIONS = VISIO_EXTENSIONS | STENCIL_EXTENSIONS


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


def _parse_vsdx_shapes(page_xml: bytes) -> list[dict]:
    """Parse shapes from a Visio page XML."""
    shapes = []
    try:
        root = ET.fromstring(page_xml)
    except ET.ParseError:
        return shapes

    for shape in root.iter(f"{{{_NS['v']}}}Shape"):
        s = {"type": "shape", "x": 0, "y": 0, "w": 72, "h": 36, "text": ""}

        for cell in shape.findall(f"{{{_NS['v']}}}Cell"):
            name = cell.get("N", "")
            val = cell.get("V", "0")
            try:
                fval = float(val)
            except (ValueError, TypeError):
                fval = 0

            if name == "PinX":
                s["x"] = fval * 72
            elif name == "PinY":
                s["y"] = fval * 72
            elif name == "Width":
                s["w"] = fval * 72
            elif name == "Height":
                s["h"] = fval * 72

        text_elem = shape.find(f"{{{_NS['v']}}}Text")
        if text_elem is not None:
            s["text"] = "".join(text_elem.itertext()).strip()

        shapes.append(s)

    return shapes


def _parse_vsdx_page_names(zf: zipfile.ZipFile) -> list[str]:
    """Parse page names from pages.xml inside a .vsdx/.vssx ZIP."""
    names = []
    try:
        pages_xml = zf.read("visio/pages/pages.xml")
        root = ET.fromstring(pages_xml)
        for page in root.findall(f"{{{_NS['v']}}}Page"):
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


def _shapes_to_svg(shapes: list[dict], page_w: float, page_h: float) -> str:
    """Generate SVG string from parsed shapes."""
    svg_lines = [
        f'<?xml version="1.0" encoding="UTF-8"?>',
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{page_w:.0f}" height="{page_h:.0f}" '
        f'viewBox="0 0 {page_w:.0f} {page_h:.0f}">',
        f'<rect width="100%" height="100%" fill="white"/>',
    ]

    for s in shapes:
        sx = s["x"] - s["w"] / 2
        sy = page_h - s["y"] - s["h"] / 2
        sw = s["w"]
        sh = s["h"]

        svg_lines.append(
            f'<rect x="{sx:.1f}" y="{sy:.1f}" '
            f'width="{sw:.1f}" height="{sh:.1f}" '
            f'fill="#e8f0fe" stroke="#4285f4" stroke-width="1.5" rx="4"/>'
        )

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
    return "\n".join(svg_lines)


def _vsdx_to_svg(input_path: str, output_dir: str) -> list[str]:
    """Parse .vsdx (ZIP+XML) and generate SVG directly."""
    if not zipfile.is_zipfile(input_path):
        return []

    basename = Path(input_path).stem
    svg_files = []

    with zipfile.ZipFile(input_path, "r") as zf:
        page_files = _get_page_files(zf)

        for i, page_file in enumerate(page_files):
            try:
                page_xml = zf.read(page_file)
            except (KeyError, zipfile.BadZipFile):
                continue

            shapes = _parse_vsdx_shapes(page_xml)
            if not shapes:
                continue

            max_x = max((s["x"] + s["w"] / 2) for s in shapes) if shapes else 800
            max_y = max((s["y"] + s["h"] / 2) for s in shapes) if shapes else 600
            page_w = max(max_x + 50, 400)
            page_h = max(max_y + 50, 300)

            svg_content = _shapes_to_svg(shapes, page_w, page_h)
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
            page_names = _parse_vsdx_page_names(zf)
            page_files = _get_page_files(zf)

            for i, page_file in enumerate(page_files):
                try:
                    page_xml = zf.read(page_file)
                except (KeyError, zipfile.BadZipFile):
                    continue

                shapes = _parse_vsdx_shapes(page_xml)
                name = page_names[i] if i < len(page_names) else f"Page {i + 1}"
                pages.append({"name": name, "shapes": shapes, "index": i})

    return pages


def extract_all_text(input_path: str) -> str:
    """Extract all text from a Visio file."""
    ext = Path(input_path).suffix.lower()

    if ext in (".vsdx", ".vssx", ".vssm"):
        pages = get_page_info(input_path)
        lines = []
        for page in pages:
            lines.append(f"--- {page['name']} ---")
            for shape in page["shapes"]:
                if shape.get("text"):
                    lines.append(shape["text"])
            lines.append("")
        return "\n".join(lines)

    # For .vsd, try vsd2xhtml text extraction
    if ext in (".vsd", ".vss"):
        tool = find_vsd2xhtml()
        if tool:
            try:
                result = subprocess.run(
                    [tool, input_path],
                    capture_output=True, text=True, timeout=60,
                )
                if result.returncode == 0:
                    # Extract text from XHTML
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

    # Try libvisio first (handles both .vsd and .vsdx)
    svg_files = _convert_with_libvisio(input_path, output_dir)
    if svg_files:
        return svg_files

    # Fall back to built-in .vsdx/.vssx parser
    if ext in (".vsdx", ".vssx", ".vssm"):
        svg_files = _vsdx_to_svg(input_path, output_dir)
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

    # Try libvisio with --page
    if ext in (".vsd", ".vss"):
        svg_files = _convert_with_libvisio(input_path, output_dir, page=page_index + 1)
        if svg_files:
            return svg_files[0]
        return None

    # Built-in parser for .vsdx/.vssx
    if ext in (".vsdx", ".vssx", ".vssm") and zipfile.is_zipfile(input_path):
        with zipfile.ZipFile(input_path, "r") as zf:
            page_files = _get_page_files(zf)
            if page_index >= len(page_files):
                return None

            page_xml = zf.read(page_files[page_index])
            shapes = _parse_vsdx_shapes(page_xml)
            if not shapes:
                return None

            max_x = max((s["x"] + s["w"] / 2) for s in shapes)
            max_y = max((s["y"] + s["h"] / 2) for s in shapes)
            page_w = max(max_x + 50, 400)
            page_h = max(max_y + 50, 300)

            svg_content = _shapes_to_svg(shapes, page_w, page_h)
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
