"""Convert Visio files (.vsdx, .vsd) to SVG using LibreOffice headless."""

import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def find_libreoffice() -> str | None:
    """Find LibreOffice binary."""
    for name in ("libreoffice", "lowriter", "soffice"):
        path = shutil.which(name)
        if path:
            return path
    # Common paths
    for p in (
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/snap/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.isfile(p):
            return p
    return None


def convert_vsd_to_svg(input_path: str, output_dir: str | None = None) -> list[str]:
    """Convert a Visio file to SVG pages.

    Returns a list of SVG file paths (one per page).
    """
    lo = find_libreoffice()
    if not lo:
        raise RuntimeError(
            "LibreOffice not found. Install it:\n"
            "  Ubuntu/Debian: sudo apt install libreoffice\n"
            "  Fedora: sudo dnf install libreoffice\n"
            "  macOS: brew install --cask libreoffice"
        )

    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="vsdview_")

    input_path = os.path.abspath(input_path)
    basename = Path(input_path).stem

    # Convert to SVG via LibreOffice headless
    result = subprocess.run(
        [
            lo,
            "--headless",
            "--convert-to", "svg",
            "--outdir", output_dir,
            input_path,
        ],
        capture_output=True,
        text=True,
        timeout=60,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed:\n{result.stderr}"
        )

    # Collect output SVG files
    svg_files = sorted(
        str(p)
        for p in Path(output_dir).glob(f"{basename}*.svg")
    )

    if not svg_files:
        # Try PDF as intermediate
        result2 = subprocess.run(
            [
                lo,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                input_path,
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result2.returncode == 0:
            pdf_path = os.path.join(output_dir, f"{basename}.pdf")
            if os.path.exists(pdf_path):
                # Convert PDF pages to SVG using pdftocairo if available
                pdftocairo = shutil.which("pdftocairo")
                if pdftocairo:
                    subprocess.run(
                        [pdftocairo, "-svg", pdf_path,
                         os.path.join(output_dir, basename)],
                        capture_output=True,
                        timeout=60,
                    )
                    svg_files = sorted(
                        str(p)
                        for p in Path(output_dir).glob(f"{basename}*.svg")
                    )

    if not svg_files:
        raise RuntimeError(
            "Conversion produced no SVG output. "
            "The file may be corrupt or unsupported."
        )

    return svg_files


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

    # Fallback: try cairosvg
    try:
        import cairosvg
        cairosvg.svg2png(url=svg_path, write_to=output_path, output_width=width)
        return output_path
    except ImportError:
        pass

    raise RuntimeError(
        "Neither rsvg-convert nor cairosvg found. Install one:\n"
        "  Ubuntu/Debian: sudo apt install librsvg2-bin\n"
        "  pip install cairosvg"
    )
