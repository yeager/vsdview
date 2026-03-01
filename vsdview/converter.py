"""Convert Visio files (.vsdx, .vsd, .vstx, .vssx, .vss) to SVG.

Uses libvisio-ng for all Visio file parsing — supports both .vsdx (XML)
and .vsd (binary) formats natively in Python. No external dependencies needed.

Author: Daniel Nylander <daniel@danielnylander.se>
"""

from libvisio_ng import (
    ALL_EXTENSIONS,
    STENCIL_EXTENSIONS,
    TEMPLATE_EXTENSIONS,
    VISIO_EXTENSIONS,
    convert as convert_vsd_to_svg,
    convert_page as convert_vsd_page_to_svg,
    export_to_pdf,
    export_to_png,
    extract_text as extract_all_text,
    get_page_info,
)

__all__ = [
    "ALL_EXTENSIONS",
    "STENCIL_EXTENSIONS",
    "TEMPLATE_EXTENSIONS",
    "VISIO_EXTENSIONS",
    "convert_vsd_to_svg",
    "convert_vsd_page_to_svg",
    "export_to_pdf",
    "export_to_png",
    "extract_all_text",
    "get_page_info",
]
