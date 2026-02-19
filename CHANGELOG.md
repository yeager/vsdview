# Changelog

## 0.4.0

Rendering quality overhaul for the built-in .vsdx parser.

### Geometry
- Fix geometry double-scaling when merging master shapes with instance overrides
- Fix IX-based geometry section merge to track master-space vs instance-space coordinates
- Add auto-close (Z) for paths where last point matches first MoveTo
- Improve NURBS parser to handle x_type/y_type coordinate flags
- Add NoShow geometry fallback: force-show basic outline when all sections are hidden
- Fix cell merge to preserve cells with formula (F) attribute even when value is empty

### Connectors
- Fix connector arrow defaults: only add default arrows for ObjType=2 shapes without explicit EndArrow
- Fix connector geometry: require both X and Y coordinates, handle incomplete rows
- Add begin point insertion for connector geometry missing initial MoveTo
- Extend connector row type support (EllipticalArcTo, NURBSTo, SplineStart, SplineKnot)

### Styling and themes
- Fix gradient fill direction: FillBkgnd is start, FillForegnd is end
- Handle same-color gradients by using solid fill instead
- Fix fill_opacity variable used before assignment in container style block

### Text
- Reduce text abbreviation threshold from 40px to 30px width
- Improve text wrapping, auto-sizing, and vertical alignment

## 0.1.0

- Initial release
- Open and view .vsdx/.vsd files via LibreOffice headless conversion
- GTK4/Adwaita UI with SVG rendering (Rsvg + Cairo)
- Zoom: Ctrl+Plus/Minus/0, scroll wheel with Ctrl
- Light/dark theme toggle
- Keyboard shortcuts dialog (Ctrl+/)
- Export as PNG (Ctrl+E)
- Refresh current file (F5)
- Status bar with filename and zoom level
- Drag and drop support for .vsdx/.vsd files
- Recent files list (stored in ~/.local/share/vsdview/recent.json)
- Debug info in About dialog
- Desktop notifications for errors
- i18n support (gettext), .pot file for Transifex
- Man page (man/vsdview.1)
- Debian and RPM packaging
- GitHub Actions for CI/CD
