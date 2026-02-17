# VSDView

A minimal read-only viewer for Microsoft Visio files (.vsdx/.vsd), built with GTK4 and libadwaita.

## Requirements

- Python 3.10+
- GTK4, libadwaita, librsvg (GObject Introspection)
- LibreOffice (for headless conversion of Visio → SVG)

## Usage

```bash
python -m vsdview [file.vsdx]
```

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| Ctrl+O | Open file |
| Ctrl+Q | Quit |

## License

GPL-3.0-or-later. See [LICENSE](LICENSE).
