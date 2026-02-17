# VSDView

A minimal read-only viewer for Microsoft Visio files (.vsdx/.vsd), built with GTK4 and libadwaita.

## Requirements

- Python 3.10+
- GTK4, libadwaita, librsvg (GObject Introspection)
- LibreOffice (for headless conversion of Visio → SVG)

## Installation

### From .deb (Ubuntu/Debian)

```bash
sudo dpkg -i vsdview_0.1.0-1_all.deb
sudo apt-get install -f
```

### From .rpm (Fedora)

```bash
sudo dnf install vsdview-0.1.0-1.noarch.rpm
```

### From source

```bash
pip install .
```

## Usage

```bash
vsdview [file.vsdx]
```

You can also drag and drop .vsdx/.vsd files onto the window.

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| Ctrl+O | Open file |
| Ctrl+Q | Quit |
| Ctrl+Plus | Zoom in |
| Ctrl+Minus | Zoom out |
| Ctrl+0 | Fit to window |
| F5 | Refresh |
| Ctrl+E | Export as PNG |
| Ctrl+/ | Keyboard shortcuts |

## Features

- SVG rendering via librsvg + Cairo
- Zoom (keyboard, Ctrl+scroll wheel)
- Light/dark theme toggle
- Drag and drop support
- Recent files list
- Export to PNG
- Desktop notifications for errors
- i18n ready (gettext)

## Translation

Translations are managed via [Transifex](https://www.transifex.com/danielnylander/vsdview/).

## License

GPL-3.0-or-later. See [LICENSE](LICENSE).

## Trademark Notice

Microsoft Visio is a registered trademark of Microsoft Corporation. VSDView is not affiliated with or endorsed by Microsoft.
