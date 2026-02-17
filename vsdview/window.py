"""VSDView main window."""

import os
import subprocess
import tempfile

import gi

gi.require_version("Gtk", "4.0")
gi.require_version("Adw", "1")
gi.require_version("Rsvg", "2.0")

from gi.repository import Adw, Gio, Gtk, Rsvg


class VSDViewWindow(Adw.ApplicationWindow):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.set_default_size(900, 700)
        self.set_title("VSDView")

        self._svg_handle = None

        # Layout
        box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.set_content(box)

        # Header bar
        header = Adw.HeaderBar()
        box.append(header)

        open_btn = Gtk.Button(icon_name="document-open-symbolic")
        open_btn.set_tooltip_text("Open Visio file")
        open_btn.connect("clicked", lambda *_: self.show_open_dialog())
        header.pack_start(open_btn)

        menu_btn = Gtk.MenuButton(icon_name="open-menu-symbolic")
        menu = Gio.Menu()
        menu.append("About", "app.about")
        menu.append("Quit", "app.quit")
        menu_btn.set_menu_model(menu)
        header.pack_end(menu_btn)

        # Drawing area
        self._drawing_area = Gtk.DrawingArea()
        self._drawing_area.set_vexpand(True)
        self._drawing_area.set_hexpand(True)
        self._drawing_area.set_draw_func(self._on_draw)
        box.append(self._drawing_area)

    def show_open_dialog(self):
        dialog = Gtk.FileDialog()
        file_filter = Gtk.FileFilter()
        file_filter.set_name("Visio files")
        file_filter.add_pattern("*.vsdx")
        file_filter.add_pattern("*.vsd")
        filters = Gio.ListStore.new(Gtk.FileFilter)
        filters.append(file_filter)
        dialog.set_filters(filters)
        dialog.open(self, None, self._on_file_chosen)

    def _on_file_chosen(self, dialog, result):
        try:
            file = dialog.open_finish(result)
            if file:
                self.open_file(file.get_path())
        except Exception:
            pass

    def open_file(self, path):
        """Convert a Visio file to SVG via LibreOffice and display it."""
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                subprocess.run(
                    [
                        "lowriter",
                        "--headless",
                        "--convert-to",
                        "svg",
                        "--outdir",
                        tmpdir,
                        path,
                    ],
                    check=True,
                    capture_output=True,
                    timeout=30,
                )
            except (subprocess.CalledProcessError, FileNotFoundError) as e:
                self._show_error(f"Conversion failed: {e}")
                return

            base = os.path.splitext(os.path.basename(path))[0]
            svg_path = os.path.join(tmpdir, base + ".svg")
            if not os.path.exists(svg_path):
                self._show_error("No SVG output produced.")
                return

            self._svg_handle = Rsvg.Handle.new_from_file(svg_path)

        self.set_title(f"VSDView — {os.path.basename(path)}")
        self._drawing_area.queue_draw()

    def _on_draw(self, area, cr, width, height):
        if not self._svg_handle:
            return

        viewport = Rsvg.Rectangle()
        viewport.x = 0
        viewport.y = 0
        viewport.width = width
        viewport.height = height
        self._svg_handle.render_document(cr, viewport)

    def _show_error(self, message):
        dialog = Adw.MessageDialog(
            transient_for=self,
            heading="Error",
            body=message,
        )
        dialog.add_response("ok", "OK")
        dialog.present()
