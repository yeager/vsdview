"""VSDView main window."""

import gettext
import os
import tempfile

import gi

gi.require_version("Gtk", "4.0")
gi.require_version("Adw", "1")
gi.require_version("Rsvg", "2.0")

from gi.repository import Adw, Gdk, Gio, GLib, Gtk, Rsvg

from vsdview.converter import convert_vsd_to_svg, export_to_png

_ = gettext.gettext


class VSDViewWindow(Adw.ApplicationWindow):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.set_default_size(900, 700)
        self.set_title("VSDView")

        self._svg_handle = None
        self._current_file = None
        self._zoom_level = 1.0
        self._svg_dir = None  # keep converted SVGs alive

        # Layout
        box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.set_content(box)

        # Header bar
        header = Adw.HeaderBar()
        box.append(header)

        open_btn = Gtk.Button(icon_name="document-open-symbolic")
        open_btn.set_tooltip_text(_("Open Visio file"))
        open_btn.connect("clicked", lambda *_: self.show_open_dialog())
        header.pack_start(open_btn)

        # Hamburger menu
        menu_btn = Gtk.MenuButton(icon_name="open-menu-symbolic")
        menu = Gio.Menu()

        # Recent files submenu
        self._recent_menu = Gio.Menu()
        menu.append_submenu(_("Recent Files"), self._recent_menu)

        menu.append(_("Toggle Dark Theme"), "app.toggle-theme")
        menu.append(_("Keyboard Shortcuts"), "app.show-shortcuts")
        menu.append(_("Export as PNG"), "app.export-png")
        menu.append(_("About"), "app.about")
        menu.append(_("Quit"), "app.quit")
        menu_btn.set_menu_model(menu)
        header.pack_end(menu_btn)

        # Scrolled window for drawing area
        scrolled = Gtk.ScrolledWindow()
        scrolled.set_vexpand(True)
        scrolled.set_hexpand(True)
        box.append(scrolled)

        # Drawing area
        self._drawing_area = Gtk.DrawingArea()
        self._drawing_area.set_vexpand(True)
        self._drawing_area.set_hexpand(True)
        self._drawing_area.set_draw_func(self._on_draw)
        scrolled.set_child(self._drawing_area)

        # Scroll zoom controller
        scroll_ctrl = Gtk.EventControllerScroll.new(
            Gtk.EventControllerScrollFlags.VERTICAL
        )
        scroll_ctrl.connect("scroll", self._on_scroll_zoom)
        self._drawing_area.add_controller(scroll_ctrl)

        # Status bar
        self._status_bar = Gtk.Label(label=_("No file loaded"))
        self._status_bar.set_xalign(0)
        self._status_bar.add_css_class("caption")
        self._status_bar.set_margin_start(8)
        self._status_bar.set_margin_end(8)
        self._status_bar.set_margin_top(4)
        self._status_bar.set_margin_bottom(4)
        box.append(self._status_bar)

        # Drag and drop
        drop_target = Gtk.DropTarget.new(Gio.File, Gdk.DragAction.COPY)
        drop_target.connect("drop", self._on_drop)
        self.add_controller(drop_target)

        # Populate recent files menu
        self._update_recent_menu()

    def _update_recent_menu(self):
        self._recent_menu.remove_all()
        app = self.get_application()
        if not app:
            return
        for path in app.recent.get_files():
            basename = os.path.basename(path)
            # Use a detailed action with the path
            item = Gio.MenuItem.new(basename, None)
            action_name = f"win.open-recent-{hash(path) & 0xFFFFFFFF}"
            action = Gio.SimpleAction.new(action_name.replace("win.", ""), None)
            action.connect("activate", lambda a, p, fp=path: self.open_file(fp))
            self.add_action(action)
            item.set_action_and_target_value(action_name, None)
            self._recent_menu.append_item(item)

    def _update_status(self):
        if self._current_file:
            basename = os.path.basename(self._current_file)
            zoom_pct = int(self._zoom_level * 100)
            self._status_bar.set_label(f"{basename} — {zoom_pct}%")
        else:
            self._status_bar.set_label(_("No file loaded"))

    def show_open_dialog(self):
        dialog = Gtk.FileDialog()
        file_filter = Gtk.FileFilter()
        file_filter.set_name(_("Visio files"))
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
        """Convert a Visio file to SVG and display it."""

        tmpdir = tempfile.mkdtemp(prefix="vsdview_")
        try:
            svg_files = convert_vsd_to_svg(path, tmpdir)
        except RuntimeError as e:
            self._send_notification(
                _("Conversion failed"),
                str(e),
            )
            self._show_error(str(e))
            return

        if not svg_files:
            self._show_error(_("No SVG output produced."))
            return

        self._svg_dir = tmpdir  # prevent cleanup
        self._svg_handle = Rsvg.Handle.new_from_file(svg_files[0])
        self._current_file = path
        self._zoom_level = 1.0

        self.set_title(f"VSDView — {os.path.basename(path)}")
        self._update_status()
        self._drawing_area.queue_draw()

        # Add to recent files
        app = self.get_application()
        if app:
            app.recent.add_file(path)
            self._update_recent_menu()

    def refresh(self):
        """Re-convert and reload the current file."""
        if self._current_file and os.path.exists(self._current_file):
            self.open_file(self._current_file)

    def zoom_in(self):
        self._zoom_level = min(self._zoom_level * 1.25, 10.0)
        self._update_status()
        self._drawing_area.queue_draw()

    def zoom_out(self):
        self._zoom_level = max(self._zoom_level / 1.25, 0.1)
        self._update_status()
        self._drawing_area.queue_draw()

    def zoom_fit(self):
        self._zoom_level = 1.0
        self._update_status()
        self._drawing_area.queue_draw()

    def _on_scroll_zoom(self, controller, dx, dy):
        state = controller.get_current_event_state()
        if state & Gdk.ModifierType.CONTROL_MASK:
            if dy < 0:
                self.zoom_in()
            elif dy > 0:
                self.zoom_out()
            return True
        return False

    def export_png(self):
        """Export current SVG as PNG."""
        if not self._svg_handle:
            self._show_error(_("No file loaded to export."))
            return

        dialog = Gtk.FileDialog()
        dialog.set_initial_name("export.png")
        dialog.save(self, None, self._on_export_chosen)

    def _on_export_chosen(self, dialog, result):
        try:
            file = dialog.save_finish(result)
            if not file:
                return
            output_path = file.get_path()
        except Exception:
            return

        # Find the SVG file to export from
        if self._svg_dir and self._current_file:
            from pathlib import Path

            basename = Path(self._current_file).stem
            svg_files = sorted(Path(self._svg_dir).glob(f"{basename}*.svg"))
            if svg_files:
                try:
                    export_to_png(str(svg_files[0]), output_path)
                except RuntimeError as e:
                    self._show_error(str(e))

    def _on_draw(self, area, cr, width, height):
        if not self._svg_handle:
            return

        viewport = Rsvg.Rectangle()
        viewport.x = 0
        viewport.y = 0
        viewport.width = width * self._zoom_level
        viewport.height = height * self._zoom_level

        # Set drawing area size for scrolling when zoomed
        self._drawing_area.set_content_width(int(width * self._zoom_level))
        self._drawing_area.set_content_height(int(height * self._zoom_level))

        self._svg_handle.render_document(cr, viewport)

    def _on_drop(self, target, value, x, y):
        if isinstance(value, Gio.File):
            path = value.get_path()
            if path and (path.endswith(".vsdx") or path.endswith(".vsd")):
                self.open_file(path)
                return True
        return False

    def _show_error(self, message):
        dialog = Adw.MessageDialog(
            transient_for=self,
            heading=_("Error"),
            body=message,
        )
        dialog.add_response("ok", _("OK"))
        dialog.present()

    def _send_notification(self, title, body):
        """Send a desktop notification."""
        app = self.get_application()
        if not app:
            return
        notification = Gio.Notification.new(title)
        notification.set_body(body)
        app.send_notification(None, notification)
