"""VSDView main window."""

import gettext
import os
import tempfile

import gi

gi.require_version("Gtk", "4.0")
gi.require_version("Adw", "1")
gi.require_version("Rsvg", "2.0")

from gi.repository import Adw, Gdk, Gio, GLib, Gtk, Rsvg

from vsdview.converter import (
    ALL_EXTENSIONS,
    STENCIL_EXTENSIONS,
    convert_vsd_to_svg,
    convert_vsd_page_to_svg,
    export_to_png,
    export_to_pdf,
    extract_all_text,
    get_page_info,
)

_ = gettext.gettext


class VSDViewWindow(Adw.ApplicationWindow):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.set_default_size(900, 700)
        self.set_title("VSDView")

        self._svg_handle = None
        self._current_file = None
        self._zoom_level = 1.0
        self._svg_dir = None
        self._svg_files = []  # all page SVGs
        self._current_page = 0
        self._page_info = []  # page metadata with shapes
        self._search_query = ""

        # Main layout
        main_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.set_content(main_box)

        # Header bar
        header = Adw.HeaderBar()
        main_box.append(header)

        open_btn = Gtk.Button(icon_name="document-open-symbolic")
        open_btn.set_tooltip_text(_("Open Visio file"))
        open_btn.connect("clicked", lambda *_: self.show_open_dialog())
        header.pack_start(open_btn)

        # Search button
        self._search_button = Gtk.ToggleButton(icon_name="edit-find-symbolic")
        self._search_button.set_tooltip_text(_("Find text"))
        self._search_button.connect("toggled", self._on_search_toggled)
        header.pack_start(self._search_button)

        # Hamburger menu
        menu_btn = Gtk.MenuButton(icon_name="open-menu-symbolic")
        menu = Gio.Menu()

        self._recent_menu = Gio.Menu()
        menu.append_submenu(_("Recent Files"), self._recent_menu)

        export_section = Gio.Menu()
        export_section.append(_("Export as PNG"), "app.export-png")
        export_section.append(_("Export as PDF"), "app.export-pdf")
        export_section.append(_("Export as Text"), "app.export-text")
        menu.append_section(_("Export"), export_section)

        menu.append(_("Copy Text"), "app.copy-text")
        menu.append(_("Toggle Dark Theme"), "app.toggle-theme")
        menu.append(_("Keyboard Shortcuts"), "app.show-shortcuts")
        menu.append(_("About"), "app.about")
        menu.append(_("Quit"), "app.quit")
        menu_btn.set_menu_model(menu)
        header.pack_end(menu_btn)

        # Search bar
        self._search_bar = Gtk.SearchBar()
        self._search_entry = Gtk.SearchEntry()
        self._search_entry.set_hexpand(True)
        self._search_entry.connect("search-changed", self._on_search_changed)
        self._search_entry.connect("stop-search", self._on_search_stop)
        clamp = Adw.Clamp(maximum_size=500)
        clamp.set_child(self._search_entry)
        self._search_bar.set_child(clamp)
        self._search_bar.connect_entry(self._search_entry)
        self._search_bar.set_search_mode(False)
        main_box.append(self._search_bar)

        # Search results label
        self._search_results_label = Gtk.Label(label="")
        self._search_results_label.set_xalign(0)
        self._search_results_label.add_css_class("caption")
        self._search_results_label.set_margin_start(8)
        self._search_results_label.set_visible(False)
        main_box.append(self._search_results_label)

        # Scrolled window for drawing area
        scrolled = Gtk.ScrolledWindow()
        scrolled.set_vexpand(True)
        scrolled.set_hexpand(True)
        main_box.append(scrolled)

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

        # Right-click context menu
        self._setup_context_menu()

        # Page tabs (Gtk.Notebook at bottom, hidden by default)
        self._notebook = Gtk.Notebook()
        self._notebook.set_show_tabs(True)
        self._notebook.set_visible(False)
        self._notebook.connect("switch-page", self._on_page_switched)
        main_box.append(self._notebook)

        # Status bar
        self._status_bar = Gtk.Label(label=_("No file loaded"))
        self._status_bar.set_xalign(0)
        self._status_bar.add_css_class("caption")
        self._status_bar.set_margin_start(8)
        self._status_bar.set_margin_end(8)
        self._status_bar.set_margin_top(4)
        self._status_bar.set_margin_bottom(4)
        main_box.append(self._status_bar)

        # Drag and drop
        drop_target = Gtk.DropTarget.new(Gio.File, Gdk.DragAction.COPY)
        drop_target.connect("drop", self._on_drop)
        self.add_controller(drop_target)

        # Key controller for Escape to close search
        key_ctrl = Gtk.EventControllerKey.new()
        key_ctrl.connect("key-pressed", self._on_key_pressed)
        self.add_controller(key_ctrl)

        self._update_recent_menu()

    def _setup_context_menu(self):
        """Set up right-click context menu."""
        gesture = Gtk.GestureClick.new()
        gesture.set_button(3)  # right click
        gesture.connect("pressed", self._on_right_click)
        self._drawing_area.add_controller(gesture)

        self._context_menu = Gio.Menu()
        self._context_menu.append(_("Copy Text"), "app.copy-text")
        self._context_menu.append(_("Export as PNG"), "app.export-png")
        self._context_menu.append(_("Export as PDF"), "app.export-pdf")
        self._context_menu.append(_("Export as Text"), "app.export-text")

        self._popover = Gtk.PopoverMenu.new_from_model(self._context_menu)
        self._popover.set_parent(self._drawing_area)
        self._popover.set_has_arrow(False)

    def _on_right_click(self, gesture, n_press, x, y):
        rect = Gdk.Rectangle()
        rect.x = int(x)
        rect.y = int(y)
        rect.width = 1
        rect.height = 1
        self._popover.set_pointing_to(rect)
        self._popover.popup()

    # --- Search ---

    def _on_key_pressed(self, controller, keyval, keycode, state):
        if keyval == Gdk.KEY_Escape and self._search_bar.get_search_mode():
            self._close_search()
            return True
        return False

    def toggle_search(self):
        """Toggle search bar visibility."""
        if self._search_bar.get_search_mode():
            self._close_search()
        else:
            self._search_bar.set_search_mode(True)
            self._search_button.set_active(True)
            self._search_entry.grab_focus()

    def _close_search(self):
        self._search_bar.set_search_mode(False)
        self._search_button.set_active(False)
        self._search_query = ""
        self._search_results_label.set_visible(False)
        self._drawing_area.queue_draw()

    def _on_search_toggled(self, button):
        active = button.get_active()
        self._search_bar.set_search_mode(active)
        if active:
            self._search_entry.grab_focus()
        else:
            self._close_search()

    def _on_search_changed(self, entry):
        self._search_query = entry.get_text().strip().lower()
        if not self._search_query:
            self._search_results_label.set_visible(False)
            return

        # Count matches on current page
        matches = 0
        if self._page_info and self._current_page < len(self._page_info):
            for shape in self._page_info[self._current_page].get("shapes", []):
                if self._search_query in shape.get("text", "").lower():
                    matches += 1

        self._search_results_label.set_label(
            _("%d match(es) found") % matches
        )
        self._search_results_label.set_visible(True)
        # Redraw would highlight matches if we had shape-level rendering
        # For now the search result count is shown

    def _on_search_stop(self, entry):
        self._close_search()

    # --- Recent files ---

    def _update_recent_menu(self):
        self._recent_menu.remove_all()
        app = self.get_application()
        if not app:
            return
        for path in app.recent.get_files():
            basename = os.path.basename(path)
            action_name = f"win.open-recent-{hash(path) & 0xFFFFFFFF}"
            action = Gio.SimpleAction.new(action_name.replace("win.", ""), None)
            action.connect("activate", lambda a, p, fp=path: self.open_file(fp))
            self.add_action(action)
            item = Gio.MenuItem.new(basename, None)
            item.set_action_and_target_value(action_name, None)
            self._recent_menu.append_item(item)

    def _update_status(self):
        if self._current_file:
            basename = os.path.basename(self._current_file)
            zoom_pct = int(self._zoom_level * 100)
            page_info = ""
            if len(self._svg_files) > 1:
                page_info = f" — {_('Page')} {self._current_page + 1}/{len(self._svg_files)}"
            self._status_bar.set_label(f"{basename}{page_info} — {zoom_pct}%")
        else:
            self._status_bar.set_label(_("No file loaded"))

    def show_open_dialog(self):
        dialog = Gtk.FileDialog()
        file_filter = Gtk.FileFilter()
        file_filter.set_name(_("Visio files"))
        for ext in sorted(ALL_EXTENSIONS):
            file_filter.add_pattern(f"*{ext}")
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
            self._send_notification(_("Conversion failed"), str(e))
            self._show_error(str(e))
            return

        if not svg_files:
            self._show_error(_("No SVG output produced."))
            return

        self._svg_dir = tmpdir
        self._svg_files = svg_files
        self._current_file = path
        self._current_page = 0
        self._zoom_level = 1.0
        self._search_query = ""

        # Load page info for search/copy
        self._page_info = get_page_info(path)

        # Load first page
        self._load_page(0)

        # Setup page tabs
        self._setup_page_tabs()

        self.set_title(f"VSDView — {os.path.basename(path)}")
        self._update_status()

        # Add to recent files
        app = self.get_application()
        if app:
            app.recent.add_file(path)
            self._update_recent_menu()

    def _load_page(self, page_index: int):
        """Load a specific page SVG."""
        if page_index < 0 or page_index >= len(self._svg_files):
            return
        self._current_page = page_index
        self._svg_handle = Rsvg.Handle.new_from_file(self._svg_files[page_index])
        self._update_status()
        self._drawing_area.queue_draw()

    def _setup_page_tabs(self):
        """Setup page tabs in notebook."""
        # Remove old pages
        while self._notebook.get_n_pages() > 0:
            self._notebook.remove_page(0)

        if len(self._svg_files) <= 1:
            self._notebook.set_visible(False)
            return

        self._notebook.set_visible(True)

        # Block signal while populating
        self._notebook.handler_block_by_func(self._on_page_switched)

        for i in range(len(self._svg_files)):
            if i < len(self._page_info):
                name = self._page_info[i].get("name", f"Page {i + 1}")
            else:
                name = f"Page {i + 1}"
            label = Gtk.Label(label=name)
            # Notebook needs a child widget per page (just a placeholder)
            placeholder = Gtk.Box()
            self._notebook.append_page(placeholder, label)

        self._notebook.set_current_page(0)
        self._notebook.handler_unblock_by_func(self._on_page_switched)

    def _on_page_switched(self, notebook, page, page_num):
        if page_num != self._current_page:
            self._load_page(page_num)

    # --- Copy text ---

    def copy_text(self):
        """Copy all text from current page to clipboard."""
        if not self._current_file:
            return

        text = ""
        if self._page_info and self._current_page < len(self._page_info):
            texts = []
            for shape in self._page_info[self._current_page].get("shapes", []):
                t = shape.get("text", "").strip()
                if t:
                    texts.append(t)
            text = "\n".join(texts)

        if not text:
            # Fallback: extract all text
            text = extract_all_text(self._current_file)

        if text:
            clipboard = self.get_clipboard()
            clipboard.set(text)

    # --- Export ---

    def export_text(self):
        """Export all text to a .txt file."""
        if not self._current_file:
            self._show_error(_("No file loaded to export."))
            return

        dialog = Gtk.FileDialog()
        basename = os.path.splitext(os.path.basename(self._current_file))[0]
        dialog.set_initial_name(f"{basename}.txt")
        dialog.save(self, None, self._on_export_text_chosen)

    def _on_export_text_chosen(self, dialog, result):
        try:
            file = dialog.save_finish(result)
            if not file:
                return
            output_path = file.get_path()
        except Exception:
            return

        text = extract_all_text(self._current_file)
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text)
        except OSError as e:
            self._show_error(str(e))

    def export_pdf(self):
        """Export current page SVG as PDF."""
        if not self._svg_handle or not self._svg_files:
            self._show_error(_("No file loaded to export."))
            return

        dialog = Gtk.FileDialog()
        basename = os.path.splitext(os.path.basename(self._current_file))[0]
        dialog.set_initial_name(f"{basename}.pdf")
        dialog.save(self, None, self._on_export_pdf_chosen)

    def _on_export_pdf_chosen(self, dialog, result):
        try:
            file = dialog.save_finish(result)
            if not file:
                return
            output_path = file.get_path()
        except Exception:
            return

        svg_path = self._svg_files[self._current_page]
        try:
            export_to_pdf(svg_path, output_path)
        except RuntimeError as e:
            self._show_error(str(e))

    # --- Core rendering ---

    def refresh(self):
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
        if not self._svg_handle:
            self._show_error(_("No file loaded to export."))
            return

        dialog = Gtk.FileDialog()
        dialog.set_initial_name("export.png")
        dialog.save(self, None, self._on_export_png_chosen)

    def _on_export_png_chosen(self, dialog, result):
        try:
            file = dialog.save_finish(result)
            if not file:
                return
            output_path = file.get_path()
        except Exception:
            return

        if self._svg_files and self._current_page < len(self._svg_files):
            try:
                export_to_png(self._svg_files[self._current_page], output_path)
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

        self._drawing_area.set_content_width(int(width * self._zoom_level))
        self._drawing_area.set_content_height(int(height * self._zoom_level))

        self._svg_handle.render_document(cr, viewport)

    def _on_drop(self, target, value, x, y):
        if isinstance(value, Gio.File):
            path = value.get_path()
            if path:
                from pathlib import Path as P
                ext = P(path).suffix.lower()
                if ext in ALL_EXTENSIONS:
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
        app = self.get_application()
        if not app:
            return
        notification = Gio.Notification.new(title)
        notification.set_body(body)
        app.send_notification(None, notification)
