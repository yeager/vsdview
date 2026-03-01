"""VSDView main window."""

import gettext
import json as _json
import math
import os
import tempfile
import webbrowser
import importlib.util

import gi

gi.require_version("Gtk", "4.0")
gi.require_version("Adw", "1")
gi.require_version("Rsvg", "2.0")

from gi.repository import Adw, Gdk, Gio, GLib, Gtk, Rsvg

from vsdview.accessibility import AccessibilityManager
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


_CANVAS_MODES = ("system", "light", "dark")
_CANVAS_ICONS = {
    "system": "preferences-desktop-display-symbolic",
    "light": "weather-clear-symbolic",
    "dark": "weather-clear-night-symbolic",
}
_CANVAS_TOOLTIPS = {
    "system": _("Canvas: System theme"),
    "light": _("Canvas: Light background"),
    "dark": _("Canvas: Dark background"),
}

_SETTINGS_DIR = os.path.join(GLib.get_user_config_dir(), "vsdview")
_SETTINGS_PATH = os.path.join(_SETTINGS_DIR, "settings.ini")


def _load_canvas_mode() -> str:
    kf = GLib.KeyFile()
    try:
        kf.load_from_file(_SETTINGS_PATH, GLib.KeyFileFlags.NONE)
        mode = kf.get_string("General", "canvas_mode")
        if mode in _CANVAS_MODES:
            return mode
    except Exception:
        pass
    return "system"


def _save_canvas_mode(mode: str):
    os.makedirs(_SETTINGS_DIR, exist_ok=True)
    kf = GLib.KeyFile()
    try:
        kf.load_from_file(_SETTINGS_PATH, GLib.KeyFileFlags.NONE)
    except Exception:
        pass
    kf.set_string("General", "canvas_mode", mode)
    kf.save_to_file(_SETTINGS_PATH)


# ---------------------------------------------------------------------------
# Helper: compute shape bounding boxes from page_info for hit-testing
# ---------------------------------------------------------------------------
_INCH_TO_PX = 72.0


def _shape_bboxes(page_info_entry, page_h):
    """Return list of (x, y, w, h, shape_dict) in SVG pixel coords."""
    bboxes = []
    for shape in page_info_entry.get("shapes", []):
        cells = shape.get("cells", {})
        px = float(cells.get("PinX", {}).get("V", 0) or 0) * _INCH_TO_PX
        py = (page_h - float(cells.get("PinY", {}).get("V", 0) or 0)) * _INCH_TO_PX
        sw = abs(float(cells.get("Width", {}).get("V", 0) or 0)) * _INCH_TO_PX
        sh = abs(float(cells.get("Height", {}).get("V", 0) or 0)) * _INCH_TO_PX
        if sw < 1 and sh < 1:
            continue
        bboxes.append((px - sw / 2, py - sh / 2, sw, sh, shape))
    return bboxes


class VSDViewWindow(Adw.ApplicationWindow):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.set_default_size(1100, 750)
        self.set_title("VSDView")

        self._svg_handle = None
        self._current_file = None
        self._zoom_level = 1.0
        self._svg_dir = None
        self._svg_files = []
        self._current_page = 0
        self._page_info = []
        self._search_query = ""
        self._canvas_mode = _load_canvas_mode()

        # Pan state
        self._pan_x = 0.0
        self._pan_y = 0.0
        self._drag_start_pan_x = 0.0
        self._drag_start_pan_y = 0.0
        self._space_held = False

        # Shape selection
        self._selected_shape = None
        self._shape_bboxes = []

        # Measurement
        self._measure_mode = False
        self._measure_point1 = None  # (x, y) in doc coords (inches)
        self._measure_point2 = None

        # Search results navigation
        self._search_results = []  # list of (page_idx, shape_idx)
        self._search_result_idx = -1

        # Layers
        self._layers = {}  # {layer_id: {"name": str, "visible": bool}}

        # Build UI
        self._build_ui()
        self._update_recent_menu()

    def _build_ui(self):
        # Main vertical box
        main_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.set_content(main_box)

        # --- Header bar ---
        header = Adw.HeaderBar()
        main_box.append(header)

        # Open button
        open_btn = Gtk.Button(icon_name="document-open-symbolic")
        open_btn.set_tooltip_text(_("Open Visio file"))
        open_btn.connect("clicked", lambda *_: self.show_open_dialog())
        header.pack_start(open_btn)

        # Zoom controls in header
        zoom_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=2)
        zoom_box.add_css_class("linked")

        zoom_out_btn = Gtk.Button(icon_name="zoom-out-symbolic")
        zoom_out_btn.set_tooltip_text(_("Zoom out"))
        zoom_out_btn.connect("clicked", lambda *_: self.zoom_out())
        zoom_box.append(zoom_out_btn)

        self._zoom_scale = Gtk.Scale.new_with_range(
            Gtk.Orientation.HORIZONTAL, 10, 500, 10
        )
        self._zoom_scale.set_value(100)
        self._zoom_scale.set_size_request(120, -1)
        self._zoom_scale.set_draw_value(False)
        self._zoom_scale.connect("value-changed", self._on_zoom_scale_changed)
        zoom_box.append(self._zoom_scale)

        self._zoom_label = Gtk.Label(label="100%")
        self._zoom_label.set_width_chars(5)
        zoom_box.append(self._zoom_label)

        zoom_in_btn = Gtk.Button(icon_name="zoom-in-symbolic")
        zoom_in_btn.set_tooltip_text(_("Zoom in"))
        zoom_in_btn.connect("clicked", lambda *_: self.zoom_in())
        zoom_box.append(zoom_in_btn)

        zoom_fit_btn = Gtk.Button(icon_name="zoom-fit-best-symbolic")
        zoom_fit_btn.set_tooltip_text(_("Fit to window"))
        zoom_fit_btn.connect("clicked", lambda *_: self.zoom_fit())
        zoom_box.append(zoom_fit_btn)

        header.set_title_widget(zoom_box)

        # Search button
        self._search_button = Gtk.ToggleButton(icon_name="edit-find-symbolic")
        self._search_button.set_tooltip_text(_("Find text"))
        self._search_button.connect("toggled", self._on_search_toggled)
        header.pack_start(self._search_button)

        # Canvas background mode button
        self._canvas_btn = Gtk.Button(icon_name=_CANVAS_ICONS[self._canvas_mode])
        self._canvas_btn.set_tooltip_text(_CANVAS_TOOLTIPS[self._canvas_mode])
        self._canvas_btn.connect("clicked", self._on_canvas_mode_cycle)
        header.pack_start(self._canvas_btn)

        # Measurement tool button
        self._measure_btn = Gtk.ToggleButton(icon_name="preferences-desktop-display-symbolic")
        self._measure_btn.set_tooltip_text(_("Measurement tool"))
        self._measure_btn.connect("toggled", self._on_measure_toggled)
        header.pack_start(self._measure_btn)

        # Layers button
        self._layers_btn = Gtk.MenuButton(icon_name="view-list-symbolic")
        self._layers_btn.set_tooltip_text(_("Layers"))
        self._layers_popover = Gtk.Popover()
        self._layers_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=4)
        self._layers_box.set_margin_start(8)
        self._layers_box.set_margin_end(8)
        self._layers_box.set_margin_top(8)
        self._layers_box.set_margin_bottom(8)
        self._layers_popover.set_child(self._layers_box)
        self._layers_btn.set_popover(self._layers_popover)
        header.pack_start(self._layers_btn)

        # Hamburger menu
        menu_btn = Gtk.MenuButton(icon_name="open-menu-symbolic")
        menu = Gio.Menu()

        self._recent_menu = Gio.Menu()
        menu.append_submenu(_("Recent Files"), self._recent_menu)

        export_section = Gio.Menu()
        export_section.append(_("Export as PNG"), "app.export-png")
        export_section.append(_("Export as PDF"), "app.export-pdf")
        export_section.append(_("Export as SVG"), "app.export-svg")
        export_section.append(_("Export as Text"), "app.export-text")
        export_section.append(_("Export All Pages…"), "app.export-all")
        menu.append_section(_("Export"), export_section)

        menu.append(_("Copy Text"), "app.copy-text")

        view_section = Gio.Menu()
        view_section.append(_("Shape Tree Sidebar"), "win.toggle-shape-tree")
        view_section.append(_("Shape Info Panel"), "win.toggle-shape-info")
        view_section.append(_("Minimap"), "win.toggle-minimap")
        menu.append_section(_("View"), view_section)

        canvas_submenu = Gio.Menu()
        canvas_submenu.append(_("System"), "win.canvas-mode::system")
        canvas_submenu.append(_("Light"), "win.canvas-mode::light")
        canvas_submenu.append(_("Dark"), "win.canvas-mode::dark")
        menu.append_submenu(_("Canvas Background"), canvas_submenu)

        # Register the stateful action for canvas mode radio
        canvas_action = Gio.SimpleAction.new_stateful(
            "canvas-mode",
            GLib.VariantType.new("s"),
            GLib.Variant.new_string(self._canvas_mode),
        )
        canvas_action.connect("activate", self._on_canvas_mode_action)
        self.add_action(canvas_action)

        # Toggle actions for panels
        for name, default in [("toggle-shape-tree", False),
                               ("toggle-shape-info", False),
                               ("toggle-minimap", True)]:
            act = Gio.SimpleAction.new_stateful(
                name, None, GLib.Variant.new_boolean(default))
            act.connect("activate", self._on_toggle_panel)
            self.add_action(act)

        menu.append(_("Keyboard Shortcuts"), "app.show-shortcuts")
        menu.append(_("About"), "app.about")
        menu.append(_("Quit"), "app.quit")
        menu_btn.set_menu_model(menu)
        header.pack_end(menu_btn)

        # --- Search bar ---
        self._search_bar = Gtk.SearchBar()
        search_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=4)
        self._search_entry = Gtk.SearchEntry()
        self._search_entry.set_hexpand(True)
        self._search_entry.connect("search-changed", self._on_search_changed)
        self._search_entry.connect("stop-search", self._on_search_stop)
        self._search_entry.connect("activate", self._on_search_next)
        search_box.append(self._search_entry)

        prev_btn = Gtk.Button(icon_name="go-up-symbolic")
        prev_btn.set_tooltip_text(_("Previous result"))
        prev_btn.connect("clicked", lambda *_: self._on_search_prev())
        search_box.append(prev_btn)

        next_btn = Gtk.Button(icon_name="go-down-symbolic")
        next_btn.set_tooltip_text(_("Next result"))
        next_btn.connect("clicked", lambda *_: self._on_search_next())
        search_box.append(next_btn)

        self._search_all_pages_check = Gtk.CheckButton(label=_("All pages"))
        self._search_all_pages_check.set_active(True)
        self._search_all_pages_check.connect("toggled", self._on_search_changed)
        search_box.append(self._search_all_pages_check)

        clamp = Adw.Clamp(maximum_size=600)
        clamp.set_child(search_box)
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

        # --- Main content area with optional sidebars ---
        self._main_paned = Gtk.Paned(orientation=Gtk.Orientation.HORIZONTAL)
        self._main_paned.set_vexpand(True)
        self._main_paned.set_hexpand(True)
        main_box.append(self._main_paned)

        # Left sidebar: Shape tree
        self._shape_tree_revealer = Gtk.Revealer()
        self._shape_tree_revealer.set_transition_type(
            Gtk.RevealerTransitionType.SLIDE_RIGHT)
        self._shape_tree_revealer.set_reveal_child(False)

        tree_frame = Gtk.Frame()
        tree_scroll = Gtk.ScrolledWindow()
        tree_scroll.set_size_request(220, -1)
        tree_scroll.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)

        # TreeView with columns: Name, Type
        self._tree_store = Gtk.TreeStore.new([str, str, str, int])  # name, type, text, shape_idx
        self._tree_view = Gtk.TreeView(model=self._tree_store)
        self._tree_view.set_headers_visible(True)

        col_name = Gtk.TreeViewColumn(_("Shape"), Gtk.CellRendererText(), text=0)
        col_name.set_expand(True)
        self._tree_view.append_column(col_name)

        col_type = Gtk.TreeViewColumn(_("Type"), Gtk.CellRendererText(), text=1)
        self._tree_view.append_column(col_type)

        self._tree_view.connect("cursor-changed", self._on_tree_selection_changed)
        tree_scroll.set_child(self._tree_view)
        tree_frame.set_child(tree_scroll)
        self._shape_tree_revealer.set_child(tree_frame)

        # Right sidebar: Shape info
        self._shape_info_revealer = Gtk.Revealer()
        self._shape_info_revealer.set_transition_type(
            Gtk.RevealerTransitionType.SLIDE_LEFT)
        self._shape_info_revealer.set_reveal_child(False)

        info_frame = Gtk.Frame()
        info_scroll = Gtk.ScrolledWindow()
        info_scroll.set_size_request(250, -1)
        info_scroll.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        self._shape_info_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=4)
        self._shape_info_box.set_margin_start(8)
        self._shape_info_box.set_margin_end(8)
        self._shape_info_box.set_margin_top(8)
        self._shape_info_box.set_margin_bottom(8)
        self._shape_info_label = Gtk.Label(label=_("Click a shape to inspect"))
        self._shape_info_label.set_wrap(True)
        self._shape_info_label.set_xalign(0)
        self._shape_info_box.append(self._shape_info_label)
        info_scroll.set_child(self._shape_info_box)
        info_frame.set_child(info_scroll)
        self._shape_info_revealer.set_child(info_frame)

        # Center: Overlay with drawing area + minimap
        center_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)

        center_box.append(self._shape_tree_revealer)

        self._overlay = Gtk.Overlay()
        self._overlay.set_hexpand(True)
        self._overlay.set_vexpand(True)

        scrolled = Gtk.ScrolledWindow()
        scrolled.set_vexpand(True)
        scrolled.set_hexpand(True)
        self._scrolled_window = scrolled

        self._drawing_area = Gtk.DrawingArea()
        self._drawing_area.set_vexpand(True)
        self._drawing_area.set_hexpand(True)
        self._drawing_area.set_draw_func(self._on_draw)
        scrolled.set_child(self._drawing_area)

        self._overlay.set_child(scrolled)

        # Minimap overlay
        self._minimap = Gtk.DrawingArea()
        self._minimap.set_size_request(180, 140)
        self._minimap.set_halign(Gtk.Align.END)
        self._minimap.set_valign(Gtk.Align.END)
        self._minimap.set_margin_end(10)
        self._minimap.set_margin_bottom(10)
        self._minimap.set_draw_func(self._on_draw_minimap)
        self._minimap.add_css_class("card")
        self._minimap.set_visible(True)

        # Minimap click to navigate
        minimap_click = Gtk.GestureClick.new()
        minimap_click.connect("pressed", self._on_minimap_click)
        self._minimap.add_controller(minimap_click)

        self._overlay.add_overlay(self._minimap)

        center_box.append(self._overlay)
        center_box.append(self._shape_info_revealer)

        self._main_paned.set_start_child(center_box)
        self._main_paned.set_resize_start_child(True)
        self._main_paned.set_shrink_start_child(False)
        # No end child needed (sidebars are in center_box via revealers)
        self._main_paned.set_end_child(Gtk.Box())
        self._main_paned.set_resize_end_child(False)
        self._main_paned.set_shrink_end_child(True)

        # --- Controllers ---

        # Scroll zoom (Ctrl+scroll = zoom, plain scroll = vertical pan)
        scroll_ctrl = Gtk.EventControllerScroll.new(
            Gtk.EventControllerScrollFlags.VERTICAL
        )
        scroll_ctrl.connect("scroll", self._on_scroll_zoom)
        self._drawing_area.add_controller(scroll_ctrl)

        # Pinch to zoom
        zoom_gesture = Gtk.GestureZoom.new()
        zoom_gesture.connect("scale-changed", self._on_pinch_zoom)
        zoom_gesture.connect("begin", self._on_pinch_begin)
        self._drawing_area.add_controller(zoom_gesture)
        self._pinch_start_zoom = 1.0

        # Drag to pan (button 1 when Space held, or button 2/middle)
        drag_ctrl = Gtk.GestureDrag.new()
        drag_ctrl.set_button(0)  # Any button
        drag_ctrl.connect("drag-begin", self._on_drag_begin)
        drag_ctrl.connect("drag-update", self._on_drag_update)
        drag_ctrl.connect("drag-end", self._on_drag_end)
        self._drawing_area.add_controller(drag_ctrl)
        self._drag_gesture = drag_ctrl

        # Click for shape selection / measurement / hyperlinks
        click_ctrl = Gtk.GestureClick.new()
        click_ctrl.set_button(1)
        click_ctrl.connect("pressed", self._on_left_click)
        self._drawing_area.add_controller(click_ctrl)

        # Motion for cursor changes (hyperlinks)
        motion_ctrl = Gtk.EventControllerMotion.new()
        motion_ctrl.connect("motion", self._on_pointer_motion)
        self._drawing_area.add_controller(motion_ctrl)

        # Right-click context menu
        self._setup_context_menu()

        # Key controller
        key_ctrl = Gtk.EventControllerKey.new()
        key_ctrl.connect("key-pressed", self._on_key_pressed)
        key_ctrl.connect("key-released", self._on_key_released)
        self.add_controller(key_ctrl)

        # --- Page tabs ---
        self._notebook = Gtk.Notebook()
        self._notebook.set_show_tabs(True)
        self._notebook.set_visible(False)
        self._notebook.connect("switch-page", self._on_page_switched)
        main_box.append(self._notebook)

        # --- Status bar ---
        status_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        status_box.set_margin_start(8)
        status_box.set_margin_end(8)
        status_box.set_margin_top(4)
        status_box.set_margin_bottom(4)

        self._status_bar = Gtk.Label(label=_("No file loaded"))
        self._status_bar.set_xalign(0)
        self._status_bar.set_hexpand(True)
        self._status_bar.add_css_class("caption")
        status_box.append(self._status_bar)

        self._measure_label = Gtk.Label(label="")
        self._measure_label.add_css_class("caption")
        self._measure_label.set_visible(False)
        status_box.append(self._measure_label)

        main_box.append(status_box)

        # Drag and drop
        drop_target = Gtk.DropTarget.new(Gio.File, Gdk.DragAction.COPY)
        drop_target.connect("drop", self._on_drop)
        self.add_controller(drop_target)

    # =====================================================================
    # Context menu
    # =====================================================================
    def _setup_context_menu(self):
        gesture = Gtk.GestureClick.new()
        gesture.set_button(3)
        gesture.connect("pressed", self._on_right_click)
        self._drawing_area.add_controller(gesture)

        self._context_menu = Gio.Menu()
        self._context_menu.append(_("Copy Text"), "app.copy-text")
        self._context_menu.append(_("Export as PNG"), "app.export-png")
        self._context_menu.append(_("Export as PDF"), "app.export-pdf")
        self._context_menu.append(_("Export as SVG"), "app.export-svg")
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

    # =====================================================================
    # Key handling
    # =====================================================================
    def _on_key_pressed(self, controller, keyval, keycode, state):
        if keyval == Gdk.KEY_Escape:
            if self._search_bar.get_search_mode():
                self._close_search()
                return True
            if self._measure_mode:
                self._measure_mode = False
                self._measure_btn.set_active(False)
                self._measure_point1 = None
                self._measure_point2 = None
                self._measure_label.set_visible(False)
                self._drawing_area.queue_draw()
                return True
        if keyval == Gdk.KEY_space:
            self._space_held = True
            return False
        # Arrow keys for panning
        pan_step = 50
        hadj = self._scrolled_window.get_hadjustment()
        vadj = self._scrolled_window.get_vadjustment()
        if keyval == Gdk.KEY_Left:
            hadj.set_value(hadj.get_value() - pan_step)
            return True
        elif keyval == Gdk.KEY_Right:
            hadj.set_value(hadj.get_value() + pan_step)
            return True
        elif keyval == Gdk.KEY_Up:
            vadj.set_value(vadj.get_value() - pan_step)
            return True
        elif keyval == Gdk.KEY_Down:
            vadj.set_value(vadj.get_value() + pan_step)
            return True
        return False

    def _on_key_released(self, controller, keyval, keycode, state):
        if keyval == Gdk.KEY_space:
            self._space_held = False

    # =====================================================================
    # Toggle panels
    # =====================================================================
    def _on_toggle_panel(self, action, param):
        new_val = not action.get_state().get_boolean()
        action.set_state(GLib.Variant.new_boolean(new_val))
        name = action.get_name()
        if name == "toggle-shape-tree":
            self._shape_tree_revealer.set_reveal_child(new_val)
        elif name == "toggle-shape-info":
            self._shape_info_revealer.set_reveal_child(new_val)
        elif name == "toggle-minimap":
            self._minimap.set_visible(new_val)

    # =====================================================================
    # Search
    # =====================================================================
    def toggle_search(self):
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
        self._search_results = []
        self._search_result_idx = -1
        self._search_results_label.set_visible(False)
        self._drawing_area.queue_draw()

    def _on_search_toggled(self, button):
        active = button.get_active()
        self._search_bar.set_search_mode(active)
        if active:
            self._search_entry.grab_focus()
        else:
            self._close_search()

    def _on_search_changed(self, *args):
        self._search_query = self._search_entry.get_text().strip().lower()
        if not self._search_query:
            self._search_results = []
            self._search_result_idx = -1
            self._search_results_label.set_visible(False)
            return

        # Build search results
        self._search_results = []
        search_all = self._search_all_pages_check.get_active()

        pages_to_search = range(len(self._page_info)) if search_all else [self._current_page]
        for pi in pages_to_search:
            if pi >= len(self._page_info):
                continue
            for si, shape in enumerate(self._page_info[pi].get("shapes", [])):
                if self._search_query in shape.get("text", "").lower():
                    self._search_results.append((pi, si))

        total = len(self._search_results)
        if total > 0:
            self._search_result_idx = 0
            self._search_results_label.set_label(
                _("Result 1 of %d") % total
            )
        else:
            self._search_result_idx = -1
            self._search_results_label.set_label(_("No matches found"))
        self._search_results_label.set_visible(True)

    def _on_search_stop(self, entry):
        self._close_search()

    def _on_search_next(self, *args):
        if not self._search_results:
            return
        self._search_result_idx = (self._search_result_idx + 1) % len(self._search_results)
        self._navigate_to_search_result()

    def _on_search_prev(self, *args):
        if not self._search_results:
            return
        self._search_result_idx = (self._search_result_idx - 1) % len(self._search_results)
        self._navigate_to_search_result()

    def _navigate_to_search_result(self):
        if self._search_result_idx < 0 or self._search_result_idx >= len(self._search_results):
            return
        pi, si = self._search_results[self._search_result_idx]
        total = len(self._search_results)
        self._search_results_label.set_label(
            _("Result %d of %d") % (self._search_result_idx + 1, total)
        )
        if pi != self._current_page:
            self._load_page(pi)
            if self._notebook.get_visible():
                self._notebook.handler_block_by_func(self._on_page_switched)
                self._notebook.set_current_page(pi)
                self._notebook.handler_unblock_by_func(self._on_page_switched)

    # =====================================================================
    # Recent files
    # =====================================================================
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

    # =====================================================================
    # Status
    # =====================================================================
    def _update_status(self):
        if self._current_file:
            basename = os.path.basename(self._current_file)
            zoom_pct = int(self._zoom_level * 100)
            page_info = ""
            if len(self._svg_files) > 1:
                page_info = f" — {_('Page')} {self._current_page + 1}/{len(self._svg_files)}"
            self._status_bar.set_label(f"{basename}{page_info} — {zoom_pct}%")
            # Update zoom controls
            self._zoom_label.set_label(f"{zoom_pct}%")
            self._zoom_scale.handler_block_by_func(self._on_zoom_scale_changed)
            self._zoom_scale.set_value(zoom_pct)
            self._zoom_scale.handler_unblock_by_func(self._on_zoom_scale_changed)
        else:
            self._status_bar.set_label(_("No file loaded"))

    # =====================================================================
    # File opening
    # =====================================================================
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
        self._selected_shape = None
        self._measure_point1 = None
        self._measure_point2 = None

        self._page_info = get_page_info(path)
        self._load_page(0)
        self._setup_page_tabs()
        self._update_shape_tree()
        self._update_layers()

        self.set_title(f"VSDView — {os.path.basename(path)}")
        self._update_status()

        app = self.get_application()
        if app:
            app.recent.add_file(path)
            self._update_recent_menu()

    def _load_page(self, page_index: int):
        if page_index < 0 or page_index >= len(self._svg_files):
            return
        self._current_page = page_index
        self._svg_handle = Rsvg.Handle.new_from_file(self._svg_files[page_index])
        self._selected_shape = None
        self._update_shape_bboxes()
        self._update_shape_tree()
        self._update_status()
        self._drawing_area.queue_draw()
        self._minimap.queue_draw()

    def _update_shape_bboxes(self):
        """Recompute shape bounding boxes for current page."""
        self._shape_bboxes = []
        if self._page_info and self._current_page < len(self._page_info):
            pi = self._page_info[self._current_page]
            page_h = pi.get("page_h", 11.0)
            self._shape_bboxes = _shape_bboxes(pi, page_h)

    def _setup_page_tabs(self):
        while self._notebook.get_n_pages() > 0:
            self._notebook.remove_page(0)

        if len(self._svg_files) <= 1:
            self._notebook.set_visible(False)
            return

        self._notebook.set_visible(True)
        self._notebook.handler_block_by_func(self._on_page_switched)

        for i in range(len(self._svg_files)):
            if i < len(self._page_info):
                name = self._page_info[i].get("name", f"Page {i + 1}")
            else:
                name = f"Page {i + 1}"
            label = Gtk.Label(label=name)
            placeholder = Gtk.Box()
            self._notebook.append_page(placeholder, label)

        self._notebook.set_current_page(0)
        self._notebook.handler_unblock_by_func(self._on_page_switched)

    def _on_page_switched(self, notebook, page, page_num):
        if page_num != self._current_page:
            self._load_page(page_num)

    def page_next(self):
        if self._current_page < len(self._svg_files) - 1:
            new_page = self._current_page + 1
            self._load_page(new_page)
            if self._notebook.get_visible():
                self._notebook.set_current_page(new_page)

    def page_prev(self):
        if self._current_page > 0:
            new_page = self._current_page - 1
            self._load_page(new_page)
            if self._notebook.get_visible():
                self._notebook.set_current_page(new_page)

    # =====================================================================
    # Copy text
    # =====================================================================
    def copy_text(self):
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
            text = extract_all_text(self._current_file)
        if text:
            clipboard = self.get_clipboard()
            clipboard.set(text)

    # =====================================================================
    # Export
    # =====================================================================
    def export_text(self):
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

    def export_png(self):
        if not self._svg_handle:
            self._show_error(_("No file loaded to export."))
            return
        dialog = Gtk.FileDialog()
        basename = os.path.splitext(os.path.basename(self._current_file))[0]
        dialog.set_initial_name(f"{basename}.png")
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

    def export_svg(self):
        """Export current page as SVG (copy the generated SVG)."""
        if not self._svg_files or self._current_page >= len(self._svg_files):
            self._show_error(_("No file loaded to export."))
            return
        dialog = Gtk.FileDialog()
        basename = os.path.splitext(os.path.basename(self._current_file))[0]
        dialog.set_initial_name(f"{basename}.svg")
        dialog.save(self, None, self._on_export_svg_chosen)

    def _on_export_svg_chosen(self, dialog, result):
        try:
            file = dialog.save_finish(result)
            if not file:
                return
            output_path = file.get_path()
        except Exception:
            return
        import shutil
        try:
            shutil.copy2(self._svg_files[self._current_page], output_path)
        except OSError as e:
            self._show_error(str(e))

    def export_all_pages(self):
        """Export all pages — show dialog to choose format and directory."""
        if not self._svg_files:
            self._show_error(_("No file loaded to export."))
            return
        dialog = Gtk.FileDialog()
        dialog.select_folder(self, None, self._on_export_all_folder_chosen)

    def _on_export_all_folder_chosen(self, dialog, result):
        try:
            folder = dialog.select_folder_finish(result)
            if not folder:
                return
            output_dir = folder.get_path()
        except Exception:
            return

        import shutil
        basename = os.path.splitext(os.path.basename(self._current_file))[0]
        for i, svg_path in enumerate(self._svg_files):
            page_name = f"page{i+1}"
            if i < len(self._page_info):
                pn = self._page_info[i].get("name", "")
                if pn:
                    # Sanitize filename
                    page_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in pn)
            # Export SVG
            out_svg = os.path.join(output_dir, f"{basename}_{page_name}.svg")
            shutil.copy2(svg_path, out_svg)
            # Export PNG
            out_png = os.path.join(output_dir, f"{basename}_{page_name}.png")
            try:
                export_to_png(svg_path, out_png)
            except Exception:
                pass
            # Export PDF
            out_pdf = os.path.join(output_dir, f"{basename}_{page_name}.pdf")
            try:
                export_to_pdf(svg_path, out_pdf)
            except Exception:
                pass

    # =====================================================================
    # Zoom
    # =====================================================================
    def refresh(self):
        if self._current_file and os.path.exists(self._current_file):
            self.open_file(self._current_file)

    def zoom_in(self):
        self._zoom_level = min(self._zoom_level * 1.25, 5.0)
        self._apply_zoom()

    def zoom_out(self):
        self._zoom_level = max(self._zoom_level / 1.25, 0.1)
        self._apply_zoom()

    def zoom_fit(self):
        if not self._svg_handle:
            return
        has_size, svg_w, svg_h = self._svg_handle.get_intrinsic_size_in_pixels()
        if not has_size or svg_w == 0 or svg_h == 0:
            self._zoom_level = 1.0
        else:
            alloc = self._drawing_area.get_allocation()
            if alloc.width > 0 and alloc.height > 0:
                zoom_w = alloc.width / svg_w
                zoom_h = alloc.height / svg_h
                self._zoom_level = min(zoom_w, zoom_h, 1.0)
            else:
                self._zoom_level = 1.0
        self._apply_zoom()

    def _apply_zoom(self):
        self._update_status()
        self._drawing_area.queue_draw()
        self._minimap.queue_draw()

    def _on_zoom_scale_changed(self, scale):
        self._zoom_level = scale.get_value() / 100.0
        self._apply_zoom()

    def _on_scroll_zoom(self, controller, dx, dy):
        state = controller.get_current_event_state()
        if state & Gdk.ModifierType.CONTROL_MASK:
            if dy < 0:
                self.zoom_in()
            elif dy > 0:
                self.zoom_out()
            return True
        return False

    def _on_pinch_begin(self, gesture, sequence):
        self._pinch_start_zoom = self._zoom_level

    def _on_pinch_zoom(self, gesture, scale):
        self._zoom_level = max(0.1, min(5.0, self._pinch_start_zoom * scale))
        self._apply_zoom()

    # =====================================================================
    # Pan / drag
    # =====================================================================
    def _on_drag_begin(self, gesture, start_x, start_y):
        button = gesture.get_current_button()
        # Pan with middle button or left button when space is held
        if button == 2 or (button == 1 and self._space_held):
            hadj = self._scrolled_window.get_hadjustment()
            vadj = self._scrolled_window.get_vadjustment()
            self._drag_start_pan_x = hadj.get_value()
            self._drag_start_pan_y = vadj.get_value()
            gesture.set_state(Gtk.EventSequenceState.CLAIMED)
        else:
            gesture.set_state(Gtk.EventSequenceState.DENIED)

    def _on_drag_update(self, gesture, offset_x, offset_y):
        hadj = self._scrolled_window.get_hadjustment()
        vadj = self._scrolled_window.get_vadjustment()
        hadj.set_value(self._drag_start_pan_x - offset_x)
        vadj.set_value(self._drag_start_pan_y - offset_y)
        self._minimap.queue_draw()

    def _on_drag_end(self, gesture, offset_x, offset_y):
        pass

    # =====================================================================
    # Canvas mode
    # =====================================================================
    def _set_canvas_mode(self, mode):
        self._canvas_mode = mode
        self._canvas_btn.set_icon_name(_CANVAS_ICONS[mode])
        self._canvas_btn.set_tooltip_text(_CANVAS_TOOLTIPS[mode])
        action = self.lookup_action("canvas-mode")
        if action:
            action.set_state(GLib.Variant.new_string(mode))
        _save_canvas_mode(mode)
        self._drawing_area.queue_draw()

    def _on_canvas_mode_cycle(self, button):
        idx = _CANVAS_MODES.index(self._canvas_mode)
        self._set_canvas_mode(_CANVAS_MODES[(idx + 1) % len(_CANVAS_MODES)])

    def _on_canvas_mode_action(self, action, parameter):
        mode = parameter.get_string()
        if mode in _CANVAS_MODES:
            self._set_canvas_mode(mode)

    # =====================================================================
    # Drawing
    # =====================================================================
    def _on_draw(self, area, cr, width, height):
        if not self._svg_handle:
            return

        # Canvas background
        if self._canvas_mode == "light":
            cr.set_source_rgb(1, 1, 1)
            cr.paint()
        elif self._canvas_mode == "dark":
            cr.set_source_rgb(0.15, 0.15, 0.15)
            cr.paint()

        has_size, svg_w, svg_h = self._svg_handle.get_intrinsic_size_in_pixels()
        if has_size and svg_w > 0 and svg_h > 0:
            content_w = int(svg_w * self._zoom_level)
            content_h = int(svg_h * self._zoom_level)
        else:
            content_w = int(width * self._zoom_level)
            content_h = int(height * self._zoom_level)

        self._drawing_area.set_content_width(content_w)
        self._drawing_area.set_content_height(content_h)

        viewport = Rsvg.Rectangle()
        viewport.x = 0
        viewport.y = 0
        viewport.width = content_w
        viewport.height = content_h
        self._svg_handle.render_document(cr, viewport)

        # Draw measurement overlay
        if self._measure_mode and self._measure_point1:
            cr.set_source_rgba(1, 0, 0, 0.8)
            cr.set_line_width(2)
            x1, y1 = self._measure_point1
            cr.arc(x1, y1, 4, 0, 2 * math.pi)
            cr.fill()
            if self._measure_point2:
                x2, y2 = self._measure_point2
                cr.arc(x2, y2, 4, 0, 2 * math.pi)
                cr.fill()
                cr.move_to(x1, y1)
                cr.line_to(x2, y2)
                cr.stroke()

    # =====================================================================
    # Minimap
    # =====================================================================
    def _on_draw_minimap(self, area, cr, width, height):
        if not self._svg_handle:
            return

        # Draw scaled-down version of the document
        cr.set_source_rgb(1, 1, 1)
        cr.paint()

        has_size, svg_w, svg_h = self._svg_handle.get_intrinsic_size_in_pixels()
        if not has_size or svg_w <= 0 or svg_h <= 0:
            return

        scale = min(width / svg_w, height / svg_h) * 0.95
        offset_x = (width - svg_w * scale) / 2
        offset_y = (height - svg_h * scale) / 2

        cr.save()
        cr.translate(offset_x, offset_y)

        viewport = Rsvg.Rectangle()
        viewport.x = 0
        viewport.y = 0
        viewport.width = svg_w * scale
        viewport.height = svg_h * scale
        self._svg_handle.render_document(cr, viewport)
        cr.restore()

        # Draw viewport rectangle
        hadj = self._scrolled_window.get_hadjustment()
        vadj = self._scrolled_window.get_vadjustment()
        content_w = svg_w * self._zoom_level
        content_h = svg_h * self._zoom_level

        if content_w > 0 and content_h > 0:
            vx = hadj.get_value() / content_w * svg_w * scale + offset_x
            vy = vadj.get_value() / content_h * svg_h * scale + offset_y
            vw = hadj.get_page_size() / content_w * svg_w * scale
            vh = vadj.get_page_size() / content_h * svg_h * scale

            cr.set_source_rgba(0.2, 0.4, 0.8, 0.3)
            cr.rectangle(vx, vy, vw, vh)
            cr.fill()
            cr.set_source_rgba(0.2, 0.4, 0.8, 0.7)
            cr.set_line_width(1.5)
            cr.rectangle(vx, vy, vw, vh)
            cr.stroke()

    def _on_minimap_click(self, gesture, n_press, x, y):
        """Click on minimap to navigate."""
        if not self._svg_handle:
            return
        has_size, svg_w, svg_h = self._svg_handle.get_intrinsic_size_in_pixels()
        if not has_size or svg_w <= 0 or svg_h <= 0:
            return

        mm_w = self._minimap.get_width()
        mm_h = self._minimap.get_height()
        scale = min(mm_w / svg_w, mm_h / svg_h) * 0.95
        offset_x = (mm_w - svg_w * scale) / 2
        offset_y = (mm_h - svg_h * scale) / 2

        # Convert minimap coords to document fraction
        doc_x = (x - offset_x) / (svg_w * scale)
        doc_y = (y - offset_y) / (svg_h * scale)

        content_w = svg_w * self._zoom_level
        content_h = svg_h * self._zoom_level
        hadj = self._scrolled_window.get_hadjustment()
        vadj = self._scrolled_window.get_vadjustment()

        hadj.set_value(doc_x * content_w - hadj.get_page_size() / 2)
        vadj.set_value(doc_y * content_h - vadj.get_page_size() / 2)
        self._minimap.queue_draw()

    # =====================================================================
    # Shape selection / hit-testing
    # =====================================================================
    def _on_left_click(self, gesture, n_press, x, y):
        if self._measure_mode:
            self._handle_measure_click(x, y)
            return

        # Hit-test shapes
        if not self._shape_bboxes:
            return

        # Convert widget coords to content coords
        hadj = self._scrolled_window.get_hadjustment()
        vadj = self._scrolled_window.get_vadjustment()
        cx = (x + hadj.get_value()) / self._zoom_level
        cy = (y + vadj.get_value()) / self._zoom_level

        # Check for hyperlink first
        hit = self._hit_test_shape(cx, cy)
        if hit:
            shape = hit[4]
            # Check for hyperlinks
            hyperlinks = shape.get("cells", {})
            # Visio stores hyperlinks in Hyperlink section — check user data too
            href = ""
            for key, val in hyperlinks.items():
                if key.startswith("Hyperlink") or key == "HyperLink":
                    href = val.get("V", "")
                    break
            if href and href.startswith(("http://", "https://")):
                webbrowser.open(href)
                return

            self._selected_shape = shape
            self._update_shape_info_panel(shape)
            # Show info panel if not visible
            act = self.lookup_action("toggle-shape-info")
            if act and not act.get_state().get_boolean():
                act.activate(None)
        else:
            self._selected_shape = None
            self._update_shape_info_panel(None)

    def _hit_test_shape(self, doc_x, doc_y):
        """Find shape at document coordinates. Returns bbox tuple or None."""
        # Search in reverse order (top shapes first)
        for bbox in reversed(self._shape_bboxes):
            bx, by, bw, bh = bbox[:4]
            if bx <= doc_x <= bx + bw and by <= doc_y <= by + bh:
                return bbox
        return None

    def _on_pointer_motion(self, controller, x, y):
        """Change cursor for hyperlinks."""
        if not self._shape_bboxes:
            return
        hadj = self._scrolled_window.get_hadjustment()
        vadj = self._scrolled_window.get_vadjustment()
        cx = (x + hadj.get_value()) / self._zoom_level
        cy = (y + vadj.get_value()) / self._zoom_level

        hit = self._hit_test_shape(cx, cy)
        if hit:
            shape = hit[4]
            hyperlinks = shape.get("cells", {})
            has_link = any(
                (k.startswith("Hyperlink") or k == "HyperLink") and
                v.get("V", "").startswith(("http://", "https://"))
                for k, v in hyperlinks.items()
            )
            if has_link:
                self._drawing_area.set_cursor(
                    Gdk.Cursor.new_from_name("pointer", None))
                return
        if self._measure_mode:
            self._drawing_area.set_cursor(
                Gdk.Cursor.new_from_name("crosshair", None))
        else:
            self._drawing_area.set_cursor(None)

    # =====================================================================
    # Shape info panel
    # =====================================================================
    def _update_shape_info_panel(self, shape):
        """Update the shape info sidebar with shape details."""
        # Remove all children except the first label
        child = self._shape_info_box.get_first_child()
        while child:
            next_child = child.get_next_sibling()
            self._shape_info_box.remove(child)
            child = next_child

        if not shape:
            lbl = Gtk.Label(label=_("Click a shape to inspect"))
            lbl.set_wrap(True)
            lbl.set_xalign(0)
            self._shape_info_box.append(lbl)
            return

        cells = shape.get("cells", {})

        def add_row(label, value):
            if not value:
                return
            row = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=2)
            lbl = Gtk.Label(label=label)
            lbl.set_xalign(0)
            lbl.add_css_class("caption")
            lbl.add_css_class("dim-label")
            row.append(lbl)
            val_lbl = Gtk.Label(label=str(value))
            val_lbl.set_xalign(0)
            val_lbl.set_wrap(True)
            val_lbl.set_selectable(True)
            row.append(val_lbl)
            self._shape_info_box.append(row)

        # Header
        title = Gtk.Label(label=_("Shape Properties"))
        title.set_xalign(0)
        title.add_css_class("title-4")
        self._shape_info_box.append(title)

        sep = Gtk.Separator()
        self._shape_info_box.append(sep)

        add_row(_("ID"), shape.get("id", ""))
        add_row(_("Name"), shape.get("name_u", "") or shape.get("name", ""))
        add_row(_("Type"), shape.get("type", ""))
        add_row(_("Master"), shape.get("master", ""))

        text = shape.get("text", "")
        if text:
            add_row(_("Text"), text[:200])

        # Dimensions
        w = cells.get("Width", {}).get("V", "")
        h = cells.get("Height", {}).get("V", "")
        if w:
            add_row(_("Width"), f"{float(w):.3f} in ({float(w)*25.4:.1f} mm)")
        if h:
            add_row(_("Height"), f"{float(h):.3f} in ({float(h)*25.4:.1f} mm)")

        # Position
        px = cells.get("PinX", {}).get("V", "")
        py = cells.get("PinY", {}).get("V", "")
        if px and py:
            add_row(_("Position"), f"({float(px):.3f}, {float(py):.3f}) in")

        # Angle
        angle = cells.get("Angle", {}).get("V", "")
        if angle and float(angle or 0) != 0:
            add_row(_("Rotation"), f"{math.degrees(float(angle)):.1f}°")

        # Fill & line
        fill = cells.get("FillForegnd", {}).get("V", "")
        if fill:
            add_row(_("Fill"), fill)
        line_col = cells.get("LineColor", {}).get("V", "")
        if line_col:
            add_row(_("Line Color"), line_col)

        # User properties
        user = shape.get("user", {})
        if user:
            sep2 = Gtk.Separator()
            self._shape_info_box.append(sep2)
            for key, vals in user.items():
                val = vals.get("Value", vals.get("V", ""))
                if val:
                    add_row(key, val)

    # =====================================================================
    # Shape tree sidebar
    # =====================================================================
    def _update_shape_tree(self):
        """Rebuild the shape tree for current page."""
        self._tree_store.clear()
        if not self._page_info or self._current_page >= len(self._page_info):
            return

        shapes = self._page_info[self._current_page].get("shapes", [])
        for i, shape in enumerate(shapes):
            self._add_shape_to_tree(shape, None, i)

    def _add_shape_to_tree(self, shape, parent_iter, shape_idx):
        name = shape.get("name_u", "") or shape.get("name", "") or f"Shape {shape.get('id', '?')}"
        stype = shape.get("type", "Shape")
        text = (shape.get("text", "") or "")[:40]
        it = self._tree_store.append(parent_iter, [name, stype, text, shape_idx])

        for i, sub in enumerate(shape.get("sub_shapes", [])):
            self._add_shape_to_tree(sub, it, i)

    def _on_tree_selection_changed(self, tree_view):
        selection = tree_view.get_selection()
        model, treeiter = selection.get_selected()
        if treeiter:
            shape_idx = model.get_value(treeiter, 3)
            shapes = self._page_info[self._current_page].get("shapes", [])
            if 0 <= shape_idx < len(shapes):
                self._selected_shape = shapes[shape_idx]
                self._update_shape_info_panel(shapes[shape_idx])
                # Show info panel
                act = self.lookup_action("toggle-shape-info")
                if act and not act.get_state().get_boolean():
                    act.activate(None)

    # =====================================================================
    # Layers
    # =====================================================================
    def _update_layers(self):
        """Update layers panel from page info."""
        # Clear existing layer toggles
        child = self._layers_box.get_first_child()
        while child:
            next_child = child.get_next_sibling()
            self._layers_box.remove(child)
            child = next_child

        # Parse layers from the page_info if available
        # For now, extract unique LayerMember values from shapes
        if not self._page_info or self._current_page >= len(self._page_info):
            lbl = Gtk.Label(label=_("No layers"))
            self._layers_box.append(lbl)
            return

        layer_names = {}
        for shape in self._page_info[self._current_page].get("shapes", []):
            cells = shape.get("cells", {})
            lm = cells.get("LayerMember", {}).get("V", "")
            if lm:
                for lid in lm.split(";"):
                    lid = lid.strip()
                    if lid and lid not in layer_names:
                        layer_names[lid] = f"Layer {lid}"

        if not layer_names:
            lbl = Gtk.Label(label=_("No layers"))
            self._layers_box.append(lbl)
            return

        title = Gtk.Label(label=_("Layers"))
        title.add_css_class("title-4")
        self._layers_box.append(title)

        for lid, name in sorted(layer_names.items()):
            cb = Gtk.CheckButton(label=name)
            cb.set_active(True)
            cb.connect("toggled", self._on_layer_toggled, lid)
            self._layers_box.append(cb)

    def _on_layer_toggled(self, checkbutton, layer_id):
        # Layer visibility toggling would require re-rendering with layer info
        # For now, this is a placeholder that shows the UI
        pass

    # =====================================================================
    # Measurement tool
    # =====================================================================
    def _on_measure_toggled(self, button):
        self._measure_mode = button.get_active()
        self._measure_point1 = None
        self._measure_point2 = None
        if self._measure_mode:
            self._drawing_area.set_cursor(
                Gdk.Cursor.new_from_name("crosshair", None))
            self._measure_label.set_label(_("Click two points to measure"))
            self._measure_label.set_visible(True)
        else:
            self._drawing_area.set_cursor(None)
            self._measure_label.set_visible(False)
            self._drawing_area.queue_draw()

    def _handle_measure_click(self, x, y):
        hadj = self._scrolled_window.get_hadjustment()
        vadj = self._scrolled_window.get_vadjustment()
        # Convert to content coordinates
        cx = x + hadj.get_value()
        cy = y + vadj.get_value()

        if self._measure_point1 is None:
            self._measure_point1 = (cx, cy)
            self._measure_label.set_label(_("Click second point"))
            self._drawing_area.queue_draw()
        else:
            self._measure_point2 = (cx, cy)
            # Calculate distance in document units
            x1, y1 = self._measure_point1
            x2, y2 = self._measure_point2
            dist_px = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
            dist_inches = dist_px / (_INCH_TO_PX * self._zoom_level)
            dist_mm = dist_inches * 25.4
            self._measure_label.set_label(
                f"{dist_inches:.3f} in / {dist_mm:.1f} mm"
            )
            self._drawing_area.queue_draw()
            # Reset for next measurement
            self._measure_point1 = None
            self._measure_point2 = None

    # =====================================================================
    # Drag & drop
    # =====================================================================
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

    # =====================================================================
    # Utility
    # =====================================================================
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


# --- Session restore ---

def _save_session(window, app_name):
    config_dir = os.path.join(os.path.expanduser('~'), '.config', app_name)
    os.makedirs(config_dir, exist_ok=True)
    state = {'width': window.get_width(), 'height': window.get_height(),
             'maximized': window.is_maximized()}
    try:
        with open(os.path.join(config_dir, 'session.json'), 'w') as f:
            _json.dump(state, f)
    except OSError:
        pass


def _restore_session(window, app_name):
    path = os.path.join(os.path.expanduser('~'), '.config', app_name, 'session.json')
    try:
        with open(path) as f:
            state = _json.load(f)
        window.set_default_size(state.get('width', 800), state.get('height', 600))
        if state.get('maximized'):
            window.maximize()
    except (FileNotFoundError, _json.JSONDecodeError, OSError):
        pass


# --- Fullscreen toggle (F11) ---

def _setup_fullscreen(window, app):
    """Add F11 fullscreen toggle."""
    if not app.lookup_action('toggle-fullscreen'):
        action = Gio.SimpleAction.new('toggle-fullscreen', None)
        action.connect('activate', lambda a, p: (
            window.unfullscreen() if window.is_fullscreen() else window.fullscreen()
        ))
        app.add_action(action)
        app.set_accels_for_action('app.toggle-fullscreen', ['F11'])


# --- Plugin system ---

def _load_plugins(app_name):
    """Load plugins from ~/.config/<app>/plugins/."""
    plugin_dir = os.path.join(os.path.expanduser('~'), '.config', app_name, 'plugins')
    plugins = []
    if not os.path.isdir(plugin_dir):
        return plugins
    for fname in sorted(os.listdir(plugin_dir)):
        if fname.endswith('.py') and not fname.startswith('_'):
            path = os.path.join(plugin_dir, fname)
            try:
                spec = importlib.util.spec_from_file_location(fname[:-3], path)
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                plugins.append(mod)
            except Exception as e:
                print(f"Plugin {fname}: {e}")
    return plugins
