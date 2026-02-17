"""VSDView application."""

import gettext
import locale
import os
import platform
import shutil
import subprocess
import sys

import gi

gi.require_version("Gtk", "4.0")
gi.require_version("Adw", "1")

from gi.repository import Adw, Gio, GLib, Gtk

from vsdview import __version__
from vsdview.recent import RecentFiles
from vsdview.window import VSDViewWindow

# i18n
LOCALE_DIR = os.path.join(os.path.dirname(__file__), "..", "locale")
if not os.path.isdir(LOCALE_DIR):
    LOCALE_DIR = "/usr/share/locale"

gettext.bindtextdomain("vsdview", LOCALE_DIR)
gettext.textdomain("vsdview")
_ = gettext.gettext


class VSDViewApplication(Adw.Application):
    def __init__(self):
        super().__init__(
            application_id="org.nylander.vsdview",
            flags=Gio.ApplicationFlags.HANDLES_OPEN,
        )
        self.recent = RecentFiles()

    def do_activate(self):
        win = self.props.active_window or VSDViewWindow(application=self)
        win.present()

    def do_open(self, files, n_files, hint):
        self.do_activate()
        win = self.props.active_window
        if files:
            win.open_file(files[0].get_path())

    def do_startup(self):
        Adw.Application.do_startup(self)
        self._setup_actions()

    def _setup_actions(self):
        open_action = Gio.SimpleAction.new("open", None)
        open_action.connect("activate", self._on_open)
        self.add_action(open_action)
        self.set_accels_for_action("app.open", ["<Control>o"])

        about_action = Gio.SimpleAction.new("about", None)
        about_action.connect("activate", self._on_about)
        self.add_action(about_action)

        quit_action = Gio.SimpleAction.new("quit", None)
        quit_action.connect("activate", lambda *_: self.quit())
        self.add_action(quit_action)
        self.set_accels_for_action("app.quit", ["<Control>q"])

        # Zoom actions
        zoom_in = Gio.SimpleAction.new("zoom-in", None)
        zoom_in.connect("activate", self._on_zoom_in)
        self.add_action(zoom_in)
        self.set_accels_for_action("app.zoom-in", ["<Control>plus", "<Control>equal"])

        zoom_out = Gio.SimpleAction.new("zoom-out", None)
        zoom_out.connect("activate", self._on_zoom_out)
        self.add_action(zoom_out)
        self.set_accels_for_action("app.zoom-out", ["<Control>minus"])

        zoom_fit = Gio.SimpleAction.new("zoom-fit", None)
        zoom_fit.connect("activate", self._on_zoom_fit)
        self.add_action(zoom_fit)
        self.set_accels_for_action("app.zoom-fit", ["<Control>0"])

        # Refresh
        refresh_action = Gio.SimpleAction.new("refresh", None)
        refresh_action.connect("activate", self._on_refresh)
        self.add_action(refresh_action)
        self.set_accels_for_action("app.refresh", ["F5"])

        # Export PNG
        export_action = Gio.SimpleAction.new("export-png", None)
        export_action.connect("activate", self._on_export_png)
        self.add_action(export_action)
        self.set_accels_for_action("app.export-png", ["<Control>e"])

        # Shortcuts dialog
        shortcuts_action = Gio.SimpleAction.new("show-shortcuts", None)
        shortcuts_action.connect("activate", self._on_show_shortcuts)
        self.add_action(shortcuts_action)
        self.set_accels_for_action("app.show-shortcuts", ["<Control>slash"])

        # Theme toggle
        theme_action = Gio.SimpleAction.new_stateful(
            "toggle-theme",
            None,
            GLib.Variant.new_boolean(False),
        )
        theme_action.connect("activate", self._on_toggle_theme)
        self.add_action(theme_action)

    def _get_win(self):
        return self.props.active_window

    def _on_open(self, action, param):
        win = self._get_win()
        if win:
            win.show_open_dialog()

    def _on_zoom_in(self, action, param):
        win = self._get_win()
        if win:
            win.zoom_in()

    def _on_zoom_out(self, action, param):
        win = self._get_win()
        if win:
            win.zoom_out()

    def _on_zoom_fit(self, action, param):
        win = self._get_win()
        if win:
            win.zoom_fit()

    def _on_refresh(self, action, param):
        win = self._get_win()
        if win:
            win.refresh()

    def _on_export_png(self, action, param):
        win = self._get_win()
        if win:
            win.export_png()

    def _on_show_shortcuts(self, action, param):
        win = self._get_win()
        if not win:
            return
        builder = Gtk.Builder.new_from_string(SHORTCUTS_UI, -1)
        shortcuts_win = builder.get_object("shortcuts")
        shortcuts_win.set_transient_for(win)
        shortcuts_win.present()

    def _on_toggle_theme(self, action, param):
        sm = Adw.StyleManager.get_default()
        if sm.get_dark():
            sm.set_color_scheme(Adw.ColorScheme.FORCE_LIGHT)
            action.set_state(GLib.Variant.new_boolean(False))
        else:
            sm.set_color_scheme(Adw.ColorScheme.FORCE_DARK)
            action.set_state(GLib.Variant.new_boolean(True))

    def _on_about(self, action, param):
        debug_info = self._build_debug_info()
        about = Adw.AboutDialog(
            application_name="VSDView",
            application_icon="org.nylander.vsdview",
            version=__version__,
            developer_name="Daniel Nylander",
            license_type=Gtk.License.GPL_3_0,
            developers=["Daniel Nylander"],
            copyright="© 2026 Daniel Nylander",
            debug_info=debug_info,
            debug_info_filename="vsdview-debug-info.txt",
            website="https://github.com/yeager/vsdview",
            issue_url="https://github.com/yeager/vsdview/issues",
        )
        about.present(self.props.active_window)

    def _build_debug_info(self):
        lo_path = shutil.which("libreoffice") or shutil.which("lowriter") or _("not found")
        lines = [
            f"VSDView {__version__}",
            f"Python {sys.version}",
            f"GTK {Gtk.get_major_version()}.{Gtk.get_minor_version()}.{Gtk.get_micro_version()}",
            f"Adwaita {Adw.get_major_version()}.{Adw.get_minor_version()}.{Adw.get_micro_version()}",
            f"OS: {platform.system()} {platform.release()}",
            f"LibreOffice: {lo_path}",
        ]
        return "\n".join(lines)


SHORTCUTS_UI = """<?xml version="1.0" encoding="UTF-8"?>
<interface>
  <object class="GtkShortcutsWindow" id="shortcuts">
    <property name="modal">1</property>
    <child>
      <object class="GtkShortcutsSection">
        <property name="visible">1</property>
        <child>
          <object class="GtkShortcutsGroup">
            <property name="title" translatable="yes">General</property>
            <child>
              <object class="GtkShortcutsShortcut">
                <property name="title" translatable="yes">Open file</property>
                <property name="accelerator">&lt;Control&gt;o</property>
              </object>
            </child>
            <child>
              <object class="GtkShortcutsShortcut">
                <property name="title" translatable="yes">Quit</property>
                <property name="accelerator">&lt;Control&gt;q</property>
              </object>
            </child>
            <child>
              <object class="GtkShortcutsShortcut">
                <property name="title" translatable="yes">Refresh</property>
                <property name="accelerator">F5</property>
              </object>
            </child>
            <child>
              <object class="GtkShortcutsShortcut">
                <property name="title" translatable="yes">Export as PNG</property>
                <property name="accelerator">&lt;Control&gt;e</property>
              </object>
            </child>
            <child>
              <object class="GtkShortcutsShortcut">
                <property name="title" translatable="yes">Keyboard shortcuts</property>
                <property name="accelerator">&lt;Control&gt;slash</property>
              </object>
            </child>
          </object>
        </child>
        <child>
          <object class="GtkShortcutsGroup">
            <property name="title" translatable="yes">Zoom</property>
            <child>
              <object class="GtkShortcutsShortcut">
                <property name="title" translatable="yes">Zoom in</property>
                <property name="accelerator">&lt;Control&gt;plus</property>
              </object>
            </child>
            <child>
              <object class="GtkShortcutsShortcut">
                <property name="title" translatable="yes">Zoom out</property>
                <property name="accelerator">&lt;Control&gt;minus</property>
              </object>
            </child>
            <child>
              <object class="GtkShortcutsShortcut">
                <property name="title" translatable="yes">Fit to window</property>
                <property name="accelerator">&lt;Control&gt;0</property>
              </object>
            </child>
          </object>
        </child>
      </object>
    </child>
  </object>
</interface>
"""
