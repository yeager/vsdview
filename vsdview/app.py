"""VSDView application."""

import gi

gi.require_version("Gtk", "4.0")
gi.require_version("Adw", "1")

from gi.repository import Adw, Gio, Gtk

from vsdview import __version__
from vsdview.window import VSDViewWindow


class VSDViewApplication(Adw.Application):
    def __init__(self):
        super().__init__(
            application_id="org.nylander.vsdview",
            flags=Gio.ApplicationFlags.HANDLES_OPEN,
        )

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

    def _on_open(self, action, param):
        win = self.props.active_window
        if win:
            win.show_open_dialog()

    def _on_about(self, action, param):
        about = Adw.AboutDialog(
            application_name="VSDView",
            application_icon="org.nylander.vsdview",
            version=__version__,
            developer_name="Daniel Nylander",
            license_type=Gtk.License.GPL_3_0,
            developers=["Daniel Nylander"],
            copyright="© 2026 Daniel Nylander",
        )
        about.present(self.props.active_window)
