"""Entry point for VSDView."""

import os
import sys

# The Windows bundle ships GTK typelibs and GSettings schemas beside the
# executable.  GI cannot discover those paths automatically outside MSYS2.
if getattr(sys, "frozen", False):
    bundle_root = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    os.environ.setdefault(
        "GI_TYPELIB_PATH", os.path.join(bundle_root, "lib", "girepository-1.0")
    )
    os.environ.setdefault(
        "GSETTINGS_SCHEMA_DIR", os.path.join(bundle_root, "share", "glib-2.0", "schemas")
    )

from vsdview.app import VSDViewApplication


def main():
    app = VSDViewApplication()
    return app.run(sys.argv)


if __name__ == "__main__":
    sys.exit(main())
