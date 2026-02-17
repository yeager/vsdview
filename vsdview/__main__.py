"""Entry point for VSDView."""

import sys

from vsdview.app import VSDViewApplication


def main():
    app = VSDViewApplication()
    return app.run(sys.argv)


if __name__ == "__main__":
    sys.exit(main())
