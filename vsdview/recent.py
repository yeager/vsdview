"""Recent files management for VSDView."""

import json
import os

MAX_RECENT = 10


class RecentFiles:
    """Store recent files in a simple JSON file."""

    def __init__(self):
        config_dir = os.path.join(
            os.environ.get("XDG_DATA_HOME", os.path.expanduser("~/.local/share")),
            "vsdview",
        )
        os.makedirs(config_dir, exist_ok=True)
        self._path = os.path.join(config_dir, "recent.json")
        self._files = self._load()

    def _load(self):
        try:
            with open(self._path) as f:
                data = json.load(f)
            if isinstance(data, list):
                return data[:MAX_RECENT]
        except (OSError, json.JSONDecodeError, TypeError):
            pass
        return []

    def _save(self):
        try:
            with open(self._path, "w") as f:
                json.dump(self._files, f, indent=2)
        except OSError:
            pass

    def add_file(self, path):
        path = os.path.abspath(path)
        if path in self._files:
            self._files.remove(path)
        self._files.insert(0, path)
        self._files = self._files[:MAX_RECENT]
        self._save()

    def get_files(self):
        return list(self._files)
