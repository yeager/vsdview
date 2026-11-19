"""Exercise document ownership without constructing a graphical window."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from vsdview import window


def fake_window():
    obj = SimpleNamespace(_svg_tempdir=None, _current_file=None)
    for name in ('_send_notification', '_show_error', '_load_page', '_setup_page_tabs',
                 '_update_shape_tree', '_update_layers', 'set_title', '_update_status'):
        setattr(obj, name, Mock())
    obj.get_application = lambda: None
    return obj


def test_replacement_and_close_remove_owned_files(monkeypatch):
    directories = []
    def convert(path, directory):
        directories.append(Path(directory))
        svg = Path(directory) / 'page.svg'
        svg.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>')
        return [str(svg)]
    monkeypatch.setattr(window, 'convert_vsd_to_svg', convert)
    monkeypatch.setattr(window, 'get_page_info', lambda path: [])
    obj = fake_window()
    window.VSDViewWindow.open_file(obj, 'first.vsdx')
    assert directories[0].exists()
    window.VSDViewWindow.open_file(obj, 'second.vsdx')
    assert not directories[0].exists()
    assert directories[1].exists()
    assert obj._current_file == 'second.vsdx'
    assert window.VSDViewWindow._on_close_request(obj) is False
    assert not directories[1].exists()


@pytest.mark.parametrize('failure', ['read', 'empty', 'metadata', 'svg'])
def test_failed_open_cleans_temp_files_and_preserves_document(monkeypatch, failure):
    directories = []
    def convert(path, directory):
        directories.append(Path(directory))
        if failure == 'read':
            raise OSError('Cannot read')
        if failure == 'empty':
            return []
        svg = Path(directory) / 'page.svg'
        svg.write_text('invalid' if failure == 'svg' else '<svg xmlns="http://www.w3.org/2000/svg"/>')
        return [str(svg)]
    def metadata(path):
        if failure == 'metadata':
            raise ValueError('Invalid metadata')
        return []
    monkeypatch.setattr(window, 'convert_vsd_to_svg', convert)
    monkeypatch.setattr(window, 'get_page_info', metadata)
    obj = fake_window()
    obj._current_file = 'previous.vsdx'
    window.VSDViewWindow.open_file(obj, 'broken.vsdx')
    assert not directories[0].exists()
    assert obj._current_file == 'previous.vsdx'
    obj._load_page.assert_not_called()
    obj._show_error.assert_called_once()


def test_export_preserves_existing_file_and_duplicate_page_names(tmp_path, monkeypatch):
    source = tmp_path / "source"
    source.mkdir()
    output = tmp_path / "output"
    output.mkdir()
    pages = [source / "first.svg", source / "second.svg"]
    for number, page in enumerate(pages, 1):
        page.write_text(str(number))
    existing = output / "diagram_page1_Plan_A.svg"
    existing.write_text("keep")
    window_obj = SimpleNamespace(
        _current_file="diagram.vsdx", _svg_files=pages,
        _page_info=[{"name": "Plan/A"}, {"name": "Plan_A"}], _show_error=Mock(),
    )
    dialog = SimpleNamespace(select_folder_finish=lambda _: SimpleNamespace(get_path=lambda: str(output)))
    def write_export(source_path, output_path):
        Path(output_path).write_text(Path(source_path).read_text())
    monkeypatch.setattr(window, "export_to_png", write_export)
    monkeypatch.setattr(window, "export_to_pdf", write_export)
    window.VSDViewWindow._on_export_all_folder_chosen(window_obj, dialog, None)
    assert existing.read_text() == "keep"
    assert (output / "diagram_page2_Plan_A.svg").read_text() == "2"
    assert (output / "diagram_page1_Plan_A.png").read_text() == "1"
    window_obj._show_error.assert_called_once()
