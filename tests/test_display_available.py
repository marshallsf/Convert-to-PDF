from unittest.mock import patch, MagicMock

from convert_docx_to_pdf import _display_available


def test_windows():
    with patch("convert_docx_to_pdf.sys.platform", "win32"):
        assert _display_available() is True


def test_linux_no_display_no_wayland():
    with patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch.dict("os.environ", {}, clear=True):
        assert _display_available() is False


def test_linux_display_set_tkinter_ok():
    mock_tk = MagicMock()
    with patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch.dict("os.environ", {"DISPLAY": ":0"}, clear=True), \
         patch.dict("sys.modules", {"tkinter": mock_tk}):
        mock_tk.Tk.return_value.destroy.return_value = None
        assert _display_available() is True


def test_linux_display_set_tkinter_fails():
    with patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch.dict("os.environ", {"DISPLAY": ":0"}, clear=True):
        # Make the tkinter import inside _display_available raise
        import builtins
        original_import = builtins.__import__

        def mock_import(name, *args, **kwargs):
            if name == "tkinter":
                raise Exception("no display")
            return original_import(name, *args, **kwargs)

        with patch("builtins.__import__", side_effect=mock_import):
            assert _display_available() is False


def test_linux_wayland_only():
    mock_tk = MagicMock()
    with patch("convert_docx_to_pdf.sys.platform", "linux"), \
         patch.dict("os.environ", {"WAYLAND_DISPLAY": "wayland-0"}, clear=True), \
         patch.dict("sys.modules", {"tkinter": mock_tk}):
        mock_tk.Tk.return_value.destroy.return_value = None
        assert _display_available() is True
