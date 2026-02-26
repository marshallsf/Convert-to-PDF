from unittest.mock import patch

from convert_docx_to_pdf import main


def test_cli_flag():
    with patch("convert_docx_to_pdf.sys.argv", ["prog", "--cli"]), \
         patch("convert_docx_to_pdf.main_cli") as mock_cli, \
         patch("convert_docx_to_pdf.main_gui") as mock_gui:
        main()
    mock_cli.assert_called_once()
    mock_gui.assert_not_called()


def test_no_display():
    with patch("convert_docx_to_pdf.sys.argv", ["prog"]), \
         patch("convert_docx_to_pdf._display_available", return_value=False), \
         patch("convert_docx_to_pdf.main_cli") as mock_cli, \
         patch("convert_docx_to_pdf.main_gui") as mock_gui:
        main()
    mock_cli.assert_called_once()
    mock_gui.assert_not_called()


def test_display_available_no_flag():
    with patch("convert_docx_to_pdf.sys.argv", ["prog"]), \
         patch("convert_docx_to_pdf._display_available", return_value=True), \
         patch("convert_docx_to_pdf.main_cli") as mock_cli, \
         patch("convert_docx_to_pdf.main_gui") as mock_gui:
        main()
    mock_gui.assert_called_once()
    mock_cli.assert_not_called()


def test_cli_flag_with_display():
    with patch("convert_docx_to_pdf.sys.argv", ["prog", "--cli"]), \
         patch("convert_docx_to_pdf._display_available", return_value=True), \
         patch("convert_docx_to_pdf.main_cli") as mock_cli, \
         patch("convert_docx_to_pdf.main_gui") as mock_gui:
        main()
    mock_cli.assert_called_once()
    mock_gui.assert_not_called()
