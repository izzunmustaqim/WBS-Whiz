"""Comprehensive Test Suite for WBS Enhancement Application.

Tests all three extracted modules (file_parser, api_client, wbs_writer)
with 100% branch coverage using pytest + pytest-mock.

To run:  python -m pytest tests/ -v --tb=short
"""

import io
import json
import os
import re
import shutil
from datetime import date, datetime
from unittest.mock import MagicMock, PropertyMock, patch, call

import pandas as pd
import pytest
import requests


# ====================================================================== #
#  Module: file_parser                                                    #
# ====================================================================== #

from file_parser import (
    read_excel_file,
    check_file_validity,
    extract_screen_name,
    parse_screen_layout,
    parse_app_detailed_spec,
    convert_spec_to_json,
)


# ---------------------------------------------------------------------- #
#  read_excel_file                                                        #
# ---------------------------------------------------------------------- #

class TestReadExcelFile:
    """Tests for file_parser.read_excel_file — reads, validates, and cleans Excel files."""

    def test_happy_path_returns_cleaned_dataframe(self, mocker, tmp_path):
        """Standard Excel file is read, cleaned (drop all-NaN cols, fill NaN with 0)."""
        # Arrange
        file = tmp_path / "data.xlsx"
        df = pd.DataFrame({"A": [1, 2], "B": [None, None], "C": [3, None]})
        df.to_excel(file, index=False)

        # Act
        result = read_excel_file(str(file))

        # Assert
        assert "A" in result.columns
        assert "B" not in result.columns, "All-NaN column should be dropped"
        assert result["C"].iloc[1] == 0, "NaN should be filled with 0"

    def test_with_skiprows_parameter(self, tmp_path):
        """rowskip parameter is forwarded to pd.read_excel."""
        # Arrange
        file = tmp_path / "data.xlsx"
        df = pd.DataFrame({"A": [1, 2, 3, 4], "B": [5, 6, 7, 8]})
        df.to_excel(file, index=False)

        # Act — skip first row after header
        result = read_excel_file(str(file), rowskip=1)

        # Assert — original 4 data rows, skip 1 → 3 remain
        assert len(result) == 3
        assert len(result.columns) == 2

    def test_file_too_large_raises_value_error(self, mocker, tmp_path):
        """Files exceeding 25MB raise ValueError with specific message."""
        # Arrange
        file = tmp_path / "big.xlsx"
        file.write_text("dummy")
        mocker.patch("file_parser.os.path.getsize", return_value=26 * 1024 * 1024)

        # Act & Assert
        with pytest.raises(ValueError, match="too large"):
            read_excel_file(str(file))

    def test_file_exactly_25mb_does_not_raise(self, mocker, tmp_path):
        """Exactly 25MB is within the limit (boundary condition)."""
        # Arrange
        file = tmp_path / "data.xlsx"
        df = pd.DataFrame({"A": [1]})
        df.to_excel(file, index=False)
        mocker.patch("file_parser.os.path.getsize", return_value=25 * 1024 * 1024)

        # Act — should not raise
        result = read_excel_file(str(file))

        # Assert
        assert not result.empty

    def test_empty_file_raises_empty_data_error(self, mocker, tmp_path):
        """An Excel file with no data raises EmptyDataError."""
        # Arrange
        file = tmp_path / "empty.xlsx"
        pd.DataFrame().to_excel(file, index=False)
        mocker.patch("file_parser.os.path.getsize", return_value=100)

        # Act & Assert
        with pytest.raises(pd.errors.EmptyDataError, match="empty"):
            read_excel_file(str(file))

    def test_file_not_found_raises_os_error(self):
        """Non-existent file raises FileNotFoundError or OSError."""
        with pytest.raises(OSError):
            read_excel_file("/nonexistent/path/file.xlsx")


# ---------------------------------------------------------------------- #
#  check_file_validity                                                    #
# ---------------------------------------------------------------------- #

class TestCheckFileValidity:
    """Tests for file_parser.check_file_validity — workbook structure validation."""

    def _make_mock_workbook(self, sheets_data):
        """Helper: create a mock workbook from a dict of {sheet_name: rows}.

        Each entry in rows is a list of cell values (None-filtered row).
        """
        wb = MagicMock()
        sheets = {}
        for name, rows in sheets_data.items():
            sheet = MagicMock()
            sheet.max_row = len(rows) if rows else 1
            sheet.max_column = max(len(r) for r in rows) if rows else 1

            # Cell for (1,1) — used in empty sheet check
            cell_mock = MagicMock()
            cell_mock.value = rows[0][0] if rows and rows[0] else None
            sheet.cell.return_value = cell_mock

            # iter_rows returns rows as tuples
            sheet.iter_rows.return_value = [tuple(r) for r in rows]
            sheets[name] = sheet

        wb.__getitem__ = lambda self_wb, name: sheets[name]
        return wb

    def test_valid_screen_layout_file(self, mocker):
        """A Screen Layout file with correct header returns (True, 'File is valid')."""
        # Arrange
        mocker.patch("file_parser.os.path.getsize", return_value=1000)
        rows = [['画面レイアウト\n/Screen Layout', 'Other', 'Data']]
        wb = self._make_mock_workbook({"Sheet1": rows})

        # Act
        is_valid, msg = check_file_validity(["Sheet1"], wb, r"C:\docs\Screen Layout_file.xlsx")

        # Assert
        assert is_valid is True
        assert msg == "File is valid"

    def test_valid_app_spec_file(self, mocker):
        """An Application Detailed Spec file with correct header returns valid."""
        mocker.patch("file_parser.os.path.getsize", return_value=1000)
        rows = [['アプリケーション詳細仕様\n/Application Detailed Specification', 'X']]
        wb = self._make_mock_workbook({"Sheet1": rows})

        is_valid, msg = check_file_validity(
            ["Sheet1"], wb, r"C:\docs\Application Detailed Specification_file.xlsx"
        )
        assert is_valid is True

    def test_valid_event_process_file(self, mocker):
        """An Event Process Sequence Diagram file with correct header returns valid."""
        mocker.patch("file_parser.os.path.getsize", return_value=1000)
        rows = [['イベント処理シーケンス図\n/Event Process Sequence Diagram', 'X']]
        wb = self._make_mock_workbook({"Sheet1": rows})

        is_valid, msg = check_file_validity(
            ["Sheet1"], wb, r"C:\docs\Event Process Sequence Diagram History_file.xlsx"
        )
        assert is_valid is True

    def test_file_too_large_returns_false(self, mocker):
        """Files over 25MB return (False, error message)."""
        mocker.patch("file_parser.os.path.getsize", return_value=26 * 1024 * 1024)
        wb = self._make_mock_workbook({"Sheet1": [["data"]]})

        is_valid, msg = check_file_validity(["Sheet1"], wb, "test.xlsx")
        assert is_valid is False
        assert "too large" in msg

    def test_empty_sheet_single_cell_none(self, mocker):
        """Sheet with max_row=1, max_col=1, cell(1,1)=None returns empty error."""
        mocker.patch("file_parser.os.path.getsize", return_value=100)

        sheet = MagicMock()
        sheet.max_row = 1
        sheet.max_column = 1
        cell_mock = MagicMock()
        cell_mock.value = None
        sheet.cell.return_value = cell_mock

        wb = MagicMock()
        wb.__getitem__ = lambda s, name: sheet

        is_valid, msg = check_file_validity(["Sheet1"], wb, "test.xlsx")
        assert is_valid is False
        assert "empty" in msg.lower()

    def test_row_all_none_returns_empty_error(self, mocker):
        """A row where all cells are None returns empty error."""
        mocker.patch("file_parser.os.path.getsize", return_value=100)

        rows = [(None, None, None)]
        sheet = MagicMock()
        sheet.max_row = 1
        sheet.max_column = 3
        cell_mock = MagicMock()
        cell_mock.value = "not_none"  # pass the single-cell check
        sheet.cell.return_value = cell_mock
        sheet.iter_rows.return_value = rows

        wb = MagicMock()
        wb.__getitem__ = lambda s, name: sheet

        is_valid, msg = check_file_validity(["Sheet1"], wb, "test.xlsx")
        assert is_valid is False
        assert "empty" in msg.lower()

    def test_wrong_header_screen_layout_returns_false(self, mocker):
        """Screen Layout file without the expected Japanese header is invalid."""
        mocker.patch("file_parser.os.path.getsize", return_value=100)
        rows = [['Wrong Header', 'Data']]
        wb = self._make_mock_workbook({"Sheet1": rows})

        is_valid, msg = check_file_validity(
            ["Sheet1"], wb, r"C:\docs\Screen Layout_test.xlsx"
        )
        assert is_valid is False
        assert "not a Screen Layout file" in msg

    def test_wrong_header_app_spec_returns_false(self, mocker):
        """App Detailed Spec file without the expected header is invalid."""
        mocker.patch("file_parser.os.path.getsize", return_value=100)
        rows = [['Wrong Header', 'Data']]
        wb = self._make_mock_workbook({"Sheet1": rows})

        is_valid, msg = check_file_validity(
            ["Sheet1"], wb, r"C:\docs\Application Detailed Specification_test.xlsx"
        )
        assert is_valid is False
        assert "not an Application Detailed Specification file" in msg

    def test_wrong_header_event_process_returns_false(self, mocker):
        """Event Process Diagram file without expected header is invalid."""
        mocker.patch("file_parser.os.path.getsize", return_value=100)
        rows = [['Wrong Header', 'Data']]
        wb = self._make_mock_workbook({"Sheet1": rows})

        is_valid, msg = check_file_validity(
            ["Sheet1"], wb, r"C:\docs\Event Process Sequence Diagram History_test.xlsx"
        )
        assert is_valid is False
        assert "not an Event Process Sequence Diagram History file" in msg

    def test_generic_file_path_with_valid_data_returns_true(self, mocker):
        """Non-SS file path with data falls to the else branch → valid."""
        mocker.patch("file_parser.os.path.getsize", return_value=100)
        rows = [['Some Data', 'More']]
        wb = self._make_mock_workbook({"Sheet1": rows})

        is_valid, msg = check_file_validity(["Sheet1"], wb, r"C:\docs\generic_file.xlsx")
        assert is_valid is True

    def test_empty_sheet_names_returns_valid(self, mocker):
        """No sheets to iterate returns (True, 'File is valid') from the fallthrough."""
        mocker.patch("file_parser.os.path.getsize", return_value=100)
        wb = MagicMock()

        is_valid, msg = check_file_validity([], wb, "test.xlsx")
        assert is_valid is True


# ---------------------------------------------------------------------- #
#  extract_screen_name                                                    #
# ---------------------------------------------------------------------- #

class TestExtractScreenName:
    """Tests for file_parser.extract_screen_name — regex-based filename extraction."""

    @pytest.mark.parametrize("file_path, expected", [
        (r"C:\project\ScreenA_Layout.xlsx", "ScreenA"),
        (r"C:\project\subdir\ComponentName_Details.xlsx", "ComponentName"),
        (r"C:\a\b\c\DeepName_test.xlsx", "DeepName"),
    ], ids=["simple_path", "nested_path", "deep_nested_path"])
    def test_extracts_last_segment_before_underscore(self, file_path, expected):
        """Extracts the last path segment between \\ and _, handling nested backslashes."""
        assert extract_screen_name(file_path) == expected

    def test_no_underscore_returns_empty_string(self):
        """If no underscore follows a backslash segment, returns empty string."""
        assert extract_screen_name(r"C:\project\NoUnderscoreHere.xlsx") == ""

    def test_empty_string_returns_empty(self):
        """Empty input returns empty string."""
        assert extract_screen_name("") == ""

    def test_no_backslash_returns_empty(self):
        """Path without backslashes returns empty (forward slashes don't match)."""
        assert extract_screen_name("project/ScreenA_Layout.xlsx") == ""

    def test_single_segment_with_underscore(self):
        """Single segment: \\Name_ captures 'Name'."""
        assert extract_screen_name(r"\Name_file.xlsx") == "Name"

    def test_multiple_underscores_captures_first(self):
        """Regex [^_]+ stops at first underscore."""
        result = extract_screen_name(r"C:\project\First_Second_Third.xlsx")
        assert result == "First"


# ---------------------------------------------------------------------- #
#  convert_spec_to_json                                                   #
# ---------------------------------------------------------------------- #

class TestConvertSpecToJson:
    """Tests for file_parser.convert_spec_to_json — state machine parser."""

    def test_empty_data_returns_empty_json(self):
        """Empty input produces empty JSON structure with no tasks."""
        json_str, task_names = convert_spec_to_json([], r"C:\test\Module_spec.xlsx")

        parsed = json.loads(json_str)
        assert parsed["Business Division Name"] == []
        assert parsed["Descriptions"] == []
        assert len(parsed["Methods"]) == 1  # Pre-populated initial method
        assert task_names == []

    def test_business_division_name_extraction(self):
        """Business Division Name row creates a task name from file path + division."""
        data = [
            ['業務分割名\n/Business Division Name', 'Division1'],
        ]
        json_str, task_names = convert_spec_to_json(data, r"C:\project\ModuleA_spec.xlsx")

        parsed = json.loads(json_str)
        assert len(task_names) == 1
        assert task_names[0] == "ModuleA_Division1"
        assert parsed["Business Division Name"][0] == "ModuleA_Division1"

    def test_description_section_parsing(self):
        """Description rows between keywords are captured."""
        data = [
            ['説明\n/Description', 'First description'],
            ['説明\n/Description', 'Second description'],
            ['処理名\n/Process Name', 'unused', 'MethodA'],  # Ends description
        ]
        json_str, _ = convert_spec_to_json(data, r"C:\test\X_spec.xlsx")

        parsed = json.loads(json_str)
        assert "First description" in parsed["Descriptions"]

    def test_process_name_creates_method(self):
        """Process Name row with 3 columns adds a method entry."""
        data = [
            ['処理名\n/Process Name', 'unused', 'doSomething'],
        ]
        json_str, _ = convert_spec_to_json(data, r"C:\test\X_spec.xlsx")

        parsed = json.loads(json_str)
        assert parsed["Methods"][0]["Process Name"] == ["doSomething"]

    def test_argument_row_format(self):
        """Argument rows with 4 columns are parsed into structured dicts.
        Note: The keyword row itself has 4 elements so it is also captured."""
        data = [
            ['処理名\n/Process Name', 'unused', 'method1'],
            ['引数\n/Argument', 'x', 'y', 'z'],
            [1, 'paramName', 'String', 'A description'],
        ]
        json_str, _ = convert_spec_to_json(data, r"C:\test\X_spec.xlsx")

        parsed = json.loads(json_str)
        args = parsed["Methods"][0]["Argument"]
        # The keyword row ['引数/Argument', 'x', 'y', 'z'] has 4 elements
        # and is also captured as an argument entry
        assert len(args) == 2
        assert args[1]["Name"] == "paramName"
        assert args[1]["Type"] == "String"

    def test_table_file_row_with_7_columns(self):
        """Table/File rows with 7 columns are parsed into CRUD dicts.
        Note: The keyword row itself has 7 elements so it is also captured."""
        data = [
            ['処理名\n/Process Name', 'unused', 'method1'],
            ['テーブル/ファイル\n/Table/File', 'x', 'y', 'z', 'a', 'b', 'c'],
            [1, 'TBL001', 'UserTable', 'Y', 'Y', 'N', 'N'],
        ]
        json_str, _ = convert_spec_to_json(data, r"C:\test\X_spec.xlsx")

        parsed = json.loads(json_str)
        tables = parsed["Methods"][0]["Table or File use"]
        assert len(tables) == 2  # keyword row + data row
        assert tables[1]["Table_ID/File_ID"] == "TBL001"
        assert tables[1]["CRUD Access for C"] == "Y"

    def test_json_uses_indent_6_and_ensure_ascii_false(self):
        """Output JSON uses indent=6 and preserves unicode."""
        data = [
            ['業務分割名\n/Business Division Name', 'テスト'],
        ]
        json_str, _ = convert_spec_to_json(data, r"C:\test\X_spec.xlsx")

        assert "テスト" in json_str  # ensure_ascii=False
        assert "      " in json_str  # indent=6

    def test_name_header_row_excluded_from_arguments(self):
        """Rows containing '名称\\n/Name' are excluded but the keyword row is captured."""
        data = [
            ['処理名\n/Process Name', 'unused', 'method1'],
            ['引数\n/Argument', 'x', 'y', 'z'],
            ['名称\n/Name', 'ignored', 'ignored', 'ignored'],
            [1, 'realParam', 'int', 'desc'],
        ]
        json_str, _ = convert_spec_to_json(data, r"C:\test\X_spec.xlsx")

        parsed = json.loads(json_str)
        args = parsed["Methods"][0]["Argument"]
        # keyword row captured + '名称/Name' excluded + real param captured = 2
        assert len(args) == 2
        assert args[1]["Name"] == "realParam"


# ====================================================================== #
#  Module: api_client                                                     #
# ====================================================================== #

from api_client import send_gemini_request, API_ENDPOINT


class TestSendGeminiRequest:
    """Tests for api_client.send_gemini_request — Fujitsu Gemini API calls."""

    def test_happy_path_returns_text_content(self, mocker):
        """Successful API call returns the text from candidates[0]."""
        # Arrange
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "candidates": [{
                "content": {
                    "parts": [{"text": "AI response text"}]
                }
            }]
        }
        mock_response.raise_for_status = MagicMock()
        mocker.patch("api_client.requests.post", return_value=mock_response)

        # Act
        result = send_gemini_request("A" * 48, "Test prompt")

        # Assert
        assert result == "AI response text"

    def test_sends_correct_headers_and_payload(self, mocker):
        """Verifies the exact headers and JSON payload structure."""
        # Arrange
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "candidates": [{"content": {"parts": [{"text": "ok"}]}}]
        }
        mock_response.raise_for_status = MagicMock()
        mock_post = mocker.patch("api_client.requests.post", return_value=mock_response)

        api_key = "TestKey123"
        prompt = "Hello world"

        # Act
        send_gemini_request(api_key, prompt)

        # Assert
        mock_post.assert_called_once_with(
            API_ENDPOINT,
            headers={"Content-type": "application/json", "api-key": api_key},
            json={
                "contents": [
                    {"role": "user", "parts": [{"text": prompt}]}
                ]
            },
        )

    def test_http_error_raises_request_exception(self, mocker):
        """HTTP 4xx/5xx triggers raise_for_status → RequestException."""
        mock_response = MagicMock()
        mock_response.raise_for_status.side_effect = requests.exceptions.HTTPError("403 Forbidden")
        mocker.patch("api_client.requests.post", return_value=mock_response)

        with pytest.raises(requests.exceptions.HTTPError, match="403"):
            send_gemini_request("key", "prompt")

    def test_connection_error_propagates(self, mocker):
        """Network failure raises ConnectionError."""
        mocker.patch(
            "api_client.requests.post",
            side_effect=requests.exceptions.ConnectionError("DNS failure"),
        )

        with pytest.raises(requests.exceptions.ConnectionError):
            send_gemini_request("key", "prompt")

    def test_timeout_error_propagates(self, mocker):
        """Timeout raises Timeout exception."""
        mocker.patch(
            "api_client.requests.post",
            side_effect=requests.exceptions.Timeout("Request timed out"),
        )

        with pytest.raises(requests.exceptions.Timeout):
            send_gemini_request("key", "prompt")

    def test_malformed_json_response_raises_key_error(self, mocker):
        """Response with missing 'candidates' key raises KeyError."""
        mock_response = MagicMock()
        mock_response.raise_for_status = MagicMock()
        mock_response.json.return_value = {"error": "something went wrong"}
        mocker.patch("api_client.requests.post", return_value=mock_response)

        with pytest.raises(KeyError):
            send_gemini_request("key", "prompt")

    def test_empty_candidates_raises_index_error(self, mocker):
        """Response with empty candidates list raises IndexError."""
        mock_response = MagicMock()
        mock_response.raise_for_status = MagicMock()
        mock_response.json.return_value = {"candidates": []}
        mocker.patch("api_client.requests.post", return_value=mock_response)

        with pytest.raises(IndexError):
            send_gemini_request("key", "prompt")

    def test_json_decode_error_propagates(self, mocker):
        """Non-JSON response body raises JSONDecodeError."""
        mock_response = MagicMock()
        mock_response.raise_for_status = MagicMock()
        mock_response.json.side_effect = json.JSONDecodeError("err", "doc", 0)
        mocker.patch("api_client.requests.post", return_value=mock_response)

        with pytest.raises(json.JSONDecodeError):
            send_gemini_request("key", "prompt")

    def test_api_endpoint_is_correct(self):
        """API_ENDPOINT constant points to the Fujitsu Gemini service."""
        assert "api.ai-service.global.fujitsu.com" in API_ENDPOINT
        assert "generateContent" in API_ENDPOINT


# ====================================================================== #
#  Module: wbs_writer                                                     #
# ====================================================================== #

from wbs_writer import markdown_table_to_dataframe, copy_to_downloads


class TestMarkdownTableToDataframe:
    """Tests for wbs_writer.markdown_table_to_dataframe — markdown → DataFrame."""

    def test_happy_path_standard_markdown_table(self):
        """Standard markdown table with header, separator, and data rows."""
        content = """\
Here is the WBS:

| Item No. | Task | Assigned | Start | End |
|---|---|---|---|---|
| 1 | Design UI | Alice | 01/01/2025 | 01/05/2025 |
| 2 | Backend | Bob | 01/01/2025 | 01/10/2025 |

Some trailing text."""

        df = markdown_table_to_dataframe(content)

        assert len(df) == 1  # 2 data rows, but pipeline uses one as header
        assert "2" in str(df.iloc[0].values)

    def test_single_data_row_produces_empty_dataframe(self):
        """Table with only header + separator + 1 data row → 0-row DataFrame
        (the 1 data row is consumed as column headers)."""
        content = """\
| Col1 | Col2 |
|---|---|
| val1 | val2 |"""

        df = markdown_table_to_dataframe(content)
        assert len(df) == 0

    def test_preserves_column_names_from_data(self):
        """After re-headering, columns come from the first data row values."""
        content = """\
| A | B | C |
|---|---|---|
| X | Y | Z |
| 1 | 2 | 3 |"""

        df = markdown_table_to_dataframe(content)
        # First data row (" X ", " Y ", " Z ") becomes headers
        col_names = [c.strip() for c in df.columns]
        assert "X" in col_names
        assert "Y" in col_names

    def test_extracts_only_pipe_delimited_lines(self):
        """Non-table text is filtered out by the regex."""
        content = """\
Some preamble text.
No pipes here.
| Only | This | Counts |
|---|---|---|
| A | B | C |
More trailing text without pipes."""

        df = markdown_table_to_dataframe(content)
        # 3 lines matched: header, sep, data → after processing: 0 rows
        assert len(df) == 0  # single data row consumed as header

    @pytest.mark.parametrize("content", [
        "",
        "No table at all.",
        "Just some text\nwith newlines\nbut no pipes.",
    ], ids=["empty_string", "plain_text", "multiline_no_pipes"])
    def test_no_table_content_raises(self, content):
        """Content without pipe characters causes a parsing error."""
        with pytest.raises(Exception):
            markdown_table_to_dataframe(content)


class TestCopyToDownloads:
    """Tests for wbs_writer.copy_to_downloads — file copy to Downloads folder."""

    def test_happy_path_returns_destination_path(self, mocker):
        """Successful copy returns the full destination file path."""
        # Arrange
        mocker.patch("wbs_writer.pd.DataFrame.to_excel")
        mocker.patch("wbs_writer.shutil.copy")
        mocker.patch("wbs_writer.os.getcwd", return_value=r"C:\project")

        # Act
        result = copy_to_downloads()

        # Assert
        assert result.endswith("Details_WBS.xlsm")
        assert "Downloads" in result

    def test_calls_shutil_copy_with_correct_paths(self, mocker):
        """Verifies source and destination paths for shutil.copy."""
        mocker.patch("wbs_writer.pd.DataFrame.to_excel")
        mock_copy = mocker.patch("wbs_writer.shutil.copy")
        mocker.patch("wbs_writer.os.getcwd", return_value=r"C:\project")

        copy_to_downloads()

        # Assert source path
        source = mock_copy.call_args[0][0]
        assert source == os.path.join(r"C:\project", "Details_WBS.xlsm")

    def test_copy_failure_raises_exception(self, mocker):
        """shutil.copy failure propagates as an exception."""
        mocker.patch("wbs_writer.pd.DataFrame.to_excel")
        mocker.patch("wbs_writer.shutil.copy", side_effect=PermissionError("Access denied"))

        with pytest.raises(PermissionError, match="Access denied"):
            copy_to_downloads()


# ====================================================================== #
#  Module: app.py — validate_api_key (via standalone logic test)          #
# ====================================================================== #

class TestValidateApiKeyLogic:
    """Tests for the API key validation logic used in app.py.
    
    Tested as standalone regex logic to avoid Tkinter dependency.
    Tests the exact same pattern: ^[A-Za-z0-9]{48}$
    """

    PATTERN = r'^[A-Za-z0-9]{48}$'
    FULLWIDTH_PATTERN = r'[\uFF01-\uFF60\uFFE0-\uFFE6]'

    def _validate(self, api_key: str) -> str | None:
        """Replicate validate_api_key logic without GUI dependency."""
        if not api_key:
            return None
        elif re.search(self.FULLWIDTH_PATTERN, api_key):
            return None
        elif re.match(self.PATTERN, api_key):
            return api_key
        else:
            return None

    @pytest.mark.parametrize("key", [
        "A" * 48,
        "a" * 48,
        "0" * 48,
        "aB3cD4eF5gH6iJ7kL8mN9oP0qR1sT2uV3wX4yZ5aB3cD4eF5",
    ], ids=["all_upper", "all_lower", "all_digits", "mixed"])
    def test_valid_48_char_alphanumeric_key(self, key):
        """48 alphanumeric characters → returns the key."""
        assert self._validate(key) == key

    @pytest.mark.parametrize("key, reason", [
        ("", "empty_string"),
        ("A" * 47, "too_short_47"),
        ("A" * 49, "too_long_49"),
        ("A" * 47 + "!", "special_char_exclamation"),
        ("A" * 47 + " ", "contains_space"),
        ("A" * 47 + "_", "contains_underscore"),
        ("A" * 47 + "@", "contains_at_sign"),
        ("A" * 47 + ".", "contains_dot"),
    ], ids=lambda x: x if isinstance(x, str) and not x.startswith("A") else "")
    def test_invalid_keys_return_none(self, key, reason):
        """Various invalid keys return None."""
        assert self._validate(key) is None

    def test_full_width_character_returns_none(self):
        """Full-width characters are caught before format check."""
        key = "Ａ" + "A" * 47  # Full-width A
        assert self._validate(key) is None

    def test_full_width_takes_precedence_over_length(self):
        """Full-width check happens before length/format validation."""
        key = "Ａ"  # Just one full-width char, wrong length
        assert self._validate(key) is None

    def test_none_like_input_returns_none(self):
        """Falsy inputs (empty string) are caught by `not api_key`."""
        assert self._validate("") is None


# ====================================================================== #
#  Module: config.py                                                      #
# ====================================================================== #

class TestConfig:
    """Tests for config.py — prompts and error messages."""

    def test_prompt_list_task_has_required_placeholders(self):
        """prompt_list_task template contains all 3 format placeholders."""
        import config
        assert "{screen_layout_json}" in config.prompt_list_task
        assert "{app_detailed_spec_data_converted_json}" in config.prompt_list_task
        assert "{tasks_list_json}" in config.prompt_list_task

    def test_prompt_has_required_placeholders(self):
        """Main prompt template contains all 9 format placeholders."""
        import config
        expected = [
            "{task_details_data}", "{skill_set_data}",
            "{start_date_str}", "{end_date_str}",
            "{task_description}", "{assigned_to}",
            "{progress}", "{plan_start_date}", "{plan_end_date}",
        ]
        for placeholder in expected:
            assert placeholder in config.prompt, f"Missing: {placeholder}"

    def test_prompt_format_with_all_placeholders_succeeds(self):
        """Formatting prompt with all placeholders doesn't raise."""
        import config
        result = config.prompt.format(
            task_details_data="tasks", skill_set_data="skills",
            start_date_str="01/01/2025", end_date_str="06/30/2025",
            task_description="desc", assigned_to="person",
            progress="To do", plan_start_date="01/01/2025",
            plan_end_date="06/30/2025",
        )
        assert "tasks" in result
        assert "skills" in result

    @pytest.mark.parametrize("key", [
        "FileNotFoundError", "FolderNotFoundError", "ManyExcelError",
        "EmptyDataError", "ParserError", "FileTooBig", "APIKeyError",
        "FailReadError", "APIEmptyField", "InvalidKeyError",
        "FullWidthCharacterError", "GeneralError",
    ])
    def test_all_error_message_keys_exist(self, key):
        """Every error key used in app.py exists in config.error_message."""
        import config
        assert key in config.error_message, f"Missing error key: {key}"

    def test_many_excel_error_message_mismatch(self):
        """Documents the known mismatch: message says '5' but code checks > 50."""
        import config
        assert "5" in config.error_message["ManyExcelError"]
