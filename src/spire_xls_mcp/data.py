from pathlib import Path
from typing import Any, List, Dict
import logging

from spire.xls import *

from .exceptions import DataError
from .cell_utils import parse_cell_range, column_to_letter, serialize_cell
from .workbook import get_or_create_workbook

logger = logging.getLogger(__name__)


def read_excel_range(
        filepath: Path | str,
        sheet_name: str,
        cell_range: str,
        preview_only: bool = False
) -> Dict[str, Dict[int, Any]]:
    """
    Read data from Excel range with optional preview mode.
    Returns data in column-first format where cells can be accessed as data[column_letter][row_number]
    """
    try:
        wb = Workbook()
        wb.LoadFromFile(str(filepath))

        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                break
        if sheet is None:
            raise DataError(f"Sheet '{sheet_name}' not found")

        ranges = sheet.Range[cell_range]

        start_row, start_col, end_row, end_col = ranges.Row, ranges.Column, ranges.LastRow, ranges.LastColumn

        # Initialize column-based structure
        data = {}

        # Create data structure organized by columns
        for col in range(start_col, end_col + 1):
            col_letter = column_to_letter(col)
            data[col_letter] = {}

            for row in range(start_row, end_row + 1):
                cell = sheet.Range[row, col]
                data[col_letter][row] = serialize_cell(cell)

        # If preview mode, limit data (keep first 5 columns and first 5 rows)
        if preview_only:
            preview_data = {}
            col_count = 0
            for col_letter in data:
                if col_count >= 5:
                    break
                preview_data[col_letter] = {}
                row_count = 0
                for row in sorted(data[col_letter].keys())[:5]:
                    preview_data[col_letter][row] = data[col_letter][row]
                    row_count += 1
                    if row_count >= 5:
                        break
                col_count += 1
            return preview_data

        return data
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to read Excel range: {e}")
        raise DataError(str(e))


def write_data(filepath: str, sheet_name: str, data: List[List], start_cell: str = "A1") -> dict[str, Any]:
    """Write data to Excel worksheet."""
    try:
        wb = get_or_create_workbook(filepath)
        sheet = None

        # Find or create sheet
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                break
        if sheet is None:
            sheet = wb.CreateEmptySheet(sheet_name)

        # Parse start cell
        cell_range = sheet.Range[start_cell]
        start_row, start_col = cell_range.Row, cell_range.Column

        # Write data
        for i, row in enumerate(data):
            for j, value in enumerate(row):
                cell = sheet.Range[start_row + i, start_col + j]
                cell.Value = str(value)

        wb.SaveToFile(filepath)
        return {"message": "Data written successfully"}
    except Exception as e:
        logger.error(f"Failed to write data: {e}")
        raise DataError(f"Failed to write data: {e!s}")


def _looks_like_headers(row_dict):
    """Check if a data row appears to be headers (keys match values)."""
    return all(
        isinstance(value, str) and str(value).strip() == str(key).strip()
        for key, value in row_dict.items()
    )


def _check_for_headers_above(worksheet, start_row, start_col, headers):
    """Check if cells above start position contain headers."""
    if start_row <= 1:
        return False  # Nothing above row 1

    # Look for header-like content above
    for check_row in range(max(1, start_row - 5), start_row):
        # Count matches for this row
        header_count = 0
        cell_count = 0

        for i, header in enumerate(headers):
            if i >= 10:  # Limit check to first 10 columns for performance
                break

            cell = worksheet.Range[check_row, start_col + i]
            cell_count += 1

            # Check if cell is formatted like a header (bold)
            is_formatted = cell.Style.Font.IsBold

            # Check for any content that could be a header
            if cell.Text:
                # Case 1: Direct match with expected header
                if str(cell.Text).strip().lower() == str(header).strip().lower():
                    header_count += 2  # Give higher weight to exact matches
                # Case 2: Any formatted cell with content
                elif is_formatted and cell.Text:
                    header_count += 1
                # Case 3: Any cell with content in the first row we check
                elif check_row == max(1, start_row - 5):
                    header_count += 0.5

        # If we have a significant number of matching cells, consider it a header row
        if cell_count > 0 and header_count >= cell_count * 0.5:
            return True

    # No headers found above
    return False


def _determine_header_behavior(worksheet, start_row, start_col, data):
    """Determine if headers should be written based on context."""
    if not data:
        return False  # No data means no headers

    # Check if we're in the title area (rows 1-4)
    if start_row <= 4:
        return False  # Don't add headers in title area

    # If we already have data in the sheet, be cautious about adding headers
    if worksheet.LastRow > 1:
        # Check if the target row already has content
        has_content = any(
            worksheet.Range[start_row, start_col + i].Text
            for i in range(min(5, len(data[0].keys())))
        )

        if has_content:
            return False  # Don't overwrite existing content with headers

        # Check if first row appears to be headers
        first_row_is_headers = _looks_like_headers(data[0])

        # Check extensively for headers above (up to 5 rows)
        has_headers_above = _check_for_headers_above(worksheet, start_row, start_col, list(data[0].keys()))

        # Be conservative - don't add headers if we detect headers above or the data has headers
        if has_headers_above or first_row_is_headers:
            return False

        # If we're appending data immediately after existing data, don't add headers
        if any(worksheet.Range[start_row - 1, start_col + i].Text
               for i in range(min(5, len(data[0].keys())))):
            return False

    # For completely new sheets or empty areas far from content, add headers
    return True


def _write_data_to_worksheet(
        worksheet: Any,
        data: list[list],
        start_cell: str = "A1",
) -> None:
    """Write data to worksheet with intelligent header handling"""
    try:
        if not data:
            raise DataError("No data provided to write")

        try:
            start_coords = parse_cell_range(start_cell)
            if not start_coords or not all(x is not None for x in start_coords[:2]):
                raise DataError(f"Invalid start cell reference: {start_cell}")
            start_row, start_col = start_coords[0], start_coords[1]
        except ValueError as e:
            raise DataError(f"Invalid start cell format: {str(e)}")

        # Write data
        for i, row in enumerate(data):
            for j, val in enumerate(row):
                worksheet.Range[start_row + i, start_col + j].Text = str(val)
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to write worksheet data: {e}")
        raise DataError(str(e))
