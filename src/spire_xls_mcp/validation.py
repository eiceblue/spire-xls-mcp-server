import logging
import re
from typing import Any, Tuple
from spire.xls import *

from .cell_utils import parse_cell_range, validate_cell_reference_regex, column_to_letter
from .exceptions import ValidationError

logger = logging.getLogger(__name__)


def validate_formula(formula: str) -> Tuple[bool, str]:
    """Validate Excel formula syntax."""
    if not formula:
        return False, "Formula cannot be empty"

    if not formula.startswith('='):
        formula = f'={formula}'

    # Basic formula validation
    try:
        # Create temporary workbook and worksheet for validation
        wb = Workbook()
        sheet = wb.Worksheets[0]

        # Try to set formula
        sheet.Range["A1"].Formula = formula

        # If no exception was raised, formula is valid
        return True, "Formula syntax is valid"

    except Exception as e:
        return False, f"Invalid formula syntax: {str(e)}"


def validate_range_in_sheet(
        filepath: str,
        sheet_name: str,
        range_str: str
) -> dict[str, Any]:
    """Validate if range exists and is properly formatted."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)

        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                break

        if sheet is None:
            raise ValidationError(f"Sheet '{sheet_name}' not found")

        # Parse range
        try:
            if ':' in range_str:
                start_cell, end_cell = range_str.split(':')
            else:
                start_cell = range_str
                end_cell = None

            start_row, start_col, end_row, end_col = parse_cell_range(start_cell, end_cell)

            # Validate start cell is within sheet bounds
            if start_row > sheet.LastRow or start_col > sheet.LastColumn:
                raise ValidationError(
                    f"Start cell out of bounds. Sheet dimensions are "
                    f"A1:{column_to_letter(sheet.LastColumn)}{sheet.LastRow}"
                )

            # If end cell specified, validate it's within bounds and after start cell
            if end_row is not None and end_col is not None:
                if end_row > sheet.LastRow or end_col > sheet.LastColumn:
                    raise ValidationError(
                        f"End cell out of bounds. Sheet dimensions are "
                        f"A1:{column_to_letter(sheet.LastColumn)}{sheet.LastRow}"
                    )
                if end_row < start_row or end_col < start_col:
                    raise ValidationError("End cell must be after start cell")

            return {
                "message": "Range is valid",
                "range": range_str,
                "dimensions": {
                    "start_row": start_row,
                    "start_col": start_col,
                    "end_row": end_row,
                    "end_col": end_col
                }
            }

        except ValueError as e:
            raise ValidationError(f"Invalid range format: {str(e)}")

    except ValidationError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to validate range: {e}")
        raise ValidationError(str(e))


def validate_range_in_sheet_operation(
        filepath: str,
        sheet_name: str,
        cell_range: str,
) -> dict[str, Any]:
    """Validate if a range exists in a worksheet and return data range info."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)
        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                break

        if sheet is None:
            raise ValidationError(f"Sheet '{sheet_name}' not found")

        worksheet = wb.Worksheets[sheet_name]

        # Get actual data dimensions
        data_max_row = worksheet.LastRow
        data_max_col = worksheet.LastColumn

        # Validate range
        valid_range = worksheet.Range[cell_range]
        start_row, start_col = valid_range.Row, valid_range.Column
        end_row, end_col = valid_range.LastRow, valid_range.LastColumn

        # Validate bounds against maximum possible Excel limits
        is_valid, message = validate_range_bounds(
            worksheet, start_row, start_col, end_row, end_col
        )
        if not is_valid:
            raise ValidationError(message)

        range_str = valid_range.RangeAddressLocal
        data_range_str = f"A1:{column_to_letter(data_max_col)}{data_max_row}"

        # Check if range is within data or extends beyond
        extends_beyond_data = (
                end_row > data_max_row or
                end_col > data_max_col
        )

        return {
            "message": (
                f"Range '{range_str}' is valid. "
                f"Sheet contains data in range '{data_range_str}'"
            ),
            "valid": True,
            "range": range_str,
            "data_range": data_range_str,
            "extends_beyond_data": extends_beyond_data,
            "data_dimensions": {
                "max_row": data_max_row,
                "max_col": data_max_col,
                "max_col_letter": column_to_letter(data_max_col)
            }
        }
    except ValidationError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to validate range: {e}")
        raise ValidationError(str(e))


def validate_range_bounds(
        worksheet: Worksheet,
        start_row: int,
        start_col: int,
        end_row: int | None = None,
        end_col: int | None = None,
) -> tuple[bool, str]:
    """Validate that cell range is within worksheet bounds"""
    max_row = worksheet.LastRow
    max_col = worksheet.LastColumn

    try:
        # Check start cell bounds
        if start_row < 1 or start_row > max_row:
            return False, f"Start row {start_row} out of bounds (1-{max_row})"
        if start_col < 1 or start_col > max_col:
            return False, (
                f"Start column {column_to_letter(start_col)} "
                f"out of bounds (A-{column_to_letter(max_col)})"
            )

        # If end cell specified, check its bounds
        if end_row is not None and end_col is not None:
            if end_row < start_row:
                return False, "End row cannot be before start row"
            if end_col < start_col:
                return False, "End column cannot be before start column"
            if end_row > max_row:
                return False, f"End row {end_row} out of bounds (1-{max_row})"
            if end_col > max_col:
                return False, (
                    f"End column {column_to_letter(end_col)} "
                    f"out of bounds (A-{column_to_letter(max_col)})"
                )

        return True, "Range is valid"
    except Exception as e:
        return False, f"Invalid range: {e!s}"
