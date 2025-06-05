import logging
from typing import Any
import base64

from spire.xls import *

from .cell_utils import parse_cell_range, column_to_letter, EnumMapper, create_spire_object
from .exceptions import SheetError, ValidationError

logger = logging.getLogger(__name__)


def copy_sheet(filepath: str, source_sheet: str, target_sheet: str) -> dict[str, Any]:
    """Copy a worksheet within the same workbook."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)

        source = None
        for ws in wb.Worksheets:
            if ws.Name == source_sheet:
                source = ws
                break

        if source is None:
            raise SheetError(f"Source sheet '{source_sheet}' not found")

        for ws in wb.Worksheets:
            if ws.Name == target_sheet:
                raise SheetError(f"Target sheet '{target_sheet}' already exists")

        # Copy sheet
        new_sheet = wb.Worksheets.AddCopy(source)
        new_sheet.Name = target_sheet

        wb.SaveToFile(filepath)
        return {"message": f"Sheet '{source_sheet}' copied to '{target_sheet}'"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to copy sheet: {e}")
        raise SheetError(str(e))


def delete_sheet(filepath: str, sheet_name: str) -> dict[str, Any]:
    """Delete a worksheet from the workbook."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)

        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                break

        if sheet is None:
            raise SheetError(f"Sheet '{sheet_name}' not found")

        if wb.Worksheets.Count == 1:
            raise SheetError("Cannot delete the only sheet in workbook")

        wb.Worksheets.Remove(sheet)
        wb.SaveToFile(filepath)
        return {"message": f"Sheet '{sheet_name}' deleted"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to delete sheet: {e}")
        raise SheetError(str(e))


def rename_sheet(filepath: str, old_name: str, new_name: str) -> dict[str, Any]:
    """Rename a worksheet."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)

        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == old_name:
                sheet = ws
                break

        if sheet is None:
            raise SheetError(f"Sheet '{old_name}' not found")

        for ws in wb.Worksheets:
            if ws.Name == new_name:
                raise SheetError(f"Sheet '{new_name}' already exists")

        sheet.Name = new_name
        wb.SaveToFile(filepath)
        return {"message": f"Sheet renamed from '{old_name}' to '{new_name}'"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to rename sheet: {e}")
        raise SheetError(str(e))


def format_range_string(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """Format range string from row and column indices."""
    return f"{column_to_letter(start_col)}{start_row}:{column_to_letter(end_col)}{end_row}"


def copy_range(
        filepath: str,
        sheet_name: str,
        source_start: str,
        source_end: str,
        target_start: str,
        target_sheet: str = None
) -> dict[str, Any]:
    """Copy a range of cells to another location."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)

        source_ws = None
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                source_ws = ws
                break

        if source_ws is None:
            raise SheetError(f"Source sheet '{sheet_name}' not found")

        # Get target worksheet
        target_ws = source_ws
        if target_sheet:
            target_ws = None
            for ws in wb.Worksheets:
                if ws.Name == target_sheet:
                    target_ws = ws
                    break
            if target_ws is None:
                raise SheetError(f"Target sheet '{target_sheet}' not found")

        # Parse ranges
        source_start_row, source_start_col, source_end_row, source_end_col = parse_cell_range(source_start, source_end)
        target_start_row, target_start_col, _, _ = parse_cell_range(target_start)

        if source_end_row is None or source_end_col is None:
            raise SheetError("Source range must specify both start and end cells")

        # Calculate dimensions
        rows = source_end_row - source_start_row + 1
        cols = source_end_col - source_start_col + 1

        # Get source range
        source_range = source_ws.Range[source_start_row, source_start_col, source_end_row, source_end_col]

        # Get target range
        target_range = target_ws.Range[target_start_row, target_start_col,
        target_start_row + rows - 1, target_start_col + cols - 1]
        # Copy range
        CellRange(source_range.Ptr).Copy(CellRange(target_range.Ptr))

        wb.SaveToFile(filepath)

        return {
            "message": f"Range copied successfully",
            "details": {
                "source_sheet": sheet_name,
                "source_range": f"{source_start}:{source_end}",
                "target_sheet": target_sheet or sheet_name,
                "target_start": target_start
            }
        }
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to copy range: {e}")
        raise SheetError(str(e))


def delete_range(
        filepath: str,
        sheet_name: str,
        cell_range: str,
        shift_direction: str = "up"
) -> dict[str, Any]:
    """Delete a range of cells and shift remaining cells."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)

        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                break

        if sheet is None:
            raise SheetError(f"Sheet '{sheet_name}' not found")

        # Get range to delete
        range_to_delete = sheet.Range[cell_range]

        # Delete range and shift cells
        if shift_direction.lower() == "up":
            sheet.DeleteRange(range_to_delete, DeleteOption.MoveUp)
        else:
            sheet.DeleteRange(range_to_delete, DeleteOption.MoveLeft)

        wb.SaveToFile(filepath)

        return {
            "message": f"Range deleted and cells shifted {shift_direction}",
            "range": f"{range_to_delete.RangeAddressLocal}"
        }
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to delete range: {e}")
        raise SheetError(str(e))


def merge_range(filepath: str, sheet_name: str, cell_range_list: List[str]) -> dict[str, Any]:
    """Merge a range of cells."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)

        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                break

        if sheet is None:
            raise SheetError(f"Sheet '{sheet_name}' not found")

        for cell in cell_range_list:
            sheet.Range[cell].Merge()

        wb.SaveToFile(filepath)
        return {"message": f"Range merged success in sheet '{sheet_name}'"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to merge range: {e}")
        raise SheetError(str(e))


def unmerge_range(filepath: str, sheet_name: str, cell_range: str) -> dict[str, Any]:
    """Unmerge a range of cells."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)

        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                break

        if sheet is None:
            raise SheetError(f"Sheet '{sheet_name}' not found")

        # Create range string
        range_to_unmerge = sheet.Range[cell_range]
        range_to_unmerge.UnMerge()

        wb.SaveToFile(filepath)
        return {"message": f"Range '{range_to_unmerge.RangeAddressLocal}' unmerged in sheet '{sheet_name}'"}
    except SheetError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to unmerge range: {e}")
        raise SheetError(str(e))


def copy_range_operation(
        filepath: str,
        sheet_name: str,
        source_range: str,
        target_range: str,
        target_sheet: str = None
) -> dict:
    """Copy a range of cells to another location."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)
        sheet = None
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                break

        if sheet is None:
            logger.error(f"Sheet '{sheet_name}' not found")
            raise ValidationError(f"Sheet '{sheet_name}' not found")

        source_ws = wb.Worksheets[sheet_name]
        target_ws = wb.Worksheets[target_sheet] if target_sheet else source_ws

        source_ws.Range[source_range].Copy(target_ws.Range[target_range], True, True)

        wb.SaveToFile(filepath)
        return {"message": f"Range copied successfully"}

    except (ValidationError, SheetError):
        raise
    except Exception as e:
        logger.error(f"Failed to copy range: {e}")
        raise SheetError(f"Failed to copy range: {str(e)}")


def apply_autofilter(
        filepath: str,
        sheet_name: str,
        cell_range: str,
        filter_criteria: dict = None
) -> dict:
    """
    Apply autofilter to a range of cells and optionally set filter criteria.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        cell_range: Range to apply autofilter (e.g. "A1:D10")
        filter_criteria: Optional dictionary of filter criteria
            Key: Column index (0-based)
            Value: Dictionary with filter settings

    Returns:
        Dictionary with result message
    """
    try:
        # Load workbook
        workbook = Workbook()
        workbook.LoadFromFile(filepath)
        
        # Ensure worksheet exists
        if sheet_name not in [sheet.Name for sheet in workbook.Worksheets]:
            raise SheetError(f"Worksheet '{sheet_name}' does not exist")
        
        sheet = workbook.Worksheets[sheet_name]
        
        # Apply auto filter
        try:
            auto_filters = sheet.AutoFilters
            auto_filters.Range = sheet.Range[cell_range]
            if filter_criteria:
                for col_index, criteria in filter_criteria.items():
                    filter_column = auto_filters[col_index]
                    
                    if criteria.get("type") == "value":
                        filter_values = criteria.get("values", [])
                        for value in filter_values:
                            auto_filters.AddFilter(filter_column, str(value))
                    
                    elif criteria.get("type") == "custom":
                        operator = criteria.get("operator")
                        criteria_value = criteria.get("criteria")
                        filter_operator = EnumMapper.get_filter_operator_enum(operator)
                        
                        spire_value = create_spire_object(criteria_value)
                        auto_filters.CustomFilter(filter_column, filter_operator, spire_value)

                    elif criteria.get("type") == "top10":
                        count = criteria.get("count", 10)
                        percent = criteria.get("percent", False)
                        bottom = criteria.get("bottom", False)

                        auto_filters.FilterTop10(filter_column, not bottom, percent, count)
                    auto_filters.Filter()
        except Exception as e:
            raise ValidationError(f"Error applying autofilter: {str(e)}")
        
        # Save workbook
        workbook.SaveToFile(filepath)
        workbook.Dispose()
        
        return {"message": "Autofilter successfully applied"}
    except Exception as e:
        raise SheetError(f"Failed to apply autofilter: {str(e)}")


def get_shape_image_base64(filepath, sheet_name, shape_name=None, shape_index=None):
    """
    Export a specified Shape object from an Excel worksheet as an image and return its base64 string.

    Args:
        filepath (str): Path to the Excel file.
        sheet_name (str): Name of the worksheet containing the shape.
        shape_name (str, optional): Name of the shape to export. If not provided, shape_index must be specified.
        shape_index (int, optional): Index of the shape in the worksheet (0-based). Used if shape_name is not provided.

    Returns:
        str: Base64-encoded string of the exported shape image.

    Raises:
        ValueError: If neither shape_name nor shape_index is provided, or if the shape is not found.
        Exception: For other errors during export or file operations.
    """

    workbook = Workbook()
    workbook.LoadFromFile(filepath)
    sheet = workbook.Worksheets[sheet_name]

    # Get Shape
    if shape_name:
        shape = next((s for s in sheet.PrstGeomShapes if s.Name == shape_name),
                     next((s for s in sheet.Pictures if s.Name == shape_name), None))
        if shape is None:
            raise ValueError(f"Shape '{shape_name}' not found in sheet '{sheet_name}'")
    elif shape_index is not None:
        if shape_index < sheet.PrstGeomShapes.Count:
            shape = sheet.PrstGeomShapes[shape_index]
        elif shape_index < sheet.Pictures.Count:
            shape = sheet.Pictures[shape_index]
        else:
            raise ValueError(f"Shape '{shape_index}' not found in sheet '{sheet_name}'")
    else:
        raise ValueError("Must provide shape_name or shape_index")

    # Export as image to memory stream
    img_bytes = shape.SaveToImage().ToArray()
    workbook.Dispose()

    # Convert to base64
    base64_str = base64.b64encode(img_bytes).decode("utf-8")
    return base64_str
