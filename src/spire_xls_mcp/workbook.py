import logging
from pathlib import Path
from typing import Any

from spire.xls import *

from .cell_utils import column_to_letter
from .exceptions import WorkbookError

logger = logging.getLogger(__name__)


def create_workbook(filepath: str, sheet_name: str = None) -> dict[str, Any]:
    """Create a new Excel workbook with optional custom sheet name"""
    try:
        wb = Workbook()
        wb.Worksheets.RemoveAt(1)
        wb.Worksheets.RemoveAt(1)
        # Rename default sheet
        if sheet_name is not None:
            wb.Worksheets.Clear()
            sheet = wb.Worksheets.Add(sheet_name)

        save_path = Path(filepath)
        save_path.parent.mkdir(parents=True, exist_ok=True)
        wb.SaveToFile(str(save_path))
        return {
            "message": f"Created workbook: {filepath}",
            "active_sheet": sheet_name,
            "workbook": wb
        }
    except Exception as e:
        logger.error(f"Failed to create workbook: {e}")
        raise WorkbookError(f"Failed to create workbook: {e!s}")


def get_or_create_workbook(filepath: str) -> Workbook:
    """Get existing workbook or create new one if it doesn't exist"""
    try:
        wb = Workbook()
        if Path(filepath).exists():
            wb.LoadFromFile(filepath)
        else:
            wb = create_workbook(filepath)["workbook"]
        return wb
    except Exception as e:
        logger.error(f"Failed to get or create workbook: {e}")
        raise WorkbookError(f"Failed to get or create workbook: {e!s}")


def create_sheet(filepath: str, sheet_name: str) -> dict:
    """Create a new worksheet in the workbook if it doesn't exist."""
    try:
        wb = Workbook()
        wb.LoadFromFile(filepath)

        # Check if sheet already exists
        for sheet in wb.Worksheets:
            if sheet.Name == sheet_name:
                raise WorkbookError(f"Sheet {sheet_name} already exists")

        # Create new sheet
        wb.CreateEmptySheet(sheet_name)
        wb.SaveToFile(filepath)
        return {"message": f"Sheet {sheet_name} created successfully"}
    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to create sheet: {e}")
        raise WorkbookError(str(e))


def get_workbook_info(filepath: str, include_ranges: bool = False) -> dict[str, Any]:
    """Get metadata about workbook including sheets, ranges, etc."""
    try:
        path = Path(filepath)
        if not path.exists():
            raise WorkbookError(f"File not found: {filepath}")

        wb = Workbook()
        wb.LoadFromFile(filepath)

        sheets = [sheet.Name for sheet in wb.Worksheets]
        info = {
            "filename": path.name,
            "sheets": sheets,
            "size": path.stat().st_size,
            "modified": path.stat().st_mtime
        }

        if include_ranges:
            # Add used ranges for each sheet
            ranges = {}
            for sheet in wb.Worksheets:
                if sheet.LastRow > 0 and sheet.LastColumn > 0:
                    last_col = column_to_letter(sheet.LastColumn)
                    ranges[sheet.Name] = f"A1:{last_col}{sheet.LastRow}"
            info["used_ranges"] = ranges

        return info

    except WorkbookError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to get workbook info: {e}")
        raise WorkbookError(str(e))
