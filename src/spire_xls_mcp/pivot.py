from typing import Any
import logging

from spire.xls import *

from .cell_utils import EnumMapper
from .exceptions import ValidationError, PivotError
from .workbook import get_or_create_workbook

logger = logging.getLogger(__name__)


def create_pivot_table(
        filepath: str,
        sheet_name: str,
        pivot_name: str,
        data_range: str,
        locate_range: str,
        rows: List[str],
        values: dict[str, str],
        columns: List[str] = None,
        agg_func: str = "sum"
) -> dict[str, Any]:
    """Create pivot table in worksheet."""
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

        cache = wb.PivotCaches.Add(sheet.Range[data_range])
        # Create pivot table
        pivot_table = sheet.PivotTables.Add(pivot_name, sheet.Range[locate_range], cache)

        # Set aggregation function
        if agg_func.lower() not in EnumMapper.SUBTOTAL_MAP:
            raise PivotError(f"Unsupported aggregation function: {agg_func}")

        # Add row fields
        for row in rows:
            pivot_table.PivotFields[row].Axis = AxisTypes.Row

        # Add column fields
        if columns:
            for col in columns:
                pivot_table.PivotFields[col].Axis = AxisTypes.Column

        # Add value fields
        for value, name in values.items():
            field = pivot_table.PivotFields[value]
            subtotal = EnumMapper.get_subtotal_enum(agg_func.lower())
            # Drag the field to the data area.
            pivot_table.DataFields.Add(field, name, subtotal)
            # Save workbook
            wb.SaveToFile(filepath)

            return {
                "message": "Pivot table created successfully",
                "details": {
                    "source_range": data_range,
                    "pivot_sheet": sheet_name,
                    "rows": rows,
                    "columns": columns or [],
                    "values": values,
                    "aggregation": agg_func
                }
            }

    except (ValidationError, PivotError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to create pivot table: {e}")
        raise PivotError(f"Failed to create pivot table: {e!s}")
