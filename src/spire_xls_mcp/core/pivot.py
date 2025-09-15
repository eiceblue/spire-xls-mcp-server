from typing import Any, List, Dict
import logging

from spire.xls import *

from spire_xls_mcp.utils.cell_utils import EnumMapper
from spire_xls_mcp.utils.exceptions import ValidationError, PivotError
from .workbook import get_or_create_workbook

logger = logging.getLogger(__name__)


def create_pivot_table(
        input_filepath: str,
        output_filepath: str,
        data_sheet_name: str,
        pivot_sheet_name: str, 
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
        if not input_filepath and output_filepath:
            input_filepath = output_filepath
        elif input_filepath and not output_filepath:
            output_filepath = input_filepath
        elif not input_filepath and not output_filepath:
            raise PivotError("input_filepath or output_filepath must be provided")
        wb = get_or_create_workbook(input_filepath)
        data_sheet = None
        pivot_sheet = None

        # Find or create sheet
        for ws in wb.Worksheets:
            if ws.Name == data_sheet_name:
                data_sheet = ws
            if ws.Name == pivot_sheet_name:
                pivot_sheet = ws
        if data_sheet is None:
            raise PivotError(f"sheet '{data_sheet_name}' not found")
        if pivot_sheet is None:
            pivot_sheet = wb.CreateEmptySheet(pivot_sheet_name)

        cache = wb.PivotCaches.Add(data_sheet.Range[data_range])
        # Create pivot table
        pivot_table = pivot_sheet.PivotTables.Add(pivot_name, pivot_sheet.Range[locate_range], cache)

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
        wb.SaveToFile(output_filepath)

        return {
            "message": "Pivot table created successfully",
            "details": {
                "pivot_table_name": pivot_name,
                "pivot_sheet_name": pivot_sheet_name,
                "source_data_range": data_range,
                "location_range": locate_range,
                "structure": {
                    "rows": rows,
                    "columns": columns or [],
                    "values": values,
                    "default_aggregation": agg_func
                },
                "output_filepath": output_filepath
            }
        }

    except (ValidationError, PivotError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to create pivot table: {e}")
        raise PivotError(f"Failed to create pivot table: {e!s}")