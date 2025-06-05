import logging
from typing import Any

from spire.xls import *

from .exceptions import ValidationError, CalculationError
from .validation import validate_formula

logger = logging.getLogger(__name__)


def apply_formula(
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str
) -> dict[str, Any]:
    """Apply Excel formula to cell."""
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
            
        # Apply formula
        try:
            cell_range = sheet.Range[cell]
            cell_range.Formula = formula
            cell_range.CalculateAllValue()
            result = cell_range.FormulaValue
            wb.SaveToFile(filepath)
            
            return {
                "message": "Formula applied successfully",
                "cell": cell,
                "formula": formula,
                "result": result
            }
        except Exception as e:
            raise CalculationError(f"Failed to apply formula: {str(e)}")
            
    except ValidationError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to apply formula: {e}")
        raise CalculationError(str(e))
