import re
from typing import Tuple, Optional, Any
import datetime

from spire.xls import *


def letter_to_column(column_letter: str) -> int:
    """
    Convert Excel column letter to column number.
    Example: A -> 1, Z -> 26, AA -> 27, etc.
    
    Args:
        column_letter: Excel column letter (A-Z, AA-ZZ, etc.)
    
    Returns:
        Column number (1-based index)
    """
    if not column_letter or not column_letter.isalpha():
        try:
            # If it's already a number, just return it
            return int(column_letter)
        except ValueError:
            raise ValueError(f"Invalid column letter: {column_letter}")

    column_letter = column_letter.upper()
    result = 0
    for char in column_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def column_to_letter(column: int) -> str:
    """
    Convert column number to Excel column letter.
    Example: 1 -> A, 26 -> Z, 27 -> AA, etc.
    
    Args:
        column: Column number (1-based index)
    
    Returns:
        Excel column letter (A-Z, AA-ZZ, etc.)
    """
    if not isinstance(column, int) or column < 1:
        # If it's already a string, just return it
        if isinstance(column, str) and column.isalpha():
            return column
        raise ValueError(f"Invalid column number: {column}")

    result = ""
    while column > 0:
        remainder = (column - 1) % 26
        result = chr(ord('A') + remainder) + result
        column = (column - 1) // 26

    return result


def serialize_cell(cell):
    """Serialize a cell to a JSON-serializable dictionary object with null checks"""
    try:
        # Basic properties
        result = {
            "address": cell.RangeAddressLocal if hasattr(cell, "RangeAddressLocal") else None,
            "row": cell.Row if hasattr(cell, "Row") else None,
            "column": cell.Column if hasattr(cell, "Column") else None,
            "column_letter": column_to_letter(cell.Column) if hasattr(cell, "Column") else None,
            "value": None,
            "text": None,
            "formula": None,
            "has_formula": False
        }

        # Value handling (handle different types of values)
        try:
            result["value"] = cell.Value
        except:
            pass

        try:
            result["formula_value"] = cell.FormulaValue
        except:
            pass

        try:
            result["text"] = cell.Text
        except:
            pass

        # Formula handling
        try:
            result["has_formula"] = cell.HasFormula
            if cell.HasFormula:
                result["formula"] = cell.Formula
        except:
            pass

        # Style handling
        style_dict = {}

        # Font handling
        try:
            font_dict = {}
            if hasattr(cell, "Style") and hasattr(cell.Style, "Font"):
                font = cell.Style.Font
                if hasattr(font, "IsBold"):
                    font_dict["bold"] = font.IsBold
                if hasattr(font, "IsItalic"):
                    font_dict["italic"] = font.IsItalic
                if hasattr(font, "FontName"):
                    font_dict["name"] = font.FontName
                if hasattr(font, "Size"):
                    font_dict["size"] = font.Size
                if hasattr(font, "Color"):
                    try:
                        font_dict["color"] = {
                            "r": font.Color.R,
                            "g": font.Color.G,
                            "b": font.Color.B
                        }
                    except:
                        pass
                if hasattr(font, "Underline"):
                    font_dict["underline"] = str(font.Underline) != "None"

            if font_dict:
                style_dict["font"] = font_dict
        except:
            pass

        # Alignment handling
        try:
            if hasattr(cell, "Style"):
                if hasattr(cell.Style, "HorizontalAlignment"):
                    style_dict["horizontal_alignment"] = str(cell.Style.HorizontalAlignment)
                if hasattr(cell.Style, "VerticalAlignment"):
                    style_dict["vertical_alignment"] = str(cell.Style.VerticalAlignment)
                if hasattr(cell.Style, "WrapText"):
                    style_dict["wrap_text"] = cell.Style.WrapText
                if hasattr(cell.Style, "Rotation"):
                    style_dict["rotation"] = cell.Style.Rotation
                if hasattr(cell.Style, "IndentLevel"):
                    style_dict["indent_level"] = cell.Style.IndentLevel
        except:
            pass

        # Fill handling
        try:
            if hasattr(cell, "Style") and hasattr(cell.Style, "Interior"):
                interior_dict = {}
                interior = cell.Style.Interior
                if hasattr(interior, "Color"):
                    try:
                        interior_dict["color"] = {
                            "r": interior.Color.R,
                            "g": interior.Color.G,
                            "b": interior.Color.B
                        }
                    except:
                        pass
                if hasattr(interior, "FillPattern"):
                    interior_dict["pattern"] = str(interior.FillPattern)

                if interior_dict:
                    style_dict["fill"] = interior_dict
        except:
            pass

        # Borders handling
        try:
            if hasattr(cell, "Style") and hasattr(cell.Style, "Borders"):
                borders_dict = {}
                borders = cell.Style.Borders

                if hasattr(borders, "LineStyle"):
                    borders_dict["line_style"] = str(borders.LineStyle)
                if hasattr(borders, "Color"):
                    try:
                        borders_dict["color"] = {
                            "r": borders.Color.R,
                            "g": borders.Color.G,
                            "b": borders.Color.B
                        }
                    except:
                        pass

                # Individual borders
                border_positions = ["Top", "Bottom", "Left", "Right"]
                for pos in border_positions:
                    try:
                        if hasattr(borders, pos):
                            border = getattr(borders, pos)
                            border_info = {}
                            if hasattr(border, "LineStyle"):
                                border_info["line_style"] = str(border.LineStyle)
                            if hasattr(border, "Color"):
                                try:
                                    border_info["color"] = {
                                        "r": border.Color.R,
                                        "g": border.Color.G,
                                        "b": border.Color.B
                                    }
                                except:
                                    pass
                            if border_info:
                                borders_dict[pos.lower()] = border_info
                    except:
                        pass

                if borders_dict:
                    style_dict["borders"] = borders_dict
        except:
            pass

        # Number format
        try:
            if hasattr(cell, "Style") and hasattr(cell.Style, "NumberFormat"):
                style_dict["number_format"] = cell.Style.NumberFormat
        except:
            pass

        # Cell protection
        try:
            if hasattr(cell, "Style"):
                protection_dict = {}
                if hasattr(cell.Style, "HideFormula"):
                    protection_dict["hide_formula"] = cell.Style.HideFormula

                if protection_dict:
                    style_dict["protection"] = protection_dict
        except:
            pass

        # Add style to result
        if style_dict:
            result["style"] = style_dict

        # Cell type
        try:
            if hasattr(cell, "Type"):
                result["cell_type"] = str(cell.Type)
        except:
            pass

        # Merged cells information
        try:
            if hasattr(cell, "IsMerged"):
                result["is_merged"] = cell.IsMerged
                if cell.IsMerged and hasattr(cell, "MergeArea"):
                    try:
                        area = cell.MergeArea
                        result["merge_area"] = {
                            "first_row": area.FirstRow,
                            "first_column": area.FirstColumn,
                            "last_row": area.LastRow,
                            "last_column": area.LastColumn
                        }
                    except:
                        pass
        except:
            pass

        return result
    except Exception as e:
        # Return basic information when any exception occurs
        return {
            "error": f"Failed to serialize cell: {str(e)}",
            "row": getattr(cell, "Row", None),
            "column": getattr(cell, "Column", None),
            "value": getattr(cell, "Value", None)
        }


def parse_cell_range(
        start_cell: str,
        end_cell: Optional[str] = None,
        workbook: Workbook = Workbook()
) -> Tuple[int, int, Optional[int], Optional[int]]:
    """Parse Excel cell references into row and column numbers."""
    try:
        start_cell = workbook.Worksheets[0].Range[start_cell]
    except:
        raise ValueError(f"Invalid start cell reference: {start_cell}")
    start_col = start_cell.Column
    start_row = start_cell.Row
    # Parse end cell if provided
    end_col = None
    end_row = None
    if end_cell:
        try:
            end_cell = workbook.Worksheets[0].Range[end_cell]
        except:
            raise ValueError(f"Invalid end cell reference: {end_cell}")
        end_col = end_cell.Column
        end_row = end_cell.Row

    return start_row, start_col, end_row, end_col


def validate_cell_reference_regex(cell_ref: str) -> bool:
    """Validate Excel cell reference format."""
    if not cell_ref:
        return False

    # Basic format validation
    pattern = r'^[A-Za-z]{1,3}[1-9][0-9]*$'
    if not re.match(pattern, cell_ref):
        return False
    return True


class EnumMapper:
    SUBTOTAL_MAP = {
        "sum": SubtotalTypes.Sum,
        "average": SubtotalTypes.Average,
        "count": SubtotalTypes.Count,
        "min": SubtotalTypes.Min,
        "max": SubtotalTypes.Max
    }

    # Filter operator type mapping
    FILTER_OPERATOR_MAP = {
        "=": FilterOperatorType.Equal,
        ">": FilterOperatorType.GreaterThan,
        "<": FilterOperatorType.LessThan,
        ">=": FilterOperatorType.GreaterOrEqual,
        "<=": FilterOperatorType.LessOrEqual,
        "<>": FilterOperatorType.NotEqual
    }

    # condition format operator 
    OPERATOR_MAP = {
        "greater": ComparisonOperatorType.Greater,
        "gt": ComparisonOperatorType.Greater,
        ">": ComparisonOperatorType.Greater,
        "大于": ComparisonOperatorType.Greater,
        "greater than": ComparisonOperatorType.Greater,
        "greater_or_equal": ComparisonOperatorType.GreaterOrEqual,
        ">=": ComparisonOperatorType.GreaterOrEqual,
        "ge": ComparisonOperatorType.GreaterOrEqual,
        "大于等于": ComparisonOperatorType.GreaterOrEqual,
        "less": ComparisonOperatorType.Less,
        "lt": ComparisonOperatorType.Less,
        "<": ComparisonOperatorType.Less,
        "小于": ComparisonOperatorType.Less,
        "less than": ComparisonOperatorType.Less,
        "less_or_equal": ComparisonOperatorType.LessOrEqual,
        "le": ComparisonOperatorType.LessOrEqual,
        "<=": ComparisonOperatorType.LessOrEqual,
        "小于等于": ComparisonOperatorType.LessOrEqual,
        "equal": ComparisonOperatorType.Equal,
        "eq": ComparisonOperatorType.Equal,
        "=": ComparisonOperatorType.Equal,
        "等于": ComparisonOperatorType.Equal,
        "not_equal": ComparisonOperatorType.NotEqual,
        "ne": ComparisonOperatorType.NotEqual,
        "!=": ComparisonOperatorType.NotEqual,
        "<>": ComparisonOperatorType.NotEqual,
        "不等于": ComparisonOperatorType.NotEqual,
    }
    # alignment
    ALIGNMENT_MAP = {
        "left": HorizontalAlignType.Left,
        "居左": HorizontalAlignType.Left,
        "center": HorizontalAlignType.Center,
        "居中": HorizontalAlignType.Center,
        "right": HorizontalAlignType.Right,
        "居右": HorizontalAlignType.Right,
        "justify": HorizontalAlignType.Justify,
        "两端对齐": HorizontalAlignType.Justify,
    }
    # border line style
    BORDER_STYLE_MAP = {
        "thin": LineStyleType.Thin,
        "细线": LineStyleType.Thin,
        "medium": LineStyleType.Medium,
        "中线": LineStyleType.Medium,
        "thick": LineStyleType.Thick,
        "粗线": LineStyleType.Thick,
        "double": LineStyleType.Double,
        "双线": LineStyleType.Double,
    }
    # chart type
    CHART_TYPE_MAP = {
        "column": ExcelChartType.ColumnClustered,
        "bar": ExcelChartType.BarClustered,
        "line": ExcelChartType.Line,
        "pie": ExcelChartType.Pie,
        "area": ExcelChartType.Area,
        "scatter": ExcelChartType.ScatterLine,
        "doughnut": ExcelChartType.Doughnut,
    }
    # conditional format type
    CONDITION_TYPE_MAP = {
        "cell": ConditionalFormatType.CellValue,
        "单元格值": ConditionalFormatType.CellValue,
        "text": ConditionalFormatType.ContainsText,
        "文本": ConditionalFormatType.ContainsText,
        "date": ConditionalFormatType.TimePeriod,
        "日期": ConditionalFormatType.TimePeriod,
        "time_period": ConditionalFormatType.TimePeriod,
        "时间段": ConditionalFormatType.TimePeriod,
        "average": ConditionalFormatType.Average,
        "平均值": ConditionalFormatType.Average,
        "duplicate": ConditionalFormatType.DuplicateValues,
        "重复值": ConditionalFormatType.DuplicateValues,
        "unique": ConditionalFormatType.UniqueValues,
        "唯一值": ConditionalFormatType.UniqueValues,
        "formula": ConditionalFormatType.Formula,
        "公式": ConditionalFormatType.Formula,
        "top10": ConditionalFormatType.TopBottom,
        "前10项": ConditionalFormatType.TopBottom,
        "data_bar": ConditionalFormatType.DataBar,
        "数据条": ConditionalFormatType.DataBar,
        "color_scale": ConditionalFormatType.ColorScale,
        "色阶": ConditionalFormatType.ColorScale,
        "icon_set": ConditionalFormatType.IconSet,
        "图标集": ConditionalFormatType.IconSet
    }

    @staticmethod
    def smart_enum_map(input_str: str, mapping: dict, default):
        if not input_str:
            return default
        key = input_str.strip().lower()
        return mapping.get(key, default)

    @classmethod
    def get_operator_enum(cls, op_str: str) -> ComparisonOperatorType:
        return cls.smart_enum_map(op_str, cls.OPERATOR_MAP, ComparisonOperatorType.Greater)

    @classmethod
    def get_alignment_enum(cls, align_str: str) -> HorizontalAlignType:
        return cls.smart_enum_map(align_str, cls.ALIGNMENT_MAP, HorizontalAlignType.Left)

    @classmethod
    def get_border_style_enum(cls, style_str: str) -> LineStyleType:
        return cls.smart_enum_map(style_str, cls.BORDER_STYLE_MAP, LineStyleType.Thin)

    @classmethod
    def get_chart_type_enum(cls, chart_str: str) -> ExcelChartType:
        return cls.smart_enum_map(chart_str, cls.CHART_TYPE_MAP, ExcelChartType.ColumnClustered)

    @classmethod
    def get_condition_enum(cls, type_str: str) -> ConditionalFormatType:
        return cls.smart_enum_map(type_str, cls.CONDITION_TYPE_MAP, ConditionalFormatType.CellValue)

    @classmethod
    def get_subtotal_enum(cls, func_str: str) -> SubtotalTypes:
        return cls.smart_enum_map(func_str, cls.SUBTOTAL_MAP, SubtotalTypes.Sum)

    @classmethod
    def get_filter_operator_enum(cls, op_str: str) -> FilterOperatorType:
        return cls.smart_enum_map(op_str, cls.FILTER_OPERATOR_MAP, FilterOperatorType.Equal)


def create_spire_object(value: Any) -> Any:
    """
    Create corresponding SpireObject based on the input value's type.

    Args:
        value: The value to be converted to SpireObject

    Returns:
        SpireObject instance of corresponding type

    Examples:
        >>> create_spire_object(123)  # Returns Int32(123)
        >>> create_spire_object(123.45)  # Returns Double(123.45)
        >>> create_spire_object("text")  # Returns String("text")
        >>> create_spire_object(True)  # Returns Boolean(True)
        >>> create_spire_object(datetime.datetime(2023, 1, 1))  # Returns DateTime(2023, 1, 1)
    """
    if value is None:
        return String("")

    if isinstance(value, bool):
        return Boolean(value)

    if isinstance(value, int):
        if -2147483648 <= value <= 2147483647:
            return Int32(value)
        else:
            return Int64(value)

    if isinstance(value, float):
        return Double(value)

    if isinstance(value, str):
        return String(value)

    if isinstance(value, (datetime.datetime, datetime.date)):
        return DateTime(value.year, value.month, value.day,
                        getattr(value, 'hour', 0),
                        getattr(value, 'minute', 0),
                        getattr(value, 'second', 0),
                        getattr(value, 'microsecond', 0))

    return String(str(value))
