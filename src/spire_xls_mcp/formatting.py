import logging
from typing import Any, Dict

from spire.xls import *

from .workbook import get_or_create_workbook
from .cell_utils import parse_cell_range, validate_cell_reference_regex, EnumMapper
from .exceptions import ValidationError, FormattingError

logger = logging.getLogger(__name__)

def format_range(
    filepath: str,
    sheet_name: str,
    cell_range: str,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int = None,
    font_color: str = None,
    bg_color: str = None,
    border_style: str = None,
    border_color: str = None,
    number_format: str = None,
    alignment: str = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Dict[str, Any] = None,
    conditional_format: Dict[str, Any] = None
) -> Dict[str, Any]:
    """Apply formatting to a range of cells.
    
    This function handles all Excel formatting operations including:
    - Font properties (bold, italic, size, color, etc.)
    - Cell fill/background color
    - Borders (style and color)
    - Number formatting
    - Alignment and text wrapping
    - Cell merging
    - Protection
    - Conditional formatting
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Starting cell reference
        end_cell: Optional ending cell reference
        bold: Whether to make text bold
        italic: Whether to make text italic
        underline: Whether to underline text
        font_size: Font size in points
        font_color: Font color (hex code)
        bg_color: Background color (hex code)
        border_style: Border style (thin, medium, thick, double)
        border_color: Border color (hex code)
        number_format: Excel number format string
        alignment: Text alignment (left, center, right, justify)
        wrap_text: Whether to wrap text
        merge_cells: Whether to merge the range
        protection: Cell protection settings
        conditional_format: Conditional formatting rules
        
    Returns:
        Dictionary with operation status
    """
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
        
        # Get the range to format
        range_to_format = sheet.Range[cell_range]
        
        # Apply font formatting
        if any([bold, italic, underline, font_size, font_color]):
            style = range_to_format.Style
            if bold:
                style.Font.IsBold = True
            if italic:
                style.Font.IsItalic = True
            if underline:
                style.Font.Underline = UnderlineStyle.Continuous
            if font_size:
                style.Font.Size = font_size
            if font_color:
                # Convert hex color to RGB
                if font_color.startswith('#'):
                    font_color = font_color[1:]
                if len(font_color) == 6:
                    r = int(font_color[0:2], 16)
                    g = int(font_color[2:4], 16)
                    b = int(font_color[4:6], 16)
                    style.Font.Color = Color.FromRgb(r, g, b)
        
        # Apply fill
        if bg_color:
            style = range_to_format.Style
            if bg_color.startswith('#'):
                bg_color = bg_color[1:]
            if len(bg_color) == 6:
                r = int(bg_color[0:2], 16)
                g = int(bg_color[2:4], 16)
                b = int(bg_color[4:6], 16)
                # Set filling pattern type
                style.Interior.FillPattern = ExcelPatternType.Gradient
                # Set filling Background color
                style.Interior.Gradient.BackColor = Color.FromRgb(r, g, b)
        
        # Apply borders
        if border_style or border_color:
            style = range_to_format.Style
            border_line_style = EnumMapper.get_border_style_enum(border_style)
            if border_color:
                if border_color.startswith('#'):
                    border_color = border_color[1:]
                if len(border_color) == 6:
                    r = int(border_color[0:2], 16)
                    g = int(border_color[2:4], 16)
                    b = int(border_color[4:6], 16)
                    border_color_obj = Color.FromRgb(r, g, b)
                    style.Borders.Color = border_color_obj
            style.Borders.LineStyle = border_line_style
        
        # Apply number format
        if number_format:
            style = range_to_format.Style
            style.NumberFormat = number_format
        
        # Apply alignment
        if alignment:
            style = range_to_format.Style
            style.HorizontalAlignment = EnumMapper.get_alignment_enum(alignment)
        
        # Apply text wrapping
        if wrap_text:
            style = range_to_format.Style
            style.WrapText = True
        
        # Apply cell merging
        if merge_cells:
            range_to_format.Merge()
        
        # Apply protection
        if protection:
            style = range_to_format.Style
            if 'locked' in protection:
                style.Locked = protection['locked']
            if 'hidden' in protection:
                style.HideFormula = protection['hidden']
        
        # Apply conditional formatting
        if conditional_format:
            op_str = conditional_format.get("criteria", "greater")
            operator = EnumMapper.get_operator_enum(op_str)
            xcf = sheet.ConditionalFormats.Add()
            xcf.AddRange(range_to_format)
            fmt = xcf.AddCondition()
            type = conditional_format.get("type", "cell")
            fmt.FormatType = EnumMapper.get_condition_enum(type)
            fmt.Operator = operator
            fmt.FirstFormula = conditional_format.get("first_formula", str(conditional_format.get("value", "0")))
            fmt.SecondFormula = conditional_format.get("second_formula", str(conditional_format.get("value2", "0")))
            if "format" in conditional_format:
                format = conditional_format["format"]
                if "font_color" in format:
                    color = format["font_color"]
                    if color.startswith("#"):
                        color = color[1:]
                    if len(color) == 6:
                        r = int(color[0:2], 16)
                        g = int(color[2:4], 16)
                        b = int(color[4:6], 16)
                        fmt.FontColor = Color.FromRgb(r, g, b)
                if "bg_color" in format:
                    color = format["bg_color"]
                    if color.startswith("#"):
                        color = color[1:]
                    if len(color) == 6:
                        r = int(color[0:2], 16)
                        g = int(color[2:4], 16)
                        b = int(color[4:6], 16)
                        fmt.BackColor = Color.FromRgb(r, g, b)
        
        # Save changes
        wb.SaveToFile(filepath)
        
        return {
            "message": "Formatting applied successfully"
        }
        
    except (ValidationError, FormattingError) as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to apply formatting: {e}")
        raise FormattingError(str(e))
