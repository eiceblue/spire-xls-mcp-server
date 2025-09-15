import json
import logging
import sys
import os
import functools
from typing import Any, List, Dict, Optional

from mcp.server.fastmcp import FastMCP

# Import exceptions
from spire_xls_mcp.utils.exceptions import (
    ExcelMCPError,
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    FormattingError,
    CalculationError,
    PivotError,
    ChartError,
    ConversionError
)

# Import core functions
from spire_xls_mcp.core.calculations import apply_formula as apply_formula_impl
from spire_xls_mcp.core.chart import create_chart_in_sheet as create_chart_impl
from spire_xls_mcp.core.conversion import convert_workbook as convert_workbook_impl
from spire_xls_mcp.core.data import read_excel_range, write_data
from spire_xls_mcp.core.formatting import format_range as format_range_func
from spire_xls_mcp.core.json_operations import export_to_json as export_json_impl, import_from_json as import_json_impl
from spire_xls_mcp.core.pivot import create_pivot_table as create_pivot_table_impl
from spire_xls_mcp.core.sheet import (
    copy_sheet,
    delete_sheet,
    rename_sheet,
    merge_range,
    unmerge_range,
    copy_range_operation,
    delete_range as delete_range_operation,
    apply_autofilter as apply_autofilter_impl,
    get_shape_image_base64 as get_shape_img_b64
)
from spire_xls_mcp.core.validation import validate_range_in_sheet_operation as validate_range_impl
from spire_xls_mcp.core.workbook import get_workbook_info, create_workbook as create_workbook_impl, create_sheet as create_worksheet_impl

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("spire-xls-mcp.log")
    ],
    force=True
)

logger = logging.getLogger("spire-xls-mcp")

# Get Excel files path from environment or use default
EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")

# Initialize FastMCP server
mcp = FastMCP(
    "spire-xls-mcp",
    version="0.1.1",
    description="Spire.Xls MCP Server for manipulating Excel files",
    dependencies=["Spire.Xls.Free>=14.12.4"],
    env_vars={
        "EXCEL_FILES_PATH": {
            "description": "Path to Excel files directory",
            "required": False,
            "default": EXCEL_FILES_PATH
        }
    }
)

def get_excel_path(filename: str) -> str:
    """Get full path to Excel file.
    
    Args:
        filename: Name of Excel file
        
    Returns:
        Full path to Excel file
    """
    if not filename or os.path.isabs(filename):
        return filename
    return os.path.join(EXCEL_FILES_PATH, filename)

def tool_exception_handler(func):
    """A decorator to handle exceptions for all tool functions."""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except ExcelMCPError as e:
            logger.error(f"Error in tool '{func.__name__}': {e}")
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"An unexpected error occurred in tool '{func.__name__}': {e}")
            raise
    return wrapper

@mcp.tool()
@tool_exception_handler
def apply_formula(
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
) -> str:
    """
    Applies an Excel formula to a specified cell with verification.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet
        cell (str): Cell reference where formula will be applied (e.g., "A1")
        formula (str): Excel formula to apply (must include "=" prefix)

    Returns:
        str: Success message confirming formula application
    """
    full_path = get_excel_path(filepath)
    result = apply_formula_impl(full_path, sheet_name, cell, formula)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
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
) -> str:
    """
    Applies formatting to a range of cells.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet
        cell_range (str): Range of cells to format (e.g., "A1:C5")
        bold (bool, optional): Whether to apply bold formatting
        italic (bool, optional): Whether to apply italic formatting
        underline (bool, optional): Whether to apply underline formatting
        font_size (int, optional): Font size to apply
        font_color (str, optional): Font color as hex code (e.g., "#FF0000")
        bg_color (str, optional): Background color as hex code (e.g., "#FFFF00")
        border_style (str, optional): Border style (thin, medium, thick, double)
        border_color (str, optional): Border color as hex code (e.g., "#000000")
        number_format (str, optional): Excel number format code
        alignment (str, optional): Text alignment (left, center, right, justify)
        wrap_text (bool, optional): Whether to enable text wrapping
        merge_cells (bool, optional): Whether to merge the cells in the range
        protection (dict, optional): Cell protection settings dict with keys 'locked' and/or 'hidden'
        conditional_format (dict, optional): Conditional formatting rules dict with keys:
            - "type": Condition type, supported values:
                - "cell"/"单元格值": Cell value condition
                - "text"/"文本": Contains text condition
                - "date"/"日期": Date condition
                - "time_period"/"时间段": Time period condition
                - "average"/"平均值": Average condition
                - "duplicate"/"重复值": Duplicate values condition
                - "unique"/"唯一值": Unique values condition
                - "formula"/"公式": Formula condition
                - "top10"/"前10项": Top/Bottom items condition
                - "data_bar"/"数据条": Data bar condition
                - "color_scale"/"色阶": Color scale condition
                - "icon_set"/"图标集": Icon set condition
            - "criteria"/"operator": Comparison operator, supported values:
                - "greater"/"gt"/">"/"大于": Greater than
                - "greater_or_equal"/">="/"ge"/"大于等于": Greater than or equal
                - "less"/"lt"/"<"/"小于": Less than
                - "less_or_equal"/"<="/"le"/"小于等于": Less than or equal
                - "equal"/"eq"/"="/"/"等于": Equal
                - "not_equal"/"ne"/"!="/"<>"/"不等于": Not equal
            - "value"/"first_formula": First value/formula for comparison
            - "value2"/"second_formula": Second value/formula for comparison (optional)
            - "format": Dict with formatting to apply when condition is met:
                - "font_color": Hex color code for font
                - "bg_color": Hex color code for background

    Returns:
        str: Success message confirming formatting was applied
    """
    full_path = get_excel_path(filepath)
    result = format_range_func(
        filepath=full_path,
        sheet_name=sheet_name,
        cell_range=cell_range,
        bold=bold,
        italic=italic,
        underline=underline,
        font_size=font_size,
        font_color=font_color,
        bg_color=bg_color,
        border_style=border_style,
        border_color=border_color,
        number_format=number_format,
        alignment=alignment,
        wrap_text=wrap_text,
        merge_cells=merge_cells,
        protection=protection,
        conditional_format=conditional_format
    )
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def read_data_from_excel(
        filepath: str,
        sheet_name: str,
        cell_range: str,
        preview_only: bool = False
) -> str:
    """
    Reads data from an Excel worksheet.

    Returns data in column-first format where cells can be accessed as data[column_letter][row_number].
    Each cell contains detailed information including value, formula, style properties, etc.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet to read from
        cell_range (str): Range of cells to read (e.g., "A1:D10")
        preview_only (bool, optional): If True, returns only preview data without full styling info

    Returns:
        dict: Column-first nested dictionary with cell data
    """
    full_path = get_excel_path(filepath)
    result = read_excel_range(full_path, sheet_name, cell_range, preview_only)
    if not result:
        return "No data found in specified range"
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def write_data_to_excel(
        filepath: str,
        sheet_name: str,
        data: List[List],
        start_cell: str = "A1",
) -> str:
    """
    Writes data to an Excel worksheet.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet to write to
        data (list): List of lists containing data to write (rows of data)
        start_cell (str, optional): Cell to start writing from, default is "A1"

    Returns:
        str: Success message confirming data was written
    """
    full_path = get_excel_path(filepath)
    result = write_data(full_path, sheet_name, data, start_cell)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def create_workbook(filepath: str, sheet_name: str = None) -> str:
    """
    Creates a new Excel workbook.

    Parameters:
        filepath (str): Path where the new workbook will be saved
        sheet_name (str, optional): Name for the initial worksheet. If not provided, default name will be used.

    Returns:
        str: Success message with the created workbook path
    """
    full_path = get_excel_path(filepath)
    result = create_workbook_impl(full_path, sheet_name)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def create_worksheet(filepath: str, sheet_name: str) -> str:
    """
    Creates a new worksheet in an existing workbook.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name for the new worksheet

    Returns:
        str: Success message confirming sheet creation
    """
    full_path = get_excel_path(filepath)
    result = create_worksheet_impl(full_path, sheet_name)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def create_chart(
        input_filepath: str,
        output_filepath: Optional[str],
        data_sheet_name: Optional[str],
        chart_sheet_name: str,
        data_range: str,
        chart_type: str,
        target_cell: str,
        chart_options: Optional[Dict[str, Any]] = None
) -> str:
    """
    Creates a chart in a worksheet.

    Parameters:
        input_filepath (str): Path to the input Excel file.
        output_filepath (str,optional): Path to the output Excel file,default use input_filepath.
        data_sheet_name (str,optional): Name of the worksheet containing the data,default use chart_sheet_name.
        chart_sheet_name (str): Name of the worksheet where the chart will be inserted.
        data_range (str): Range of cells containing data for the chart (e.g., "A1:B10"). This is used for charts with a single data source and is ignored if `series` is provided in `chart_options`.
        chart_type (str): Type of chart to create ("column", "line", "pie", "bar", "scatter", 
        "doughnut", "area", "waterfall", "column_stacked", "column_100_percent_stacked", 
        "bubble", "funnel", "treemap", "sunburst", "histogram", "box_and_whisker").
        target_cell (str): Cell where the top-left corner of the chart will be positioned (e.g., "D5").
        chart_options (dict, optional): Dictionary with chart options. Supported keys include:
            - title (str): Chart title.
            - axis_titles (dict): A dictionary for axis titles, e.g., `{"x_axis_title": "X-Axis", "y_axis_title": "Y-Axis"}`.
            - series (list[dict]): A list of data series for multi-series charts. Each dict can have `name`, `values`, and `category_labels`.
            - category_labels (str): A common range for category labels for all series.
            - bubbles (str): range of the bubble chart.
            - style (dict): Chart style settings:
                - legend_position: "right", "left", "top", "bottom".
                - has_legend: bool.
                - has_data_labels: bool.
                - width: Chart width in pixels.
                - height: Chart height in pixels.

    Returns:
        str: Success message, e.g., "Chart created successfully".
    """
    input_full_path = get_excel_path(input_filepath)
    output_full_path = get_excel_path(output_filepath) if output_filepath else input_full_path
    if data_sheet_name is None:
        data_sheet_name = chart_sheet_name
        
    result = create_chart_impl(
        input_filepath=input_full_path,
        output_filepath=output_full_path,
        data_sheet_name=data_sheet_name,
        chart_sheet_name=chart_sheet_name,
        data_range=data_range,
        chart_type=chart_type,
        target_cell=target_cell,
        chart_options=chart_options
    )
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
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
) -> str:
    """
    Creates a pivot table in a worksheet.

    Parameters:
        input_filepath (str): Path to the input Excel file.
        output_filepath (str): Path to the output Excel file.
        data_sheet_name (str): Name of the worksheet containing the source data.
        pivot_sheet_name (str): Name of the worksheet where the pivot table will be inserted.        pivot_name (str): Name for the pivot table
        data_range (str): Range containing source data (e.g., "A1:D10")
        locate_range (str): Range where the pivot table will be placed
        rows (list): List of field names to use as row labels
        values (dict): Dictionary mapping field names to aggregation functions
            Key: Field name
            Value: Aggregation function ("sum", "count", "average", "max", "min",)
        columns (list, optional): List of field names to use as column labels
        agg_func (str, optional): Default aggregation function ("sum", "count", "average","min", "max".)

    Returns:
        str: Success message confirming pivot table creation
    """
    input_full_path = get_excel_path(input_filepath)
    output_full_path = get_excel_path(output_filepath)
    result = create_pivot_table_impl(
        input_filepath=input_full_path,
        output_filepath=output_full_path,
        data_sheet_name=data_sheet_name,
        pivot_sheet_name=pivot_sheet_name,
        pivot_name=pivot_name,
        data_range=data_range,
        locate_range=locate_range,
        rows=rows,
        values=values,
        columns=columns or [],
        agg_func=agg_func
    )
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def copy_worksheet(
        filepath: str,
        source_sheet: str,
        target_sheet: str,
        target_filepath: str = None
) -> str:
    """
    Copies a worksheet within the same workbook or to another workbook.

    Parameters:
        filepath (str): Path to the source Excel file
        source_sheet (str): Name of the worksheet to copy
        target_sheet (str): Name for the new worksheet copy
        target_filepath (str, optional): Path to the target Excel file if copying to another workbook

    Returns:
        str: Success message confirming sheet was copied
    """
    full_path = get_excel_path(filepath)
    target_path = get_excel_path(target_filepath) if target_filepath else full_path
    result = copy_sheet(full_path, source_sheet, target_sheet, target_path)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def delete_worksheet(
        filepath: str,
        sheet_name: str
) -> str:
    """
    Deletes a worksheet from an Excel workbook.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet to delete

    Returns:
        str: Success message confirming worksheet deletion
    """
    full_path = get_excel_path(filepath)
    result = delete_sheet(full_path, sheet_name)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def rename_worksheet(
        filepath: str,
        old_name: str,
        new_name: str
) -> str:
    """
    Renames a worksheet in an Excel workbook.

    Parameters:
        filepath (str): Path to the Excel file
        old_name (str): Current name of the worksheet
        new_name (str): New name to assign to the worksheet

    Returns:
        str: Success message confirming the rename operation
    """
    full_path = get_excel_path(filepath)
    result = rename_sheet(full_path, old_name, new_name)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def get_workbook_metadata(
        filepath: str,
        include_ranges: bool = False
) -> str:
    """
    Gets metadata about an Excel workbook including sheets, ranges, and file information.

    Parameters:
        filepath (str): Path to the Excel file
        include_ranges (bool, optional): Whether to include data about used ranges for each sheet

    Returns:
        dict: Dictionary containing workbook metadata:
            - filename: Name of the Excel file
            - sheets: List of worksheet names
            - size: File size in bytes
            - modified: Last modification timestamp
            - used_ranges: Dictionary mapping sheet names to their used data ranges (if include_ranges=True)
    """
    full_path = get_excel_path(filepath)
    result = get_workbook_info(full_path, include_ranges=include_ranges)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def merge_cells(filepath: str,
                sheet_name: str,
                cell_range_list: List[str]) -> str:
    """
    Merges multiple cell ranges in a worksheet.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet
        cell_range_list (list): List of cell ranges to merge (e.g., ["A1:C1", "A2:B2"])

    Returns:
        str: Success message confirming cells were merged
    """
    full_path = get_excel_path(filepath)
    result = merge_range(full_path, sheet_name, cell_range_list)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def unmerge_cells(filepath: str, sheet_name: str, cell_range: str) -> str:
    """
    Unmerges a range of previously merged cells.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet
        cell_range (str): Range of cells to unmerge (e.g., "A1:C1")

    Returns:
        str: Success message confirming cells were unmerged
    """
    full_path = get_excel_path(filepath)
    result = unmerge_range(full_path, sheet_name, cell_range)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def copy_range(
        filepath: str,
        sheet_name: str,
        source_range: str,
        target_range: str,
        target_sheet: str = None,
        target_filepath: str = None
) -> str:
    """
    Copies a range of cells to another location within the same workbook or to another workbook.

    Parameters:
        filepath (str): Path to the source Excel file
        sheet_name (str): Name of the source worksheet
        source_range (str): Range of cells to copy (e.g., "A1:C5")
        target_range (str): Target range where cells will be copied
        target_sheet (str, optional): Name of the target worksheet if different from source
        target_filepath (str, optional): Path to the target Excel file if copying to another workbook

    Returns:
        str: Success message confirming range was copied
    """
    full_path = get_excel_path(filepath)
    target_path = get_excel_path(target_filepath) if target_filepath else full_path
    result = copy_range_operation(
        full_path,
        sheet_name,
        source_range,
        target_range,
        target_sheet,
        target_path
    )
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def delete_range(
        filepath: str,
        sheet_name: str,
        cell_range: str,
        shift_direction: str = "up"
) -> str:
    """
    Deletes a range of cells and shifts remaining cells.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet
        cell_range (str): Range of cells to delete (e.g., "A1:C5")
        shift_direction (str, optional): Direction to shift remaining cells ("up" or "left", default "up")

    Returns:
        str: Success message describing the deletion and shift operation
    """
    full_path = get_excel_path(filepath)
    result = delete_range_operation(
        full_path,
        sheet_name,
        cell_range,
        shift_direction
    )
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def apply_autofilter(
        filepath: str,
        sheet_name: str,
        cell_range: str,
        filter_criteria: Dict[int, Dict[str, Any]] = None
) -> str:
    """
    Applies autofilter to a range of cells and optionally sets filter criteria.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet
        cell_range (str): Range to apply autofilter (e.g., "A1:D10")
        filter_criteria (dict, optional): Dictionary of filter criteria
            Key: Column index (0-based)
            Value: Dictionary with filter settings:
                "type": "value", "top10", "custom"
                "values": List of values for "value" type
                "operator": "<", ">", "=", ">=", "<=", "<>" for "custom" type
                "criteria": Criteria value for "custom" type
                "operator2": Second operator for "custom" type when using AND/OR conditions
                "criteria2": Second criteria value for "custom" type when using AND/OR conditions
                "is_and": Boolean indicating whether to use AND (True) or OR (False) for dual conditions
                "percent": True/False for "top10" type
                "count": Count for "top10" type
                "bottom": True/False for "top10" type

    Returns:
        str: Success or error message
    """
    full_path = get_excel_path(filepath)
    result = apply_autofilter_impl(
        full_path,
        sheet_name,
        cell_range,
        filter_criteria
    )
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def validate_excel_range(
        filepath: str,
        sheet_name: str,
        cell_range: str
) -> str:
    """
    Validates if a cell range exists and is properly formatted.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet
        cell_range (str): Range to validate (e.g., "A1:D10")

    Returns:
        str: Validation result including details about the actual data range in the sheet
    """
    full_path = get_excel_path(filepath)
    result = validate_range_impl(full_path, sheet_name, cell_range)
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def export_to_json(
        filepath: str,
        sheet_name: str,
        cell_range: str,
        output_filepath: str,
        include_headers: bool = True,
        options: Dict[str, Any] = None
) -> str:
    """
    Exports Excel worksheet data to a JSON file.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet
        cell_range (str): Cell range to export (e.g., "A1:D10")
        output_filepath (str): Path to the output JSON file
        include_headers (bool, optional): Whether to use the first row as headers (default True)
        options (dict, optional): Additional options including:
            - pretty_print: Whether to format JSON with indentation (default True)
            - date_format: Format for date values (default ISO format)
            - encoding: File encoding (default "utf-8")
            - array_format: Use array format when include_headers is False (default False)

    Returns:
        str: Success message with the path to the created JSON file
    """
    full_path = get_excel_path(filepath)
    output_path = get_excel_path(output_filepath)
    
    result = export_json_impl(
        full_path,
        sheet_name,
        cell_range,
        output_path,
        include_headers,
        options
    )
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def import_from_json(
        json_filepath: str,
        excel_filepath: str,
        sheet_name: str,
        start_cell: str = "A1",
        create_sheet: bool = False,
        options: Dict[str, Any] = None
) -> str:
    """
    Imports data from a JSON file to an Excel worksheet.

    Parameters:
        json_filepath (str): Path to the JSON file
        excel_filepath (str): Path to the Excel file
        sheet_name (str): Name of the target worksheet
        create_sheet (bool, optional): Whether to create the sheet if it doesn't exist
        start_cell (str, optional): Cell to start importing data (default "A1")
        options (dict, optional): Additional options:
            - encoding: File encoding (default "utf-8")
            - include_headers: Add header row for object arrays (default True)
            - date_format: Date format string for date values

    Returns:
        str: Success message with the path to the updated Excel file
    """
    json_path = get_excel_path(json_filepath)
    excel_path = get_excel_path(excel_filepath)
    
    result = import_json_impl(
        json_path,
        excel_path,
        sheet_name,
        start_cell,
        create_sheet,
        options
    )
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def convert_excel(
        filepath: str,
        output_filepath: str,
        format_type: str,  
        options: Dict[str, Any] = None,
        sheet_name: str = None,
        cell_range: str = None
) -> str:
    """
    Converts Excel file to different formats.

    Supported formats:
    - pdf: Convert to PDF document
    - csv: Convert to CSV text file (requires sheet_name)
    - txt: Convert to text file (requires sheet_name)
    - html: Convert to HTML document
    - image: Convert to image file (png, jpg)
    - xlsx/xls: Convert between Excel formats
    - xml: Convert to XML format

    Parameters:
        filepath (str): Path to the Excel file
        format_type (str): Target format type (pdf, csv, txt, html, image, xlsx, xls, xml)
        output_filepath (str): Path for the output file
        sheet_name (str, optional): Name of the worksheet (required for some formats)
        cell_range (str, optional): Range to convert (if not entire sheet)
        options (dict, optional): Format-specific options:
            For PDF:
              - orientation: "portrait" or "landscape"
              - paper_size: "a4", "letter", etc.
              - fit_to_page: true/false
            For CSV/TXT:
              - delimiter: Character to use as delimiter (default ",")
              - encoding: File encoding (default "utf-8")
            For HTML:
              - image_embedded: true/false
              - image_locationType: Controls image position mode
            For Image:
              - image_type: "png", "jpg", "original"

    Returns:
        str: Success message or error description
    """
    full_path = get_excel_path(filepath)
    output_path = get_excel_path(output_filepath)
    
    result = convert_workbook_impl(
        filepath=full_path,
        output_filepath=output_path,
        format_type=format_type,
        options=options,
        sheet_name=sheet_name,
        cell_range=cell_range
    )
    
    return json.dumps(result)

@mcp.tool()
@tool_exception_handler
def get_shape_image_base64(
        filepath: str,
        sheet_name: str,
        shape_name: str = None,
        shape_index: int = None
) -> str:
    """
    Gets the image of a Shape in Excel as a base64 string. temp support PrstGeom Shapes and Pictures

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the worksheet containing the shape
        shape_name (str, optional): Name of the shape to export
        shape_index (int, optional): Index of the shape in the worksheet (0-based)

    Returns:
        str: Base64 string representation of the shape image
        
    Note: Either shape_name or shape_index must be provided. If the worksheet has no 
    shapes or the specified shape doesn't exist, an error will be returned.
    """
    full_path = get_excel_path(filepath)
    result = get_shape_img_b64(
        full_path,
        sheet_name,
        shape_name,
        shape_index
    )
    return json.dumps(result)

async def run_server():
    """Run the Spire.Xls MCP Server."""
    try:
        logger.info(f"Starting Spire.Xls MCP Server (files directory: {EXCEL_FILES_PATH})")
        await mcp.run_sse_async()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        await mcp.shutdown()
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")