import logging
import sys
import os
from typing import Any, List, Dict, Optional

from mcp.server.fastmcp import FastMCP

# Import exceptions
from .exceptions import (
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


from .validation import (
    validate_range_in_sheet_operation as validate_range_impl
)
from .chart import create_chart_in_sheet as create_chart_impl
from .workbook import get_workbook_info
from .data import write_data
from .pivot import create_pivot_table as create_pivot_table_impl
from .sheet import (
    copy_sheet,
    delete_sheet,
    rename_sheet,
    merge_range,
    unmerge_range,
    apply_autofilter as apply_autofilter_impl
)
from .conversion import convert_workbook as convert_workbook_impl

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

# # Create the directory if it doesn't exist
# os.makedirs(EXCEL_FILES_PATH, exist_ok=True)

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
    # If filename is already an absolute path, return it
    if os.path.isabs(filename):
        return filename

    # Use the configured Excel files path
    return os.path.join(EXCEL_FILES_PATH, filename)


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)

        from .calculations import apply_formula as apply_formula_impl
        result = apply_formula_impl(full_path, sheet_name, cell, formula)
        return result["message"]
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error applying formula: {e}")
        raise


@mcp.tool()
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
        bg_color (str, optional): Background color as hex code
        alignment (str, optional): Text alignment (e.g., "center", "left", "right")
        number_format (str, optional): Excel number format code
        border_style (str, optional): Border style for the range
        border_color (str, optional): Border color as hex code
        wrap_text (bool, optional): Whether to enable text wrapping
        merge_cells (bool, optional): Whether to merge the cells in the range
        protection (dict, optional): Cell protection settings
        conditional_format (dict, optional): Conditional formatting rules

    Returns:
        str: Success message confirming formatting was applied
    """
    try:
        full_path = get_excel_path(filepath)
        from .formatting import format_range as format_range_func

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
        return "Range formatted successfully"
    except (ValidationError, FormattingError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error formatting range: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        from .data import read_excel_range
        result = read_excel_range(full_path, sheet_name, cell_range, preview_only)
        if not result:
            return "No data found in specified range"
        return result
    except Exception as e:
        logger.error(f"Error reading data: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        result = write_data(full_path, sheet_name, data, start_cell)
        return result["message"]
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error writing data: {e}")
        raise


@mcp.tool()
def create_workbook(filepath: str, sheet_name: str = None) -> str:
    """
    Creates a new Excel workbook.

    Parameters:
        filepath (str): Path where the new workbook will be saved
        sheet_name (str, optional): Name for the initial worksheet. If not provided, default name will be used.

    Returns:
        str: Success message with the created workbook path
    """
    try:
        full_path = get_excel_path(filepath)
        from .workbook import create_workbook as create_workbook_impl
        result = create_workbook_impl(full_path, sheet_name)
        return f"Created workbook at {full_path}"
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating workbook: {e}")
        raise


@mcp.tool()
def create_worksheet(filepath: str, sheet_name: str) -> str:
    """
    Creates a new worksheet in an existing workbook.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name for the new worksheet

    Returns:
        str: Success message confirming sheet creation
    """
    try:
        full_path = get_excel_path(filepath)
        from .workbook import create_sheet as create_worksheet_impl
        result = create_worksheet_impl(full_path, sheet_name)
        return result["message"]
    except (ValidationError, WorkbookError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating worksheet: {e}")
        raise


@mcp.tool()
def create_chart(
        filepath: str,
        sheet_name: str,
        data_range: str,
        chart_type: str,
        target_cell: str,
        title: str = "",
        x_axis: str = "",
        y_axis: str = "",
        style: Optional[Dict[str, Any]] = None
) -> str:
    """
    Creates a chart in a worksheet.

    Parameters:
        filepath (str): Path to the Excel file.
        sheet_name (str): Name of the worksheet. If the sheet does not exist, it will be created automatically.
        data_range (str): Range of cells containing data for the chart (e.g., "A1:B10").
        chart_type (str): Type of chart to create (e.g., "column", "line", "pie", "bar", "scatter", etc. See EnumMapper for supported types).
        target_cell (str): Cell where the top-left corner of the chart will be positioned (e.g., "D5").
        title (str, optional): Chart title. Default is an empty string.
        x_axis (str, optional): X-axis title. Default is an empty string.
        y_axis (str, optional): Y-axis title. Default is an empty string.
        style (dict, optional): Dictionary with chart style settings. Supported keys include:
            - legend_position: Position of the legend ("right", "left", "top", "bottom").
            - has_legend: Whether to display the legend (bool).
            - has_data_labels: Whether to display data labels for the first series (bool).
            - width: Chart width in pixels (default 480).
            - height: Chart height in pixels (default 300).

    Returns:
        str: Success message, e.g., "Chart created successfully".

    Raises:
        ChartError: If chart creation fails or parameters are invalid.
    """
    try:
        full_path = get_excel_path(filepath)
        result = create_chart_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            data_range=data_range,
            chart_type=chart_type,
            target_cell=target_cell,
            title=title,
            x_axis=x_axis,
            y_axis=y_axis,
            style=style
        )
        return result["message"]
    except (ValidationError, ChartError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating chart: {e}")
        raise


@mcp.tool()
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
) -> str:
    """
    Creates a pivot table in a worksheet.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the source worksheet
        pivot_name (str): Name for the pivot table
        data_range (str): Range containing source data (e.g., "A1:D10")
        locate_range (str): Range where the pivot table will be placed
        rows (list): List of field names to use as row labels
        values (dict): Dictionary mapping field names to aggregation functions
            Key: Field name
            Value: Aggregation function ("sum", "count", "average", "max", "min", etc.)
        columns (list, optional): List of field names to use as column labels
        agg_func (str, optional): Default aggregation function ("sum", "count", etc.)

    Returns:
        str: Success message confirming pivot table creation
    """
    try:
        full_path = get_excel_path(filepath)
        result = create_pivot_table_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            pivot_name=pivot_name,
            data_range=data_range,
            locate_range=locate_range,
            rows=rows,
            values=values,
            columns=columns or [],
            agg_func=agg_func
        )
        return result["message"]
    except (ValidationError, PivotError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating pivot table: {e}")
        raise


@mcp.tool()
def copy_worksheet(
        filepath: str,
        source_sheet: str,
        target_sheet: str
) -> str:
    """
    Copies a worksheet within the same workbook.

    Parameters:
        filepath (str): Path to the Excel file
        source_sheet (str): Name of the worksheet to copy
        target_sheet (str): Name for the new worksheet copy

    Returns:
        str: Success message confirming sheet was copied
    """
    try:
        full_path = get_excel_path(filepath)
        result = copy_sheet(full_path, source_sheet, target_sheet)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error copying worksheet: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        result = delete_sheet(full_path, sheet_name)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting worksheet: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        result = rename_sheet(full_path, old_name, new_name)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error renaming worksheet: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        result = get_workbook_info(full_path, include_ranges=include_ranges)
        return str(result)
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting workbook metadata: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        result = merge_range(full_path, sheet_name, cell_range_list)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error merging cells: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        result = unmerge_range(full_path, sheet_name, cell_range)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error unmerging cells: {e}")
        raise


@mcp.tool()
def copy_range(
        filepath: str,
        sheet_name: str,
        source_range: str,
        target_range: str,
        target_sheet: str = None
) -> str:
    """
    Copies a range of cells to another location.

    Parameters:
        filepath (str): Path to the Excel file
        sheet_name (str): Name of the source worksheet
        source_range (str): Range of cells to copy (e.g., "A1:C5")
        target_range (str): Target range where cells will be copied
        target_sheet (str, optional): Name of the target worksheet if different from source

    Returns:
        str: Success message confirming range was copied
    """
    try:
        full_path = get_excel_path(filepath)
        from .sheet import copy_range_operation
        result = copy_range_operation(
            full_path,
            sheet_name,
            source_range,
            target_range,
            target_sheet
        )
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error copying range: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        from .sheet import delete_range as delete_range_operation
        result = delete_range_operation(
            full_path,
            sheet_name,
            cell_range,
            shift_direction
        )
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting range: {e}")
        raise


@mcp.tool()
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
                "type": "value", "top10", "custom", "dynamic"
                "values": List of values for "value" type
                "operator": "<", ">", "=", ">=", "<=", "<>" for "custom" type
                "criteria": Criteria value for "custom" type
                "percent": True/False for "top10" type
                "count": Count for "top10" type
                "bottom": True/False for "top10" type

    Returns:
        str: Success or error message
    """
    try:
        full_path = get_excel_path(filepath)
        result = apply_autofilter_impl(
            full_path,
            sheet_name,
            cell_range,
            filter_criteria
        )
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error applying autofilter: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        result = validate_range_impl(full_path, sheet_name, cell_range)
        return result["message"]
    except ValidationError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error validating range: {e}")
        raise


@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        output_path = get_excel_path(output_filepath)
        
        from .json_operations import export_to_json as export_json_impl
        result = export_json_impl(
            full_path,
            sheet_name,
            cell_range,
            output_path,
            include_headers,
            options
        )
        return result["message"]
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error exporting to JSON: {e}")
        raise


@mcp.tool()
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
    try:
        json_path = get_excel_path(json_filepath)
        excel_path = get_excel_path(excel_filepath)
        
        from .json_operations import import_from_json as import_json_impl
        result = import_json_impl(
            json_path,
            excel_path,
            sheet_name,
            start_cell,
            create_sheet,
            options
        )
        return result["message"]
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error importing from JSON: {e}")
        raise


@mcp.tool()
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
    try:
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
        
        return result["message"]
    except ConversionError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error converting file: {e}")
        raise ConversionError(f"Failed to convert Excel file: {str(e)}")

@mcp.tool()
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
    try:
        full_path = get_excel_path(filepath)
        from .sheet import get_shape_image_base64 as get_shape_img_b64
        return get_shape_img_b64(
            full_path,
            sheet_name,
            shape_name,
            shape_index
        )
    except Exception as e:
        logger.error(f"Error getting shape image base64: {e}")
        raise


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
