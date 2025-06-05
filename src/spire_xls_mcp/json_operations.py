import json
import logging
from typing import Dict, Any

from spire.xls import *
from .exceptions import DataError, ValidationError

logger = logging.getLogger(__name__)

def export_to_json(
    filepath: str,
    sheet_name: str,
    cell_range: str,
    output_filepath: str,
    include_headers: bool = True,
    options: Dict[str, Any] = None
) -> Dict[str, str]:
    """
    Export Excel worksheet data to a JSON file.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        cell_range: Cell range to export
        output_filepath: Path to output JSON file
        include_headers: Whether to use the first row as headers
        options: Additional options such as:
            - pretty_print: Whether to pretty print JSON (default True)
            - date_format: Date format (default ISO format)
            - encoding: File encoding (default utf-8)
            - array_format: Use array format when include_headers is False (default False)
    
    Returns:
        Dictionary with result message
    """
    try:

        # Set default options
        options = options or {}
        pretty_print = options.get("pretty_print", True)
        encoding = options.get("encoding", "utf-8")
        array_format = options.get("array_format", False)
        
        # Load workbook
        workbook = Workbook()
        workbook.LoadFromFile(filepath)
        
        # Get worksheet
        if sheet_name not in [sheet.Name for sheet in workbook.Worksheets]:
            raise ValidationError(f"Worksheet '{sheet_name}' does not exist")
        
        sheet = workbook.Worksheets[sheet_name]
        
        # Get range
        try:
            range_data = sheet.Range[cell_range]
        except Exception as e:
            raise ValidationError(f"Invalid cell range '{cell_range}': {str(e)}")
        
        # Get data
        rows = range_data.LastRow - range_data.Row + 1
        columns = range_data.LastColumn - range_data.Column + 1
        data = []
        
        if include_headers:
            # Use first row as headers
            headers = []
            for j in range(1, columns + 1):
                cell_value = range_data[1, j].Value
                headers.append(str(cell_value) if cell_value is not None else f"Column{j}")
            
            # Read data starting from second row (skip header row)
            for i in range(2, rows + 1):
                row_data = {}
                for j in range(1, columns + 1):
                    cell_value = range_data[i, j].Value
                    row_data[headers[j - 1]] = cell_value
                data.append(row_data)
        else:
            # All rows as data
            if array_format:
                # Simple 2D array
                for i in range(1, rows + 1):
                    row_data = []
                    for j in range(1, columns + 1):
                        cell_value = range_data[i, j].Value
                        row_data.append(cell_value)
                    data.append(row_data)
            else:
                # Use position index as keys
                for i in range(1, rows + 1):
                    row_data = {}
                    for j in range(1, columns + 1):
                        cell_value = range_data[i, j].Value
                        row_data[f"Column{j}"] = cell_value
                    data.append(row_data)
        
        # Write to JSON file
        with open(output_filepath, 'w', encoding=encoding) as f:
            indent = 4 if pretty_print else None
            json.dump(data, f, indent=indent, ensure_ascii=False, default=str)
        
        workbook.Dispose()
        
        return {"message": f"Excel data successfully exported to JSON file: {output_filepath}"}
    
    except Exception as e:
        logger.error(f"Failed to export JSON: {str(e)}")
        raise DataError(f"Failed to export JSON: {str(e)}")


def import_from_json(
    json_filepath: str,
    excel_filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    create_sheet: bool = False,
    options: Dict[str, Any] = None
) -> Dict[str, str]:
    """
    Import data from a JSON file to an Excel worksheet.
    
    Args:
        json_filepath: Path to JSON file
        excel_filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Starting cell
        create_sheet: Whether to create the sheet if it doesn't exist
        options: Additional options such as:
            - encoding: File encoding (default utf-8)
            - include_headers: Add header row for object arrays (default True)
            - date_format: Date format (default ISO format)
    
    Returns:
        Dictionary with result message
    """
    try:

        # Set default options
        options = options or {}
        encoding = options.get("encoding", "utf-8")
        include_headers = options.get("include_headers", True)
        
        # Read JSON file
        with open(json_filepath, 'r', encoding=encoding) as f:
            json_data = json.load(f)
        
        if not json_data:
            raise DataError("No data found in JSON file")
        
        # Load workbook
        workbook = Workbook()
        if os.path.exists(excel_filepath):
            workbook.LoadFromFile(excel_filepath)
        
        # Get or create worksheet
        if sheet_name not in [sheet.Name for sheet in workbook.Worksheets]:
            if create_sheet:
                sheet = workbook.Worksheets.Add(sheet_name)
            else:
                raise ValidationError(f"Worksheet '{sheet_name}' does not exist, and create flag not set")
        else:
            sheet = workbook.Worksheets[sheet_name]
        
        # Parse starting cell
        try:
            start_range = sheet.Range[start_cell]
            start_row = start_range.Row
            start_col = start_range.Column
        except Exception as e:
            raise ValidationError(f"Invalid start cell '{start_cell}': {str(e)}")
        
        # Write data to Excel
        current_row = start_row
        
        # Check JSON data type
        if isinstance(json_data, list):
            if json_data and isinstance(json_data[0], dict):
                # Array of objects
                if include_headers:
                    # Write header row
                    headers = list(json_data[0].keys())
                    for col_idx, header in enumerate(headers):
                        sheet.Range[current_row, start_col + col_idx].Text = header
                    current_row += 1
                
                # Write data rows
                for row_data in json_data:
                    if isinstance(row_data, dict):
                        if include_headers:
                            # Use headers as keys
                            headers = list(json_data[0].keys())
                            for col_idx, header in enumerate(headers):
                                value = row_data.get(header)
                                sheet.Range[current_row, start_col + col_idx].Value = str(value)
                        else:
                            # Write values in order
                            for col_idx, (_, value) in enumerate(row_data.items()):
                                sheet.Range[current_row, start_col + col_idx].Value = str(value)
                        current_row += 1
            elif json_data and isinstance(json_data[0], list):
                # 2D array
                for row_data in json_data:
                    if isinstance(row_data, list):
                        for col_idx, value in enumerate(row_data):
                            sheet.Range[current_row, start_col + col_idx].Value = str(value)
                        current_row += 1
            else:
                # Simple array
                for row_idx, value in enumerate(json_data):
                    sheet.Range[current_row + row_idx, start_col].Value = str(value)
        else:
            # Single object
            if include_headers:
                # Write keys as headers, values as data
                for col_idx, (key, value) in enumerate(json_data.items()):
                    sheet.Range[current_row, start_col + col_idx].Text = str(key)
                    sheet.Range[current_row + 1, start_col + col_idx].Value = str(value)
            else:
                # Write values only
                for col_idx, (_, value) in enumerate(json_data.items()):
                    sheet.Range[current_row, start_col + col_idx].Value = str(value)
        
        # Save workbook
        workbook.SaveToFile(excel_filepath)
        workbook.Dispose()
        
        return {"message": f"JSON data successfully imported to Excel file: {excel_filepath}"}
    
    except Exception as e:
        logger.error(f"Failed to import JSON: {str(e)}")
        raise DataError(f"Failed to import JSON: {str(e)}") 