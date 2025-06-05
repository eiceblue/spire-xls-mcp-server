# Spire.Xls MCP Server Tools

This document provides detailed information about all available tools in the Spire.Xls MCP Server.

## Workbook Operations

### create_workbook

Creates a new Excel workbook.

```python
create_workbook(filepath: str, sheet_name: str = None) -> str
```

- `filepath`: Path where the new workbook will be saved
- `sheet_name`: Optional name for the initial worksheet
- Returns: Success message with the created workbook path

### create_worksheet

Creates a new worksheet in an existing workbook.

```python
create_worksheet(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name for the new worksheet
- Returns: Success message confirming sheet creation

### get_workbook_metadata

Gets metadata about an Excel workbook including sheets, ranges, and file information.

```python
get_workbook_metadata(filepath: str, include_ranges: bool = False) -> dict
```

- `filepath`: Path to Excel file
- `include_ranges`: Whether to include data about used ranges for each sheet
- Returns: Dictionary containing workbook metadata:
  - filename: Name of the Excel file
  - sheets: List of worksheet names
  - size: File size in bytes
  - modified: Last modification timestamp
  - used_ranges: Dictionary mapping sheet names to their used data ranges (if include_ranges=True)

## Data Operations

### write_data_to_excel

Writes data to an Excel worksheet.

```python
write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: str = "A1"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet to write to
- `data`: List of lists containing data to write (rows of data)
- `start_cell`: Cell to start writing from (default: "A1")
- Returns: Success message confirming data was written

### read_data_from_excel

Reads data from an Excel worksheet.

```python
read_data_from_excel(
    filepath: str,
    sheet_name: str,
    cell_range: str,
    preview_only: bool = False
) -> dict
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet to read from
- `cell_range`: Range of cells to read (e.g., "A1:D10")
- `preview_only`: If True, returns only preview data without full styling info
- Returns: Column-first nested dictionary with cell data

## Formatting Operations

### format_range

Applies formatting to a range of cells.

```python
format_range(
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
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet
- `cell_range`: Range of cells to format (e.g., "A1:C5")
- Various formatting options (see parameters)
- Returns: Success message confirming formatting was applied

### merge_cells

Merges multiple cell ranges in a worksheet.

```python
merge_cells(filepath: str, sheet_name: str, cell_range_list: List[str]) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet
- `cell_range_list`: List of cell ranges to merge (e.g., ["A1:C1", "A2:B2"])
- Returns: Success message confirming cells were merged

### unmerge_cells

Unmerges a range of previously merged cells.

```python
unmerge_cells(filepath: str, sheet_name: str, cell_range: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet
- `cell_range`: Range of cells to unmerge (e.g., "A1:C1")
- Returns: Success message confirming cells were unmerged

## Formula Operations

### apply_formula

Applies an Excel formula to a specified cell with verification.

```python
apply_formula(filepath: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet
- `cell`: Cell reference where formula will be applied (e.g., "A1")
- `formula`: Excel formula to apply (must include "=" prefix)
- Returns: Success message confirming formula application

## Chart Operations

### create_chart

Creates a chart in a worksheet.

```python
create_chart(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
    style: Optional[Dict[str, Any]] = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet
- `data_range`: Range of cells containing data for the chart (e.g., "A1:B10")
- `chart_type`: Type of chart to create (e.g., "column", "line", "pie", "bar", "scatter")
- `target_cell`: Cell where the top-left corner of the chart will be positioned
- `title`: Chart title (optional)
- `x_axis`: X-axis title (optional)
- `y_axis`: Y-axis title (optional)
- `style`: Dictionary with chart style settings (optional)
- Returns: Success message confirming chart creation

## Pivot Table Operations

### create_pivot_table

Creates a pivot table in a worksheet.

```python
create_pivot_table(
    filepath: str,
    sheet_name: str,
    pivot_name: str,
    data_range: str,
    locate_range: str,
    rows: List[str],
    values: dict[str, str],
    columns: List[str] = None,
    agg_func: str = "sum"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the source worksheet
- `pivot_name`: Name for the pivot table
- `data_range`: Range containing source data (e.g., "A1:D10")
- `locate_range`: Range where the pivot table will be placed
- `rows`: List of field names to use as row labels
- `values`: Dictionary mapping field names to aggregation functions
- `columns`: List of field names to use as column labels (optional)
- `agg_func`: Default aggregation function ("sum", "count", etc.) (optional)
- Returns: Success message confirming pivot table creation

## Worksheet Operations

### copy_worksheet

Copies a worksheet within the same workbook.

```python
copy_worksheet(filepath: str, source_sheet: str, target_sheet: str) -> str
```

- `filepath`: Path to Excel file
- `source_sheet`: Name of the worksheet to copy
- `target_sheet`: Name for the new worksheet copy
- Returns: Success message confirming sheet was copied

### delete_worksheet

Deletes a worksheet from an Excel workbook.

```python
delete_worksheet(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet to delete
- Returns: Success message confirming worksheet deletion

### rename_worksheet

Renames a worksheet in an Excel workbook.

```python
rename_worksheet(filepath: str, old_name: str, new_name: str) -> str
```

- `filepath`: Path to Excel file
- `old_name`: Current name of the worksheet
- `new_name`: New name to assign to the worksheet
- Returns: Success message confirming the rename operation

## Range Operations

### copy_range

Copies a range of cells to another location.

```python
copy_range(
    filepath: str,
    sheet_name: str,
    source_range: str,
    target_range: str,
    target_sheet: str = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the source worksheet
- `source_range`: Range of cells to copy (e.g., "A1:C5")
- `target_range`: Target range where cells will be copied
- `target_sheet`: Name of the target worksheet if different from source (optional)
- Returns: Success message confirming range was copied

### delete_range

Deletes a range of cells and shifts remaining cells.

```python
delete_range(
    filepath: str,
    sheet_name: str,
    cell_range: str,
    shift_direction: str = "up"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet
- `cell_range`: Range of cells to delete (e.g., "A1:C5")
- `shift_direction`: Direction to shift remaining cells ("up" or "left", default "up")
- Returns: Success message describing the deletion and shift operation

### validate_excel_range

Validates if a cell range exists and is properly formatted.

```python
validate_excel_range(
    filepath: str,
    sheet_name: str,
    cell_range: str
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet
- `cell_range`: Range to validate (e.g., "A1:D10")
- Returns: Validation result including details about the actual data range in the sheet

## Export and Import Operations

### export_to_json

Exports Excel worksheet data to a JSON file.

```python
export_to_json(
    filepath: str,
    sheet_name: str,
    cell_range: str,
    output_filepath: str,
    include_headers: bool = True,
    options: Dict[str, Any] = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet
- `cell_range`: Cell range to export (e.g., "A1:D10")
- `output_filepath`: Path to the output JSON file
- `include_headers`: Whether to use the first row as headers (default True)
- `options`: Additional options including:
  - pretty_print: Whether to format JSON with indentation (default True)
  - date_format: Format for date values (default ISO format)
  - encoding: File encoding (default "utf-8")
  - array_format: Use array format when include_headers is False (default False)
- Returns: Success message with the path to the created JSON file

### import_from_json

Imports data from a JSON file to an Excel worksheet.

```python
import_from_json(
    json_filepath: str,
    excel_filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    create_sheet: bool = False,
    options: Dict[str, Any] = None
) -> str
```

- `json_filepath`: Path to JSON file
- `excel_filepath`: Path to Excel file
- `sheet_name`: Name of the target worksheet
- `start_cell`: Cell to start importing data (default "A1")
- `create_sheet`: Whether to create the sheet if it doesn't exist
- `options`: Additional options:
  - encoding: File encoding (default "utf-8")
  - include_headers: Add header row for object arrays (default True)
  - date_format: Date format string for date values
- Returns: Success message with the path to the updated Excel file

## Conversion Operations

### convert_excel

Converts Excel file to different formats.

```python
convert_excel(
    filepath: str,
    output_filepath: str,
    format_type: str,  
    options: Dict[str, Any] = None,
    sheet_name: str = None,
    cell_range: str = None
) -> str
```

- `filepath`: Path to Excel file
- `format_type`: Target format type (pdf, csv, txt, html, image, xlsx, xls, xml)
- `output_filepath`: Path for the output file
- `sheet_name`: Name of the worksheet (required for some formats)
- `cell_range`: Range to convert (if not entire sheet)
- `options`: Format-specific options:
  - For PDF:
    - orientation: "portrait" or "landscape"
    - paper_size: "a4", "letter", etc.
    - fit_to_page: true/false
  - For CSV/TXT:
    - delimiter: Character to use as delimiter (default ",")
    - encoding: File encoding (default "utf-8")
  - For HTML:
    - image_embedded: true/false
    - image_locationType: Controls image position mode
  - For Image:
    - image_type: "png", "jpg", "original"
- Returns: Success message or error description

## Shape Operations

### get_shape_image_base64

Gets the image of a Shape in Excel as a base64 string.

```python
get_shape_image_base64(
    filepath: str,
    sheet_name: str,
    shape_name: str = None,
    shape_index: int = None,
    image_type: str = "png"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet containing the shape
- `shape_name`: Name of the shape to export
- `shape_index`: Index of the shape in the worksheet (0-based)
- `image_type`: Image format, either 'png' or 'jpg' (default 'png')
- Returns: Base64 string representation of the shape image

Note: Either shape_name or shape_index must be provided. If the worksheet has no shapes or the specified shape doesn't exist, an error will be returned.

## Filter Operations

### apply_autofilter

Applies autofilter to a range of cells and optionally sets filter criteria.

```python
apply_autofilter(
    filepath: str,
    sheet_name: str,
    cell_range: str,
    filter_criteria: Dict[int, Dict[str, Any]] = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of the worksheet
- `cell_range`: Range to apply autofilter (e.g., "A1:D10")
- `filter_criteria`: Dictionary of filter criteria
  - Key: Column index (0-based)
  - Value: Dictionary with filter settings:
    - "type": "value", "top10", "custom", "dynamic"
    - "values": List of values for "value" type
    - "operator": "<", ">", "=", ">=", "<=", "<>" for "custom" type
    - "criteria": Criteria value for "custom" type
    - "percent": True/False for "top10" type
    - "count": Count for "top10" type
    - "bottom": True/False for "top10" type
- Returns: Success or error message
