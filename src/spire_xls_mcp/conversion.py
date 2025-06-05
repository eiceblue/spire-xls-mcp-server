import logging
from typing import Any, Dict, Optional

from spire.xls import *

from .exceptions import ConversionError

logger = logging.getLogger(__name__)


def convert_workbook(
        filepath: str,
        output_filepath: str,
        format_type: str,
        options: Dict[str, Any] = None,
        sheet_name: str = None,
        cell_range: str = None
) -> dict[str, Any]:
    """
    Convert Excel workbook to different formats.
    
    Args:
        filepath: Source Excel file path
        output_filepath: Target output file path
        format_type: Target format (pdf, csv, html, image, txt, xml, xlsx, xls, etc.)
        options: Format-specific options
        sheet_name: Specific sheet to convert (if applicable)
        cell_range: Specific range to convert (if applicable)
        
    Returns:
        Dictionary with operation status
    """
    try:
        # Load the workbook
        wb = Workbook()
        wb.LoadFromFile(filepath)

        # Ensure output directory exists
        output_dir = os.path.dirname(output_filepath)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        # Handle format-specific conversion
        format_type = format_type.lower()

        # Find specific sheet if needed
        target_sheet = None
        if sheet_name:
            for ws in wb.Worksheets:
                if ws.Name == sheet_name:
                    target_sheet = ws
                    break

            if target_sheet is None and (format_type == 'csv' or format_type == 'txt' or cell_range):
                raise ConversionError(f"Sheet '{sheet_name}' not found")

        # Process different format types
        if format_type == 'pdf':
            # Configure PDF options
            if options:
                # Apply page setup options
                for sheet in wb.Worksheets:
                    if sheet_name and sheet.Name != sheet_name:
                        continue

                    if 'orientation' in options:
                        orientation = options['orientation'].lower()
                        if orientation == 'landscape':
                            sheet.PageSetup.Orientation = PageOrientationType.Landscape
                        elif orientation == 'portrait':
                            sheet.PageSetup.Orientation = PageOrientationType.Portrait

                    if 'paper_size' in options:
                        paper_size = options['paper_size'].lower()
                        if paper_size == 'a4':
                            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4
                        elif paper_size == 'letter':
                            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter

                    if 'fit_to_page' in options and options['fit_to_page']:
                        sheet.PageSetup.FitToPagesWide = 1
                        sheet.PageSetup.FitToPagesTall = 1

            # Convert to PDF
            if sheet_name and target_sheet:
                target_sheet.SaveToPdf(output_filepath)
            else:
                wb.SaveToFile(output_filepath, FileFormat.PDF)

        elif format_type == 'csv' or format_type == 'txt':
            if target_sheet is None:
                raise ConversionError("Sheet name is required for CSV/TXT conversion")

            if not options or not options['delimiter'] or not options['encoding']:
                raise ConversionError("options delimiter、encoding is required for CSV/TXT conversion")

            separator = options['delimiter']
            encoding = Encoding.GetEncoding(str(options['encoding']))
            # Convert to CSV/TXT
            target_sheet.SaveToFile(output_filepath, separator, encoding)

        elif format_type == 'html':
            # Configure HTML options
            html_options = HTMLOptions()

            if options:
                if 'image_embedded' in options and not options['image_embedded']:
                    html_options.ImageEmbedded = False

                if 'image_locationType' in options and options['image_locationType'] == 0:
                    html_options.ImageLocationType = ImageLocationTypes.GlobalAbsolute
                else:
                    html_options.ImageLocationType = ImageLocationTypes.TableRelative

            # Convert to HTML
            if sheet_name and target_sheet:
                target_sheet.SaveToHtml(output_filepath, html_options)
            else:
                wb.SaveToFile(output_filepath, FileFormat.HTML)

        elif format_type == 'image':
            # Determine image format
            # image_format = ImageFormatType.Png  # Default
            #
            # if options and 'image_type' in options:
            #     img_type = options['image_type'].lower()
            #     if img_type == 'jpg' or img_type == 'jpeg':
            #         image_format = ImageFormatType.Jpeg
            #     elif img_type == 'original':
            #         image_format = ImageFormatType.Original

            # Convert to image
            if target_sheet is None:
                if sheet_name:
                    raise ConversionError(f"Sheet '{sheet_name}' not found")
                target_sheet = wb.Worksheets[0]

            if cell_range:
                save_range = target_sheet.Range[cell_range]
                row, column, end_row, end_column = (
                    save_range.Row, save_range.Column, save_range.LastRow, save_range.LastColumn)
                stream = target_sheet.ToImage(row, column, end_row, end_column)
                stream.Save(output_filepath)
            else:
                lastColumn = 0
                lastRow = 0
                # TODO replace with target_sheet.SaveToImage(filePath)
                for pic in target_sheet.Pictures:
                    lastColumn = max(pic.RightColumn, lastColumn)
                    lastRow = max(pic.BottomRow, lastRow)
                for pic in target_sheet.PrstGeomShapes:
                    lastColumn = max(pic.RightColumn, lastColumn)
                    lastRow = max(pic.BottomRow, lastRow)
                stream = target_sheet.ToImage(target_sheet.FirstRow, target_sheet.FirstColumn,
                                              lastRow, lastColumn)
                stream.Save(output_filepath)

        elif format_type in ['xlsx', 'xls', 'ods', 'xml', 'uos']:
            # Map format type to file format
            format_map = {
                'xlsx': FileFormat.Version2013,
                'xls': FileFormat.Version97to2003,
                'ods': FileFormat.ODS,
                'xml': FileFormat.XML,
                'uos': FileFormat.UOS
            }

            if format_type not in format_map:
                raise ConversionError(f"Unsupported Excel format: {format_type}")

            # Convert to the specified format
            wb.SaveToFile(output_filepath, format_map[format_type])

        else:
            raise ConversionError(f"Unsupported format type: {format_type}")

        return {
            "message": f"Excel file successfully converted to {format_type.upper()}: {output_filepath}",
            "source_file": filepath,
            "output_file": output_filepath,
            "format": format_type
        }

    except ConversionError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to convert Excel file: {e}")
        raise ConversionError(f"Failed to convert Excel file: {str(e)}")
