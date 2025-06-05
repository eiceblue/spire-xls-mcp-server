import logging
from typing import Any, Dict, Optional
from spire.xls import *

from .exceptions import ChartError
from .workbook import get_or_create_workbook
from .cell_utils import  EnumMapper

logger = logging.getLogger(__name__)


def create_chart_in_sheet(
        filepath: str,
        sheet_name: str,
        data_range: str,
        chart_type: str,
        target_cell: str,
        title: str = "",
        x_axis: str = "",
        y_axis: str = "",
        style: Optional[Dict] = None
) -> dict[str, Any]:
    """Create chart in sheet with enhanced styling options"""
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

        # Parse ranges
        target_range = sheet.Range[target_cell]
        target_row, target_col = target_range.Row, target_range.Column

        # Create chart
        chart = sheet.Charts.Add()

        # Set chart type
        chart.ChartType = EnumMapper.get_chart_type_enum(chart_type)

        # Set data range
        chart.DataRange = sheet.Range[data_range]

        # Set chart position
        chart.LeftColumn = target_col
        chart.TopRow = target_row

        # Set chart title
        if title:
            chart.ChartTitle = title

        # Set axis labels
        if x_axis:
            chart.PrimaryCategoryAxis.Title = x_axis
        if y_axis:
            chart.PrimaryValueAxis.Title = y_axis

        width = 480
        height = 300
        # Apply style if provided
        if style:
            if 'legend_position' in style:
                positions = {
                    'right': LegendPositionType.Right,
                    'left': LegendPositionType.Left,
                    'top': LegendPositionType.Top,
                    'bottom': LegendPositionType.Bottom
                }
                if style['legend_position'] in positions:
                    chart.Legend.Position = positions[style['legend_position']]

            if 'has_legend' in style:
                chart.Legend.Visible = style['has_legend']

            if 'has_data_labels' in style:
                chart.Series[0].DataLabels.HasValue = style['has_data_labels']

            # Set default size if not specified in style
            width = style.get('width', width)  # Default width: 480 pixels
            height = style.get('height', height)  # Default height: 300 pixels

        chart.Width = width
        chart.Height = height

        # Save workbook
        wb.SaveToFile(filepath)
        return {"message": "Chart created successfully"}

    except Exception as e:
        logger.error(f"Failed to create chart: {e}")
        raise ChartError(f"Failed to create chart: {e!s}")
