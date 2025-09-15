import logging
from typing import Any, Dict, Optional, Callable

from spire.xls import Workbook, Worksheet, Chart, ExcelChartType, LegendPositionType

from spire_xls_mcp.utils.exceptions import ChartError
from .workbook import get_or_create_workbook
from spire_xls_mcp.utils.cell_utils import EnumMapper

logger = logging.getLogger(__name__)




def _configure_waterfall_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a waterfall chart."""
    chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.ShowCategoryName = True
    chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.ShowSeriesName = True


def _configure_bubble_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a bubble chart."""
    bubbles = chart_options.get('bubbles')
    if bubbles:
        chart.Series[0].Bubbles = sheet.Range[bubbles]


# Chart configuration mapping
CHART_CONFIGURATORS: Dict[ExcelChartType, Callable[[Chart, Worksheet, Dict[str, Any]], None]] = {
    ExcelChartType.WaterFall: _configure_waterfall_chart,
    ExcelChartType.Bubble: _configure_bubble_chart,
}

def _configure_default_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a default chart with axis titles."""
    pass


def create_chart_in_sheet(
        input_filepath: str,
        output_filepath: str,
        data_sheet_name: str,
        chart_sheet_name: str,
        data_range: str,
        chart_type: str,
        target_cell: str,
        chart_options: Optional[Dict[str, Any]] = None
) -> dict[str, Any]:
    """Create chart in sheet with enhanced styling options"""
    chart_options = chart_options or {}
    try:
        if not input_filepath and output_filepath:
            input_filepath = output_filepath
        elif input_filepath and not output_filepath:
            output_filepath = input_filepath
        elif not input_filepath and not output_filepath:
            raise ChartError("input_filepath or output_filepath must be provided")
        wb = get_or_create_workbook(input_filepath)
        data_sheet = None
        chart_sheet = None
        # Find or create sheet
        for ws in wb.Worksheets:
            if ws.Name == data_sheet_name:
                data_sheet = ws
            if ws.Name == chart_sheet_name:
                chart_sheet = ws
        if data_sheet is None:
            raise ChartError(f"sheet '{data_sheet_name}' not found")
        if chart_sheet is None:
            chart_sheet = wb.CreateEmptySheet(chart_sheet_name)

        chart = chart_sheet.Charts.Add()
        chart.ChartType = EnumMapper.get_chart_type_enum(chart_type)

        if axis_titles := chart_options.get('axis_titles'):
            if x_title := axis_titles.get('x_axis_title'):
                chart.PrimaryCategoryAxis.Title = x_title
                chart.PrimaryCategoryAxis.HasTitle = True
            if y_title := axis_titles.get('y_axis_title'):
                chart.PrimaryValueAxis.Title = y_title
                chart.PrimaryValueAxis.HasTitle = True
            
        if series_data := chart_options.get('series'):
            chart.Series.Clear()
            chart.SeriesDataFromRange = False

            common_category_labels_range = chart_options.get('category_labels')

            for s in series_data:
                series_name = s.get('name', '')
                values_range = s.get('values')

                if not values_range:
                    continue

                cs = chart.Series.Add(series_name)
                cs.Values = data_sheet.Range[values_range]

                category_labels_range = s.get('category_labels', common_category_labels_range)
                if category_labels_range:
                    cs.CategoryLabels = data_sheet.Range[category_labels_range]
        else:
            chart.DataRange = data_sheet.Range[data_range]

        target_range = chart_sheet.Range[target_cell]
        chart.LeftColumn = target_range.Column
        chart.TopRow = target_range.Row

        if title := chart_options.get("title"):
            chart.ChartTitle = title

        # Get the specific configurator for the chart type, or the default one
        configurator = CHART_CONFIGURATORS.get(chart.ChartType, _configure_default_chart)
        # Pass all relevant options to the configurator
        options_for_configurator = chart_options.copy()
        configurator(chart, data_sheet, options_for_configurator)
        

        # Apply common styling
        style = chart_options.get('style', {})
        width = style.get('width', 480)
        height = style.get('height', 300)

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
            for series in chart.Series:
                series.Format.Options.IsVaryColor = True
                series.DataPoints.DefaultDataPoint.DataLabels.HasValue = style['has_data_labels']

        chart.Width = width
        chart.Height = height

        wb.SaveToFile(output_filepath)
        chart_title = None
        # HasChartTitle has bug
        try:
            chart_title = chart.ChartTitle if chart.HasChartTitle else None
        except Exception as e:
            pass
        
        return {
            "message": "Chart created successfully",
            "details": {
                "chart_name": chart_title,
                "chart_type": chart_type,
                "sheet_name": chart_sheet_name,
                "data_sheet_name": data_sheet_name,
                "data_range": data_range,
                "position": {
                    "target_cell": target_cell,
                    "top_row": chart.TopRow,
                    "left_column": chart.LeftColumn
                },
                "dimensions": {
                    "width_pixels": chart.Width,
                    "height_pixels": chart.Height
                },
                "output_filepath": output_filepath
            }
        }

    except Exception as e:
        logger.error(f"Failed to create chart: {e}")
        raise ChartError(f"Failed to create chart: {e!s}")
