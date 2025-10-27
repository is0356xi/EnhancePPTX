from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Pt

from ..utils import setup_chart_title, setup_chart_legend, setup_chart_data_labels
from ..schemas.bar_chart import BarChartSchema


def render(slide, data: BarChartSchema, geom: dict, context: dict):
    """棒グラフを描画する"""
    chart_data = CategoryChartData()
    chart_data.categories = data.categories

    for series_data in data.series:
        chart_data.add_series(
            series_data.name, series_data.values
        )

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        geom["left"],
        geom["top"],
        geom["width"],
        geom["height"],
        chart_data,
    )

    chart = graphic_frame.chart

    # グラフタイトル
    setup_chart_title(chart, data.title)

    # 凡例
    setup_chart_legend(chart, data.show_legend, XL_LEGEND_POSITION.BOTTOM)

    # データラベル
    if data.data_labels:
        setup_chart_data_labels(chart, True)