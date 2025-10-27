from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION
from pptx.util import Pt

from ..utils import setup_chart_title, setup_chart_legend, setup_chart_data_labels
from ..schemas.pie_chart import PieChartSchema


def render(slide, data: PieChartSchema, geom: dict, context: dict):
    """円グラフを描画する"""
    chart_data = CategoryChartData()
    chart_data.categories = [item.label for item in data.items]
    chart_data.add_series("Share", [item.value for item in data.items])

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
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
    setup_chart_legend(chart, data.show_legend, XL_LEGEND_POSITION.RIGHT)

    # データラベル
    data_labels = setup_chart_data_labels(chart, data.data_labels, font_size=12)
    if data_labels:
        data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END
        if data.data_labels == "percent":
            data_labels.show_percentage = True
            data_labels.show_value = False
        elif data.data_labels == "value":
            data_labels.show_percentage = False
            data_labels.show_value = True