from pydantic import BaseModel, Field
from typing import List, Optional


class AxisSpec(BaseModel):
    """軸の仕様"""

    label: str = Field(..., description="軸のラベル（例: 競争上の戦略変数の数）")
    low_label: str = Field(default="低", description="低い側のラベル")
    high_label: str = Field(default="高", description="高い側のラベル")


class QuadrantSpec(BaseModel):
    """象限の仕様"""

    title: str = Field(..., description="象限のタイトル")
    description: Optional[str] = Field(None, description="象限の説明（オプション）")
    color: Optional[str] = Field(default="#E8F4FD", description="象限の背景色")


class PlotPoint(BaseModel):
    """マトリクス上のプロットポイント（オプション機能）"""

    label: str = Field(..., description="ポイントのラベル")
    x: float = Field(..., ge=0, le=1, description="X座標（0.0〜1.0）")
    y: float = Field(..., ge=0, le=1, description="Y座標（0.0〜1.0）")
    color: Optional[str] = Field(default="#FF6B6B", description="ポイントの色")


class StyleConfig(BaseModel):
    """スタイル設定"""

    axis_label_pt: int = Field(default=14, description="軸ラベルのフォントサイズ")
    axis_end_label_pt: int = Field(default=12, description="軸端のラベルフォントサイズ")
    quadrant_title_pt: int = Field(
        default=16, description="象限タイトルのフォントサイズ"
    )
    quadrant_desc_pt: int = Field(default=12, description="象限説明のフォントサイズ")
    axis_color: str = Field(default="#333333", description="軸の色")
    axis_width_pt: float = Field(default=2.0, description="軸の線の太さ")
    grid_color: str = Field(default="#CCCCCC", description="グリッド線の色")
    grid_width_pt: float = Field(default=1.0, description="グリッド線の太さ")
    padding_pct: float = Field(default=5.0, description="全体のパディング（%）")
    point_size_pt: float = Field(default=12.0, description="プロットポイントのサイズ")


class Matrix2x2Schema(BaseModel):
    """2x2マトリクスのスキーマ"""

    title: Optional[str] = Field(None, description="マトリクス全体のタイトル")
    x_axis: AxisSpec = Field(..., description="X軸の仕様")
    y_axis: AxisSpec = Field(..., description="Y軸の仕様")

    # 象限の定義（左下から反時計回り）
    bottom_left: QuadrantSpec = Field(..., description="左下象限（低X、低Y）")
    bottom_right: QuadrantSpec = Field(..., description="右下象限（高X、低Y）")
    top_right: QuadrantSpec = Field(..., description="右上象限（高X、高Y）")
    top_left: QuadrantSpec = Field(..., description="左上象限（低X、高Y）")

    # オプション機能
    plot_points: List[PlotPoint] = Field(
        default_factory=list, description="プロットするポイント（オプション）"
    )

    style: StyleConfig = Field(default_factory=StyleConfig, description="スタイル設定")


Schema = Matrix2x2Schema
