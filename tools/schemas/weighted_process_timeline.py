from pydantic import BaseModel, Field
from typing import List, Optional


class Process(BaseModel):
    """単一のプロセスを表現する"""

    name: str = Field(..., description="プロセス名")
    weight: float = Field(..., description="プロセスの工数・重要度（幅に反映される）")
    description: Optional[str] = Field(None, description="プロセスの説明")


class ProcessLane(BaseModel):
    """プロセスレーンを表現する"""

    name: str = Field(..., description="レーン名（例：チームA、チームB）")
    processes: List[Process] = Field(
        default_factory=list, description="レーン内のプロセスリスト"
    )
    color: Optional[str] = Field(None, description="レーンの色（HEXカラー）")


class StyleConfig(BaseModel):
    """スタイル設定"""

    lane_height_pt: float = Field(default=40.0, description="レーンの高さ（pt）")
    lane_margin_pt: float = Field(default=5.0, description="レーン間のマージン（pt）")
    process_margin_pt: float = Field(
        default=2.0, description="プロセス間のマージン（pt）"
    )
    show_labels: bool = Field(default=True, description="プロセス名を表示するか")
    show_weights: bool = Field(default=True, description="重みを表示するか")
    label_font_size: float = Field(
        default=10.0, description="ラベルのフォントサイズ（pt）"
    )
    weight_font_size: float = Field(
        default=8.0, description="重みのフォントサイズ（pt）"
    )


class WeightedProcessTimelineSchema(BaseModel):
    """工数付きプロセスタイムラインのスキーマ"""

    title: Optional[str] = Field(None, description="タイムラインのタイトル")
    lanes: List[ProcessLane] = Field(
        default_factory=list, description="プロセスレーンのリスト"
    )
    style: StyleConfig = Field(default_factory=StyleConfig, description="スタイル設定")


# Export as Schema for consistent naming
Schema = WeightedProcessTimelineSchema
