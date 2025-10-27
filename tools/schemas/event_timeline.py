from pydantic import BaseModel, Field
from typing import List, Optional


class EventStep(BaseModel):
    """タイムラインの単一ステップ"""

    number: Optional[str] = Field(None, description="ステップ番号（例：'1', '2', ...）")
    title: Optional[str] = Field(None, description="ステップの見出し・タイトル")
    description: Optional[str] = Field(None, description="ステップの詳細説明")
    color: Optional[str] = Field(default="#0D6EFD", description="ステップの色")


class EventTimelineSchema(BaseModel):
    """イベント・プロセスタイムラインのスキーマ"""

    steps: List[EventStep] = Field(..., description="タイムラインのステップリスト")
    orientation: Optional[str] = Field(
        default="horizontal", description="タイムラインの向き（horizontal/vertical）"
    )
    show_numbers: Optional[bool] = Field(
        default=True, description="ステップ番号を表示するか"
    )
    card_bg: Optional[str] = Field(default="#FFFFFF", description="カードの背景色")
    card_border: Optional[str] = Field(default="#E6E6E6", description="カードの枠線色")


Schema = EventTimelineSchema
