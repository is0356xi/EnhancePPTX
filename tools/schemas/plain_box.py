# src/tools/schemas/plain_box.py
# -*- coding: utf-8 -*-
from typing import Literal, Optional
from pydantic import BaseModel, Field

class PlainBoxStyle(BaseModel):
    """テキストボックスのスタイル定義"""
    font_size: int = 18
    font_color: str = "#000000"
    background_color: str = "#F0F0F0"  # 薄い灰色
    align: Literal["left", "center", "right", "justify"] = "left"
    vertical_align: Literal["top", "middle", "bottom"] = "middle"

class PlainBoxData(BaseModel):
    """
    plain_box ツールのデータスキーマ
    - 枠線なしのシンプルな角丸四角形にテキストを描画します。
    """
    text: str = Field(..., description="ボックス内に表示するテキスト")
    style: Optional[PlainBoxStyle] = Field(default_factory=PlainBoxStyle)

Schema = PlainBoxData  # Export as Schema for consistent naming