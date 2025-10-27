from pydantic import BaseModel, Field
from typing import Optional


class DividerStyle(BaseModel):
    on: bool = Field(default=False)
    color: str = Field(default="#E6E6E6")
    height_pct: float = Field(default=1.2)
    margin_top_pt: float = Field(default=6.0)


class SlideStyle(BaseModel):
    title_pt: int = Field(default=32)
    subtitle_scale: float = Field(default=0.80)
    subtitle_pt: Optional[int] = None
    color: str = Field(default="#000000")
    divider: Optional[DividerStyle] = None


class SlideStyleSchema(BaseModel):
    title: str = Field(...)
    subtitle: Optional[str] = None
    style: Optional[SlideStyle] = Field(default_factory=SlideStyle)


# Export as Schema for consistent naming
Schema = SlideStyleSchema