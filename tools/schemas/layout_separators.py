from pydantic import BaseModel, Field
from typing import List, Optional


class SeparatorStyle(BaseModel):
    font_size: int = Field(default=12)
    bold: bool = Field(default=True)
    font_color: str = Field(default="#000000")
    color: str = Field(default="#888888")  # Separator line color


class LayoutSeparatorsSchema(BaseModel):
    sections: List[str] = Field(default_factory=list)
    style: SeparatorStyle = Field(default_factory=SeparatorStyle)


# Export as Schema for consistent naming
Schema = LayoutSeparatorsSchema