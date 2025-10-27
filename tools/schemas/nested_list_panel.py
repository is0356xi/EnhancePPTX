from pydantic import BaseModel, Field
from typing import List, Optional, Literal


class ChildItem(BaseModel):
    text: str


class ParentItem(BaseModel):
    text: str
    color: Optional[str] = None
    icon: Optional[Literal["DOT", "TRIANGLE", "NONE"]] = None
    children: List[ChildItem] = Field(default_factory=list)


class NestedListPanelSchema(BaseModel):
    title: Optional[str] = None
    items: List[ParentItem] = Field(default_factory=list)
    parent_pt: int = Field(default=14)
    child_pt: int = Field(default=12)
    title_pt: int = Field(default=16)
    indent_em: float = Field(default=0.3)
    icon_size_pt: int = Field(default=12)
    gap_item_pt: int = Field(default=8)


# Export as Schema for consistent naming
Schema = NestedListPanelSchema