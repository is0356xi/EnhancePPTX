from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any


class PyramidLevel(BaseModel):
    name: str
    subtitle: Optional[str] = None
    items: Optional[List[str]] = None


class OutlineConfig(BaseModel):
    width_pt: float = Field(default=1.0)
    color: str = Field(default="#000000")


class FontConfig(BaseModel):
    bold: bool = Field(default=True)
    align: str = Field(default="center")
    size_pt: Optional[int] = None


class StyleConfig(BaseModel):
    palette: Optional[List[str]] = None
    outline: OutlineConfig = Field(default_factory=OutlineConfig)
    font: FontConfig = Field(default_factory=FontConfig)


class PyramidHierarchySchema(BaseModel):
    title: Optional[str] = None
    levels: List[PyramidLevel] = Field(default_factory=list)
    style: StyleConfig = Field(default_factory=StyleConfig)


# Export as Schema for consistent naming
Schema = PyramidHierarchySchema