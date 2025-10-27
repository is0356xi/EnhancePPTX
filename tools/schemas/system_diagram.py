from pydantic import BaseModel, Field
from typing import List, Optional, Literal


class GridConfig(BaseModel):
    rows: int = Field(default=3)
    cols: int = Field(default=5)


class NodePosition(BaseModel):
    row: int
    col: int


class SystemNode(BaseModel):
    id: str
    type: Optional[str] = None
    label: str
    pos: NodePosition


class Connector(BaseModel):
    from_: str = Field(alias="from")
    to: str
    label: Optional[str] = None
    style: Optional[str] = None
    arrow_head: Optional[Literal["start", "end", "both", "none"]] = None

    model_config = {"populate_by_name": True}


class Boundary(BaseModel):
    label: str
    nodes: List[str] = Field(default_factory=list)
    style: Optional[str] = None
    color: Optional[str] = None


class SystemDiagramSchema(BaseModel):
    title: Optional[str] = None
    grid: GridConfig = Field(default_factory=GridConfig)
    nodes: List[SystemNode] = Field(default_factory=list)
    connectors: List[Connector] = Field(default_factory=list)
    boundaries: List[Boundary] = Field(default_factory=list)


# Export as Schema for consistent naming
Schema = SystemDiagramSchema