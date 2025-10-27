from pydantic import BaseModel, Field
from typing import List, Optional, Literal


class Position(BaseModel):
    x: float
    y: float
    w: float
    h: float


class NodeStyle(BaseModel):
    fill: Optional[str] = None
    stroke: Optional[str] = None
    stroke_pt: float = Field(default=1.2)
    font_color: Optional[str] = Field(default="#000000")
    font_size: Optional[int] = None


class ComponentNode(BaseModel):
    id: str
    kind: Literal["user", "rect"] = Field(default="rect")
    label: Optional[str] = None
    pos: Position
    style: NodeStyle = Field(default_factory=NodeStyle)


class ConnectorStyle(BaseModel):
    color: str = Field(default="#888888")
    pt: float = Field(default=1.2)
    dash: Optional[str] = None


class ComponentConnector(BaseModel):
    from_: str = Field(alias="from")
    to: str
    label: Optional[str] = None
    style: ConnectorStyle = Field(default_factory=ConnectorStyle)

    model_config = {"populate_by_name": True}


class ComponentsSchema(BaseModel):
    nodes: List[ComponentNode] = Field(default_factory=list)
    connectors: List[ComponentConnector] = Field(default_factory=list)


# Export as Schema for consistent naming
Schema = ComponentsSchema
