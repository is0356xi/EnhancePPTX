from pydantic import BaseModel, Field
from typing import List, Optional, Literal


class PieChartItem(BaseModel):
    label: str
    value: float


class PieChartSchema(BaseModel):
    title: Optional[str] = None
    items: List[PieChartItem] = Field(default_factory=list)
    show_legend: bool = Field(default=False)
    data_labels: Literal["none", "percent", "value"] = Field(default="none")


# Export as Schema for consistent naming
Schema = PieChartSchema