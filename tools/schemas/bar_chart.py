from pydantic import BaseModel, Field
from typing import List, Optional


class SeriesData(BaseModel):
    name: str
    values: List[float]


class BarChartSchema(BaseModel):
    title: Optional[str] = None
    categories: List[str] = Field(default_factory=list)
    series: List[SeriesData] = Field(default_factory=list)
    show_legend: bool = Field(default=False)
    data_labels: bool = Field(default=False)


# Export as Schema for consistent naming
Schema = BarChartSchema