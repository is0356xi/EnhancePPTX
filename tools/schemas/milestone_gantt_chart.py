from pydantic import BaseModel, Field
from typing import List, Optional, Union, Literal


class Milestone(BaseModel):
    label: str
    time_index: float


class Task(BaseModel):
    name: str
    start: float
    end: float
    owner: Optional[str] = None


class Phase(BaseModel):
    name: str
    tasks: List[Task] = Field(default_factory=list)


class Owner(BaseModel):
    name: str
    color: Optional[str] = None


class LegendConfig(BaseModel):
    show: bool = Field(default=True)
    position: Literal["below", "above", "right"] = Field(default="below")
    include_task_owners: bool = Field(default=True)


class OptionsConfig(BaseModel):
    label_col_pct: float = Field(default=22.0)
    header_fill: str = Field(default="#F2F2F2")
    show_time_labels: bool = Field(default=True)
    row_height_pt: float = Field(default=22.0)
    cell_shade: Literal["owner", "phase", "none"] = Field(default="owner")
    milestone_row_fill: str = Field(default="#9FC289")
    legend: LegendConfig = Field(default_factory=LegendConfig)
    assert_no_spanned_txbody: bool = Field(default=False)


class MilestoneGanttChartSchema(BaseModel):
    time_axis: List[str] = Field(default_factory=list)
    milestones: List[Milestone] = Field(default_factory=list)
    phases: List[Phase] = Field(default_factory=list)  # For backward compatibility
    owners: Optional[List[Owner]] = None
    options: OptionsConfig = Field(default_factory=OptionsConfig)


# Export as Schema for consistent naming
Schema = MilestoneGanttChartSchema
