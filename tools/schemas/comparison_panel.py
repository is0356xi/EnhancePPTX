from pydantic import BaseModel, Field, model_validator
from typing import List, Optional, Literal


class SectionSpec(BaseModel):
    title: Optional[str] = None
    items: List[str] = Field(default_factory=list)
    color: Optional[str] = Field(
        default="#E8F4FD"
    )  # 薄い青色（黒いテキストが読みやすい）
    chip_color: Optional[str] = Field(default="#F0F8FF")  # さらに薄い青色


class StyleConfig(BaseModel):
    section_gap_pct: float = Field(default=3.0)
    padding_pct: float = Field(default=4.0)
    panel_title_pt: int = Field(default=18)
    section_title_pt: int = Field(default=16)
    item_pt: int = Field(default=14)
    chip_radius_pct: float = Field(default=0.30)
    chip_vpad_pt: float = Field(default=6.0)
    chip_spacing_pt: float = Field(default=6.0)


class ComparisonPanelSchema(BaseModel):
    title: Optional[str] = None
    direction: Literal["horizontal", "vertical"] = Field(default="horizontal")
    left: Optional[SectionSpec] = None
    right: Optional[SectionSpec] = None
    top: Optional[SectionSpec] = None
    bottom: Optional[SectionSpec] = None
    style: StyleConfig = Field(default_factory=StyleConfig)

    @model_validator(mode="after")
    def validate_sections_by_direction(self):
        if self.direction == "horizontal":
            if not self.left or not self.right:
                raise ValueError(
                    "direction='horizontal' requires both 'left' and 'right'"
                )
            if self.top or self.bottom:
                raise ValueError("direction='horizontal' cannot have 'top' or 'bottom'")
        elif self.direction == "vertical":
            if not self.top or not self.bottom:
                raise ValueError(
                    "direction='vertical' requires both 'top' and 'bottom'"
                )
            if self.left or self.right:
                raise ValueError("direction='vertical' cannot have 'left' or 'right'")
        return self


Schema = ComparisonPanelSchema
