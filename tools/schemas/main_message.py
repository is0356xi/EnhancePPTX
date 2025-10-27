from pydantic import BaseModel, Field
from typing import Optional, Literal


class MainMessageStyle(BaseModel):
    font_size_body: int = Field(default=22)
    align: Literal["left", "center", "right"] = Field(default="left")


class MainMessageSchema(BaseModel):
    main_message: str = Field(default="")
    color: Optional[str] = None
    corner_radius: float = Field(default=0.2)
    style: Optional[MainMessageStyle] = Field(default_factory=MainMessageStyle)


# Export as Schema for consistent naming
Schema = MainMessageSchema
