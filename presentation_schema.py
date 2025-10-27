from pydantic import BaseModel, Field
from typing import List, Optional, Union, Any, Dict
from enum import Enum


class SlideToolType(str, Enum):
    """Known tool names.

    Keeps a canonical list but components may use arbitrary strings.
    """

    BAR_CHART = "bar_chart"
    COMPARISON_PANEL = "comparison_panel"
    COMPONENT_DIAGRAM = "component_diagram"
    DECOMPOSE_BOXES = "decompose_boxes"
    MAIN_MESSAGE = "main_message"
    LAYOUT_SEPARATORS = "layout_separators"
    MATRIX_2X2 = "matrix_2x2"
    MILESTONE_GANTT_CHART = "milestone_gantt_chart"
    NESTED_LIST_PANEL = "nested_list_panel"
    PIE_CHART = "pie_chart"
    PROGRESS_BAR = "progress_bar"
    PYRAMID_HIERARCHY = "pyramid_hierarchy"
    SYSTEM_DIAGRAM = "system_diagram"
    WEIGHTED_PROCESS_TIMELINE = "weighted_process_timeline"
    EVENT_TIMELINE = "event_timeline"


class Pos(BaseModel):
    """Absolute percentage position/size as {x,y,w,h} (percent 0..100).

    The LayoutEngine will convert this percent mapping to EMU.
    """

    x: Optional[float] = Field(None, description="left percentage (0..100)")
    y: Optional[float] = Field(None, description="top percentage (0..100)")
    w: Optional[float] = Field(None, description="width percentage (0..100)")
    h: Optional[float] = Field(None, description="height percentage (0..100)")


class Component(BaseModel):
    """Single component entry in `components[]`.

    Mirrors the IR used by `render.py`.
    """

    tool: Union[SlideToolType, str] = Field(
        ...,
        description="Tool name (e.g. 'bar_chart')",
    )
    id: Optional[str] = Field(None, description="Component id")
    pos: Optional[Pos] = Field(None, description="Percent box (x,y,w,h)")
    anchor: Optional[str] = Field(
        None,
        description=("Anchor (left/right/top/bottom/title/center)"),
    )
    z_index: Optional[int] = Field(None, description="Z-order integer")
    data: Optional[Dict[str, Any]] = Field(
        default_factory=dict,
        description="Tool data payload",
    )
    style: Optional[Dict[str, Any]] = Field(
        default_factory=dict,
        description="Style overrides",
    )
    group: Optional[bool] = Field(
        True,
        description="Render into GroupShape",
    )


class Slide(BaseModel):
    """Slide representation compatible with render.py.

    Optional fields keep the model permissive to match the IR normalizer.
    """

    id: Optional[str] = Field(
        None,
        description=(
            "Slide identifier " "(optional, auto-generated from title if not provided)"
        ),
    )
    title: str = Field(..., description="Slide title (required)")
    background: Optional[str] = Field(
        None,
        description="Background color (hex)",
    )
    components: List[Component] = Field(
        default_factory=list,
        description="List of components",
    )

    def __post_init__(self):
        """Post-init to auto-generate ID from title if not provided."""
        if not self.id and self.title:
            # Generate a simple ID from title (remove special chars, lowercase)
            import re

            clean_title = re.sub(r"[^\w\s-]", "", self.title)
            clean_title = re.sub(r"[-\s]+", "_", clean_title)
            self.id = clean_title.lower().strip("_")

    @classmethod
    def __pydantic_init_subclass__(cls, **kwargs):
        """Enable post_init for Pydantic v2."""
        super().__pydantic_init_subclass__(**kwargs)
        cls.model_rebuild()

    def model_post_init(self, __context) -> None:
        """Pydantic v2 equivalent of __post_init__."""
        if not self.id and self.title:
            import re

            clean_title = re.sub(r"[^\w\s-]", "", self.title)
            clean_title = re.sub(r"[-\s]+", "_", clean_title)
            self.id = clean_title.lower().strip("_")


class PresentationMetadata(BaseModel):
    """Top-level metadata object matching render.py's `meta` map."""

    title: Optional[str] = Field(None, description="Presentation title")
    author: Optional[str] = Field(None, description="Presentation author")
    description: Optional[str] = Field(None, description="Brief description")
    slide_size: Optional[Dict[str, Any]] = Field(
        default_factory=dict,
        description="Slide size config (e.g. preset: '16x9')",
    )
    created_date: Optional[str] = Field(None, description="Creation date")


class PresentationSchema(BaseModel):
    """Top-level IR schema for the presentation used by the renderer.

    Matches the IR shape expected by `render.py`.
    Expected top-level keys: version, meta, theme, slides
    """

    version: int = Field(1, description="IR version")
    meta: PresentationMetadata = Field(
        default_factory=PresentationMetadata,
        description="Presentation metadata (meta)",
    )
    theme: Optional[Dict[str, Any]] = Field(
        default_factory=dict, description="Theme mapping (name->color)"
    )
    slides: List[Slide] = Field(
        default_factory=list,
        description="Ordered slides",
    )

    class Config:
        """Pydantic config:
        be permissive for the IR and allow population by name.
        """

        allow_population_by_field_name = True
        extra = "allow"

    def get_slide_count(self) -> int:
        return len(self.slides)

    def get_tools_used(self) -> List[str]:
        tools = set()
        for slide in self.slides:
            for comp in slide.components:
                t = (
                    comp.tool.value
                    if isinstance(comp.tool, SlideToolType)
                    else str(comp.tool)
                )
                tools.add(t)
        return sorted(list(tools))


# Backwards-compatible export name
Schema = PresentationSchema
