from pydantic import BaseModel, Field, validator
from typing import Optional, Union


class ProgressBarSchema(BaseModel):
    title: str = Field(default="")
    text: Optional[str] = None
    current_pct: Optional[float] = None
    numerator: Optional[Union[int, float]] = None
    denominator: Optional[Union[int, float]] = None
    target_pct: float = Field(default=100.0)
    title_pt: int = Field(default=18)
    suffix_pt: int = Field(default=12)

    @validator('current_pct', pre=True, always=True)
    def validate_current_pct(cls, v, values):
        if v is not None:
            return float(v)
        
        # Try to calculate from numerator/denominator
        num = values.get('numerator')
        den = values.get('denominator')
        if num is not None and den is not None:
            try:
                if float(den) != 0:
                    return (float(num) / float(den)) * 100.0
            except (ValueError, TypeError):
                pass
        return 0.0

    @validator('target_pct')
    def validate_target_pct(cls, v):
        return max(0.1, float(v))  # Ensure positive


# Export as Schema for consistent naming
Schema = ProgressBarSchema