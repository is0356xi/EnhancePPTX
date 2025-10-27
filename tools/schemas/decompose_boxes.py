from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any, Union


class DecomposeBoxNode(BaseModel):
    name: str
    children: Optional[List["DecomposeBoxNode"]] = None

    # Allow forward reference
    model_config = {"from_attributes": True}


class DecomposeBoxesSchema(BaseModel):
    # root may be a single node or a list of top-level nodes
    root: Union[DecomposeBoxNode, List[DecomposeBoxNode]]
    column_headers: Optional[List[str]] = None


# Update forward references
DecomposeBoxNode.model_rebuild()

# Export as Schema for consistent naming
Schema = DecomposeBoxesSchema
