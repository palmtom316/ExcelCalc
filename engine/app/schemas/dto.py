from pydantic import BaseModel, Field
from typing import Dict, Literal, Optional, List

class CalcRequest(BaseModel):
    expr: str
    variables: Dict[str, float] = Field(default_factory=dict)
    precision: int = 6

class OutputDetail(BaseModel):
    key: str
    name: str | None = None
    unit: str | None = None
    value: float
    primary: bool = False

class CalcResponse(BaseModel):
    ok: bool
    value: float | None = None
    error: str | None = None
    outputs: Dict[str, float] | None = None
    primary_key: str | None = None
    outputs_detail: List[OutputDetail] | None = None
    component_name: str | None = None

class ImportRequest(BaseModel):
    kind: Literal["boq","quota","enterprise","components"]
    file_path: str
    sheet_name: str = ""
    dedup_mode: Literal["overwrite","skip"] = "overwrite"
