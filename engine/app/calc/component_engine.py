from __future__ import annotations
import re
from dataclasses import dataclass
from typing import Dict, Any, Tuple, List
from .safe_eval import safe_eval, SafeEvalError

_COMPONENT_RE = re.compile(r"^@(?P<name>[A-Za-z_][A-Za-z0-9_]*)\((?P<args>.*)\)\s*$")

class ComponentError(Exception):
    pass

def parse_component_call(expr: str) -> Tuple[str, Dict[str, float]]:
    m = _COMPONENT_RE.match(expr.strip())
    if not m:
        raise ComponentError("Invalid component syntax. Use @Name(p1=...,p2=...)")
    name = m.group("name")
    args_str = m.group("args").strip()
    params: Dict[str, float] = {}

    if args_str == "":
        return name, params

    parts = [p.strip() for p in args_str.split(",") if p.strip()]
    for p in parts:
        if "=" not in p:
            raise ComponentError(f"Invalid param: {p}")
        k, v = p.split("=", 1)
        k = k.strip(); v = v.strip()
        try:
            params[k] = float(safe_eval(v, {}))
        except SafeEvalError as e:
            raise ComponentError(f"Invalid param value for {k}: {e}") from e
    return name, params

@dataclass
class ComponentDef:
    name: str
    params: List[Dict[str, Any]]
    outputs: List[Dict[str, Any]]
    formulas: Dict[str, str]
    version: str | None = None
    description: str | None = None

def evaluate_component(comp: ComponentDef, input_params: Dict[str, float]) -> Dict[str, float]:
    vars_: Dict[str, float] = {}
    for p in comp.params:
        key = p["key"]
        if key in input_params:
            vars_[key] = float(input_params[key])
        else:
            if p.get("required", False) and "default" not in p:
                raise ComponentError(f"Missing required param: {key}")
            if "default" in p:
                vars_[key] = float(p["default"])
    results: Dict[str, float] = {}
    for out in comp.outputs:
        ok = out["key"]
        formula = comp.formulas.get(ok)
        if not formula:
            raise ComponentError(f"Missing formula for output: {ok}")
        env = {**vars_, **results}
        try:
            results[ok] = float(safe_eval(formula, env))
        except SafeEvalError as e:
            raise ComponentError(f"Error evaluating {comp.name}.{ok}: {e}") from e
    return results

def pick_primary_output(comp: ComponentDef) -> str:
    for o in comp.outputs:
        if o.get("primary"):
            return o["key"]
    return comp.outputs[0]["key"] if comp.outputs else ""
