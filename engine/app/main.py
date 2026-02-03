from fastapi import FastAPI, Header, HTTPException
from .core.config import settings
from .db.db_init import init_db
from .schemas.dto import CalcRequest, CalcResponse, ImportRequest, OutputDetail
from .calc.safe_eval import safe_eval, SafeEvalError
from .calc.component_engine import parse_component_call, evaluate_component, pick_primary_output, ComponentError
from .services.components import get_component, list_components
from .services.importers import import_boq_excel, import_quota_excel, import_enterprise_excel, import_components_excel

app = FastAPI(title="ExcelSmartCost Engine", version="0.1.0")

@app.on_event("startup")
def _startup():
    init_db(settings.db_path)

def _auth(x_token: str | None):
    if not x_token or x_token != settings.token:
        raise HTTPException(status_code=401, detail="Unauthorized")

@app.get("/healthz")
def healthz():
    return {"status":"ok","version":app.version}

@app.get("/api/components")
def components(x_token: str | None = Header(default=None, alias="X-Token")):
    _auth(x_token)
    return list_components()

@app.post("/api/calc/eval", response_model=CalcResponse)
def calc_eval(req: CalcRequest, x_token: str | None = Header(default=None, alias="X-Token")):
    _auth(x_token)
    expr = (req.expr or "").strip()
    if not expr:
        return CalcResponse(ok=False, error="Empty expression")
    if expr.startswith("@"):
        try:
            cname, params = parse_component_call(expr)
            comp = get_component(cname)
            if not comp:
                return CalcResponse(ok=False, error=f"Component not found: {cname}")
            outputs = evaluate_component(comp, params)
            primary = pick_primary_output(comp)
            if primary not in outputs:
                return CalcResponse(ok=False, error="Primary output missing")
            value = round(outputs[primary], req.precision)
            outputs_r = {k: round(v, req.precision) for k, v in outputs.items()}
            details = []
            for out in comp.outputs:
                key = out.get("key")
                if key not in outputs:
                    continue
                details.append(OutputDetail(
                    key=key,
                    name=out.get("name"),
                    unit=out.get("unit"),
                    value=round(outputs[key], req.precision),
                    primary=bool(out.get("primary", False))
                ))
            return CalcResponse(
                ok=True,
                value=value,
                outputs=outputs_r,
                primary_key=primary,
                outputs_detail=details,
                component_name=comp.name,
            )
        except (ComponentError, SafeEvalError) as e:
            return CalcResponse(ok=False, error=str(e))
    try:
        v = safe_eval(expr, req.variables)
        return CalcResponse(ok=True, value=round(v, req.precision))
    except SafeEvalError as e:
        return CalcResponse(ok=False, error=str(e))

@app.post("/api/import")
def import_any(req: ImportRequest, x_token: str | None = Header(default=None, alias="X-Token")):
    _auth(x_token)
    kind = req.kind
    sheet = req.sheet_name or kind
    if kind == "boq":
        return import_boq_excel(req.file_path, sheet, req.dedup_mode)
    if kind == "quota":
        return import_quota_excel(req.file_path, sheet, req.dedup_mode)
    if kind == "enterprise":
        return import_enterprise_excel(req.file_path, sheet, req.dedup_mode)
    if kind == "components":
        return import_components_excel(req.file_path, sheet, req.dedup_mode)
    raise HTTPException(status_code=400, detail="Invalid kind")
