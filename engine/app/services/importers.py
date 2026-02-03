from __future__ import annotations
import pandas as pd
from datetime import datetime
from typing import Dict, Any, Literal, List
import json
from ..db.conn import connect

DedupMode = Literal["overwrite","skip"]

def _s(x) -> str:
    return "" if x is None else str(x).strip()

def import_boq_excel(path: str, sheet_name: str = "boq", dedup: DedupMode = "overwrite") -> Dict[str, Any]:
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]
    if not {"code","name"}.issubset(df.columns):
        raise ValueError("BoQ requires columns: code,name")
    total, inserted, updated, skipped = len(df), 0, 0, 0
    warnings: List[str] = []
    with connect() as conn:
        for _, r in df.iterrows():
            code, name = _s(r.get("code")), _s(r.get("name"))
            if not code or not name:
                skipped += 1; continue
            unit, desc = _s(r.get("unit")), _s(r.get("description"))
            if dedup == "skip" and conn.execute("SELECT 1 FROM lib_boq WHERE code=?", (code,)).fetchone():
                skipped += 1; continue
            exists = conn.execute("SELECT 1 FROM lib_boq WHERE code=?", (code,)).fetchone()
            conn.execute("INSERT OR REPLACE INTO lib_boq(code,name,unit,description) VALUES (?,?,?,?)", (code,name,unit,desc))
            updated += 1 if exists else 0
            inserted += 0 if exists else 1
        conn.commit()
    return {"total_rows": total, "inserted": inserted, "updated": updated, "skipped": skipped, "warnings": warnings}

def import_quota_excel(path: str, sheet_name: str = "quota", dedup: DedupMode = "overwrite") -> Dict[str, Any]:
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]
    if not {"code","name"}.issubset(df.columns):
        raise ValueError("Quota requires columns: code,name")
    total, inserted, updated, skipped = len(df), 0, 0, 0
    warnings: List[str] = []
    def fnum(x, field, code):
        try:
            if pd.isna(x) or x == "":
                return 0.0
            return float(x)
        except Exception:
            warnings.append(f"Invalid {field} for {code}, set 0")
            return 0.0
    with connect() as conn:
        for _, r in df.iterrows():
            code, name = _s(r.get("code")), _s(r.get("name"))
            if not code or not name:
                skipped += 1; continue
            unit = _s(r.get("unit"))
            labor = fnum(r.get("labor_unit_price"), "labor_unit_price", code)
            mat = fnum(r.get("material_unit_price"), "material_unit_price", code)
            mac = fnum(r.get("machine_unit_price"), "machine_unit_price", code)
            desc = _s(r.get("description"))
            if dedup == "skip" and conn.execute("SELECT 1 FROM lib_quota WHERE code=?", (code,)).fetchone():
                skipped += 1; continue
            exists = conn.execute("SELECT 1 FROM lib_quota WHERE code=?", (code,)).fetchone()
            conn.execute(
                "INSERT OR REPLACE INTO lib_quota(code,name,unit,labor_unit_price,material_unit_price,machine_unit_price,description) "
                "VALUES (?,?,?,?,?,?,?)",
                (code,name,unit,labor,mat,mac,desc)
            )
            updated += 1 if exists else 0
            inserted += 0 if exists else 1
        conn.commit()
    return {"total_rows": total, "inserted": inserted, "updated": updated, "skipped": skipped, "warnings": warnings}

def import_enterprise_excel(path: str, sheet_name: str = "enterprise", dedup: DedupMode = "overwrite") -> Dict[str, Any]:
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]
    if not {"resource_name","category","unit_price"}.issubset(df.columns):
        raise ValueError("Enterprise requires columns: resource_name,category,unit_price")
    total, inserted, updated, skipped = len(df), 0, 0, 0
    warnings: List[str] = []
    now = datetime.now().isoformat(timespec="seconds")
    with connect() as conn:
        for _, r in df.iterrows():
            rn = _s(r.get("resource_name"))
            cat = _s(r.get("category")).lower()
            if cat not in ("labor","material","machine"):
                warnings.append(f"Invalid category for {rn}, skipped")
                skipped += 1; continue
            spec, unit, src = _s(r.get("spec")), _s(r.get("unit")), _s(r.get("source"))
            try:
                up = float(r.get("unit_price"))
            except Exception:
                warnings.append(f"Invalid unit_price for {rn}, skipped")
                skipped += 1; continue
            if not rn:
                skipped += 1; continue
            if dedup == "skip" and conn.execute(
                "SELECT 1 FROM lib_enterprise_price WHERE resource_name=? AND spec=? AND category=?",
                (rn,spec,cat)).fetchone():
                skipped += 1; continue
            exists = conn.execute(
                "SELECT 1 FROM lib_enterprise_price WHERE resource_name=? AND spec=? AND category=?",
                (rn,spec,cat)).fetchone()
            conn.execute(
                "INSERT OR REPLACE INTO lib_enterprise_price(resource_name,spec,category,unit,unit_price,source,updated_at) "
                "VALUES (?,?,?,?,?,?,?)",
                (rn,spec,cat,unit,up,src,now)
            )
            updated += 1 if exists else 0
            inserted += 0 if exists else 1
        conn.commit()
    return {"total_rows": total, "inserted": inserted, "updated": updated, "skipped": skipped, "warnings": warnings}

def import_components_excel(path: str, sheet_name: str = "components", dedup: DedupMode = "overwrite") -> Dict[str, Any]:
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]
    if not {"name","params_json","outputs_json","formulas_json"}.issubset(df.columns):
        raise ValueError("Components requires columns: name,params_json,outputs_json,formulas_json")
    total, inserted, updated, skipped = len(df), 0, 0, 0
    warnings: List[str] = []
    with connect() as conn:
        for _, r in df.iterrows():
            name = _s(r.get("name"))
            if not name:
                skipped += 1; continue
            try:
                params = json.loads(r.get("params_json"))
                outputs = json.loads(r.get("outputs_json"))
                formulas = json.loads(r.get("formulas_json"))
            except Exception:
                warnings.append(f"{name}: invalid json, skipped")
                skipped += 1; continue
            ver, desc = _s(r.get("version")), _s(r.get("description"))
            if dedup == "skip" and conn.execute("SELECT 1 FROM lib_components_v2 WHERE name=?", (name,)).fetchone():
                skipped += 1; continue
            exists = conn.execute("SELECT 1 FROM lib_components_v2 WHERE name=?", (name,)).fetchone()
            conn.execute(
                "INSERT OR REPLACE INTO lib_components_v2(name,params_json,outputs_json,formulas_json,version,description) "
                "VALUES (?,?,?,?,?,?)",
                (name, json.dumps(params,ensure_ascii=False), json.dumps(outputs,ensure_ascii=False),
                 json.dumps(formulas,ensure_ascii=False), ver or None, desc or None)
            )
            updated += 1 if exists else 0
            inserted += 0 if exists else 1
        conn.commit()
    return {"total_rows": total, "inserted": inserted, "updated": updated, "skipped": skipped, "warnings": warnings}
