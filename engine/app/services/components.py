from __future__ import annotations
import json
from typing import Dict, Any, Optional, List
from ..db.conn import connect
from ..calc.component_engine import ComponentDef

def get_component(name: str) -> Optional[ComponentDef]:
    with connect() as conn:
        row = conn.execute("SELECT * FROM lib_components_v2 WHERE name=?", (name,)).fetchone()
        if not row:
            return None
        return ComponentDef(
            name=row["name"],
            params=json.loads(row["params_json"]),
            outputs=json.loads(row["outputs_json"]),
            formulas=json.loads(row["formulas_json"]),
            version=row["version"],
            description=row["description"],
        )

def list_components() -> List[Dict[str, Any]]:
    with connect() as conn:
        rows = conn.execute("SELECT name, version, description FROM lib_components_v2 ORDER BY name").fetchall()
        return [dict(r) for r in rows]
