# ExcelSmartCost — Route A Scaffold (VSTO C# + Python Engine)

Generated: 2026-02-02 14:43:45

## Contents
- engine/: Python FastAPI engine + SQLite schema + safe eval + components (multi-output) + importers
- add-in/: VSTO skeleton files (copy into a VS VSTO project)
- docs/: architecture + import template specs

## Quick start (engine)
```bash
cd engine
python -m venv .venv
. .venv/Scripts/activate
pip install -r requirements.txt
set EXCELSMARTCOST_TOKEN=DEV_TOKEN_SET_ME
python -m uvicorn app.main:app --host 127.0.0.1 --port 17831
```

## Seeded components
- Beam (完整)
- 7 个行业构件（当前为占位公式；待大样图后补齐）

## Add-in
Create a VSTO Excel Add-in project in Visual Studio and copy add-in/ExcelSmartCostAddIn/*.cs + Ribbon.xml.
Set EngineClient.Token = DEV_TOKEN_SET_ME and run Excel.
