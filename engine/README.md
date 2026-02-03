# ExcelSmartCost Engine (Python)

## Setup
```bash
cd engine
python -m venv .venv
. .venv/Scripts/activate
pip install -r requirements.txt
```

## Run
```bash
set EXCELSMARTCOST_TOKEN=DEV_TOKEN_SET_ME
python -m uvicorn app.main:app --host 127.0.0.1 --port 17831
```

## DB
SQLite is stored at `engine/data/cost_data.db` by default (auto created & seeded).

## API
- GET /healthz
- POST /api/calc/eval (X-Token required)
- POST /api/import (X-Token required)
- GET /api/components (X-Token required)

## Tests
```bash
pytest -q
```
