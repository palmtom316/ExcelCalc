from pydantic import BaseModel
from pathlib import Path
import os, secrets

class Settings(BaseModel):
    host: str = "127.0.0.1"
    port: int = int(os.getenv("EXCELSMARTCOST_PORT", "17831"))
    db_path: str = os.getenv("EXCELSMARTCOST_DB", str(Path(__file__).resolve().parents[2] / "data" / "cost_data.db"))
    token: str = os.getenv("EXCELSMARTCOST_TOKEN", secrets.token_urlsafe(24))

settings = Settings()
