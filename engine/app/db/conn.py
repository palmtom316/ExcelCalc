from __future__ import annotations
import sqlite3
from contextlib import contextmanager
from ..core.config import settings

@contextmanager
def connect(db_path: str | None = None):
    conn = sqlite3.connect(db_path or settings.db_path)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()
