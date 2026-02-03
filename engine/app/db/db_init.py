"""SQLite schema and seed for ExcelSmartCost."""
from __future__ import annotations
import sqlite3, json
from pathlib import Path
from datetime import datetime

SCHEMA_SQL = """
PRAGMA journal_mode=WAL;

CREATE TABLE IF NOT EXISTS lib_boq (
  code TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  unit TEXT,
  description TEXT
);
CREATE INDEX IF NOT EXISTS idx_lib_boq_name ON lib_boq(name);

CREATE TABLE IF NOT EXISTS lib_quota (
  code TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  unit TEXT,
  labor_unit_price REAL NOT NULL DEFAULT 0,
  material_unit_price REAL NOT NULL DEFAULT 0,
  machine_unit_price REAL NOT NULL DEFAULT 0,
  description TEXT
);
CREATE INDEX IF NOT EXISTS idx_lib_quota_name ON lib_quota(name);

CREATE TABLE IF NOT EXISTS lib_enterprise_price (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  resource_name TEXT NOT NULL,
  spec TEXT NOT NULL DEFAULT '',
  category TEXT NOT NULL CHECK(category IN ('labor','material','machine')),
  unit TEXT,
  unit_price REAL NOT NULL,
  source TEXT,
  updated_at TEXT NOT NULL,
  UNIQUE(resource_name, spec, category)
);
CREATE INDEX IF NOT EXISTS idx_ent_name_spec_cat ON lib_enterprise_price(resource_name, spec, category);

CREATE TABLE IF NOT EXISTS lib_components_v2 (
  name TEXT PRIMARY KEY,
  params_json TEXT NOT NULL,
  outputs_json TEXT NOT NULL,
  formulas_json TEXT NOT NULL,
  version TEXT,
  description TEXT
);

CREATE TABLE IF NOT EXISTS pricing_rules (
  rule_name TEXT PRIMARY KEY,
  rate REAL NOT NULL,
  applies_to TEXT NOT NULL
);
"""

DEFAULT_COMPONENTS = [
  {
    'name': 'Beam',
    'params': [
      {'key':'width','unit':'m','required':True},
      {'key':'height','unit':'m','required':True},
      {'key':'len','unit':'m','required':True}
    ],
    'outputs': [
      {'key':'volume_m3','name':'混凝土体积','unit':'m3','primary':True}
    ],
    'formulas': {'volume_m3':'width*height*len'},
    'version':'1.0',
    'description':'示例：梁体积'
  },
  {
    'name':'PowerCableTrench',
    'params':[
      {'key':'L','unit':'m','required':True},
      {'key':'W','unit':'m','required':True},
      {'key':'H','unit':'m','required':True},
      {'key':'k_overbreak','unit':'-','required':False,'default':1.0}
    ],
    'outputs':[
      {'key':'excavation_m3','name':'土方开挖','unit':'m3','primary':False},
      {'key':'concrete_m3','name':'混凝土（占位）','unit':'m3','primary':True},
      {'key':'backfill_m3','name':'回填（占位）','unit':'m3','primary':False}
    ],
    'formulas':{
      'excavation_m3':'L*W*H*k_overbreak',
      'concrete_m3':'0',
      'backfill_m3':'excavation_m3 - concrete_m3'
    },
    'version':'0.1',
    'description':'电力管沟（占位公式：待按大样图补齐）'
  },
  {
    'name':'PipeEncasement',
    'params':[{'key':'L','unit':'m','required':True},{'key':'W','unit':'m','required':True},{'key':'H','unit':'m','required':True}],
    'outputs':[{'key':'encase_m3','name':'包封混凝土','unit':'m3','primary':True}],
    'formulas':{'encase_m3':'L*W*H'},
    'version':'0.1',
    'description':'管道包封（简单体积占位）'
  },
  {
    'name':'PowerCableManhole',
    'params':[{'key':'L','unit':'m','required':True},{'key':'W','unit':'m','required':True},{'key':'H','unit':'m','required':True}],
    'outputs':[{'key':'concrete_m3','name':'井体混凝土','unit':'m3','primary':True}],
    'formulas':{'concrete_m3':'L*W*H'},
    'version':'0.1',
    'description':'电力井（简单体积占位）'
  },
  {
    'name':'PowerPipeJacking',
    'params':[{'key':'L','unit':'m','required':True}],
    'outputs':[{'key':'length_m','name':'顶管长度','unit':'m','primary':True}],
    'formulas':{'length_m':'L'},
    'version':'0.1',
    'description':'电力顶管（长度）'
  },
  {
    'name':'PowerJackingWorkingShaft',
    'params':[{'key':'L','unit':'m','required':True},{'key':'W','unit':'m','required':True},{'key':'H','unit':'m','required':True}],
    'outputs':[{'key':'volume_m3','name':'工作井体积','unit':'m3','primary':True}],
    'formulas':{'volume_m3':'L*W*H'},
    'version':'0.1',
    'description':'顶管工作井（体积占位）'
  },
  {
    'name':'EquipmentFoundation',
    'params':[{'key':'L','unit':'m','required':True},{'key':'W','unit':'m','required':True},{'key':'H','unit':'m','required':True}],
    'outputs':[{'key':'volume_m3','name':'基础混凝土','unit':'m3','primary':True}],
    'formulas':{'volume_m3':'L*W*H'},
    'version':'0.1',
    'description':'设备基础（体积占位）'
  },
  {
    'name':'RingMainUnitFoundation',
    'params':[{'key':'L','unit':'m','required':True},{'key':'W','unit':'m','required':True},{'key':'H','unit':'m','required':True}],
    'outputs':[{'key':'volume_m3','name':'基础混凝土','unit':'m3','primary':True}],
    'formulas':{'volume_m3':'L*W*H'},
    'version':'0.1',
    'description':'环网柜基础（体积占位）'
  }
]

DEFAULT_RULES = [
  ('management_fee_rate', 0.15, 'boq_item_total'),
  ('profit_rate', 0.00, 'boq_item_total'),
]

def init_db(db_path: str) -> None:
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(SCHEMA_SQL)
        conn.executemany(
            'INSERT OR REPLACE INTO pricing_rules(rule_name, rate, applies_to) VALUES (?,?,?)',
            DEFAULT_RULES
        )
        for comp in DEFAULT_COMPONENTS:
            conn.execute(
              'INSERT OR REPLACE INTO lib_components_v2(name, params_json, outputs_json, formulas_json, version, description) '
              'VALUES (?,?,?,?,?,?)',
              (
                comp['name'],
                json.dumps(comp['params'], ensure_ascii=False),
                json.dumps(comp['outputs'], ensure_ascii=False),
                json.dumps(comp['formulas'], ensure_ascii=False),
                comp.get('version'),
                comp.get('description'),
              )
            )
        conn.commit()
    finally:
        conn.close()

if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('--db', required=True)
    args = p.parse_args()
    init_db(args.db)
    print(f'Initialized DB at {args.db}')
