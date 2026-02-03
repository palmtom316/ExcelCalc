# Import Templates (Excel)

Create an .xlsx with sheets named:
- boq
- quota
- enterprise
- components

## boq columns
code, name, unit, description

## quota columns
code, name, unit, labor_unit_price, material_unit_price, machine_unit_price, description

## enterprise columns
resource_name, spec, category(labor/material/machine), unit, unit_price, source

## components columns (方案B 多输出)
name, params_json, outputs_json, formulas_json, version, description
