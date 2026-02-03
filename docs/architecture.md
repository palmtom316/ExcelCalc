# ExcelSmartCost Route A — Architecture

- Windows only
- Excel Add-in: VSTO / C# (Ribbon + WPF dialogs + Excel automation)
- Python Engine: FastAPI (localhost) + SQLite

## Why this route
- C# is best for robust Excel events and high-performance range I/O.
- Python is best for safe expression evaluation, component formulas, and SQLite library operations.
