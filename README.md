# ExcelCLI - Universal Excel Data Extraction Tool

A robust command-line tool for extracting data from Excel files (.xlsx/.xlsm). Designed for both humans and LLMs with beautiful formatted output and structured JSON.

## Features

- ✅ Read sheets, cells, ranges
- ✅ Search for values
- ✅ Extract formulas
- ✅ JSON output for LLMs
- ✅ Beautiful tables for humans
- ✅ Clean stdout for pipes (JSON only)
- ✅ Human UX on stderr (with --quiet support)
- ✅ Works with .xlsx and .xlsm files
- ✅ Single standalone .exe

## Installation

Download the latest `excel-cli.exe` from releases or build from source:

```bash
cd ExcelCLI
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

## Usage

```bash
excel-cli <command> [arguments] [options]
```

### Commands

#### `info` - List all sheets

```bash
excel-cli info data.xlsx
# Output:
# ┌─────────────┐
# │ Sheet Name  │
# ├─────────────┤
# │ Sales       │
# │ Customers   │
# └─────────────┘
```

#### `inspect` - Show sheet dimensions

```bash
excel-cli inspect data.xlsx Sales
# Shows: rows, columns, first/last cell
```

#### `read` - Read sheet data

```bash
# As table
excel-cli read data.xlsx Sales --limit 10

# As JSON (LLM-friendly)
excel-cli read data.xlsx Sales --json

# Specific range
excel-cli read data.xlsx Sales --range A1:C10
```

#### `cell` - Read single cell

```bash
excel-cli cell data.xlsx Sales B5
# Output: Cell B5: 1500.0

# JSON output
excel-cli cell data.xlsx Sales B5 --json
# {"cell":"B5","value":1500.0}
```

#### `search` - Search for value

```bash
excel-cli search data.xlsx Sales "João"
# Shows all cells containing "João"
```

#### `formulas` - List all formulas

```bash
excel-cli formulas data.xlsx Sales
# Shows cells with formulas, their formulas, and calculated values
```

### Options

- `--json`, `-j` - Output in JSON format (perfect for LLMs)
- `--range <range>` - Specify cell range (e.g., A1:C10)
- `--limit <n>` - Limit number of rows displayed
- `--quiet`, `-q` - Suppress human UX output (stderr)

## stdout vs stderr

- stdout: JSON only (clean, deterministic)
- stderr: tables, colors, human messages
- `--quiet`: suppresses stderr output

Examples:

```bash
# Keep JSON only on stdout
excel-cli read data.xlsx Sales --json > data.json

# Keep human tables/messages separate
excel-cli read data.xlsx Sales --limit 10 2> human.log

# Suppress human output completely
excel-cli read data.xlsx Sales --json --quiet
```

## For LLMs

Use `--json` flag for structured output:

```bash
# List sheets as JSON
excel-cli info data.xlsx --json
# {
#   "schemaVersion": "1.0",
#   "toolVersion": "0.1.0",
#   "command": "info",
#   "success": true,
#   "data": { "sheets": ["Sales","Customers","Summary"] },
#   "warnings": [],
#   "errorCode": null,
#   "message": null
# }

# Read data as JSON
excel-cli read data.xlsx Sales --json
# {
#   "schemaVersion": "1.0",
#   "toolVersion": "0.1.0",
#   "command": "read",
#   "success": true,
#   "data": {
#     "rows": [{"ID":1,"Name":"João","Value":1500.0}],
#     "count": 150,
#     "displayed": 150
#   },
#   "warnings": [],
#   "errorCode": null,
#   "message": null
# }

# Error (example)
excel-cli inspect data.xlsx MissingSheet --json
# {
#   "schemaVersion": "1.0",
#   "toolVersion": "0.1.0",
#   "command": "inspect",
#   "success": false,
#   "data": null,
#   "warnings": [],
#   "errorCode": "SHEET_NOT_FOUND",
#   "message": "Sheet 'MissingSheet' not found. Available sheets: Sales, Customers"
# }
```

## Examples

```bash
# Quick overview
excel-cli info budget.xlsx

# Inspect specific sheet
excel-cli inspect budget.xlsx "January"

# Read first 20 rows
excel-cli read budget.xlsx "January" --limit 20

# Extract specific range
excel-cli read budget.xlsx "January" --range A1:E50

# Find all cells with "Total"
excel-cli search budget.xlsx "January" "Total"

# Export formulas for analysis
excel-cli formulas budget.xlsx "January" --json

# Get single cell value
excel-cli cell budget.xlsx "January" D15
```

## Building from Source

Requirements:

- .NET 9.0+ SDK

```bash
# Debug build
dotnet build

# Release build (optimized)
dotnet build -c Release

# Publish as single .exe (~25MB)
dotnet publish -c Release -r win-x64 --self-contained \
  -p:PublishSingleFile=true \
  -p:PublishTrimmed=true

# Output: bin/Release/net10.0/win-x64/publish/excel-cli.exe
```

## Dependencies

- **ClosedXML** - Excel file handling
- **Spectre.Console** - Beautiful CLI tables and formatting

## License

MIT License

## Contributing

Issues and PRs welcome!
