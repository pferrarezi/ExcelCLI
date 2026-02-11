---
name: excel-data-extractor
description: Extract, read, search, and analyze data from Excel files (.xlsx/.xlsm) using the ExcelCLI command-line tool. Use this skill whenever the user mentions reading Excel files, extracting data from spreadsheets, searching for values in Excel, analyzing Excel data, getting sheet information, reading cells, finding formulas, or working with .xlsx/.xlsm files. Also trigger when the user wants to integrate Excel data with other systems, process Excel exports from ERP systems like Linx DMS, or needs structured JSON output from Excel files. This skill is essential for any task involving reading or querying Excel data programmatically.
---

# Excel Data Extractor

## Overview

This skill enables Claude to extract and analyze data from Excel files using **ExcelCLI**, a robust command-line tool designed for both human and LLM interaction. The tool provides structured JSON output with typed cell metadata, making it perfect for integrating Excel data into automated workflows, analyzing reports, and extracting information from spreadsheets.

**Key capabilities:**
- List sheets and file metadata
- Read entire sheets or specific ranges
- Search for values (text or regex)
- Extract individual cells with type information
- List all formulas in a sheet
- Get structured JSON output for downstream processing

## When to Use This Skill

Use this skill when the user wants to:
- **Read Excel files**: "What's in this spreadsheet?", "Show me the data from Sheet1"
- **Extract specific data**: "Get cell B5 from the Sales sheet", "Read rows 10-20"
- **Search for values**: "Find all cells containing 'Total'", "Search for customer names matching a pattern"
- **Analyze structure**: "What sheets are in this file?", "How many rows of data?"
- **List formulas**: "Show me all the formulas in this sheet"
- **Process Excel exports**: "Read this Linx DMS report", "Extract data from this monthly export"
- **Get JSON data**: "Convert this Excel to JSON", "I need machine-readable output"

## Tool Location

The ExcelCLI tool is located at: `C:\apps\clis\ExcelCLI.exe`

Always use the full path when calling the tool unless it's confirmed to be in the system PATH.

## Core Commands

### 1. List Sheets (`info`)

Use this to discover what sheets exist in an Excel file and get basic file metadata.

**Command:**
```bash
C:\apps\clis\ExcelCLI.exe info <file.xlsx> --json
```

**When to use:**
- User asks "What sheets are in this file?"
- Before reading data, to show available sheets
- To validate the file can be opened

**Output:** JSON envelope with sheet names list

### 2. Inspect Sheet (`inspect`)

Get sheet dimensions, row/column counts, and cell range information.

**Command:**
```bash
C:\apps\clis\ExcelCLI.exe inspect <file.xlsx> <SheetName> --json
```

**When to use:**
- User asks "How many rows of data?"
- To understand sheet structure before reading
- To determine if a range is valid

**Output:** JSON with dimensions, first/last cell addresses

### 3. Read Sheet Data (`read`)

The primary command for extracting data from sheets. Supports full sheets, ranges, and custom headers.

**Command:**
```bash
# Full sheet as JSON
C:\apps\clis\ExcelCLI.exe read <file.xlsx> <SheetName> --json

# Specific range
C:\apps\clis\ExcelCLI.exe read <file.xlsx> <SheetName> --range A1:C10 --json

# Limit rows
C:\apps\clis\ExcelCLI.exe read <file.xlsx> <SheetName> --limit 50 --json

# Custom header row (use row 3 as headers, data starts at row 4)
C:\apps\clis\ExcelCLI.exe read <file.xlsx> <SheetName> --header-row 3 --json

# Compact JSON (single line)
C:\apps\clis\ExcelCLI.exe read <file.xlsx> <SheetName> --json-compact

# NDJSON streaming (one row per line)
C:\apps\clis\ExcelCLI.exe read <file.xlsx> <SheetName> --ndjson --limit 100
```

**When to use:**
- User wants to see data from a sheet
- Extracting data for analysis or processing
- Converting Excel to JSON for other tools
- Reading specific ranges or limited rows

**Important notes:**
- Empty headers are replaced with `ColA`, `ColB`, `ColC`, etc.
- Duplicate headers are deduplicated: `Total`, `Total_2`, `Total_3`
- Use `--header-row` if the actual headers are not in row 1
- For large sheets, use `--limit` to avoid overwhelming output

### 4. Read Single Cell (`cell`)

Extract a single cell value with type information.

**Command:**
```bash
C:\apps\clis\ExcelCLI.exe cell <file.xlsx> <SheetName> <CellAddress> --json
```

**When to use:**
- User asks for a specific cell value: "What's in cell B5?"
- Need type information (string, number, date, boolean, etc.)
- Validating a specific value

**Output:** JSON with `cell`, `value`, `type`, and `raw` fields

**Cell types:**
- `string`: Text values
- `number`: Numeric values
- `boolean`: TRUE/FALSE
- `date`: Date values
- `timespan`: Time duration
- `blank`: Empty cells
- `error`: Excel errors (#DIV/0!, #REF!, etc.)

### 5. Search for Values (`search`)

Find cells containing specific text or matching a pattern.

**Command:**
```bash
# Simple text search (case-insensitive)
C:\apps\clis\ExcelCLI.exe search <file.xlsx> <SheetName> "search term" --json

# Regex pattern
C:\apps\clis\ExcelCLI.exe search <file.xlsx> <SheetName> "pattern.*" --regex --json
```

**When to use:**
- "Find all cells with 'Total'"
- "Search for customer names"
- "Locate cells matching a pattern"

**Output:** JSON list of matching cells with addresses and values

### 6. List Formulas (`formulas`)

Extract all formulas from a sheet.

**Command:**
```bash
C:\apps\clis\ExcelCLI.exe formulas <file.xlsx> <SheetName> --json
```

**When to use:**
- User wants to see calculations
- "What formulas are used?"
- Debugging spreadsheet logic
- Understanding data transformations

**Output:** JSON list with `cell`, `formula`, and `value` for each formula

## Workflow Guide

### Typical Workflow

1. **Start with `info`** to list available sheets
2. **Use `inspect`** to understand sheet structure
3. **Read data** with appropriate command:
   - `read` for bulk data extraction
   - `cell` for specific values
   - `search` to find specific content
   - `formulas` to understand calculations
4. **Process the JSON output** as needed

### Example: Complete Sheet Analysis

```bash
# Step 1: What sheets exist?
C:\apps\clis\ExcelCLI.exe info report.xlsx --json

# Step 2: How big is the Sales sheet?
C:\apps\clis\ExcelCLI.exe inspect report.xlsx Sales --json

# Step 3: Read the data (first 20 rows to preview)
C:\apps\clis\ExcelCLI.exe read report.xlsx Sales --limit 20 --json

# Step 4: Search for specific values
C:\apps\clis\ExcelCLI.exe search report.xlsx Sales "Total" --json

# Step 5: Check formulas
C:\apps\clis\ExcelCLI.exe formulas report.xlsx Sales --json
```

### Example: Extracting Specific Data

```bash
# User: "Get the value in cell D15 from the Budget sheet"
C:\apps\clis\ExcelCLI.exe cell financials.xlsx Budget D15 --json

# User: "Read rows 5-25 from columns A through F"
C:\apps\clis\ExcelCLI.exe read data.xlsx Sheet1 --range A5:F25 --json

# User: "Find all cells with customer names matching 'Silva'"
C:\apps\clis\ExcelCLI.exe search customers.xlsx Dados "Silva" --json
```

## Understanding JSON Output

All commands with `--json` flag return a **stable envelope** structure:

```json
{
  "schemaVersion": "1.0",
  "toolVersion": "1.0.0+abc123",
  "command": "read",
  "success": true,
  "data": { /* command-specific payload */ },
  "warnings": [],
  "errorCode": null,
  "message": null
}
```

**Always check `success` first:**
- If `true`: Process the `data` field
- If `false`: Check `errorCode` and `message` for the error

**Common error codes:**
- `MISSING_ARGUMENT`: Required argument not provided
- `INVALID_ARGUMENT`: Invalid flag value
- `UNKNOWN_COMMAND`: Command not recognized
- `SHEET_NOT_FOUND`: Sheet doesn't exist
- `UNHANDLED_ERROR`: Unexpected exception

## Output Modes

| Mode | Flag | Use Case |
|------|------|----------|
| **Pretty JSON** | `--json` | Human-readable, indented JSON |
| **Compact JSON** | `--json-compact` | Single-line JSON (good for piping) |
| **NDJSON** | `--ndjson` | Streaming format (one row per line) |

**NDJSON format:**
```
{"schemaVersion":"1.0","command":"read","success":true,"count":150,"displayed":3}
{"ID":1,"Name":"João","Value":1500.0}
{"ID":2,"Name":"Maria","Value":2300.0}
{"ID":3,"Name":"Pedro","Value":900.0}
```

First line is metadata envelope, subsequent lines are data rows.

## Best Practices

### 1. Always Use JSON Mode

For LLM/Agent usage, **always** include `--json`, `--json-compact`, or `--ndjson`:

```bash
# Good ✓
C:\apps\clis\ExcelCLI.exe read data.xlsx Sales --json

# Avoid (human table format, harder to parse)
C:\apps\clis\ExcelCLI.exe read data.xlsx Sales
```

### 2. Check Success Before Processing

```python
import subprocess
import json

result = subprocess.run(
    [r"C:\apps\clis\ExcelCLI.exe", "read", "data.xlsx", "Sales", "--json"],
    capture_output=True,
    text=True
)

response = json.loads(result.stdout)

if response["success"]:
    data = response["data"]
    # Process data
else:
    error_code = response["errorCode"]
    error_msg = response["message"]
    # Handle error
```

### 3. Use Limits for Large Sheets

For exploration or preview, limit output to avoid overwhelming data:

```bash
# Preview first 50 rows
C:\apps\clis\ExcelCLI.exe read large_file.xlsx Data --limit 50 --json
```

### 4. Specify Header Row When Needed

If headers are not in row 1, specify the correct row:

```bash
# Headers in row 3, data starts at row 4
C:\apps\clis\ExcelCLI.exe read report.xlsx Sales --header-row 3 --json
```

### 5. Use Regex for Complex Searches

For pattern matching, use `--regex` flag:

```bash
# Find cells starting with "Total" followed by anything
C:\apps\clis\ExcelCLI.exe search data.xlsx Sales "Total.*" --regex --json

# Find email addresses
C:\apps\clis\ExcelCLI.exe search contacts.xlsx Emails "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}" --regex --json
```

## Integration with Linx DMS / ERP Workflows

The skill is particularly useful for processing Excel exports from Linx DMS:

**Example: Processing Monthly Sales Report**
```bash
# 1. List sheets in the export
C:\apps\clis\ExcelCLI.exe info "Linx_Export_January_2026.xlsx" --json

# 2. Read sales data
C:\apps\clis\ExcelCLI.exe read "Linx_Export_January_2026.xlsx" "Vendas" --json

# 3. Search for specific dealership data
C:\apps\clis\ExcelCLI.exe search "Linx_Export_January_2026.xlsx" "Vendas" "VW Matriz" --json
```

**Example: Extracting Efetivo Data**
```bash
# Read employee data from specific range
C:\apps\clis\ExcelCLI.exe read "Efetivo_Novos.xlsx" "Dados" --range A1:J100 --json

# Get specific month's data (if month is in cell B2)
C:\apps\clis\ExcelCLI.exe cell "Efetivo_Novos.xlsx" "Dados" B2 --json
```

## Error Handling

When commands fail, the JSON envelope will have `success: false` and provide error details.

**Example error response:**
```json
{
  "schemaVersion": "1.0",
  "command": "read",
  "success": false,
  "data": null,
  "errorCode": "SHEET_NOT_FOUND",
  "message": "Sheet 'SalesData' not found in workbook. Available sheets: Sales, Customers, Products"
}
```

**Common issues and solutions:**

| Error | Cause | Solution |
|-------|-------|----------|
| `SHEET_NOT_FOUND` | Sheet name typo or doesn't exist | Run `info` to see available sheets |
| `INVALID_ARGUMENT` | Bad flag value | Check command syntax |
| File not found | Wrong file path | Verify file path is correct |
| Permission denied | File is open | Close the Excel file first |

## Advanced Usage

### Chaining Commands

You can use ExcelCLI output as input to other tools:

```bash
# Extract data and pipe to jq for further processing
C:\apps\clis\ExcelCLI.exe read data.xlsx Sales --json | jq '.data.rows[] | select(.Total > 1000)'

# Convert to CSV using Python
C:\apps\clis\ExcelCLI.exe read data.xlsx Sales --ndjson > temp.ndjson
python process_ndjson_to_csv.py temp.ndjson output.csv
```

### Working with Multiple Sheets

```python
import subprocess
import json

# Get all sheets
info_result = subprocess.run(
    [r"C:\apps\clis\ExcelCLI.exe", "info", "report.xlsx", "--json"],
    capture_output=True, text=True
)
sheets = json.loads(info_result.stdout)["data"]["sheets"]

# Read data from each sheet
all_data = {}
for sheet in sheets:
    read_result = subprocess.run(
        [r"C:\apps\clis\ExcelCLI.exe", "read", "report.xlsx", sheet, "--json"],
        capture_output=True, text=True
    )
    all_data[sheet] = json.loads(read_result.stdout)["data"]
```

## Limitations

- **Read-only**: ExcelCLI can only read data, not write or modify Excel files
- **Windows only**: The provided executable is for Windows (win-x64)
- **File access**: The file must be closed in Excel before reading
- **Large files**: Very large sheets may take time to process; use `--limit` for previews

## Quick Reference

| Task | Command Pattern |
|------|----------------|
| List sheets | `ExcelCLI.exe info <file> --json` |
| Sheet dimensions | `ExcelCLI.exe inspect <file> <sheet> --json` |
| Read full sheet | `ExcelCLI.exe read <file> <sheet> --json` |
| Read range | `ExcelCLI.exe read <file> <sheet> --range A1:C10 --json` |
| Limit rows | `ExcelCLI.exe read <file> <sheet> --limit 50 --json` |
| Custom headers | `ExcelCLI.exe read <file> <sheet> --header-row 3 --json` |
| Single cell | `ExcelCLI.exe cell <file> <sheet> <cell> --json` |
| Search text | `ExcelCLI.exe search <file> <sheet> "term" --json` |
| Search regex | `ExcelCLI.exe search <file> <sheet> "pattern" --regex --json` |
| List formulas | `ExcelCLI.exe formulas <file> <sheet> --json` |

## Troubleshooting

**Q: The command isn't working**
- Verify the file path is correct and file exists
- Ensure the Excel file is closed (not open in Excel)
- Check that you're using the full path: `C:\apps\clis\ExcelCLI.exe`

**Q: Getting "Sheet not found" error**
- Run `info` command to see exact sheet names
- Sheet names are case-sensitive
- Check for spaces in sheet names (use quotes: `"Sheet Name"`)

**Q: Output is truncated or incomplete**
- For large sheets, use `--limit` to control output size
- Consider reading specific ranges instead of full sheet
- Use `--ndjson` for streaming large datasets

**Q: Headers look wrong (ColA, ColB, etc.)**
- This means row 1 has empty headers
- If headers are in a different row, use `--header-row <N>`
- This is expected behavior when headers are missing

## Example Conversations

**User:** "What's in the Excel file sales_2026.xlsx?"

**Response:**
```bash
# First, list all sheets
C:\apps\clis\ExcelCLI.exe info sales_2026.xlsx --json
```
[Parse JSON, show sheets to user, then:]
```bash
# Inspect the main sheet
C:\apps\clis\ExcelCLI.exe inspect sales_2026.xlsx "Sales" --json
```
[Show dimensions, then:]
```bash
# Preview first 20 rows
C:\apps\clis\ExcelCLI.exe read sales_2026.xlsx "Sales" --limit 20 --json
```

---

**User:** "Find all rows where the customer name contains 'Silva'"

**Response:**
```bash
C:\apps\clis\ExcelCLI.exe search customers.xlsx "Dados" "Silva" --json
```
[Parse results and present to user]

---

**User:** "Get the value from cell E10 in the Budget sheet"

**Response:**
```bash
C:\apps\clis\ExcelCLI.exe cell financials.xlsx "Budget" E10 --json
```
[Parse and show: "Cell E10 contains: 45,230.50 (number)"]

## Summary

Use the Excel Data Extractor skill whenever you need to read, search, or analyze Excel files. The tool provides structured JSON output perfect for integration with other systems, automated workflows, and data processing tasks. Always start with `info` to understand the file structure, then use the appropriate command (`read`, `cell`, `search`, `formulas`) based on what the user needs. Remember to include `--json` for machine-readable output and check the `success` field before processing results.
