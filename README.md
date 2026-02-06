# ExcelCLI — Universal Excel Data Extraction Tool

A robust command-line tool for extracting data from Excel files (`.xlsx` / `.xlsm`).  
Designed for **both humans and LLMs** — beautiful formatted tables on stderr, structured JSON on stdout, and a stdio loop ready for agent integration.

## Features

| Category          | Details                                                                                  |
| ----------------- | ---------------------------------------------------------------------------------------- |
| **Data commands** | `info`, `inspect`, `read`, `cell`, `search`, `formulas`                                  |
| **Output modes**  | `--json` (pretty), `--json-compact` (single line), `--ndjson` (streaming)                |
| **Human UX**      | Spectre.Console tables & colours on stderr, `--quiet` / `--no-color`                     |
| **Agent / LLM**   | Stable JSON envelope, typed cell metadata, tool manifest (`tools`), stdio loop (`serve`) |
| **Parsing**       | Header fallback (`ColA`, `ColB`…), `--header-row`, header deduplication                  |
| **Search**        | Case-insensitive text match, `--regex` for pattern matching                              |
| **Portability**   | Single self-contained `.exe` (win-x64), no runtime required                              |

## Installation

Download the latest `ExcelCLI.exe` from releases, or build from source:

```bash
cd ExcelCLI
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
# → bin\Release\net10.0\win-x64\publish\ExcelCLI.exe
```

---

## Quick Start

```bash
excel-cli info data.xlsx                          # list sheets (table)
excel-cli read data.xlsx Sales --limit 10         # first 10 rows (table)
excel-cli read data.xlsx Sales --json             # full sheet (JSON envelope)
excel-cli search data.xlsx Sales "Total" --json   # find cells containing "Total"
excel-cli tools                                   # print tool manifest for agents
```

---

## Commands

### `info <file>` — List sheets and file metadata

```bash
excel-cli info data.xlsx
# ┌─────────────┐
# │ Sheet Name  │
# ├─────────────┤
# │ Sales       │
# │ Customers   │
# └─────────────┘

excel-cli info data.xlsx --json
# → JSON envelope with { "sheets": ["Sales","Customers"] }
```

### `inspect <file> <sheet>` — Sheet dimensions and preview

```bash
excel-cli inspect data.xlsx Sales
# Rows, columns, first/last cell address

excel-cli inspect data.xlsx Sales --json-compact
# → single-line JSON envelope
```

### `read <file> <sheet>` — Read sheet data

```bash
# Human table
excel-cli read data.xlsx Sales --limit 20

# JSON (indented)
excel-cli read data.xlsx Sales --json

# Compact JSON (one line — ideal for pipes)
excel-cli read data.xlsx Sales --json-compact

# NDJSON streaming (one JSON object per row)
excel-cli read data.xlsx Sales --ndjson --limit 100

# Specific range
excel-cli read data.xlsx Sales --range A1:C10 --json

# Custom header row (row 3 becomes column names, data starts at row 4)
excel-cli read data.xlsx Sales --header-row 3 --json
```

### `cell <file> <sheet> <cell>` — Read a single cell with type info

```bash
excel-cli cell data.xlsx Sales B5
# Cell B5: 1500.0

excel-cli cell data.xlsx Dados A2 --json
# {
#   "schemaVersion": "1.0",
#   "command": "cell",
#   "success": true,
#   "data": {
#     "cell": "A2",
#     "value": "Olá Mundo",
#     "type": "string",
#     "raw": "Olá Mundo"
#   }
# }
```

In JSON mode, `type` is one of: `string`, `number`, `boolean`, `date`, `timespan`, `blank`, `error`.  
`raw` contains the unformatted stored value.

### `search <file> <sheet> <term>` — Find values

```bash
# Simple text search (case-insensitive)
excel-cli search data.xlsx Sales "João"

# Regex pattern
excel-cli search data.xlsx Sales "Total.*" --regex --json
```

### `formulas <file> <sheet>` — List formulas

```bash
excel-cli formulas data.xlsx Dados --json
# Each entry: { "cell": "B3", "formula": "=B2*2", "value": "84" }
```

### `tools` — Print tool manifest (JSON)

Outputs a structured definition of every command, its parameters, and types.  
Useful for MCP clients, LLM tool-calling frameworks, or self-documenting agents.

```bash
excel-cli tools
# [
#   {
#     "name": "info",
#     "description": "List all sheet names...",
#     "parameters": [ { "name": "file", "type": "string", "required": true } ... ]
#   },
#   ...
# ]
```

### `serve` — Agent stdio loop

Reads JSON commands from stdin (one per line) and writes JSON responses to stdout.  
Designed for agent orchestrators and future MCP integration.

```bash
echo '{"command":"info","file":"data.xlsx"}' | excel-cli serve
```

Request format:

```json
{ "command": "info", "file": "data.xlsx" }
{ "command": "read", "file": "data.xlsx", "sheet": "Sales", "limit": 10, "json": true }
{ "command": "search", "file": "data.xlsx", "sheet": "Sales", "term": "Total", "regex": true }
```

Supported fields: `command`, `file`, `sheet`, `cell`, `term`, `range`, `limit`, `headerRow`, `json`, `jsonCompact`, `ndjson`, `regex`, `quiet`.

---

## Options Reference

| Flag               | Short | Description                                           |
| ------------------ | ----- | ----------------------------------------------------- |
| `--json`           | `-j`  | JSON output, indented (pretty)                        |
| `--json-compact`   |       | JSON output, single line                              |
| `--ndjson`         |       | Newline-delimited JSON — one row per line (streaming) |
| `--range <range>`  |       | Cell range, e.g. `A1:C10` (for `read` command)        |
| `--limit <n>`      |       | Max number of rows to display                         |
| `--header-row <n>` |       | Use row N as header (1-based); data starts at N+1     |
| `--regex`          |       | Treat search term as .NET regular expression          |
| `--quiet`          | `-q`  | Suppress all human output on stderr                   |
| `--no-color`       |       | Disable ANSI colours on stderr                        |

---

## stdout vs stderr

| Channel    | Content                                                 |
| ---------- | ------------------------------------------------------- |
| **stdout** | JSON data only — clean, deterministic, machine-readable |
| **stderr** | Tables, colours, progress info, human messages          |

```bash
# Pipe JSON to a file while seeing the human UX
excel-cli read data.xlsx Sales --json > data.json

# Suppress human output entirely
excel-cli read data.xlsx Sales --json --quiet

# Redirect stderr to a log
excel-cli read data.xlsx Sales --limit 10 2> human.log
```

---

## JSON Envelope

Every JSON response follows a stable envelope contract:

```jsonc
{
  "schemaVersion": "1.0",        // envelope structure version
  "toolVersion": "1.0.0+abc123", // assembly version + git hash
  "command": "read",             // command that was executed
  "success": true,               // boolean: did it succeed?
  "data": { ... },               // command-specific payload (null on error)
  "warnings": [],                // non-fatal messages
  "errorCode": null,             // e.g. "MISSING_ARGUMENT", "SHEET_NOT_FOUND"
  "message": null                // human-readable error description
}
```

### Error codes

| Code               | Meaning                                             |
| ------------------ | --------------------------------------------------- |
| `MISSING_ARGUMENT` | Required positional argument not provided           |
| `INVALID_ARGUMENT` | Flag value could not be parsed (e.g. `--limit abc`) |
| `UNKNOWN_COMMAND`  | Command name not recognised                         |
| `SHEET_NOT_FOUND`  | Requested sheet does not exist                      |
| `UNHANDLED_ERROR`  | Unexpected exception                                |

Exit codes: `0` success, `1` unhandled error, `2` invalid arguments.

---

## NDJSON Streaming

With `--ndjson`, output is streamed as **newline-delimited JSON** — one line per record:

```bash
excel-cli read data.xlsx Sales --ndjson --limit 3
```

```jsonc
{"schemaVersion":"1.0","command":"read","success":true,"count":150,"displayed":3}
{"ID":1,"Name":"João","Value":1500.0}
{"ID":2,"Name":"Maria","Value":2300.0}
{"ID":3,"Name":"Pedro","Value":900.0}
```

The **first line** is the envelope metadata; subsequent lines are data rows.  
This format is ideal for streaming parsers and piping to tools like `jq`.

---

## Header Fallback

When a sheet has empty or missing header cells, ExcelCLI generates positional column names: `ColA`, `ColB`, `ColC`, etc.  
Duplicate headers are automatically deduplicated: `Total`, `Total_2`, `Total_3`.

You can override header detection with `--header-row <n>` to pick any row as the header.

---

## For LLMs / Agents

1. Always use `--json` (or `--json-compact` / `--ndjson`) for structured output.
2. Parse `success` first; if `false`, read `errorCode` and `message`.
3. Use `excel-cli tools` to discover available commands and parameters.
4. Use `serve` for multi-turn conversation without spawning a process per request.
5. Cell type information is available via the `cell` command in JSON mode (`type`, `raw`).
6. Empty headers are replaced with `ColA`/`ColB`/… — always stable keys.

### Tool Discovery

```bash
excel-cli tools | jq '.[].name'
# "info"
# "inspect"
# "read"
# "cell"
# "search"
# "formulas"
```

---

## Building from Source

Requirements: **.NET 10.0+ SDK**

```bash
# Debug build
dotnet build

# Release build
dotnet build -c Release

# Single self-contained .exe
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
# → bin\Release\net10.0\win-x64\publish\ExcelCLI.exe
```

## Dependencies

| Package                    | Purpose                                   |
| -------------------------- | ----------------------------------------- |
| **ClosedXML** 0.105.0      | Excel file reading (.xlsx / .xlsm)        |
| **Spectre.Console** 0.54.0 | Rich terminal tables, colours, formatting |

## License

MIT License

## Contributing

Issues and PRs welcome!
