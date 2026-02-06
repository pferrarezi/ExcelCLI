using ExcelCLI.Services;
using Spectre.Console;
using System.Globalization;
using System.Reflection;
using System.Text.Json;

namespace ExcelCLI.Formatters;

/// <summary>Output format mode for JSON serialization.</summary>
public enum JsonMode { Off, Pretty, Compact, Ndjson }

public static class OutputFormatter
{
    private const string SchemaVersion = "1.0";

    private static readonly JsonSerializerOptions JsonPrettyOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private static readonly JsonSerializerOptions JsonCompactOptions = new()
    {
        WriteIndented = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private static readonly string ToolVersion = GetToolVersion();

    private static IAnsiConsole? _errorConsole;

    /// <summary>Get or create the stderr console, optionally with no-color.</summary>
    public static IAnsiConsole GetErrorConsole(bool noColor = false)
    {
        if (_errorConsole == null || noColor)
        {
            var settings = new AnsiConsoleSettings
            {
                Out = new AnsiConsoleOutput(Console.Error)
            };
            if (noColor)
            {
                settings.ColorSystem = ColorSystemSupport.NoColors;
            }
            _errorConsole = AnsiConsole.Create(settings);
        }
        return _errorConsole;
    }

    // ── Sheet list ──────────────────────────────────────────────────────

    public static void PrintSheetList(List<string> sheets, JsonMode jsonMode, bool quiet, bool noColor, string command)
    {
        if (jsonMode != JsonMode.Off)
        {
            WriteJsonSuccess(command, new { sheets }, jsonMode: jsonMode);
            return;
        }

        if (quiet) return;

        var con = GetErrorConsole(noColor);
        var table = new Table();
        table.AddColumn(new TableColumn("[bold]Sheet Name[/]").Centered());
        foreach (var sheet in sheets) table.AddRow(sheet);
        con.Write(table);
        con.MarkupLine($"\n[green]Found {sheets.Count} sheet(s)[/]");
    }

    // ── Sheet info / inspect ────────────────────────────────────────────

    public static void PrintSheetInfo(string sheetName, (int Rows, int Cols, string FirstCell, string LastCell) dims, JsonMode jsonMode, bool quiet, bool noColor, string command)
    {
        if (jsonMode != JsonMode.Off)
        {
            WriteJsonSuccess(command, new
            {
                sheet = sheetName,
                rows = dims.Rows,
                columns = dims.Cols,
                range = new { first = dims.FirstCell, last = dims.LastCell }
            }, jsonMode: jsonMode);
            return;
        }

        if (quiet) return;

        var con = GetErrorConsole(noColor);
        var table = new Table().Border(TableBorder.Rounded);
        table.AddColumn("[bold]Property[/]");
        table.AddColumn("[bold]Value[/]");
        table.AddRow("Sheet Name", sheetName);
        table.AddRow("Rows", dims.Rows.ToString());
        table.AddRow("Columns", dims.Cols.ToString());
        table.AddRow("First Cell", dims.FirstCell);
        table.AddRow("Last Cell", dims.LastCell);
        con.Write(table);
    }

    // ── Read data ───────────────────────────────────────────────────────

    public static void PrintData(List<Dictionary<string, object?>> data, JsonMode jsonMode, bool quiet, bool noColor, string command, int? limit = null)
    {
        var displayData = limit.HasValue ? data.Take(limit.Value).ToList() : data;

        if (jsonMode != JsonMode.Off)
        {
            var warnings = new List<string>();
            if (limit.HasValue && data.Count > limit.Value)
                warnings.Add($"Showing {displayData.Count} of {data.Count} rows (use --limit to show more)");

            if (jsonMode == JsonMode.Ndjson)
            {
                // NDJSON: one JSON object per line, no envelope per row
                var envelope = BuildSuccessEnvelope(command, new { count = data.Count, displayed = displayData.Count }, warnings);
                Console.Out.WriteLine(JsonSerializer.Serialize(envelope, JsonCompactOptions));
                foreach (var row in displayData)
                    Console.Out.WriteLine(JsonSerializer.Serialize(row, JsonCompactOptions));
                return;
            }

            WriteJsonSuccess(command, new
            {
                rows = displayData,
                count = data.Count,
                displayed = displayData.Count
            }, warnings, jsonMode);
            return;
        }

        if (quiet) return;

        var con = GetErrorConsole(noColor);

        if (data.Count == 0)
        {
            con.MarkupLine("[yellow]No data found[/]");
            return;
        }

        var table = new Table().Border(TableBorder.Rounded);
        var headers = data[0].Keys.ToList();
        foreach (var header in headers) table.AddColumn(new TableColumn($"[bold]{header}[/]"));
        foreach (var row in displayData)
        {
            var values = headers.Select(h => FormatValue(row.GetValueOrDefault(h))).ToArray();
            table.AddRow(values);
        }
        con.Write(table);

        if (limit.HasValue && data.Count > limit.Value)
            con.MarkupLine($"\n[yellow]Showing {displayData.Count} of {data.Count} rows (use --limit to show more)[/]");
        else
            con.MarkupLine($"\n[green]Total: {data.Count} row(s)[/]");
    }

    // ── Search results ──────────────────────────────────────────────────

    public static void PrintSearchResults(List<(string Cell, object? Value)> results, JsonMode jsonMode, bool quiet, bool noColor, string command)
    {
        if (jsonMode != JsonMode.Off)
        {
            WriteJsonSuccess(command, new
            {
                results = results.Select(r => new { cell = r.Cell, value = r.Value }).ToList(),
                count = results.Count
            }, jsonMode: jsonMode);
            return;
        }

        if (quiet) return;

        var con = GetErrorConsole(noColor);

        if (results.Count == 0)
        {
            con.MarkupLine("[yellow]No matches found[/]");
            return;
        }

        var table = new Table();
        table.AddColumn("[bold]Cell[/]");
        table.AddColumn("[bold]Value[/]");
        foreach (var (cell, value) in results) table.AddRow(cell, FormatValue(value));
        con.Write(table);
        con.MarkupLine($"\n[green]Found {results.Count} match(es)[/]");
    }

    // ── Formulas ────────────────────────────────────────────────────────

    public static void PrintFormulas(List<(string Cell, string Formula, object? Value)> formulas, JsonMode jsonMode, bool quiet, bool noColor, string command)
    {
        if (jsonMode != JsonMode.Off)
        {
            WriteJsonSuccess(command, new
            {
                formulas = formulas.Select(f => new { cell = f.Cell, formula = f.Formula, value = f.Value }).ToList(),
                count = formulas.Count
            }, jsonMode: jsonMode);
            return;
        }

        if (quiet) return;

        var con = GetErrorConsole(noColor);

        if (formulas.Count == 0)
        {
            con.MarkupLine("[yellow]No formulas found[/]");
            return;
        }

        var table = new Table();
        table.AddColumn("[bold]Cell[/]");
        table.AddColumn("[bold]Formula[/]");
        table.AddColumn("[bold]Value[/]");
        foreach (var (cell, formula, value) in formulas)
            table.AddRow(cell, $"[cyan]{formula}[/]", FormatValue(value));
        con.Write(table);
        con.MarkupLine($"\n[green]Found {formulas.Count} formula(s)[/]");
    }

    // ── Cell value ──────────────────────────────────────────────────────

    public static void PrintCellValue(string cell, object? value, JsonMode jsonMode, bool quiet, bool noColor, string command)
    {
        if (jsonMode != JsonMode.Off)
        {
            WriteJsonSuccess(command, new { cell, value }, jsonMode: jsonMode);
            return;
        }

        PrintSuccess($"Cell {cell}: {FormatValue(value)}", quiet, noColor);
    }

    public static void PrintCellInfo(CellInfo info, JsonMode jsonMode, bool quiet, bool noColor, string command)
    {
        if (jsonMode != JsonMode.Off)
        {
            WriteJsonSuccess(command, new { cell = info.Address, value = info.Value, type = info.Type, raw = info.Raw }, jsonMode: jsonMode);
            return;
        }

        PrintSuccess($"Cell {info.Address} [{info.Type}]: {FormatValue(info.Value)} (raw: {info.Raw})", quiet, noColor);
    }

    // ── Errors / helpers ────────────────────────────────────────────────

    public static void WriteError(string command, string errorCode, string message, JsonMode jsonMode, bool quiet, bool noColor = false)
    {
        if (jsonMode != JsonMode.Off)
            WriteJsonError(command, errorCode, message, jsonMode);

        if (!quiet)
            GetErrorConsole(noColor).MarkupLine($"[red]Error:[/] {message}");
    }

    public static void PrintSuccess(string message, bool quiet, bool noColor = false)
    {
        if (quiet) return;
        GetErrorConsole(noColor).MarkupLine($"[green]{message}[/]");
    }

    public static void PrintHint(string message, bool quiet, bool noColor = false)
    {
        if (quiet) return;
        GetErrorConsole(noColor).MarkupLine(message);
    }

    // ── Tool manifest ───────────────────────────────────────────────────

    public static void PrintToolManifest()
    {
        var manifest = new
        {
            name = "excel-cli",
            version = ToolVersion,
            schemaVersion = SchemaVersion,
            description = "Universal Excel data extraction tool for humans and LLMs",
            tools = new object[]
            {
                new {
                    name = "info",
                    description = "List all sheets in a workbook",
                    parameters = new object[] {
                        new { name = "file", type = "string", required = true, description = "Path to .xlsx/.xlsm file" }
                    }
                },
                new {
                    name = "inspect",
                    description = "Show sheet dimensions and range",
                    parameters = new object[] {
                        new { name = "file", type = "string", required = true, description = "Path to .xlsx/.xlsm file" },
                        new { name = "sheet", type = "string", required = true, description = "Sheet name" }
                    }
                },
                new {
                    name = "read",
                    description = "Read sheet data as structured records",
                    parameters = new object[] {
                        new { name = "file", type = "string", required = true, description = "Path to .xlsx/.xlsm file" },
                        new { name = "sheet", type = "string", required = true, description = "Sheet name" },
                        new { name = "--range", type = "string", required = false, description = "Cell range (e.g. A1:C10)" },
                        new { name = "--limit", type = "integer", required = false, description = "Max rows to return" },
                        new { name = "--header-row", type = "integer", required = false, description = "Row number to use as header" }
                    }
                },
                new {
                    name = "cell",
                    description = "Read a single cell value with type info",
                    parameters = new object[] {
                        new { name = "file", type = "string", required = true, description = "Path to .xlsx/.xlsm file" },
                        new { name = "sheet", type = "string", required = true, description = "Sheet name" },
                        new { name = "cell", type = "string", required = true, description = "Cell address (e.g. B5)" }
                    }
                },
                new {
                    name = "search",
                    description = "Search for a value across all cells in a sheet",
                    parameters = new object[] {
                        new { name = "file", type = "string", required = true, description = "Path to .xlsx/.xlsm file" },
                        new { name = "sheet", type = "string", required = true, description = "Sheet name" },
                        new { name = "term", type = "string", required = true, description = "Search term (substring or regex with --regex)" },
                        new { name = "--regex", type = "boolean", required = false, description = "Treat term as regex pattern" }
                    }
                },
                new {
                    name = "formulas",
                    description = "List all cells with formulas in a sheet",
                    parameters = new object[] {
                        new { name = "file", type = "string", required = true, description = "Path to .xlsx/.xlsm file" },
                        new { name = "sheet", type = "string", required = true, description = "Sheet name" }
                    }
                }
            },
            outputModes = new[] { "--json", "--json-compact", "--ndjson" },
            globalOptions = new[] { "--quiet", "--no-color" }
        };

        Console.Out.WriteLine(JsonSerializer.Serialize(manifest, JsonPrettyOptions));
    }

    // ── JSON internals ──────────────────────────────────────────────────

    private static JsonEnvelope<T> BuildSuccessEnvelope<T>(string command, T data, List<string>? warnings = null)
    {
        return new JsonEnvelope<T>
        {
            SchemaVersion = SchemaVersion,
            ToolVersion = ToolVersion,
            Command = command,
            Success = true,
            Data = data,
            Warnings = warnings ?? new List<string>(),
            ErrorCode = null,
            Message = null
        };
    }

    private static void WriteJsonSuccess<T>(string command, T data, List<string>? warnings = null, JsonMode jsonMode = JsonMode.Pretty)
    {
        var envelope = BuildSuccessEnvelope(command, data, warnings);
        var opts = jsonMode == JsonMode.Compact ? JsonCompactOptions : JsonPrettyOptions;
        Console.Out.WriteLine(JsonSerializer.Serialize(envelope, opts));
    }

    private static void WriteJsonError(string command, string errorCode, string message, JsonMode jsonMode = JsonMode.Pretty)
    {
        var envelope = new JsonEnvelope<object>
        {
            SchemaVersion = SchemaVersion,
            ToolVersion = ToolVersion,
            Command = command,
            Success = false,
            Data = null,
            Warnings = new List<string>(),
            ErrorCode = errorCode,
            Message = message
        };

        var opts = jsonMode == JsonMode.Compact ? JsonCompactOptions : JsonPrettyOptions;
        Console.Out.WriteLine(JsonSerializer.Serialize(envelope, opts));
    }

    private static string FormatValue(object? value)
    {
        return value switch
        {
            null => "[dim]null[/]",
            double d => d.ToString("N2", CultureInfo.CurrentCulture),
            DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture),
            TimeSpan ts => ts.ToString(),
            bool b => b ? "[green]true[/]" : "[red]false[/]",
            _ => value.ToString() ?? "[dim]null[/]"
        };
    }

    private static string GetToolVersion()
    {
        var assembly = Assembly.GetEntryAssembly();
        if (assembly == null)
            return "0.0.0";

        var info = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
        if (!string.IsNullOrWhiteSpace(info?.InformationalVersion))
            return info!.InformationalVersion;

        return assembly.GetName().Version?.ToString() ?? "0.0.0";
    }
}
