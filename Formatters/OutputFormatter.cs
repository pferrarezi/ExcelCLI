using Spectre.Console;
using System.Globalization;
using System.Reflection;
using System.Text.Json;

namespace ExcelCLI.Formatters;

public static class OutputFormatter
{
    private const string SchemaVersion = "1.0";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private static readonly IAnsiConsole ErrorConsole = AnsiConsole.Create(new AnsiConsoleSettings
    {
        Out = new AnsiConsoleOutput(Console.Error)
    });

    private static readonly string ToolVersion = GetToolVersion();

    public static void PrintSheetList(List<string> sheets, bool asJson, bool quiet, string command)
    {
        if (asJson)
        {
            WriteJsonSuccess(command, new { sheets });
            return;
        }

        if (quiet)
            return;

        var table = new Table();
        table.AddColumn(new TableColumn("[bold]Sheet Name[/]").Centered());

        foreach (var sheet in sheets)
        {
            table.AddRow(sheet);
        }

        ErrorConsole.Write(table);
        ErrorConsole.MarkupLine($"\n[green]Found {sheets.Count} sheet(s)[/]");
    }

    public static void PrintSheetInfo(string sheetName, (int Rows, int Cols, string FirstCell, string LastCell) dimensions, bool asJson, bool quiet, string command)
    {
        if (asJson)
        {
            WriteJsonSuccess(command, new
            {
                sheet = sheetName,
                rows = dimensions.Rows,
                columns = dimensions.Cols,
                range = new { first = dimensions.FirstCell, last = dimensions.LastCell }
            });
            return;
        }

        if (quiet)
            return;

        var table = new Table();
        table.Border(TableBorder.Rounded);
        table.AddColumn("[bold]Property[/]");
        table.AddColumn("[bold]Value[/]");

        table.AddRow("Sheet Name", sheetName);
        table.AddRow("Rows", dimensions.Rows.ToString());
        table.AddRow("Columns", dimensions.Cols.ToString());
        table.AddRow("First Cell", dimensions.FirstCell);
        table.AddRow("Last Cell", dimensions.LastCell);

        ErrorConsole.Write(table);
    }

    public static void PrintData(List<Dictionary<string, object?>> data, bool asJson, bool quiet, string command, int? limit = null)
    {
        var displayData = limit.HasValue ? data.Take(limit.Value).ToList() : data;

        if (asJson)
        {
            var warnings = new List<string>();
            if (limit.HasValue && data.Count > limit.Value)
            {
                warnings.Add($"Showing {displayData.Count} of {data.Count} rows (use --limit to show more)");
            }

            WriteJsonSuccess(command, new
            {
                rows = displayData,
                count = data.Count,
                displayed = displayData.Count
            }, warnings);
            return;
        }

        if (quiet)
            return;

        if (data.Count == 0)
        {
            ErrorConsole.MarkupLine("[yellow]No data found[/]");
            return;
        }

        var table = new Table();
        table.Border(TableBorder.Rounded);

        // Headers
        var headers = data[0].Keys.ToList();
        foreach (var header in headers)
        {
            table.AddColumn(new TableColumn($"[bold]{header}[/]"));
        }

        // Rows
        foreach (var row in displayData)
        {
            var values = headers.Select(h => FormatValue(row.GetValueOrDefault(h))).ToArray();
            table.AddRow(values);
        }

        ErrorConsole.Write(table);

        if (limit.HasValue && data.Count > limit.Value)
        {
            ErrorConsole.MarkupLine($"\n[yellow]Showing {displayData.Count} of {data.Count} rows (use --limit to show more)[/]");
        }
        else
        {
            ErrorConsole.MarkupLine($"\n[green]Total: {data.Count} row(s)[/]");
        }
    }

    public static void PrintSearchResults(List<(string Cell, object? Value)> results, bool asJson, bool quiet, string command)
    {
        if (asJson)
        {
            WriteJsonSuccess(command, new
            {
                results = results.Select(r => new { cell = r.Cell, value = r.Value }).ToList(),
                count = results.Count
            });
            return;
        }

        if (quiet)
            return;

        if (results.Count == 0)
        {
            ErrorConsole.MarkupLine("[yellow]No matches found[/]");
            return;
        }

        var table = new Table();
        table.AddColumn("[bold]Cell[/]");
        table.AddColumn("[bold]Value[/]");

        foreach (var (cell, value) in results)
        {
            table.AddRow(cell, FormatValue(value));
        }

        ErrorConsole.Write(table);
        ErrorConsole.MarkupLine($"\n[green]Found {results.Count} match(es)[/]");
    }

    public static void PrintFormulas(List<(string Cell, string Formula, object? Value)> formulas, bool asJson, bool quiet, string command)
    {
        if (asJson)
        {
            WriteJsonSuccess(command, new
            {
                formulas = formulas.Select(f => new { cell = f.Cell, formula = f.Formula, value = f.Value }).ToList(),
                count = formulas.Count
            });
            return;
        }

        if (quiet)
            return;

        if (formulas.Count == 0)
        {
            ErrorConsole.MarkupLine("[yellow]No formulas found[/]");
            return;
        }

        var table = new Table();
        table.AddColumn("[bold]Cell[/]");
        table.AddColumn("[bold]Formula[/]");
        table.AddColumn("[bold]Value[/]");

        foreach (var (cell, formula, value) in formulas)
        {
            table.AddRow(cell, $"[cyan]{formula}[/]", FormatValue(value));
        }

        ErrorConsole.Write(table);
        ErrorConsole.MarkupLine($"\n[green]Found {formulas.Count} formula(s)[/]");
    }

    public static void WriteError(string command, string errorCode, string message, bool asJson, bool quiet)
    {
        if (asJson)
        {
            WriteJsonError(command, errorCode, message);
        }

        if (!quiet)
        {
            ErrorConsole.MarkupLine($"[red]Error:[/] {message}");
        }
    }

    public static void PrintSuccess(string message, bool quiet)
    {
        if (quiet)
            return;

        ErrorConsole.MarkupLine($"[green]{message}[/]");
    }

    public static void PrintHint(string message, bool quiet)
    {
        if (quiet)
            return;

        ErrorConsole.MarkupLine(message);
    }

    public static void PrintCellValue(string cell, object? value, bool asJson, bool quiet, string command)
    {
        if (asJson)
        {
            WriteJsonSuccess(command, new { cell, value });
            return;
        }

        PrintSuccess($"Cell {cell}: {FormatValue(value)}", quiet);
    }

    private static void WriteJsonSuccess<T>(string command, T data, List<string>? warnings = null)
    {
        var envelope = new JsonEnvelope<T>
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

        Console.Out.WriteLine(JsonSerializer.Serialize(envelope, JsonOptions));
    }

    private static void WriteJsonError(string command, string errorCode, string message)
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

        Console.Out.WriteLine(JsonSerializer.Serialize(envelope, JsonOptions));
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
