using ExcelCLI.Services;
using ExcelCLI.Formatters;
using System.Text.Json;

namespace ExcelCLI;

class Program
{
    private const int ExitSuccess = 0;
    private const int ExitUnhandledError = 1;
    private const int ExitInvalidArguments = 2;

    private const string ErrorCodeMissingArgument = "MISSING_ARGUMENT";
    private const string ErrorCodeInvalidArgument = "INVALID_ARGUMENT";
    private const string ErrorCodeUnknownCommand = "UNKNOWN_COMMAND";
    private const string ErrorCodeUnhandled = "UNHANDLED_ERROR";

    static int Main(string[] args)
    {
        if (args.Length == 0)
        {
            ShowHelp(false);
            return ExitSuccess;
        }

        var command = args[0].ToLowerInvariant();

        // Fast-path meta commands (no file needed)
        if (command is "--tools" or "tools")
        {
            OutputFormatter.PrintToolManifest();
            return ExitSuccess;
        }

        var parseResult = ParseOptions(args, command);

        if (!string.IsNullOrWhiteSpace(parseResult.Error))
        {
            OutputFormatter.WriteError(command, parseResult.ErrorCode ?? ErrorCodeInvalidArgument, parseResult.Error!, parseResult.Options.JsonMode, parseResult.Options.Quiet, parseResult.Options.NoColor);
            return ExitInvalidArguments;
        }

        var opts = parseResult.Options;

        // serve --stdio loop
        if (command is "serve")
            return RunStdioLoop(opts);

        try
        {
            return command switch
            {
                "info" => HandleInfo(opts),
                "inspect" => HandleInspect(opts),
                "read" => HandleRead(opts),
                "cell" => HandleCell(opts),
                "search" => HandleSearch(opts),
                "formulas" => HandleFormulas(opts),
                "help" or "--help" or "-h" => ShowHelp(opts.Quiet),
                _ => ShowError(command, opts)
            };
        }
        catch (Exception ex)
        {
            OutputFormatter.WriteError(command, ErrorCodeUnhandled, ex.Message, opts.JsonMode, opts.Quiet, opts.NoColor);
            return ExitUnhandledError;
        }
    }

    // ── Command handlers ────────────────────────────────────────────────

    static int HandleInfo(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File))
        {
            OutputFormatter.WriteError("info", ErrorCodeMissingArgument, "Missing required argument: file", opts.JsonMode, opts.Quiet, opts.NoColor);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var sheets = service.GetSheetNames(opts.File);
        OutputFormatter.PrintSheetList(sheets, opts.JsonMode, opts.Quiet, opts.NoColor, "info");
        return ExitSuccess;
    }

    static int HandleInspect(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet))
        {
            OutputFormatter.WriteError("inspect", ErrorCodeMissingArgument, "Missing required arguments: file and sheet", opts.JsonMode, opts.Quiet, opts.NoColor);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var dims = service.GetSheetDimensions(opts.File, opts.Sheet);
        OutputFormatter.PrintSheetInfo(opts.Sheet, dims, opts.JsonMode, opts.Quiet, opts.NoColor, "inspect");
        return ExitSuccess;
    }

    static int HandleRead(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet))
        {
            OutputFormatter.WriteError("read", ErrorCodeMissingArgument, "Missing required arguments: file and sheet", opts.JsonMode, opts.Quiet, opts.NoColor);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var data = service.ReadSheet(opts.File, opts.Sheet, opts.Range, opts.HeaderRow);
        OutputFormatter.PrintData(data, opts.JsonMode, opts.Quiet, opts.NoColor, "read", opts.Limit);
        return ExitSuccess;
    }

    static int HandleCell(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet) || string.IsNullOrEmpty(opts.Cell))
        {
            OutputFormatter.WriteError("cell", ErrorCodeMissingArgument, "Missing required arguments: file, sheet, and cell", opts.JsonMode, opts.Quiet, opts.NoColor);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();

        // JSON mode returns rich type info; human mode returns simple value
        if (opts.JsonMode != JsonMode.Off)
        {
            var info = service.ReadCellInfo(opts.File, opts.Sheet, opts.Cell);
            OutputFormatter.PrintCellInfo(info, opts.JsonMode, opts.Quiet, opts.NoColor, "cell");
        }
        else
        {
            var value = service.ReadCell(opts.File, opts.Sheet, opts.Cell);
            OutputFormatter.PrintCellValue(opts.Cell, value, opts.JsonMode, opts.Quiet, opts.NoColor, "cell");
        }
        return ExitSuccess;
    }

    static int HandleSearch(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet) || string.IsNullOrEmpty(opts.SearchTerm))
        {
            OutputFormatter.WriteError("search", ErrorCodeMissingArgument, "Missing required arguments: file, sheet, and term", opts.JsonMode, opts.Quiet, opts.NoColor);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var results = service.SearchValue(opts.File, opts.Sheet, opts.SearchTerm, opts.Regex);
        OutputFormatter.PrintSearchResults(results, opts.JsonMode, opts.Quiet, opts.NoColor, "search");
        return ExitSuccess;
    }

    static int HandleFormulas(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet))
        {
            OutputFormatter.WriteError("formulas", ErrorCodeMissingArgument, "Missing required arguments: file and sheet", opts.JsonMode, opts.Quiet, opts.NoColor);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var formulas = service.GetAllFormulas(opts.File, opts.Sheet);
        OutputFormatter.PrintFormulas(formulas, opts.JsonMode, opts.Quiet, opts.NoColor, "formulas");
        return ExitSuccess;
    }

    // ── serve --stdio ───────────────────────────────────────────────────

    static int RunStdioLoop(Options baseOpts)
    {
        // Acknowledge startup
        Console.Error.WriteLine("[excel-cli] serve mode ready. Send JSON commands on stdin, one per line.");

        var jsonOpts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

        while (true)
        {
            var line = Console.In.ReadLine();
            if (line == null) break; // EOF

            line = line.Trim();
            if (line.Length == 0) continue;

            StdioRequest? request;
            try
            {
                request = JsonSerializer.Deserialize<StdioRequest>(line, jsonOpts);
            }
            catch
            {
                OutputFormatter.WriteError("serve", ErrorCodeInvalidArgument, "Invalid JSON request", JsonMode.Compact, true);
                continue;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Command))
            {
                OutputFormatter.WriteError("serve", ErrorCodeMissingArgument, "Missing 'command' field", JsonMode.Compact, true);
                continue;
            }

            var syntheticArgs = new List<string> { request.Command };
            if (!string.IsNullOrEmpty(request.File)) syntheticArgs.Add(request.File);
            if (!string.IsNullOrEmpty(request.Sheet)) syntheticArgs.Add(request.Sheet);

            // Third positional depends on command
            if (request.Command is "cell" && !string.IsNullOrEmpty(request.Cell))
                syntheticArgs.Add(request.Cell);
            else if (request.Command is "search" && !string.IsNullOrEmpty(request.Term))
                syntheticArgs.Add(request.Term);

            syntheticArgs.Add("--json-compact");
            syntheticArgs.Add("--quiet");

            if (request.Regex == true) syntheticArgs.Add("--regex");
            if (!string.IsNullOrEmpty(request.Range)) { syntheticArgs.Add("--range"); syntheticArgs.Add(request.Range); }
            if (request.Limit.HasValue) { syntheticArgs.Add("--limit"); syntheticArgs.Add(request.Limit.Value.ToString()); }
            if (request.HeaderRow.HasValue) { syntheticArgs.Add("--header-row"); syntheticArgs.Add(request.HeaderRow.Value.ToString()); }

            var cmd = request.Command.ToLowerInvariant();
            var parseResult = ParseOptions(syntheticArgs.ToArray(), cmd);
            var opts = parseResult.Options;

            if (!string.IsNullOrWhiteSpace(parseResult.Error))
            {
                OutputFormatter.WriteError(cmd, parseResult.ErrorCode ?? ErrorCodeInvalidArgument, parseResult.Error!, opts.JsonMode, opts.Quiet, opts.NoColor);
                continue;
            }

            try
            {
                _ = cmd switch
                {
                    "info" => HandleInfo(opts),
                    "inspect" => HandleInspect(opts),
                    "read" => HandleRead(opts),
                    "cell" => HandleCell(opts),
                    "search" => HandleSearch(opts),
                    "formulas" => HandleFormulas(opts),
                    _ => ShowError(cmd, opts)
                };
            }
            catch (Exception ex)
            {
                OutputFormatter.WriteError(cmd, ErrorCodeUnhandled, ex.Message, opts.JsonMode, opts.Quiet, opts.NoColor);
            }

            Console.Out.Flush();
        }

        return ExitSuccess;
    }

    // ── Help / error ────────────────────────────────────────────────────

    static int ShowHelp(bool quiet)
    {
        if (quiet)
            return ExitSuccess;

        Console.Error.WriteLine(@"
ExcelCLI - Universal Excel data extraction tool for humans and LLMs

USAGE:
    excel-cli <command> [arguments] [options]

COMMANDS:
    info <file>                   List all sheets and file metadata
    inspect <file> <sheet>        Show sheet dimensions and preview
    read <file> <sheet>           Read sheet data as structured records
    cell <file> <sheet> <cell>    Read a single cell value (with type info in JSON)
    search <file> <sheet> <term>  Search for a value in a sheet
    formulas <file> <sheet>       List all cells with formulas
    serve                         Start stdio loop for agent integration
    tools                         Print tool manifest (JSON)

OPTIONS:
    --json, -j                 JSON output (indented)
    --json-compact             JSON output (single line)
    --ndjson                   Newline-delimited JSON (streaming)
    --range <range>            Cell range (e.g., A1:C10) for read command
    --limit <n>                Limit number of rows displayed
    --header-row <n>           Use row N as header (1-based absolute row number)
    --regex                    Treat search term as regex pattern
    --quiet, -q                Suppress human output (stderr)
    --no-color                 Disable ANSI colors on stderr

EXAMPLES:
    excel-cli info data.xlsx
    excel-cli read data.xlsx Sheet1 --limit 10 --json
    excel-cli read data.xlsx Sheet1 --range A1:C10 --ndjson
    excel-cli read data.xlsx Sheet1 --header-row 3 --json
    excel-cli cell data.xlsx Sheet1 B5 --json
    excel-cli search data.xlsx Sheet1 ""João""
    excel-cli search data.xlsx Sheet1 ""Total.*"" --regex --json
    excel-cli formulas data.xlsx Sheet1 --json-compact
    excel-cli tools
    echo '{""command"":""info"",""file"":""data.xlsx""}' | excel-cli serve
");
        return ExitSuccess;
    }

    static int ShowError(string command, Options opts)
    {
        OutputFormatter.WriteError(command, ErrorCodeUnknownCommand, $"Unknown command: {command}", opts.JsonMode, opts.Quiet, opts.NoColor);

        if (opts.JsonMode == JsonMode.Off)
            OutputFormatter.PrintHint("Use 'excel-cli help' for usage information.", opts.Quiet, opts.NoColor);

        return ExitInvalidArguments;
    }

    // ── Option parsing ──────────────────────────────────────────────────

    static ParseResult ParseOptions(string[] args, string command)
    {
        var opts = new Options();
        string? error = null;
        string? errorCode = null;

        // Pre-scan for output mode flags (needed before errors)
        for (int i = 1; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg is "--json" or "-j") opts.JsonMode = JsonMode.Pretty;
            else if (arg == "--json-compact") opts.JsonMode = JsonMode.Compact;
            else if (arg == "--ndjson") opts.JsonMode = JsonMode.Ndjson;
            else if (arg is "--quiet" or "-q") opts.Quiet = true;
            else if (arg == "--no-color") opts.NoColor = true;
            else if (arg == "--regex") opts.Regex = true;
        }

        var positionals = new List<string>();

        for (int i = 1; i < args.Length; i++)
        {
            var arg = args[i];

            if (arg is "--json" or "-j" or "--json-compact" or "--ndjson" or "--quiet" or "-q" or "--no-color" or "--regex")
                continue; // already handled

            if (arg == "--range" && i + 1 < args.Length)
            {
                opts.Range = args[++i];
            }
            else if (arg == "--range")
            {
                error = "Missing value for --range";
                errorCode = ErrorCodeMissingArgument;
                break;
            }
            else if (arg == "--limit" && i + 1 < args.Length)
            {
                var raw = args[++i];
                if (!int.TryParse(raw, out var parsed) || parsed < 0)
                {
                    error = $"Invalid value for --limit: '{raw}'";
                    errorCode = ErrorCodeInvalidArgument;
                    break;
                }
                opts.Limit = parsed;
            }
            else if (arg == "--limit")
            {
                error = "Missing value for --limit";
                errorCode = ErrorCodeMissingArgument;
                break;
            }
            else if (arg == "--header-row" && i + 1 < args.Length)
            {
                var raw = args[++i];
                if (!int.TryParse(raw, out var parsed) || parsed < 1)
                {
                    error = $"Invalid value for --header-row: '{raw}' (must be >= 1)";
                    errorCode = ErrorCodeInvalidArgument;
                    break;
                }
                opts.HeaderRow = parsed;
            }
            else if (arg == "--header-row")
            {
                error = "Missing value for --header-row";
                errorCode = ErrorCodeMissingArgument;
                break;
            }
            else
            {
                positionals.Add(arg);
            }
        }

        if (string.IsNullOrEmpty(error))
            AssignPositionals(opts, command, positionals);

        return new ParseResult(opts, error, errorCode);
    }

    static void AssignPositionals(Options opts, string command, List<string> positionals)
    {
        switch (command)
        {
            case "info":
                opts.File = positionals.ElementAtOrDefault(0);
                break;
            case "inspect":
            case "read":
            case "formulas":
                opts.File = positionals.ElementAtOrDefault(0);
                opts.Sheet = positionals.ElementAtOrDefault(1);
                break;
            case "cell":
                opts.File = positionals.ElementAtOrDefault(0);
                opts.Sheet = positionals.ElementAtOrDefault(1);
                opts.Cell = positionals.ElementAtOrDefault(2);
                break;
            case "search":
                opts.File = positionals.ElementAtOrDefault(0);
                opts.Sheet = positionals.ElementAtOrDefault(1);
                opts.SearchTerm = positionals.ElementAtOrDefault(2);
                break;
            default:
                opts.File = positionals.ElementAtOrDefault(0);
                opts.Sheet = positionals.ElementAtOrDefault(1);
                break;
        }
    }
}

// ── Models ──────────────────────────────────────────────────────────────

class ParseResult
{
    public ParseResult(Options options, string? error, string? errorCode)
    {
        Options = options;
        Error = error;
        ErrorCode = errorCode;
    }

    public Options Options { get; }
    public string? Error { get; }
    public string? ErrorCode { get; }
}

class Options
{
    public string? File { get; set; }
    public string? Sheet { get; set; }
    public string? Cell { get; set; }
    public string? SearchTerm { get; set; }
    public string? Range { get; set; }
    public int? Limit { get; set; }
    public int? HeaderRow { get; set; }
    public JsonMode JsonMode { get; set; } = JsonMode.Off;
    public bool Quiet { get; set; }
    public bool NoColor { get; set; }
    public bool Regex { get; set; }
}

class StdioRequest
{
    public string? Command { get; set; }
    public string? File { get; set; }
    public string? Sheet { get; set; }
    public string? Cell { get; set; }
    public string? Term { get; set; }
    public string? Range { get; set; }
    public int? Limit { get; set; }
    public int? HeaderRow { get; set; }
    public bool? Regex { get; set; }
}
