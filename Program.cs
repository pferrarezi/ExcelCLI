using ExcelCLI.Services;
using ExcelCLI.Formatters;

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
        var parseResult = ParseOptions(args, command);

        if (!string.IsNullOrWhiteSpace(parseResult.Error))
        {
            OutputFormatter.WriteError(command, parseResult.ErrorCode ?? ErrorCodeInvalidArgument, parseResult.Error!, parseResult.Options.Json, parseResult.Options.Quiet);
            return ExitInvalidArguments;
        }

        var opts = parseResult.Options;

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
            OutputFormatter.WriteError(command, ErrorCodeUnhandled, ex.Message, opts.Json, opts.Quiet);
            return ExitUnhandledError;
        }
    }

    static int HandleInfo(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File))
        {
            OutputFormatter.WriteError("info", ErrorCodeMissingArgument, "Missing required argument: file", opts.Json, opts.Quiet);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var sheets = service.GetSheetNames(opts.File);
        OutputFormatter.PrintSheetList(sheets, opts.Json, opts.Quiet, "info");
        return ExitSuccess;
    }

    static int HandleInspect(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet))
        {
            OutputFormatter.WriteError("inspect", ErrorCodeMissingArgument, "Missing required arguments: file and sheet", opts.Json, opts.Quiet);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var dims = service.GetSheetDimensions(opts.File, opts.Sheet);
        OutputFormatter.PrintSheetInfo(opts.Sheet, dims, opts.Json, opts.Quiet, "inspect");
        return ExitSuccess;
    }

    static int HandleRead(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet))
        {
            OutputFormatter.WriteError("read", ErrorCodeMissingArgument, "Missing required arguments: file and sheet", opts.Json, opts.Quiet);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var data = service.ReadSheet(opts.File, opts.Sheet, opts.Range);
        OutputFormatter.PrintData(data, opts.Json, opts.Quiet, "read", opts.Limit);
        return ExitSuccess;
    }

    static int HandleCell(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet) || string.IsNullOrEmpty(opts.Cell))
        {
            OutputFormatter.WriteError("cell", ErrorCodeMissingArgument, "Missing required arguments: file, sheet, and cell", opts.Json, opts.Quiet);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var value = service.ReadCell(opts.File, opts.Sheet, opts.Cell);
        OutputFormatter.PrintCellValue(opts.Cell, value, opts.Json, opts.Quiet, "cell");
        return ExitSuccess;
    }

    static int HandleSearch(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet) || string.IsNullOrEmpty(opts.SearchTerm))
        {
            OutputFormatter.WriteError("search", ErrorCodeMissingArgument, "Missing required arguments: file, sheet, and term", opts.Json, opts.Quiet);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var results = service.SearchValue(opts.File, opts.Sheet, opts.SearchTerm);
        OutputFormatter.PrintSearchResults(results, opts.Json, opts.Quiet, "search");
        return ExitSuccess;
    }

    static int HandleFormulas(Options opts)
    {
        if (string.IsNullOrEmpty(opts.File) || string.IsNullOrEmpty(opts.Sheet))
        {
            OutputFormatter.WriteError("formulas", ErrorCodeMissingArgument, "Missing required arguments: file and sheet", opts.Json, opts.Quiet);
            return ExitInvalidArguments;
        }

        var service = new ExcelService();
        var formulas = service.GetAllFormulas(opts.File, opts.Sheet);
        OutputFormatter.PrintFormulas(formulas, opts.Json, opts.Quiet, "formulas");
        return ExitSuccess;
    }

    static int ShowHelp(bool quiet)
    {
        if (quiet)
            return ExitSuccess;

        Console.Error.WriteLine(@"

ExcelCLI - Universal Excel data extraction tool for humans and LLMs

USAGE:
    excel-cli <command> [arguments] [options]

COMMANDS:
    info <file>                List all sheets and file metadata
    inspect <file> <sheet>     Show sheet dimensions and preview
    read <file> <sheet>        Read sheet data as structured records
    cell <file> <sheet> <cell> Read a single cell value
    search <file> <sheet> <term>  Search for a value in a sheet
    formulas <file> <sheet>    List all cells with formulas

OPTIONS:
    --json, -j                 Output in JSON format (LLM-friendly)
    --range <range>            Cell range (e.g., A1:C10) for read command
    --limit <n>                Limit number of rows displayed
    --quiet, -q                Suppress progress output

EXAMPLES:
    excel-cli info data.xlsx
    excel-cli inspect data.xlsx Sheet1
    excel-cli read data.xlsx Sheet1 --limit 10
    excel-cli read data.xlsx Sheet1 --range A1:C10 --json
    excel-cli cell data.xlsx Sheet1 B5
    excel-cli search data.xlsx Sheet1 ""João""
    excel-cli formulas data.xlsx Sheet1 --json

");
        return ExitSuccess;
    }

    static int ShowError(string command, Options opts)
    {
        OutputFormatter.WriteError(command, ErrorCodeUnknownCommand, $"Unknown command: {command}", opts.Json, opts.Quiet);

        if (!opts.Json)
        {
            OutputFormatter.PrintHint("Use 'excel-cli help' for usage information.", opts.Quiet);
        }

        return ExitInvalidArguments;
    }

    static ParseResult ParseOptions(string[] args, string command)
    {
        var opts = new Options();
        string? error = null;
        string? errorCode = null;

        var positionals = new List<string>();

        for (int i = 1; i < args.Length; i++)
        {
            var arg = args[i];

            if (arg == "--json" || arg == "-j")
            {
                opts.Json = true;
            }
            else if (arg == "--quiet" || arg == "-q")
            {
                opts.Quiet = true;
            }
        }

        for (int i = 1; i < args.Length; i++)
        {
            var arg = args[i];

            if (arg == "--json" || arg == "-j")
            {
                opts.Json = true;
            }
            else if (arg == "--quiet" || arg == "-q")
            {
                opts.Quiet = true;
            }
            else if (arg == "--range" && i + 1 < args.Length)
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
            else
            {
                positionals.Add(arg);
            }
        }

        if (string.IsNullOrEmpty(error))
        {
            AssignPositionals(opts, command, positionals);
        }

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
                opts.Cell = positionals.ElementAtOrDefault(2);
                opts.SearchTerm = positionals.ElementAtOrDefault(3);
                break;
        }
    }
}

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
    public bool Json { get; set; }
    public bool Quiet { get; set; }
}
