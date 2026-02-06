using ClosedXML.Excel;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelCLI.Services;

public record CellInfo(string Address, object? Value, string Type, string? Raw);

public class ExcelService
{
    public List<string> GetSheetNames(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);
        return workbook.Worksheets.Select(ws => ws.Name).ToList();
    }

    public (int Rows, int Cols, string FirstCell, string LastCell) GetSheetDimensions(string filePath, string sheetName)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var usedRange = worksheet.RangeUsed();

        if (usedRange == null)
            return (0, 0, string.Empty, string.Empty);

        return (
            usedRange.RowCount(),
            usedRange.ColumnCount(),
            usedRange.FirstCell().Address.ToString()!,
            usedRange.LastCell().Address.ToString()!
        );
    }

    public List<Dictionary<string, object?>> ReadSheet(string filePath, string sheetName, string? range = null, int? headerRow = null)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);

        var dataRange = string.IsNullOrEmpty(range)
            ? worksheet.RangeUsed()
            : worksheet.Range(range);

        if (dataRange == null)
            return new List<Dictionary<string, object?>>();

        var rows = dataRange.RowsUsed().ToList();
        if (rows.Count == 0)
            return new List<Dictionary<string, object?>>();

        // Determinar headers
        List<string> headers;

        if (headerRow.HasValue)
        {
            var absRow = headerRow.Value;
            headers = new List<string>();
            var firstCol = dataRange.FirstCell().Address.ColumnNumber;
            var lastCol = dataRange.LastCell().Address.ColumnNumber;
            for (int col = firstCol; col <= lastCol; col++)
            {
                var val = worksheet.Cell(absRow, col).GetString();
                headers.Add(string.IsNullOrWhiteSpace(val) ? GetColumnFallbackName(col) : val);
            }
            // Keep only rows after the header row
            rows = rows.Where(r => r.RowNumber() > absRow).ToList();
        }
        else
        {
            // Primeira linha como cabeçalho, com fallback para colunas vazias
            var rawHeaders = rows[0].Cells().Select(c => c.GetString()).ToList();
            headers = new List<string>();
            for (int i = 0; i < rawHeaders.Count; i++)
            {
                var h = rawHeaders[i];
                headers.Add(string.IsNullOrWhiteSpace(h)
                    ? GetColumnFallbackName(rows[0].Cell(i + 1).Address.ColumnNumber)
                    : h);
            }
            rows = rows.Skip(1).ToList();
        }

        // Desduplicar headers
        headers = DeduplicateHeaders(headers);

        var result = new List<Dictionary<string, object?>>();

        foreach (var row in rows)
        {
            var dict = new Dictionary<string, object?>();
            var cells = row.Cells().ToList();

            for (int i = 0; i < headers.Count && i < cells.Count; i++)
            {
                dict[headers[i]] = GetCellValue(cells[i]);
            }

            result.Add(dict);
        }

        return result;
    }

    public CellInfo ReadCellInfo(string filePath, string sheetName, string cellAddress)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var cell = worksheet.Cell(cellAddress);
        return BuildCellInfo(cell);
    }

    public object? ReadCell(string filePath, string sheetName, string cellAddress)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var cell = worksheet.Cell(cellAddress);
        return GetCellValue(cell);
    }

    public List<(string Cell, object? Value)> SearchValue(string filePath, string sheetName, string searchTerm, bool isRegex = false)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var usedRange = worksheet.RangeUsed();

        if (usedRange == null)
            return new List<(string, object?)>();

        Regex? regex = null;
        if (isRegex)
        {
            try { regex = new Regex(searchTerm, RegexOptions.IgnoreCase | RegexOptions.Compiled); }
            catch (ArgumentException ex)
            {
                throw new ArgumentException($"Invalid regex pattern: {ex.Message}");
            }
        }

        var results = new List<(string Cell, object? Value)>();

        foreach (var cell in usedRange.CellsUsed())
        {
            var stringValue = cell.GetString();
            var typedValue = GetCellValue(cell);
            var typedString = ConvertToInvariantString(typedValue);

            bool match = isRegex
                ? (MatchesRegex(regex!, stringValue) || MatchesRegex(regex!, typedString))
                : (ContainsSearchTerm(stringValue, searchTerm) || ContainsSearchTerm(typedString, searchTerm));

            if (match)
            {
                results.Add((cell.Address.ToString()!, typedValue));
            }
        }

        return results;
    }

    public (string? Formula, object? Value) GetFormula(string filePath, string sheetName, string cellAddress)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var cell = worksheet.Cell(cellAddress);

        var formula = cell.HasFormula ? cell.FormulaA1 : null;
        var value = GetCellValue(cell);

        return (formula, value);
    }

    public List<(string Cell, string Formula, object? Value)> GetAllFormulas(string filePath, string sheetName)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var usedRange = worksheet.RangeUsed();

        if (usedRange == null)
            return new List<(string Cell, string Formula, object? Value)>();

        return usedRange.CellsUsed()
            .Where(c => c.HasFormula)
            .Select(c => (c.Address.ToString()!, c.FormulaA1, GetCellValue(c)))
            .ToList();
    }

    private IXLWorksheet GetWorksheet(XLWorkbook workbook, string sheetName)
    {
        var worksheet = workbook.Worksheets.FirstOrDefault(ws =>
            ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

        if (worksheet == null)
            throw new ArgumentException($"Sheet '{sheetName}' not found. Available sheets: {string.Join(", ", workbook.Worksheets.Select(ws => ws.Name))}");

        return worksheet;
    }

    private CellInfo BuildCellInfo(IXLCell cell)
    {
        var address = cell.Address.ToString()!;
        if (cell.IsEmpty())
            return new CellInfo(address, null, "null", null);

        var value = GetCellValue(cell);
        var type = cell.DataType switch
        {
            XLDataType.Number => "number",
            XLDataType.Boolean => "boolean",
            XLDataType.DateTime => "datetime",
            XLDataType.TimeSpan => "timespan",
            _ => "string"
        };
        var raw = cell.GetString();

        return new CellInfo(address, value, type, raw);
    }

    private object? GetCellValue(IXLCell cell)
    {
        if (cell.IsEmpty())
            return null;

        return cell.DataType switch
        {
            XLDataType.Number => cell.GetDouble(),
            XLDataType.Boolean => cell.GetBoolean(),
            XLDataType.DateTime => cell.GetDateTime(),
            XLDataType.TimeSpan => cell.GetTimeSpan(),
            _ => cell.GetString()
        };
    }

    private static string GetColumnFallbackName(int columnNumber)
    {
        var name = string.Empty;
        var n = columnNumber;
        while (n > 0)
        {
            n--;
            name = (char)('A' + n % 26) + name;
            n /= 26;
        }
        return $"Col{name}";
    }

    private static List<string> DeduplicateHeaders(List<string> headers)
    {
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();

        foreach (var h in headers)
        {
            if (seen.TryGetValue(h, out var count))
            {
                seen[h] = count + 1;
                result.Add($"{h}_{count + 1}");
            }
            else
            {
                seen[h] = 1;
                result.Add(h);
            }
        }

        return result;
    }

    private static bool ContainsSearchTerm(string? value, string searchTerm)
    {
        if (string.IsNullOrEmpty(value))
            return false;

        return value.Contains(searchTerm, StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesRegex(Regex regex, string? value)
    {
        if (string.IsNullOrEmpty(value))
            return false;

        return regex.IsMatch(value);
    }

    private static string? ConvertToInvariantString(object? value)
    {
        if (value == null)
            return null;

        return value switch
        {
            DateTime dt => dt.ToString("O", CultureInfo.InvariantCulture),
            double d => d.ToString(CultureInfo.InvariantCulture),
            bool b => b ? "true" : "false",
            TimeSpan ts => ts.ToString(),
            _ => value.ToString()
        };
    }
}
