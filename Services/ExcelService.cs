using ClosedXML.Excel;
using System.Globalization;

namespace ExcelCLI.Services;

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
            usedRange.FirstCell().Address.ToString(),
            usedRange.LastCell().Address.ToString()
        );
    }

    public List<Dictionary<string, object?>> ReadSheet(string filePath, string sheetName, string? range = null)
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

        // Primeira linha como cabeçalho
        var headers = rows[0].Cells().Select(c => c.GetString()).ToList();
        var result = new List<Dictionary<string, object?>>();

        foreach (var row in rows.Skip(1))
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

    public object? ReadCell(string filePath, string sheetName, string cellAddress)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var cell = worksheet.Cell(cellAddress);

        return GetCellValue(cell);
    }

    public List<(string Cell, object? Value)> SearchValue(string filePath, string sheetName, string searchTerm)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = GetWorksheet(workbook, sheetName);
        var usedRange = worksheet.RangeUsed();

        if (usedRange == null)
            return new List<(string, object?)>();

        var results = new List<(string Cell, object? Value)>();

        foreach (var cell in usedRange.CellsUsed())
        {
            var stringValue = cell.GetString();
            var typedValue = GetCellValue(cell);
            var typedString = ConvertToInvariantString(typedValue);

            if (ContainsSearchTerm(stringValue, searchTerm) || ContainsSearchTerm(typedString, searchTerm))
            {
                results.Add((cell.Address.ToString(), typedValue));
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
            .Select(c => (c.Address.ToString(), c.FormulaA1, GetCellValue(c)))
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

    private static bool ContainsSearchTerm(string? value, string searchTerm)
    {
        if (string.IsNullOrEmpty(value))
            return false;

        return value.Contains(searchTerm, StringComparison.OrdinalIgnoreCase);
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
