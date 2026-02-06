using System.Text.Json.Serialization;

namespace ExcelCLI.Formatters;

public class JsonEnvelope<T>
{
    [JsonPropertyOrder(1)]
    public string SchemaVersion { get; init; } = "1.0";

    [JsonPropertyOrder(2)]
    public string ToolVersion { get; init; } = "0.0.0";

    [JsonPropertyOrder(3)]
    public string Command { get; init; } = string.Empty;

    [JsonPropertyOrder(4)]
    public bool Success { get; init; }

    [JsonPropertyOrder(5)]
    public T? Data { get; init; }

    [JsonPropertyOrder(6)]
    public List<string> Warnings { get; init; } = new();

    [JsonPropertyOrder(7)]
    public string? ErrorCode { get; init; }

    [JsonPropertyOrder(8)]
    public string? Message { get; init; }
}
