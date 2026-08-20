namespace ExcelConverter.Api.Contracts;

public sealed record ConvertToXlsxRequest(
    string? FileName,
    string? SourceFormat,
    string? ContentBase64,
    string? CsvEncoding = null,
    string? CsvDelimiter = null);
