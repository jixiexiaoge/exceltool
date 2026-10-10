namespace ExcelConverter.Api.Contracts;

public sealed record ConvertToXlsxResponse(
    string FileName,
    string ContentType,
    string ContentBase64,
    long SizeBytes,
    string CorrelationId);
