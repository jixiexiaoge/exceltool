namespace ExcelConverter.Api.Contracts;

public sealed record ErrorResponse(
    string Code,
    string Message,
    string CorrelationId);
