namespace ExcelConverter.Api.Services;

public sealed class ConversionException : Exception
{
    public ConversionException(string code, string message, int statusCode)
        : base(message)
    {
        Code = code;
        StatusCode = statusCode;
    }

    public string Code { get; }

    public int StatusCode { get; }
}
