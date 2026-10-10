namespace ExcelConverter.Api.Services;

public static class ConverterOptions
{
    public const int MaxSourceBytes = 10 * 1024 * 1024;
    public const long MaxJsonRequestBytes = 15 * 1024 * 1024;
    public const int MaxExcelRows = 1_048_576;
    public const int MaxExcelColumns = 16_384;
}
