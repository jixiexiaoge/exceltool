using System.Text;
using ClosedXML.Excel;
using ExcelConverter.Api.Contracts;
using ExcelConverter.Api.Services;
using Microsoft.Extensions.Logging.Abstractions;
using NPOI.HSSF.UserModel;

namespace ExcelConverter.Api.Tests;

public sealed class WorkbookConverterTests
{
    private readonly WorkbookConverter converter = new(NullLogger<WorkbookConverter>.Instance);

    static WorkbookConverterTests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Theory]
    [InlineData("UTF-8", ",", "姓名,数量\r\n测试,12", "UTF-8")]
    [InlineData("UTF-16", ";", "姓名;数量\r\n测试;12", "UTF-16")]
    [InlineData("GB18030", "|", "姓名|数量\r\n测试|12", "GB18030")]
    public async Task ConvertsCsvEncodingsAndDelimiters(
        string encodingName,
        string delimiter,
        string csv,
        string requestEncoding)
    {
        var encoding = encodingName switch
        {
            "UTF-8" => new UTF8Encoding(true),
            "UTF-16" => new UnicodeEncoding(false, true),
            _ => Encoding.GetEncoding("GB18030")
        };
        var request = CreateRequest("data.csv", encoding.GetBytes(csv), requestEncoding, delimiter);

        var result = await converter.ConvertAsync(request, CancellationToken.None);

        using var workbook = OpenWorkbook(result.Content);
        var sheet = workbook.Worksheet(1);
        Assert.Equal("姓名", sheet.Cell(1, 1).GetString());
        Assert.Equal("数量", sheet.Cell(1, 2).GetString());
        Assert.Equal("测试", sheet.Cell(2, 1).GetString());
        Assert.Equal("12", sheet.Cell(2, 2).GetString());
        Assert.Equal("data.xlsx", result.FileName);
    }

    [Fact]
    public async Task AutoDetectsCsvDelimiter()
    {
        var request = CreateRequest("pipe.csv", Encoding.UTF8.GetBytes("A|B\none|two"));

        var result = await converter.ConvertAsync(request, CancellationToken.None);

        using var workbook = OpenWorkbook(result.Content);
        Assert.Equal("two", workbook.Worksheet(1).Cell(2, 2).GetString());
    }

    [Fact]
    public async Task ConvertsXlsWithMultipleSheetsAndValues()
    {
        byte[] source;
        using (var stream = new MemoryStream())
        {
            var sourceWorkbook = new HSSFWorkbook();
            var first = sourceWorkbook.CreateSheet("First");
            var row = first.CreateRow(0);
            row.CreateCell(0).SetCellValue("text");
            row.CreateCell(1).SetCellValue(42.5);
            sourceWorkbook.CreateSheet("Second").CreateRow(0).CreateCell(0).SetCellValue("value");
            sourceWorkbook.Write(stream, leaveOpen: true);
            sourceWorkbook.Close();
            source = stream.ToArray();
        }

        var result = await converter.ConvertAsync(
            CreateRequest("legacy.xls", source),
            CancellationToken.None);

        using var workbook = OpenWorkbook(result.Content);
        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.Equal("text", workbook.Worksheet("First").Cell(1, 1).GetString());
        Assert.Equal(42.5, workbook.Worksheet("First").Cell(1, 2).GetDouble());
        Assert.Equal("value", workbook.Worksheet("Second").Cell(1, 1).GetString());
    }

    [Fact]
    public async Task ConvertsXlsbFixture()
    {
        var fixture = await File.ReadAllBytesAsync(
            Path.Combine(AppContext.BaseDirectory, "Fixtures", "10x10.xlsb"),
            CancellationToken.None);

        var result = await converter.ConvertAsync(
            CreateRequest("binary.xlsb", fixture),
            CancellationToken.None);

        using var workbook = OpenWorkbook(result.Content);
        Assert.NotEmpty(workbook.Worksheets);
        Assert.True(workbook.Worksheet(1).LastRowUsed()?.RowNumber() > 0);
        Assert.Equal("binary.xlsx", result.FileName);
    }

    [Fact]
    public async Task RejectsInvalidBase64()
    {
        var exception = await Assert.ThrowsAsync<ConversionException>(() =>
            converter.ConvertAsync(
                new ConvertToXlsxRequest("bad.csv", null, "not-base64!"),
                CancellationToken.None));

        Assert.Equal("base64_invalid", exception.Code);
        Assert.Equal(400, exception.StatusCode);
    }

    [Fact]
    public async Task RejectsFormatMismatch()
    {
        var exception = await Assert.ThrowsAsync<ConversionException>(() =>
            converter.ConvertAsync(
                CreateRequest("wrong.xls", Encoding.UTF8.GetBytes("a,b")),
                CancellationToken.None));

        Assert.Equal("format_mismatch", exception.Code);
    }

    [Fact]
    public async Task RejectsFilesLargerThanLimit()
    {
        var content = new byte[ConverterOptions.MaxSourceBytes + 1];

        var exception = await Assert.ThrowsAsync<ConversionException>(() =>
            converter.ConvertAsync(
                CreateRequest("large.csv", content),
                CancellationToken.None));

        Assert.Equal("file_too_large", exception.Code);
        Assert.Equal(413, exception.StatusCode);
    }

    [Fact]
    public async Task RejectsInvalidCsvDelimiterWithStableError()
    {
        var request = CreateRequest(
            "data.csv",
            Encoding.UTF8.GetBytes("a,b"),
            csvDelimiter: ",;");

        var exception = await Assert.ThrowsAsync<ConversionException>(() =>
            converter.ConvertAsync(request, CancellationToken.None));

        Assert.Equal("csv_delimiter_invalid", exception.Code);
        Assert.Equal(400, exception.StatusCode);
    }

    private static ConvertToXlsxRequest CreateRequest(
        string fileName,
        byte[] content,
        string? csvEncoding = null,
        string? csvDelimiter = null) =>
        new(fileName, null, Convert.ToBase64String(content), csvEncoding, csvDelimiter);

    private static XLWorkbook OpenWorkbook(byte[] content) =>
        new(new MemoryStream(content, writable: false));
}
