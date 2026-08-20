using System.IO.Compression;
using System.Text;
using ClosedXML.Excel;
using ExcelConverter.Api.Contracts;
using ExcelDataReader;

namespace ExcelConverter.Api.Services;

public sealed class WorkbookConverter(ILogger<WorkbookConverter> logger)
{
    private static readonly byte[] OleSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    private static readonly char[] InvalidSheetNameCharacters = ['\\', '/', '?', '*', '[', ']', ':'];

    public Task<ConversionResult> ConvertAsync(
        ConvertToXlsxRequest request,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var format = ResolveFormat(request);
        var source = DecodeContent(request.ContentBase64);
        ValidateFileSignature(source, format);

        var startedAt = DateTime.UtcNow;
        using var sourceStream = new MemoryStream(source, writable: false);
        using var reader = CreateReader(sourceStream, format, request);
        using var workbook = new XLWorkbook();

        var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sheetNumber = 0;

        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheetNumber++;
            var sheetName = CreateUniqueSheetName(reader.Name, sheetNumber, usedSheetNames);
            var worksheet = workbook.Worksheets.Add(sheetName);
            var rowNumber = 0;

            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                rowNumber++;
                if (rowNumber > ConverterOptions.MaxExcelRows)
                {
                    throw BadRequest("excel_row_limit", $"Worksheet '{sheetName}' exceeds the XLSX row limit.");
                }

                if (reader.FieldCount > ConverterOptions.MaxExcelColumns)
                {
                    throw BadRequest("excel_column_limit", $"Worksheet '{sheetName}' exceeds the XLSX column limit.");
                }

                for (var column = 0; column < reader.FieldCount; column++)
                {
                    WriteValue(worksheet.Cell(rowNumber, column + 1), reader.GetValue(column));
                }
            }
        }
        while (reader.NextResult());

        if (sheetNumber == 0)
        {
            workbook.Worksheets.Add("Sheet1");
        }

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        var result = output.ToArray();

        logger.LogInformation(
            "Converted {Format} workbook: {InputBytes} input bytes, {OutputBytes} output bytes, {SheetCount} sheets, {ElapsedMs} ms",
            format,
            source.LongLength,
            result.LongLength,
            sheetNumber,
            (DateTime.UtcNow - startedAt).TotalMilliseconds);

        return Task.FromResult(new ConversionResult(CreateOutputFileName(request.FileName), result));
    }

    private static IExcelDataReader CreateReader(
        Stream stream,
        SourceFileFormat format,
        ConvertToXlsxRequest request)
    {
        try
        {
            var configuration = new ExcelReaderConfiguration
            {
                FallbackEncoding = ResolveCsvEncoding(request.CsvEncoding),
                AutodetectSeparators = ResolveCsvSeparators(request.CsvDelimiter),
                LeaveOpen = false
            };

            return format == SourceFileFormat.Csv
                ? ExcelReaderFactory.CreateCsvReader(stream, configuration)
                : ExcelReaderFactory.CreateReader(stream, configuration);
        }
        catch (DecoderFallbackException)
        {
            throw new ConversionException(
                "csv_encoding_invalid",
                "The CSV content cannot be decoded with the selected encoding.",
                StatusCodes.Status422UnprocessableEntity);
        }
        catch (Exception exception) when (
            exception is InvalidDataException or
            NotSupportedException or
            ArgumentException)
        {
            throw new ConversionException(
                "workbook_invalid",
                "The source workbook is damaged, encrypted, or unsupported.",
                StatusCodes.Status422UnprocessableEntity);
        }
    }

    private static SourceFileFormat ResolveFormat(ConvertToXlsxRequest request)
    {
        var candidate = request.SourceFormat;
        if (string.IsNullOrWhiteSpace(candidate))
        {
            candidate = Path.GetExtension(request.FileName ?? string.Empty).TrimStart('.');
        }

        return candidate?.Trim().ToLowerInvariant() switch
        {
            "csv" => SourceFileFormat.Csv,
            "xls" => SourceFileFormat.Xls,
            "xlsb" => SourceFileFormat.Xlsb,
            _ => throw BadRequest(
                "source_format_unsupported",
                "SourceFormat or FileName must identify a CSV, XLS, or XLSB file.")
        };
    }

    private static byte[] DecodeContent(string? contentBase64)
    {
        if (string.IsNullOrWhiteSpace(contentBase64))
        {
            throw BadRequest("content_required", "ContentBase64 is required.");
        }

        var maximumEncodedLength = ((ConverterOptions.MaxSourceBytes + 2L) / 3L) * 4L;
        if (contentBase64.Length > maximumEncodedLength)
        {
            throw TooLarge();
        }

        try
        {
            var bytes = Convert.FromBase64String(contentBase64);
            if (bytes.Length == 0)
            {
                throw BadRequest("file_empty", "The source file is empty.");
            }

            if (bytes.Length > ConverterOptions.MaxSourceBytes)
            {
                throw TooLarge();
            }

            return bytes;
        }
        catch (FormatException)
        {
            throw BadRequest("base64_invalid", "ContentBase64 is not valid Base64 data.");
        }
    }

    private static void ValidateFileSignature(byte[] source, SourceFileFormat format)
    {
        var isOle = source.AsSpan().StartsWith(OleSignature);
        var isZip = source.Length >= 4 &&
                    source[0] == 0x50 &&
                    source[1] == 0x4B &&
                    source[2] is 0x03 or 0x05 or 0x07 &&
                    source[3] is 0x04 or 0x06 or 0x08;

        if (format == SourceFileFormat.Xls && !isOle)
        {
            throw BadRequest("format_mismatch", "The content is not an XLS binary workbook.");
        }

        if (format == SourceFileFormat.Xlsb && (!isZip || !ContainsXlsbWorkbookPart(source)))
        {
            throw BadRequest("format_mismatch", "The content is not an XLSB binary workbook.");
        }

        if (format == SourceFileFormat.Csv && (isOle || isZip))
        {
            throw BadRequest("format_mismatch", "Binary Excel content cannot be processed as CSV.");
        }
    }

    private static bool ContainsXlsbWorkbookPart(byte[] source)
    {
        try
        {
            using var stream = new MemoryStream(source, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            return archive.Entries.Any(entry =>
                entry.FullName.Equals("xl/workbook.bin", StringComparison.OrdinalIgnoreCase));
        }
        catch (InvalidDataException)
        {
            return false;
        }
    }

    private static Encoding ResolveCsvEncoding(string? encodingName)
    {
        if (string.IsNullOrWhiteSpace(encodingName))
        {
            return Encoding.GetEncoding("GB18030");
        }

        var normalized = encodingName.Trim().Replace("_", "-", StringComparison.Ordinal).ToLowerInvariant();
        try
        {
            return normalized switch
            {
                "utf-8" or "utf8" => new UTF8Encoding(false, true),
                "utf-16" or "utf16" or "unicode" or "utf-16le" => new UnicodeEncoding(false, true, true),
                "utf-16be" => new UnicodeEncoding(true, true, true),
                "gb18030" or "gb-18030" => Encoding.GetEncoding(
                    "GB18030",
                    EncoderFallback.ExceptionFallback,
                    DecoderFallback.ExceptionFallback),
                _ => throw BadRequest(
                    "csv_encoding_unsupported",
                    "CsvEncoding must be UTF-8, UTF-16, UTF-16BE, or GB18030.")
            };
        }
        catch (ArgumentException)
        {
            throw BadRequest(
                "csv_encoding_unsupported",
                "CsvEncoding must be UTF-8, UTF-16, UTF-16BE, or GB18030.");
        }
    }

    private static char[] ResolveCsvSeparators(string? delimiter)
    {
        if (string.IsNullOrEmpty(delimiter))
        {
            return [',', ';', '\t', '|'];
        }

        if (delimiter.Equals("\\t", StringComparison.Ordinal))
        {
            return ['\t'];
        }

        if (delimiter.Length != 1)
        {
            throw BadRequest("csv_delimiter_invalid", "CsvDelimiter must contain exactly one character or \\\\t.");
        }

        return [delimiter[0]];
    }

    private static string CreateUniqueSheetName(
        string? sourceName,
        int sheetNumber,
        ISet<string> usedNames)
    {
        var name = string.IsNullOrWhiteSpace(sourceName) ? $"Sheet{sheetNumber}" : sourceName.Trim();
        foreach (var invalidCharacter in InvalidSheetNameCharacters)
        {
            name = name.Replace(invalidCharacter, '_');
        }

        name = name.Trim('\'');
        if (string.IsNullOrWhiteSpace(name))
        {
            name = $"Sheet{sheetNumber}";
        }

        name = name[..Math.Min(name.Length, 31)];
        var baseName = name;
        var suffix = 2;
        while (!usedNames.Add(name))
        {
            var suffixText = $" ({suffix++})";
            name = baseName[..Math.Min(baseName.Length, 31 - suffixText.Length)] + suffixText;
        }

        return name;
    }

    private static void WriteValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null:
            case DBNull:
                break;
            case string text:
                cell.Value = text;
                break;
            case bool boolean:
                cell.Value = boolean;
                break;
            case DateTime dateTime:
                cell.Value = dateTime;
                break;
            case TimeSpan timeSpan:
                cell.Value = timeSpan;
                break;
            case byte number:
                cell.Value = number;
                break;
            case short number:
                cell.Value = number;
                break;
            case int number:
                cell.Value = number;
                break;
            case long number:
                cell.Value = number;
                break;
            case float number:
                cell.Value = number;
                break;
            case double number:
                cell.Value = number;
                break;
            case decimal number:
                cell.Value = number;
                break;
            default:
                cell.Value = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                break;
        }
    }

    private static string CreateOutputFileName(string? sourceFileName)
    {
        var baseName = Path.GetFileNameWithoutExtension(sourceFileName);
        if (string.IsNullOrWhiteSpace(baseName))
        {
            baseName = "converted";
        }

        foreach (var invalidCharacter in Path.GetInvalidFileNameChars())
        {
            baseName = baseName.Replace(invalidCharacter, '_');
        }

        return $"{baseName}.xlsx";
    }

    private static ConversionException BadRequest(string code, string message) =>
        new(code, message, StatusCodes.Status400BadRequest);

    private static ConversionException TooLarge() =>
        new(
            "file_too_large",
            $"The decoded source file exceeds the {ConverterOptions.MaxSourceBytes / 1024 / 1024} MB limit.",
            StatusCodes.Status413PayloadTooLarge);
}
