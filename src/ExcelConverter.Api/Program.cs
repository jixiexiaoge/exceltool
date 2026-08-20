using System.Text;
using ExcelConverter.Api.Contracts;
using ExcelConverter.Api.Services;
using Microsoft.AspNetCore.Diagnostics;
using Microsoft.AspNetCore.Http.Features;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = ConverterOptions.MaxJsonRequestBytes;
});
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = ConverterOptions.MaxJsonRequestBytes;
});
builder.Services.AddProblemDetails();
builder.Services.AddHealthChecks();
builder.Services.AddRequestTimeouts(options =>
{
    options.DefaultPolicy = new()
    {
        Timeout = TimeSpan.FromSeconds(110),
        TimeoutStatusCode = StatusCodes.Status504GatewayTimeout,
        WriteTimeoutResponse = async context =>
        {
            context.Response.ContentType = "application/json";
            await context.Response.WriteAsJsonAsync(
                new ErrorResponse(
                    "conversion_timeout",
                    "The conversion did not finish within 110 seconds.",
                    context.TraceIdentifier));
        }
    };
});
builder.Services.AddSingleton<WorkbookConverter>();

var app = builder.Build();

app.UseExceptionHandler(errorApp =>
{
    errorApp.Run(async context =>
    {
        var exception = context.Features.Get<IExceptionHandlerFeature>()?.Error;
        var conversionException = exception as ConversionException;
        var status = conversionException?.StatusCode ?? StatusCodes.Status500InternalServerError;
        var code = conversionException?.Code ?? "internal_error";
        var message = conversionException?.Message ?? "An unexpected conversion error occurred.";

        context.Response.StatusCode = status;
        context.Response.ContentType = "application/json";
        await context.Response.WriteAsJsonAsync(
            new ErrorResponse(code, message, context.TraceIdentifier),
            context.RequestAborted);
    });
});

app.UseRequestTimeouts();
app.MapHealthChecks("/health");

app.MapPost("/api/v1/convert-to-xlsx", async (
    ConvertToXlsxRequest request,
    WorkbookConverter converter,
    HttpContext context,
    CancellationToken cancellationToken) =>
{
    var result = await converter.ConvertAsync(request, cancellationToken);
    return Results.Ok(new ConvertToXlsxResponse(
        result.FileName,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        Convert.ToBase64String(result.Content),
        result.Content.LongLength,
        context.TraceIdentifier));
});

app.Run();

public partial class Program;
