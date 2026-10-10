using System.Net;
using System.Net.Http.Json;
using ExcelConverter.Api.Contracts;
using Microsoft.AspNetCore.Mvc.Testing;

namespace ExcelConverter.Api.Tests;

public sealed class ApiContractTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly HttpClient client;

    public ApiContractTests(WebApplicationFactory<Program> factory)
    {
        client = factory.CreateClient();
    }

    [Fact]
    public async Task HealthEndpointReturnsOk()
    {
        var response = await client.GetAsync("/health", CancellationToken.None);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
    }

    [Fact]
    public async Task ConvertEndpointReturnsXlsxContract()
    {
        var request = new ConvertToXlsxRequest(
            "sample.csv",
            null,
            Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("A,B\n1,2")));

        var response = await client.PostAsJsonAsync(
            "/api/v1/convert-to-xlsx",
            request,
            CancellationToken.None);
        var result = await response.Content.ReadFromJsonAsync<ConvertToXlsxResponse>(
            cancellationToken: CancellationToken.None);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.NotNull(result);
        Assert.Equal("sample.xlsx", result.FileName);
        Assert.NotEmpty(Convert.FromBase64String(result.ContentBase64));
        Assert.False(string.IsNullOrWhiteSpace(result.CorrelationId));
    }

    [Fact]
    public async Task ConvertEndpointReturnsStableErrorContract()
    {
        var request = new ConvertToXlsxRequest("sample.csv", null, "invalid!");

        var response = await client.PostAsJsonAsync(
            "/api/v1/convert-to-xlsx",
            request,
            CancellationToken.None);
        var error = await response.Content.ReadFromJsonAsync<ErrorResponse>(
            cancellationToken: CancellationToken.None);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal("base64_invalid", error?.Code);
        Assert.False(string.IsNullOrWhiteSpace(error?.CorrelationId));
    }
}
