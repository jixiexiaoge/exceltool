FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY ExcelConverter.sln ./
COPY src/ExcelConverter.Api/ExcelConverter.Api.csproj src/ExcelConverter.Api/
RUN dotnet restore src/ExcelConverter.Api/ExcelConverter.Api.csproj
COPY src/ExcelConverter.Api/ src/ExcelConverter.Api/
RUN dotnet publish src/ExcelConverter.Api/ExcelConverter.Api.csproj -c Release -o /app --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final
WORKDIR /app
ENV ASPNETCORE_HTTP_PORTS=8080
EXPOSE 8080
USER $APP_UID
COPY --from=build /app .
ENTRYPOINT ["dotnet", "ExcelConverter.Api.dll"]
