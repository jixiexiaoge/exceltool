# Excel 格式转换自定义连接器

这是一个面向 Power Automate 云端流的 Excel 格式转换工具：

- 输入：CSV、XLS 或 XLSB 文件的 Base64 数据。
- 输出：标准 XLSX 文件的 Base64 数据。
- 保留：工作表和单元格值。
- 不保留：样式、公式表达式、批注、图表、宏及其他工作簿对象。

转换逻辑运行在自托管 ASP.NET Core `.NET 8` 服务中。Power Automate 自定义连接器负责调用服务，不使用连接器自定义 C# 脚本，因为该脚本运行时不能加载 ExcelDataReader、ClosedXML 等第三方程序集，无法可靠解析 XLS/XLSB。

## 重要限制

服务端允许解码后的源文件最大为 **10 MB**，但微软当前公布的 Power Automate 自定义连接器网关请求上限是 **3,182,218 字节**。Base64 还会增加约三分之一体积，因此通过自定义连接器单次调用时，建议将原始文件控制在 **2 MB 以内**，并为 JSON 字段留出空间。

因此：

- 直接调用自托管 API：源文件上限 10 MB。
- 通过 Power Automate 自定义连接器：实际建议上限 2 MB。
- 若必须在云端流中处理 2–10 MB 文件，需要改为文件 URL/对象存储传递或分块协议；单次 Base64 请求无法绕过连接器网关限制。

相关微软文档：

- [从 OpenAPI 定义创建自定义连接器](https://learn.microsoft.com/zh-cn/connectors/custom-connectors/define-openapi-definition)
- [使用 paconn CLI](https://learn.microsoft.com/zh-cn/connectors/custom-connectors/paconn-cli)
- [编写自定义连接器代码及限制](https://learn.microsoft.com/en-us/connectors/custom-connectors/write-code)
- [自定义连接器 FAQ 与请求大小限制](https://learn.microsoft.com/connectors/custom-connectors/faq)

## 架构

```text
Power Automate
    |
    | HTTPS JSON (Base64)
    v
Power Automate 自定义连接器
    |
    v
自托管 ExcelConverter.Api
    |-- ExcelDataReader: CSV / XLS / XLSB
    `-- ClosedXML: 创建 XLSX
```

服务未配置应用层鉴权。公网部署时必须在反向代理或网络基础设施中配置 TLS、访问控制和限流，否则任何可访问域名的人都能调用转换接口。

## API

`POST /api/v1/convert-to-xlsx`

请求示例：

```json
{
  "fileName": "source.xlsb",
  "sourceFormat": "xlsb",
  "contentBase64": "<BASE64>",
  "csvEncoding": null,
  "csvDelimiter": null
}
```

`sourceFormat` 可省略，服务会从 `fileName` 的 `.csv`、`.xls` 或 `.xlsb` 扩展名推断。CSV 支持：

- BOM/UTF-8 自动识别，默认使用 GB18030 回退。
- `csvEncoding` 可指定 `UTF-8`、`UTF-16`、`UTF-16BE` 或 `GB18030`。
- 默认自动检测逗号、分号、制表符和竖线。
- `csvDelimiter` 可指定单个字符，制表符可传 `\t`。

成功响应：

```json
{
  "fileName": "source.xlsx",
  "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "contentBase64": "<BASE64>",
  "sizeBytes": 12345,
  "correlationId": "..."
}
```

错误响应包含稳定的 `code`、可读的 `message` 和用于排查的 `correlationId`。服务会拒绝无效 Base64、空文件、格式与内容不匹配、损坏或加密工作簿、超过限制的文件，以及超过 XLSX 行列上限的数据。

## 本地运行

需要 [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)。

```powershell
dotnet restore
dotnet run --project .\src\ExcelConverter.Api
```

健康检查：

```powershell
Invoke-WebRequest http://localhost:5000/health
```

实际端口以 `dotnet run` 输出为准。

## Docker 部署

```powershell
docker build -t excel-converter:1.0 .
docker run --rm -p 8080:8080 excel-converter:1.0
```

在公网反向代理上完成 HTTPS 终止，并将请求转发到容器 `8080` 端口。建议：

- 只允许 Power Platform 所需网络来源或企业出口访问。
- 配置每 IP/客户端限流。
- 代理请求体上限至少为 15 MB。
- 代理超时高于服务的 110 秒转换超时。
- 不记录请求和响应中的 Base64 文件数据。

## 导入 Power Automate 自定义连接器

导入前，将
[`connector/apiDefinition.swagger.json`](connector/apiDefinition.swagger.json)
中的：

```json
"host": "excel-converter.example.com"
```

替换为实际公网 HTTPS 域名，不要包含 `https://` 或路径。

### Power Automate UI

1. 打开 Power Automate，进入 **数据 > 自定义连接器**。
2. 选择 **新建自定义连接器 > 导入 OpenAPI 文件**。
3. 上传 `connector/apiDefinition.swagger.json`。
4. 确认主机、基础 URL 和“Convert a file to XLSX”操作。
5. 安全性保持“无身份验证”。
6. 创建连接器，然后在“测试”页创建连接并测试。

### paconn CLI

安装并登录：

```powershell
pip install paconn
paconn login
```

验证 OpenAPI：

```powershell
paconn validate --api-def .\connector\apiDefinition.swagger.json
```

复制 `connector/settings.example.json` 为 `connector/settings.json`，填写 Power Platform 环境 GUID，然后创建：

```powershell
Set-Location .\connector
paconn create -s .\settings.json
```

创建成功后，将返回的连接器 ID 写入 `settings.json` 的 `connectorId`，后续更新：

```powershell
paconn update -s .\settings.json
```

`apiProperties.json` 已配置为无身份验证，项目不需要 `script.csx`。

## 在云端流中使用

1. 使用 SharePoint、OneDrive 等连接器获取源文件内容。
2. 调用自定义连接器的 **Convert a file to XLSX** 操作。
3. `fileName` 传入带扩展名的源文件名。
4. `contentBase64` 传入文件内容对应的 Base64。若动态内容已被连接器识别为二进制，可直接选择该文件内容；否则使用流设计器提供的 Base64 值。不要用表达式处理大于平台表达式限制的内容。
5. 使用 OneDrive 或 SharePoint 的“创建文件”操作：
   - 文件名：转换操作返回的 `fileName`
   - 文件内容：`base64ToBinary(body('ConvertToXlsx')?['contentBase64'])`

操作名称可能被设计器本地化，以实际步骤名替换表达式中的 `ConvertToXlsx`。

## 测试

```powershell
dotnet test .\ExcelConverter.sln
```

测试覆盖 UTF-8、UTF-16、GB18030 CSV，多种分隔符，XLS，多工作表，真实 XLSB 样本，输出 XLSX 可读性，API 契约和主要错误边界。

## 许可证说明

测试使用了 ExcelDataReader 项目 MIT 许可下的 `10x10.xlsb` 测试样本，来源说明位于
[`tests/ExcelConverter.Api.Tests/Fixtures/README.md`](tests/ExcelConverter.Api.Tests/Fixtures/README.md)。
