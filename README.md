# SanChong.Excel
A lightweight library built on OpenXML SDK to easily generate Excel files.

## Features
- Generate Excel files from Dapper queries or any `IEnumerable<dynamic>`.
- Automatically create headers from property names.
- Support multiple sheets with simple APIs.

## Installation
Currently available on GitHub.

**Note**: If you need this project to be available on NuGet, feel free to let me know, and I’ll publish it.

## Usage
```csharp
var data = Conn.Query("SELECT Role, Account, Name FROM Users;");
var excelStream = Excel.Create(data);
