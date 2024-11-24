# SanChong.Excel
A lightweight library built on OpenXML SDK to easily generate Excel files.

## Features
- Generate Excel files from Dapper queries or any `IEnumerable<dynamic>`.
- Automatically create headers from property names.
- Support multiple sheets with simple APIs.

## Installation
Currently available on GitHub. NuGet release coming soon.

## Usage
```csharp
var data = Conn.Query("SELECT Role, Account, Name FROM Users;");
var excelStream = Excel.Create(data);
