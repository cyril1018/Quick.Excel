
    using System.IO;
    using Xunit;
    using SanChong.Excel.API;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

public class ExcelTests
{
    [Fact]
    public void Create_ShouldGenerateExcelWithSingleWorksheet()
    {
        // Arrange
        var data = new[]
        {
            new { Name = "Alice", Age = 30 },
            new { Name = "Bob", Age = 25 }
        };
        string sheetName = "TestSheet";

        // Act
        var stream = Excel.Create(data, sheetName);

        // Assert
        Assert.NotNull(stream); // Ensure the stream is created
        Assert.True(stream.Length > 0); // Ensure the stream contains data

        // Validate the content
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var sheets = document.WorkbookPart.Workbook.Sheets;
            Assert.Single(sheets); // Ensure there is only one worksheet
            Assert.Equal(sheetName, sheets.GetFirstChild<Sheet>().Name); // Check sheet name
        }
    }
}

