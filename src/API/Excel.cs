using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SanChong.Excel.Core;
using SanChong.Excel.Core.Helpers;
using SanChong.Excel.Models;
using System.Collections;

namespace SanChong.Excel.API;

/// <summary>Generate Excel</summary>
public static class Excel

{
    /// <summary>Generate Excel with a single worksheet</summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="data">Data</param>
    /// <returns>Memory stream</returns>
    public static MemoryStream Create(IEnumerable data, string sheetName = "Sheet1")
    {
        var sheetDescriptor = CreateSheetDescriptor(sheetName, data);
        return Create(sheetDescriptor);
    }

    /// <summary>Generate Excel</summary>
    /// <param name="sheetDescriptors">Worksheet descriptors</param>
    /// <returns>Memory stream</returns>
    public static MemoryStream Create(params SheetDescriptor[] sheetDescriptors)
    {
        var ms = new MemoryStream();
        var doc = Create(ms, sheetDescriptors);
        doc.Dispose();
        ms.Seek(0, SeekOrigin.Begin);
        return ms;
    }

    /// <summary>Generate Excel</summary>
    /// <param name="stream">File storage location</param>
    /// <param name="sheetDescriptors">Worksheet descriptors</param>
    public static SpreadsheetDocument Create(Stream stream, params SheetDescriptor[] sheetDescriptors)
    {
        var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        UInt32Value sheetId = 1;
        foreach (var dto in sheetDescriptors)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();
            worksheetPart.Worksheet.AppendChild(dto.Columns);
            sheets.AppendChild(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = dto.Name
            });
            worksheetPart.Worksheet.AddChild(dto.Data);
        }

        // Check if there are cells formatted as Date without a style set
        if (sheetDescriptors.Any(
            x => x.Data.Descendants<Cell>().Any(
                x => x.DataType == CellValues.Date && x.StyleIndex == null)
            ))
        {
            // Set default date style
            doc.ApplyDefaultDateFormat(out UInt32Value cellStyleIndex);

            foreach (var sheetData in sheetDescriptors.Select(x => x.Data))
                foreach (var cell in sheetData.Descendants<Cell>().Where(x => x.DataType == CellValues.Date && x.StyleIndex == null))
                    cell.StyleIndex = cellStyleIndex;
        }

        return doc;
    }

    /// <summary>Create worksheet descriptor</summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="data">Data</param>
    /// <param name="renderTitleRow">Whether to render title row</param>
    /// <returns>Worksheet descriptor</returns>
    public static SheetDescriptor CreateSheetDescriptor(string sheetName, IEnumerable data, bool renderTitleRow = true)
    {
        var creator = renderTitleRow
            ? new SheetDescriptorCreator(new TitleDataSheetCreator(data))
                : new SheetDescriptorCreator(new DataSheetCreator(data));
        return creator.Create(sheetName);
    }

    /// <summary>Set cell value</summary>
    /// <param name="doc">Excel document</param>
    /// <param name="rowIndex">Row index (zero-based)</param>
    /// <param name="columnIndex">Column index (zero-based)</param>
    /// <param name="value">Value to set</param>
    /// <param name="sheetIndex">Worksheet index (zero-based)</param>
    public static void BindCellValue<T>(this SpreadsheetDocument doc, uint rowIndex, uint columnIndex, T value, int sheetIndex = 0)
    {
        SheetEditor.BindCellValue(doc, rowIndex, columnIndex, value, sheetIndex);
    }

    /// <summary>Set cell value</summary>
    /// <typeparam name="T">Type of value</typeparam>
    /// <param name="doc">Excel document</param>
    /// <param name="rowIndex">Row index (zero-based)</param>
    /// <param name="columnIndex">Column index (zero-based)</param>
    /// <param name="value">Value to set</param>
    /// <param name="sheetName">Worksheet name</param>
    public static void BindCellValue<T>(this SpreadsheetDocument doc, string sheetName, uint rowIndex, uint columnIndex, T value)
    {
        SheetEditor.BindCellValue(doc, sheetName, rowIndex, columnIndex, value);
    }

    // <summary>Set cell value</summary>
    /// <typeparam name="T">Type of value</typeparam>
    /// <param name="doc">Excel document</param>
    /// <param name="sheet">Worksheet</param>
    /// <param name="rowIndex">Row index (zero-based)</param>
    /// <param name="columnIndex">Column index (zero-based)</param>
    /// <param name="value">Value to set</param>
    private static void BindCellValue<T>(this SpreadsheetDocument doc, Sheet sheet, uint rowIndex, uint columnIndex, T value)
    {
        SheetEditor.BindCellValue(doc, sheet, rowIndex, columnIndex, value);
    }

    /// <summary>Set cell value</summary>
    /// <param name="doc">Excel document</param>
    /// <param name="reference">Position e.g. A1, B3,...</param>
    /// <param name="value">Value to set</param>
    /// <param name="sheetIndex">Worksheet index (zero-based)</param>
    public static void BindCellValue<T>(this SpreadsheetDocument doc, string reference, T value, int sheetIndex = 0)
    {
        SheetEditor.BindCellValue(doc, reference, value, sheetIndex);
    }

    /// <summary>Set cell value</summary>
    /// <param name="doc">Excel document</param>
    /// <param name="reference">Position</param>
    /// <param name="value">Value to set</param>
    /// <param name="sheetName">Worksheet name</param>
    public static void BindCellValue<T>(this SpreadsheetDocument doc, string reference, T value, string sheetName)
    {
        SheetEditor.BindCellValue(doc, reference, value, sheetName);
    }

    /// <summary>Get cell from Excel document</summary>
    /// <param name="doc">Excel document</param>
    /// <param name="sheetIndex">Worksheet index (zero-based)</param>
    /// <param name="rowIndex">Row index (zero-based)</param>
    /// <param name="columnIndex">Column index (zero-based)</param>
    /// <returns></returns>
    public static Cell GetCell(this SpreadsheetDocument doc, int rowIndex, int columnIndex, int sheetIndex = 0)
    {
        return CellLocator.GetCell(doc, rowIndex, columnIndex, sheetIndex);
    }

    /// <summary>Find cell</summary>
    /// <param name="worksheet">Worksheet</param>
    /// <param name="columnIndex">Column index (zero-based)</param>
    /// <param name="rowIndex">Row index (zero-based)</param>
    /// <param name="cell">Found cell</param>
    /// <returns></returns>
    public static bool FindCell(this Worksheet worksheet, uint columnIndex, uint rowIndex, out Cell cell)
        => CellLocator.FindCell(worksheet, columnIndex, rowIndex, out cell);
}
