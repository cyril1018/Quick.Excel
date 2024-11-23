using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Quick.Excel.Core;
using Quick.Excel.Core.Helpers;
using Quick.Excel.Models;
using System.Collections;

namespace Quick.Excel.API;

/// <summary>產生 Excel</summary>
public static class QuickExcel

{
	/// <summary>產生單一工作表之 Excel</summary>
	/// <param name="sheetName">工作表名稱</param>
	/// <param name="data">資料</param
	/// <returns>記憶體串流</returns>
	public static MemoryStream Create(IEnumerable data, string sheetName = "Sheet1")
	{
		var sheetDescriptor = CreateSheetDescriptor(sheetName, data);
		return Create(sheetDescriptor);
	}

	/// <summary>產生 Excel</summary>
	/// <param name="sheetDescriptors">工作表描述</param>
	/// <returns>記憶體串流</returns>
	public static MemoryStream Create(params SheetDescriptor[] sheetDescriptors)
	{
		var ms = new MemoryStream();
		var doc = Create(ms, sheetDescriptors);
		doc.Dispose();
		ms.Seek(0, SeekOrigin.Begin);
		return ms;
	}

	/// <summary>產生 Excel</summary>
	/// <param name="stream">檔案存放位置</param>
	/// <param name="sheetDescriptors">工作表描述</param>
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

		// 是否有 儲存格格式為 日期 且未設定 樣式
		if (sheetDescriptors.Any(
			x => x.Data.Descendants<Cell>().Any(
				x => x.DataType == CellValues.Date && x.StyleIndex == null)
			))
		{
            // 設定預設日期樣式
            doc.ApplyDefaultDateFormat(out UInt32Value cellStyleIndex);

			foreach (var sheetData in sheetDescriptors.Select(x => x.Data))
				foreach (var cell in sheetData.Descendants<Cell>().Where(x => x.DataType == CellValues.Date && x.StyleIndex == null))
					cell.StyleIndex = cellStyleIndex;
		}

		return doc;
	}

	/// <summary>建立工作表描述</summary>
	/// <param name="sheetName">工作表名稱</param>
	/// <param name="data">資料</param>
	/// <param name="data">是否輸出標題列</param>
	/// <returns>工作表描述</returns>
	public static SheetDescriptor CreateSheetDescriptor(string sheetName, IEnumerable data, bool renderTitleRow = true)
	{
		var creator = renderTitleRow
			? new SheetDescriptorCreator(new TitleDataSheetCreator(data))
				: new SheetDescriptorCreator(new DataSheetCreator(data));
		return creator.Create(sheetName);
	}

	/// <summary>設定儲存格值</summary>
	/// <param name="doc">Excel 文件</param>
	/// <param name="rowIndex">列索引(從零開始)</param>
	/// <param name="columnIndex">欄索引(從零開始)</param>
	/// <param name="value">要設定的值</param>
	/// <param name="sheetIndex">工作表索引(從零開始)</param>
	public static void BindCellValue<T>(this SpreadsheetDocument doc, uint rowIndex, uint columnIndex, T value, int sheetIndex = 0)
	{
         SheetEditor.BindCellValue(doc, rowIndex, columnIndex, value, sheetIndex);
    }

	/// <summary>設定儲存格值</summary>
	/// <typeparam name="T">值的類型</typeparam>
	/// <param name="doc">Excel 文件</param>
	/// <param name="rowIndex">列索引(從零開始)</param>
	/// <param name="columnIndex">欄索引(從零開始)</param>
	/// <param name="value">要設定的值</param>
	/// <param name="sheetName">工作表名稱</param>
	public static void BindCellValue<T>(this SpreadsheetDocument doc, string sheetName,uint rowIndex, uint columnIndex, T value)
	{
		SheetEditor.BindCellValue(doc, sheetName, rowIndex, columnIndex, value);
    }

	// <summary>設定儲存格值</summary>
	/// <typeparam name="T">值的類型</typeparam>
	/// <param name="doc">Excel 文件</param>
	/// <param name="sheet">工作表</param>
	/// <param name="rowIndex">列索引(從零開始)</param>
	/// <param name="columnIndex">欄索引(從零開始)</param>
	/// <param name="value">要設定的值</param>
	private static void BindCellValue<T>(SpreadsheetDocument doc, Sheet sheet, uint rowIndex, uint columnIndex, T value)
	{
		SheetEditor.BindCellValue(doc, sheet, rowIndex, columnIndex, value);
    }

	/// <summary>設定儲存格值</summary>
	/// <param name="doc">Excel 文件</param>
	/// <param name="reference">位置 e.g. A1, B3,...</param>
	/// <param name="value">要設定的值</param>
	/// <param name="sheetIndex">工作表索引(從零開始)</param>
	public static void BindCellValue<T>(SpreadsheetDocument doc, string reference, T value, int sheetIndex = 0)
	{
		SheetEditor.BindCellValue(doc, reference, value, sheetIndex);
    }

	/// <summary>設定儲存格值</summary>
	/// <param name="doc">Excel 文件</param>
	/// <param name="reference">位置</param>
	/// <param name="value">要設定的值</param>
	/// <param name="sheetName">工作表索引(從零開始)</param>
	public static void BindCellValue<T>(SpreadsheetDocument doc, string reference, T value, string sheetName)
	{
       SheetEditor.BindCellValue(doc, reference, value, sheetName);
    }

	/// <summary>取得excel文件的儲存格</summary>
	/// <param name="document">excel文件</param>
	/// <param name="sheetIndex">工作表索引(從零開始)</param>
	/// <param name="rowIndex">列索引(從零開始)</param>
	/// <param name="columnIndex">欄索引(從零開始)</param>
	/// <returns></returns>
	public static Cell GetCell(this SpreadsheetDocument document, int rowIndex, int columnIndex, int sheetIndex = 0)
	{
		return CellLocator.GetCell(document, rowIndex, columnIndex, sheetIndex);
    }

	/// <summary>尋找儲存格</summary>
	/// <param name="worksheet">工作表</param>
	/// <param name="columnIndex">欄索引(從零開始)</param>
	/// <param name="rowIndex">列索引(從零開始)</param>
	/// <param name="cell">找到的儲存格</param>
	/// <returns></returns>
	public static bool FindCell(Worksheet worksheet, uint columnIndex, uint rowIndex, out Cell cell)
		=> CellLocator.FindSpreadsheetCell(worksheet, columnIndex, rowIndex, out cell);
}