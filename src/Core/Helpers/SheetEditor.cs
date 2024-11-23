using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Quick.Excel.Core.Helpers
{
    static internal class SheetEditor
    {
        /// <summary>設定儲存格值</summary>
        /// <param name="doc">Excel 文件</param>
        /// <param name="rowIndex">列索引(從零開始)</param>
        /// <param name="columnIndex">欄索引(從零開始)</param>
        /// <param name="value">要設定的值</param>
        /// <param name="sheetIndex">工作表索引(從零開始)</param>
        public static void BindCellValue<T>(SpreadsheetDocument doc, uint rowIndex, uint columnIndex, T value, int sheetIndex = 0)
        {
            var sheet = doc.WorkbookPart.Workbook
                .GetFirstChild<Sheets>()
                .Elements<Sheet>()
                .Skip(sheetIndex)
                .FirstOrDefault();
            if (sheet == null)
                return;
            BindCellValue(doc, sheet, rowIndex, columnIndex, value);
        }

        /// <summary>設定儲存格值</summary>
        /// <param name="doc">Excel 文件</param>
        /// <param name="reference">位置 e.g. A1, B3,...</param>
        /// <param name="value">要設定的值</param>
        /// <param name="sheetIndex">工作表索引(從零開始)</param>
        public static void BindCellValue<T>(SpreadsheetDocument doc, string reference, T value, int sheetIndex = 0)
        {
            var pos = CellReferenceConverter.Convert(reference);
            BindCellValue(doc, pos.rowIndex, pos.columnIndex, value, sheetIndex);
        }

        /// <summary>設定儲存格值</summary>
        /// <param name="document">Excel 文件</param>
        /// <param name="position">位置</param>
        /// <param name="value">要設定的值</param>
        /// <param name="sheetName">工作表索引(從零開始)</param>
        public static void BindCellValue<T>(SpreadsheetDocument document, string position, T value, string sheetName)
        {
            var pos = CellReferenceConverter.Convert(position);
            var sheet = FindSheetByName(document, sheetName);
            BindCellValue(document, sheet, pos.rowIndex, pos.columnIndex, value);
        }

        /// <summary>設定儲存格值</summary>
        /// <typeparam name="T">值的類型</typeparam>
        /// <param name="doc">Excel 文件</param>
        /// <param name="rowIndex">列索引(從零開始)</param>
        /// <param name="columnIndex">欄索引(從零開始)</param>
        /// <param name="value">要設定的值</param>
        /// <param name="sheetName">工作表名稱</param>
        public static void BindCellValue<T>(SpreadsheetDocument doc, string sheetName, uint rowIndex, uint columnIndex, T value)
        {
            var sheet = FindSheetByName(doc, sheetName);
            if (sheet == null)
                return;
            BindCellValue(doc, sheet, rowIndex, columnIndex, value);
        }

        // <summary>設定儲存格值</summary>
        /// <typeparam name="T">值的類型</typeparam>
        /// <param name="document">Excel 文件</param>
        /// <param name="sheet">工作表</param>
        /// <param name="rowIndex">列索引(從零開始)</param>
        /// <param name="columnIndex">欄索引(從零開始)</param>
        /// <param name="value">要設定的值</param>
        public static void BindCellValue<T>(SpreadsheetDocument document, Sheet sheet, uint rowIndex, uint columnIndex, T value)
        {
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
            var cell = CellLocator.GetOrCreateCell(worksheetPart.Worksheet, columnIndex, rowIndex);
            CellBinder.BindValue(cell, value);
        }


        /// <summary>取得工作表</summary>
        /// <param name="doc">文件</param>
        /// <param name="sheetName">工作表名稱</param>
        /// <returns></returns>
        private static Sheet FindSheetByName(SpreadsheetDocument doc, string sheetName)
        => doc.WorkbookPart.Workbook
            .GetFirstChild<Sheets>()
            ?.Elements<Sheet>()
            .FirstOrDefault(x => x.Name == sheetName);
    }
}
