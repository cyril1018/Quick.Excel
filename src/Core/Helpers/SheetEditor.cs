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
        /// <summary>Set cell value</summary>
        /// <param name="doc">Excel document</param>
        /// <param name="rowIndex">Row index (zero-based)</param>
        /// <param name="columnIndex">Column index (zero-based)</param>
        /// <param name="value">Value to set</param>
        /// <param name="sheetIndex">Sheet index (zero-based)</param>
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

        /// <summary>Set cell value</summary>
        /// <param name="doc">Excel document</param>
        /// <param name="reference">Position e.g. A1, B3,...</param>
        /// <param name="value">Value to set</param>
        /// <param name="sheetIndex">Sheet index (zero-based)</param>
        public static void BindCellValue<T>(SpreadsheetDocument doc, string reference, T value, int sheetIndex = 0)
        {
            var pos = CellReferenceConverter.Convert(reference);
            BindCellValue(doc, pos.rowIndex, pos.columnIndex, value, sheetIndex);
        }

        /// <summary>Set cell value</summary>
        /// <param name="document">Excel document</param>
        /// <param name="position">Position</param>
        /// <param name="value">Value to set</param>
        /// <param name="sheetName">Sheet name</param>
        public static void BindCellValue<T>(SpreadsheetDocument document, string position, T value, string sheetName)
        {
            var pos = CellReferenceConverter.Convert(position);
            var sheet = FindSheetByName(document, sheetName);
            BindCellValue(document, sheet, pos.rowIndex, pos.columnIndex, value);
        }

        /// <summary>Set cell value</summary>
        /// <typeparam name="T">Type of value</typeparam>
        /// <param name="doc">Excel document</param>
        /// <param name="rowIndex">Row index (zero-based)</param>
        /// <param name="columnIndex">Column index (zero-based)</param>
        /// <param name="value">Value to set</param>
        /// <param name="sheetName">Sheet name</param>
        public static void BindCellValue<T>(SpreadsheetDocument doc, string sheetName, uint rowIndex, uint columnIndex, T value)
        {
            var sheet = FindSheetByName(doc, sheetName);
            if (sheet == null)
                return;
            BindCellValue(doc, sheet, rowIndex, columnIndex, value);
        }

        /// <summary>Set cell value</summary>
        /// <typeparam name="T">Type of value</typeparam>
        /// <param name="document">Excel document</param>
        /// <param name="sheet">Sheet</param>
        /// <param name="rowIndex">Row index (zero-based)</param>
        /// <param name="columnIndex">Column index (zero-based)</param>
        /// <param name="value">Value to set</param>
        public static void BindCellValue<T>(SpreadsheetDocument document, Sheet sheet, uint rowIndex, uint columnIndex, T value)
        {
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
            var cell = CellLocator.GetOrCreateCell(worksheetPart.Worksheet, columnIndex, rowIndex);
            CellBinder.BindValue(cell, value);
        }

        /// <summary>Get sheet</summary>
        /// <param name="doc">Document</param>
        /// <param name="sheetName">Sheet name</param>
        /// <returns></returns>
        private static Sheet FindSheetByName(SpreadsheetDocument doc, string sheetName)
        => doc.WorkbookPart.Workbook
            .GetFirstChild<Sheets>()
            ?.Elements<Sheet>()
            .FirstOrDefault(x => x.Name == sheetName);
    }
}
