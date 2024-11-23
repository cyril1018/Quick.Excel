using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quick.Excel.Core.Helpers
{
    static internal class CellLocator
    {
        /// <summary>尋找儲存格</summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="columnIndex">欄索引(從零開始)</param>
        /// <param name="rowIndex">列索引(從零開始)</param>
        /// <param name="cell">找到的儲存格</param>
        /// <returns></returns>
        public static bool FindSpreadsheetCell(Worksheet worksheet, uint columnIndex, uint rowIndex, out Cell cell)
            => FindSpreadsheetCell(worksheet, $"{CellReferenceConverter.NumberToAlphabet(columnIndex + 1)}", rowIndex, out cell);



        /// <summary>尋找儲存格</summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="columnName">欄位名稱</param>
        /// <param name="rowIndex">列索引(從零開始)</param>
        /// <param name="cell">找到的儲存格</param>
        /// <returns></returns>
        private static bool FindSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex, out Cell cell)
        {
            cell = null;
            Row _Row;
            if (!FindRow(worksheet, rowIndex, out _Row))
                return false;
            cell = _Row.Elements<Cell>()
                .FirstOrDefault(c => string.Compare(c.CellReference.Value, $"{columnName}{rowIndex + 1}", true) == 0);
            return cell != null;
        }

        /// <summary>尋找資料列</summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="rowIndex">列索引(從零開始)</param>
        /// <param name="row">5找到的資料列</param>
        private static bool FindRow(Worksheet worksheet, uint rowIndex, out Row row)
        {
            row = worksheet.GetFirstChild<SheetData>()
           .Elements<Row>()
           .FirstOrDefault(r => r.RowIndex == rowIndex + 1);// Index starts from 1 in OpenXml
            return row != null;
        }



        /// <summary>取得excel文件的儲存格</summary>
        /// <param name="document">excel文件</param>
        /// <param name="sheetIndex">工作表索引(從零開始)</param>
        /// <param name="rowIndex">列索引(從零開始)</param>
        /// <param name="columnIndex">欄索引(從零開始)</param>
        /// <returns></returns>
        public static Cell GetCell(SpreadsheetDocument document, int sheetIndex, int rowIndex, int columnIndex)
        {
            WorkbookPart workbookPart = document.WorkbookPart;

            Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().ElementAt(sheetIndex);

            if (sheet == null)
            {
                throw new InvalidOperationException($"找不到 第'{sheetIndex + 1}'個工作表");
            }

            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetDataList = worksheetPart.Worksheet.Elements<SheetData>().First();

            Row row = sheetDataList.Elements<Row>().ElementAt(rowIndex);

            Cell cell = row.Elements<Cell>().ElementAt(columnIndex);

            return cell;
        }

        /// <summary>取得儲存格(若不存在則建立)</summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="columnIndex">欄索引(從零開始)</param>
        /// <param name="rowIndex">列索引(從零開始)</param>
        /// <returns></returns>
        public static Cell GetOrCreateCell(Worksheet worksheet, uint columnIndex, uint rowIndex)
        {
            Cell _Cell;
            if (FindSpreadsheetCell(worksheet, columnIndex, rowIndex, out _Cell))
                return _Cell;

            _Cell = new Cell
            {
                CellReference = new StringValue($"{CellReferenceConverter.NumberToAlphabet(columnIndex + 1)}{rowIndex + 1}")
            };

            var _Row = GetOrCreateRow(worksheet, rowIndex);
            _Row.AppendChild(_Cell);
            return _Cell;
        }

        /// <summary>取得資料列(若不存在則建立)</summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="rowIndex">列索引(從零開始)</param>
        private static Row GetOrCreateRow(Worksheet worksheet, uint rowIndex)
        {
            Row _Row;
            if (FindRow(worksheet, rowIndex, out _Row))
                return _Row;
            _Row = new Row
            {
                RowIndex = new UInt32Value(rowIndex + 1)// Index starts from 1 in OpenXml
            };
            worksheet.GetFirstChild<SheetData>().AppendChild(_Row);
            return _Row;
        }
    }
}
