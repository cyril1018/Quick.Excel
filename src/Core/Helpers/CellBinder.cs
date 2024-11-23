using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Quick.Excel.Core.Helpers
{
    static internal class CellBinder
    {
        /// <summary>設定儲存格值</summary>
        /// <param name="cell">儲存格</param>
        /// <param name="value">值</param>
        public static void BindValue<T>(Cell cell, T value)
        {
            if (value is string strVal)
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(strVal);
                return;
            }
            if (value is int intVal)
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(intVal);
                return;
            }
            if (value is double doubleVal)
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(doubleVal);
                return;
            }
            if (value is decimal decimalVal)
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(decimalVal);
                return;
            }
            if (value is DateTime datetimeVal)
            {
                cell.DataType = CellValues.Date;
                cell.CellValue = new CellValue(datetimeVal);
                return;
            }
            if (value is DateTimeOffset datetimeOffsetVal)
            {
                cell.DataType = CellValues.Date;
                cell.CellValue = new CellValue(datetimeOffsetVal);
                return;
            }
            if (value is bool boolVal)
            {
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(boolVal);
                return;
            }

            cell.DataType = CellValues.String;
            if (value == null)
                cell.CellValue = null;
            else
                cell.CellValue = new CellValue(value.ToString());
        }
    }
}
