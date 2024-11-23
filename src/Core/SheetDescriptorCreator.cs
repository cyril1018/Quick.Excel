using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Quick.Excel.Models;

namespace Quick.Excel.Core
{
    /// <summary>工作表描述產生器</summary>
    internal class SheetDescriptorCreator
    {
        /// <summary>欄位資訊列表</summary>
        private readonly List<Column> _ColumnList;

        /// <summary>工作表產生器</summary>
        private readonly SheetCreatorBase _SheetCreator;

        public SheetDescriptorCreator(SheetCreatorBase sheetCreator)
        {
            _SheetCreator = sheetCreator;
            _SheetCreator.CellCreated += SheetCreator_CellCreated;
            _ColumnList = new List<Column>();
        }

        /// <summary>欄位資訊</summary>
        private Columns Columns
        {
            get
            {
                Columns cols = new Columns();
                foreach (var col in _ColumnList)
                    cols.AppendChild(col);
                return cols;
            }
        }

        /// <summary>當 Cell 建立後，於此依資料內容設定欄位</summary>
        private void SheetCreator_CellCreated(object sender, SheetCreatorBase.CellCreatedEventArgs e)
        {
            if (e.Cell.CellValue == null)
            {
                BindColumn(e.ColumnIndex, 0);
                return;
            }

            if (e.Cell.DataType == CellValues.Date)
                return;
            var _Str = e.Cell.CellValue.Text;
            BindColumn(e.ColumnIndex, _Str.Length);
        }

        /// <summary>設定欄位</summary>
        /// <param name="columnIndex">欄位索引</param>
        /// <param name="valLength">資料長度</param>
        protected void BindColumn(int columnIndex, int valLength)
        {
            var created = _ColumnList.Count > columnIndex;
            if (created)
            {
                var width = CalculateColumnWidth(valLength);
                var col = _ColumnList[columnIndex];
                if (width > DoubleValue.ToDouble(col.Width))
                    col.Width = DoubleValue.FromDouble(width);
                return;
            }

            // index start from 1 in OpenXML
            var uInt32ColumnIndex = new UInt32Value(Convert.ToUInt32(columnIndex)) + 1;
            _ColumnList.Add(new Column
            {
                Min = uInt32ColumnIndex,
                Max = uInt32ColumnIndex,
                CustomWidth = true,
                Width = CalculateColumnWidth(valLength)
            });
        }

        /// <summary>計算欄位寬度</summary>
        /// <param name="valLength">字串長度</param>
        /// <returns>欄位寬度</returns>
        private double CalculateColumnWidth(int valLength)
         => valLength * 2 + 5;

        /// <summary>建立工作表描述</summary>
        /// <param name="sheetName">工作表名稱</param>
        /// <returns>工作表描述</returns>
        public SheetDescriptor Create(string sheetName)
            => new SheetDescriptor { Name = sheetName, Data = _SheetCreator.CreateSheetData(), Columns = Columns };
    }
}
