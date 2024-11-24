using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SanChong.Excel.Models;

namespace SanChong.Excel.Core
{
    /// <summary>Sheet descriptor creator</summary>
    internal class SheetDescriptorCreator
    {
        /// <summary>List of column information</summary>
        private readonly List<Column> _ColumnList;

        /// <summary>Sheet creator</summary>
        private readonly SheetCreatorBase _SheetCreator;

        public SheetDescriptorCreator(SheetCreatorBase sheetCreator)
        {
            _SheetCreator = sheetCreator;
            _SheetCreator.CellCreated += SheetCreator_CellCreated;
            _ColumnList = new List<Column>();
        }

        /// <summary>Column information</summary>
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

        /// <summary>Set column based on data content when Cell is created</summary>
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

        /// <summary>Set column</summary>
        /// <param name="columnIndex">Column index</param>
        /// <param name="valLength">Data length</param>
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

        /// <summary>Calculate column width</summary>
        /// <param name="valLength">String length</param>
        /// <returns>Column width</returns>
        private double CalculateColumnWidth(int valLength)
         => valLength * 2 + 5;

        /// <summary>Create sheet descriptor</summary>
        /// <param name="sheetName">Sheet name</param>
        /// <returns>Sheet descriptor</returns>
        public SheetDescriptor Create(string sheetName)
            => new SheetDescriptor { Name = sheetName, Data = _SheetCreator.CreateSheetData(), Columns = Columns };
    }
}
