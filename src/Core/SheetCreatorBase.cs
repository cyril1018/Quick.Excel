using DocumentFormat.OpenXml.Spreadsheet;
using Quick.Excel.Models;
using System.Collections;
using System.Reflection;

namespace Quick.Excel.Core;

/// <summary>Sheet Creator</summary>
public abstract class SheetCreatorBase
{
    protected SheetCreatorBase()
    {
    }

    /// <summary>Create Sheet</summary>
    public abstract SheetData CreateSheetData();

    /// <summary>Create Sheet</summary>
    /// <param name="columns">Number of columns</param>
    /// <param name="rows">Number of rows</param>
    /// <returns>Sheet data</returns>
    protected SheetData CreateSheetData(int columns, int rows)
    {
        var _SheetData = new SheetData() { };
        for (var _RowIndex = 0; _RowIndex < rows; _RowIndex++)
        {
            var _Row = CreateRow();
            for (var _ColumnIndex = 0; _ColumnIndex < columns; _ColumnIndex++)
            {
                var _Cell = CreateCell();
                _Row.AppendChild(_Cell);
                CellCreated?.Invoke(this, new CellCreatedEventArgs(_Cell, _RowIndex, _ColumnIndex));
            }
            _SheetData.AppendChild(_Row);
            RowCreated?.Invoke(this, new RowCreatedEventArgs(_Row, _RowIndex));
        }
        return _SheetData;
    }

    /// <summary>Create Row</summary>
    /// <returns>Row</returns>
    Row CreateRow() => new Row();

    /// <summary>Create Cell</summary>
    /// <returns>Cell</returns>
    Cell CreateCell()
    {
        var _Cell = new Cell();
        return _Cell;
    }

    /// <summary>Row Created Event Args</summary>
    public class RowCreatedEventArgs : EventArgs
    {
        /// <summary>Row Created Event Args Constructor</summary>
        /// <param name="row">Created row</param>
        /// <param name="rowIndex">Row index</param>
        public RowCreatedEventArgs(Row row, int rowIndex)
        {
            Row = row;
            RowIndex = rowIndex;
        }

        /// <summary>Row</summary>
        public Row Row { get; private set; }

        /// <summary>Row index</summary>
        public int RowIndex { get; private set; }
    }

    /// <summary>Cell Created Event Args</summary>
    public class CellCreatedEventArgs : EventArgs
    {
        /// <summary>Cell Created Event Args Constructor</summary>
        /// <param name="cell">Created cell</param>
        /// <param name="rowIndex">Row index</param>
        /// <param name="columnIndex">Column index</param>
        public CellCreatedEventArgs(Cell cell, int rowIndex, int columnIndex)
        {
            Cell = cell;
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
        }

        /// <summary>Cell</summary>
        public Cell Cell { get; private set; }

        /// <summary>Row index</summary>
        public int RowIndex { get; private set; }

        /// <summary>Column index</summary>
        public int ColumnIndex { get; private set; }
    }

    /// <summary>Row Created Event Handler</summary>
    public event EventHandler<RowCreatedEventArgs> RowCreated;

    /// <summary>Cell Created Event Handler</summary>
    public event EventHandler<CellCreatedEventArgs> CellCreated;
}
