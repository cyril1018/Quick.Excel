using DocumentFormat.OpenXml.Spreadsheet;
using Quick.Excel.Models;
using System.Collections;
using System.Reflection;

namespace Quick.Excel.Core;

/// <summary>工作表產生器</summary>
public abstract class SheetCreatorBase
{
	protected SheetCreatorBase()
	{
	}

	/// <summary>建立工作表</summary>
	public abstract SheetData CreateSheetData();

	/// <summary>建立工作表</summary>
	/// <param name="columns">欄位數量</param>
	/// <param name="rows">列數量</param>
	/// <returns>工作表資料</returns>
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

	/// <summary>建立資料列</summary>
	/// <returns>資料列</returns>
	Row CreateRow() => new Row();

	/// <summary>建立資料 Cell</summary>
	/// <returns>資料 Cell</returns>
	Cell CreateCell()
	{
		var _Cell = new Cell();
		return _Cell;
	}

	/// <summary>資料列建立事件參數</summary>
	public class RowCreatedEventArgs : EventArgs
	{
		/// <summary>資料列建立事件參數建構式</summary>
		/// <param name="row">建立之資料列</param>
		/// <param name="rowIndex">列索引</param>
		public RowCreatedEventArgs(Row row, int rowIndex)
		{
			Row = row;
			RowIndex = rowIndex;
		}

		/// <summary>資料列</summary>
		public Row Row { get; private set; }

		/// <summary>列索引</summary>
		public int RowIndex { get; private set; }
	}

	/// <summary>資料 Cell 建立事件參數</summary>
	public class CellCreatedEventArgs : EventArgs
	{
		/// <summary>資料 Cell 建立事件參數建構式</summary>
		/// <param name="cell">建立之資料 Cell</param>
		/// <param name="rowIndex">列索引</param>
		/// <param name="columnIndex">欄索引</param>
		public CellCreatedEventArgs(Cell cell, int rowIndex, int columnIndex)
		{
			Cell = cell;
			ColumnIndex = columnIndex;
			RowIndex = rowIndex;
		}

		/// <summary>資料 Cell</summary>
		public Cell Cell { get; private set; }

		/// <summary>列索引</summary>
		public int RowIndex { get; private set; }

		/// <summary>欄索引</summary>
		public int ColumnIndex { get; private set; }
	}

	/// <summary>資料列建立事件處理器</summary>
	public event EventHandler<RowCreatedEventArgs> RowCreated;

	/// <summary>資料 Cell 建立事件處理器</summary>
	public event EventHandler<CellCreatedEventArgs> CellCreated;
}