﻿using DocumentFormat.OpenXml.Spreadsheet;
using SanChong.Excel.Core.Helpers;
using System.Collections;
using System.Reflection;

namespace SanChong.Excel.Core;

/// <summary>Generate worksheet using enumerable data</summary>
public class DataSheetCreator : SheetCreatorBase
{
    /// <summary>Output data</summary>
    protected readonly IEnumerable Data;

    /// <summary>Starting row index for output</summary>
    protected readonly int StartRowIndex;

    /// <summary>Starting column index for output</summary>
    protected readonly int StartColumnIndex;

    public DataSheetCreator(IEnumerable data, int startRowIndex = 0, int startColumnIndex = 0)
    {
        Data = data;
        StartRowIndex = startRowIndex;
        StartColumnIndex = startColumnIndex;
        CellCreated += EnumerableDataSheetCreator_CellCreated;
    }

    /// <summary>
    /// Whether the data is empty
    /// </summary>
    protected bool IsEmptyData => Data == null || !Data.Cast<object>().Any();

    /// <summary>Number of rows to generate in the worksheet</summary>
    int RowsCount => IsEmptyData
        ? 1
        : StartRowIndex + Data.Cast<object>().Count();

    /// <summary>Number of columns to generate in the worksheet</summary>
    int ColumnsCount => IsEmptyData
        ? 1
        : StartColumnIndex + PropertyNames.Length;

    /// <summary>Data enumerator index</summary>
    int? DataEnumeratorIndex = null;

    /// <summary>Data enumerator</summary>
    IEnumerator DataEnumerator;

    /// <summary>Data property names cache</summary>
    string[] _PropertyNames;

    /// <summary>Data property names</summary>
    protected string[] PropertyNames
    {
        get
        {
            if (_PropertyNames != null)
                return _PropertyNames;

            _PropertyNames = Data.Cast<object>().FirstOrDefault() is IDictionary<string, object> dict
                ? dict.Keys.ToArray()
                : DataProperties?.Select(x => x.Name).ToArray();

            return _PropertyNames;
        }
    }

    /// <summary>Set cell data value</summary>
    /// <param name="cell">Cell to set</param>
    /// <param name="rowIndex">Data row index</param>
    /// <param name="columnIndex">Data column index</param>
    private void SetDataCell(Cell cell, int rowIndex, int columnIndex)
    {
        if(IsEmptyData)
        {
            CellBinder.BindValue(cell, (string)null);
            return;
        }    
        var row = GetDataRow(rowIndex);
        var name = PropertyNames[columnIndex];
        var value = GetValue(row, name);
        CellBinder.BindValue(cell, value);
    }

    /// <summary>Get value by property name</summary>
    /// <param name="data">Data</param>
    /// <param name="propertyName">Property name</param>
    /// <returns></returns>
    object GetValue(object data, string propertyName)
    => data is IDictionary<string, object> dict
        ? dict[propertyName]
        : (data.GetType().GetProperty(propertyName)?.GetValue(data));

    /// <summary>Data property information</summary>
    IEnumerable<PropertyInfo> DataProperties
        => Data.Cast<object>().FirstOrDefault()?
        .GetType()
        .GetProperties()
        .Where(x => x.DeclaringType.Name != "DynamicClass");

    /// <summary>Set data value when cell is created</summary>
    /// <param name="sender">Event sender</param>
    /// <param name="e">Event arguments</param>
    private void EnumerableDataSheetCreator_CellCreated(object sender, CellCreatedEventArgs e)
    {
        if (e.RowIndex < StartRowIndex || e.ColumnIndex < StartColumnIndex)
            return;
        var _DataColumnIndex = e.ColumnIndex - StartColumnIndex;
        var _DataRowIndex = e.RowIndex - StartRowIndex;
        SetDataCell(e.Cell, _DataRowIndex, _DataColumnIndex);
    }

    /// <summary>Create worksheet</summary>
    public override SheetData CreateSheetData()
        => CreateSheetData(ColumnsCount, RowsCount);

    /// <summary>Get single data row</summary>
    /// <param name="index">Data index</param>
    /// <returns></returns>
    object GetDataRow(int index)
    {
        if (DataEnumeratorIndex == null)
            InitDataEnumerator();

        while (DataEnumeratorIndex < index)
        {
            DataEnumerator.MoveNext();
            DataEnumeratorIndex++;
        }

        if (DataEnumeratorIndex == index)
            return DataEnumerator.Current;

        DataEnumeratorIndex = null;
        DataEnumerator.Reset();
        return GetDataRow(index);
    }

    /// <summary>Initialize data enumerator</summary>
    void InitDataEnumerator()
    {
        DataEnumerator = Data.GetEnumerator();
        DataEnumerator.MoveNext();
        DataEnumeratorIndex = 0;
    }
}
