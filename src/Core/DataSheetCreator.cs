using DocumentFormat.OpenXml.Spreadsheet;
using Quick.Excel.API;
using Quick.Excel.Core.Helpers;
using Quick.Excel.Models;
using System.Collections;
using System.Reflection;

namespace Quick.Excel.Core;

/// <summary>使用列舉資料產生 工作表</summary>
public class DataSheetCreator
    : SheetCreatorBase
{
    /// <summary>輸出資料</summary>
    protected readonly IEnumerable Data;

    /// <summary>起始輸出列位置索引</summary>
    protected readonly int StartRowIndex;

    /// <summary>起始輸出欄位置索引</summary>
    protected readonly int StartColumnIndex;
    public DataSheetCreator
        (IEnumerable data, int startRowIndex = 0, int startColumnIndex = 0)
    {
        Data = data;
        StartRowIndex = startRowIndex;
        StartColumnIndex = startColumnIndex;
        CellCreated += EnumerableDataSheetCreator_CellCreated;
    }

    /// <summary>產生工作表列數</summary>
    int RowsCount => StartRowIndex + Data.Cast<object>().Count();

    /// <summary>產生工作表欄數</summary>
    int ColumnsCount => StartColumnIndex + PropertyNames.Length;

    /// <summary>資料列舉操作索引</summary>
    int? DataEnumeratorIndex = null;

    /// <summary>資料列舉操作器</summary>
    IEnumerator DataEnumerator;

    /// <summary>資料屬性名稱暫存</summary>
    string[] _PropertyNames;

    /// <summary>資料屬性名稱</summary>
    protected string[] PropertyNames
    {
        get
        {
            if (_PropertyNames != null)
                return _PropertyNames;

            _PropertyNames = Data.Cast<object>().FirstOrDefault() is IDictionary<string, object> dict
                ? dict.Keys.ToArray()
                : DataProperties.Select(x => x.Name).ToArray();

            return _PropertyNames;
        }
    }

    /// <summary>設定儲存格資料值</summary>
    /// <param name="cell">待設定儲存格</param>
    /// <param name="rowIndex">資料列索引</param>
    /// <param name="columnIndex">資料欄位索引</param>
    private void SetDataCell(Cell cell, int rowIndex, int columnIndex)
    {
        var row = GetDataRow(rowIndex);
        var name = PropertyNames[columnIndex];
        var value = GetValue(row, name);
        CellBinder.BindValue(cell, value);
    }

    /// <summary>依屬性名稱取值 </summary>
    /// <param name="data">資料</param>
    /// <param name="propertyName">屬性名稱</param>
    /// <returns></returns>
    object GetValue(object data, string propertyName)
    => data is IDictionary<string, object> dict
        ? dict[propertyName]
        : (data.GetType().GetProperty(propertyName)?.GetValue(data));

    /// <summary>資料屬性資訊</summary>
    IEnumerable<PropertyInfo> DataProperties
        => Data.Cast<object>().FirstOrDefault()?
        .GetType()
        .GetProperties()
        .Where(x => x.DeclaringType.Name != "DynamicClass");

    /// <summary> Cell 建立事件時設定資料值</summary>
    /// <param name="sender">事件發動者</param>
    /// <param name="e">事件參數</param>
    private void EnumerableDataSheetCreator_CellCreated(object sender, CellCreatedEventArgs e)
    {
        if (e.RowIndex < StartRowIndex || e.ColumnIndex < StartColumnIndex)
            return;
        var _DataColumnIndex = e.ColumnIndex - StartColumnIndex;
        var _DataRowIndex = e.RowIndex - StartRowIndex;
        SetDataCell(e.Cell, _DataRowIndex, _DataColumnIndex);
    }

    /// <summary>建立工作表</summary>
    public override SheetData CreateSheetData()
        => CreateSheetData(ColumnsCount, RowsCount);

    /// <summary>取得單筆資料</summary>
    /// <param name="index">資料索引</param>
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

    /// <summary>初始化資料列舉操作器</summary>
    void InitDataEnumerator()
    {
        DataEnumerator = Data.GetEnumerator();
        DataEnumerator.MoveNext();
        DataEnumeratorIndex = 0;
    }
}
