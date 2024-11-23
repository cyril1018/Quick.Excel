using Quick.Excel.API;
using Quick.Excel.Core.Helpers;
using System.Collections;

namespace Quick.Excel.Core;

/// <summary>使用列舉資料產生 工作表(第一列輸出資料屬性名稱作為標題)</summary>
public class TitleDataSheetCreator : DataSheetCreator
{
    /// <summary>標題列索引</summary>
    readonly int TitleRowIndex;
    public TitleDataSheetCreator(IEnumerable data, int startRowIndex = 0, int startColumnIndex = 0)
        : base(data, startRowIndex + 1, startColumnIndex)
    {
        TitleRowIndex = startRowIndex;
        CellCreated += EnumerableDataWithTitleSheetCreator_CellCreated;
    }

    /// <summary> Cell 建立事件時 設定標題列</summary>
    /// <param name="sender">事件發動者</param>
    /// <param name="e">事件參數</param>
    private void EnumerableDataWithTitleSheetCreator_CellCreated(object sender, CellCreatedEventArgs e)
    {
        if (e.RowIndex != TitleRowIndex)
            return;
            CellBinder.BindValue(e.Cell, PropertyNames[e.ColumnIndex - StartColumnIndex]);
    }
}
