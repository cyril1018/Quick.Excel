using Quick.Excel.API;
using Quick.Excel.Core.Helpers;
using System.Collections;

namespace Quick.Excel.Core;

/// <summary>Generates a worksheet using enumerable data (outputs property names as titles in the first row)</summary>
public class TitleDataSheetCreator : DataSheetCreator
{
    /// <summary>Title row index</summary>
    readonly int TitleRowIndex;
    public TitleDataSheetCreator(IEnumerable data, int startRowIndex = 0, int startColumnIndex = 0)
        : base(data, startRowIndex + 1, startColumnIndex)
    {
        TitleRowIndex = startRowIndex;
        CellCreated += EnumerableDataWithTitleSheetCreator_CellCreated;
    }

    /// <summary> Sets the title row when the Cell Created event is triggered</summary>
    /// <param name="sender">Event sender</param>
    /// <param name="e">Event arguments</param>
    private void EnumerableDataWithTitleSheetCreator_CellCreated(object sender, CellCreatedEventArgs e)
    {
        if (e.RowIndex != TitleRowIndex)
            return;
        CellBinder.BindValue(e.Cell, PropertyNames[e.ColumnIndex - StartColumnIndex]);
    }
}
