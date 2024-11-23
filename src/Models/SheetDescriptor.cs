using DocumentFormat.OpenXml.Spreadsheet;

namespace Quick.Excel.Models
{
    /// <summary>
    /// 工作表描述
    /// </summary>
    public class SheetDescriptor
    {
        /// <summary>工作表名稱</summary>
		public string Name { get; set; }

        /// <summary>工作表資料</summary>
        public SheetData Data { get; set; }

        /// <summary>工作表欄位設定</summary>
        public Columns Columns { get; set; }
    }
}
