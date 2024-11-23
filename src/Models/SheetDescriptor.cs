using DocumentFormat.OpenXml.Spreadsheet;

namespace Quick.Excel.Models
{
    /// <summary>
    /// Sheet descriptor
    /// </summary>
    public class SheetDescriptor
    {
        /// <summary>Sheet name</summary>
        public string Name { get; set; }

        /// <summary>Sheet data</summary>
        public SheetData Data { get; set; }

        /// <summary>Sheet columns settings</summary>
        public Columns Columns { get; set; }
    }
}
