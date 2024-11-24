using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SanChong.Excel.Models;

namespace SanChong.Excel.Core
{
    /// <summary>
    /// Static internal class for building workbooks
    /// </summary>
    static internal class WorkbookBuilder
    {
        /// <summary>
        /// Adds sheets to the workbook part
        /// </summary>
        /// <param name="workbookPart">The workbook part to add sheets to</param>
        /// <param name="sheetDescriptors">Array of sheet descriptors containing sheet details</param>
        internal static void AddSheets(WorkbookPart workbookPart, SheetDescriptor[] sheetDescriptors)
        {
            // Append a new Sheets collection to the workbook
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            UInt32Value sheetId = 1;

            // Iterate through each sheet descriptor and add corresponding sheets
            foreach (var dto in sheetDescriptors)
            {
                // Create a new worksheet part and set its worksheet
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // Append columns settings to the worksheet
                worksheetPart.Worksheet.AppendChild(dto.Columns);

                // Append the sheet to the sheets collection
                sheets.AppendChild(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId++,
                    Name = dto.Name
                });

                // Add sheet data to the worksheet
                worksheetPart.Worksheet.AddChild(dto.Data);
            }
        }
    }
}
