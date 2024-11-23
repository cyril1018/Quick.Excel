using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quick.Excel.Core.Helpers
{
    static internal class StyleHelper
    {
        /// <summary>設定預設日期格式</summary>
        /// <param name="doc">文件</param>
        /// <param name="cellStyleIndex">日期格式索引</param>
        /// <remarks>參考來源：http://polymathprogrammer.com/2009/11/09/how-to-create-stylesheet-in-excel-open-xml/ </remarks>
        public static void ApplyDefaultDateFormat(this SpreadsheetDocument doc, out UInt32Value cellStyleIndex)
        {
            var workbookStylesPart = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = new Stylesheet();
            var stylesheet = workbookStylesPart.Stylesheet;

            var fonts = new Fonts();
            fonts.AppendChild(new Font
            {
                FontName = new FontName { Val = "Calibri" },
                FontSize = new FontSize { Val = 11 }
            });
            fonts.Count = (uint)fonts.ChildElements.Count;

            var fills = new Fills();
            fills.AppendChild(new Fill
            {
                PatternFill = new PatternFill { PatternType = PatternValues.None }
            });
            fills.AppendChild(new Fill
            {
                PatternFill = new PatternFill { PatternType = PatternValues.Gray125 }
            });
            fills.Count = (uint)fills.ChildElements.Count;

            var borders = new Borders();
            borders.AppendChild(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Count = (uint)borders.ChildElements.Count;

            var cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.AppendChild(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            });
            cellStyleFormats.Count = (uint)cellStyleFormats.ChildElements.Count;

            uint excelIndex = 164;
            var numberingFormats = new NumberingFormats();
            var cellFormats = new CellFormats();
            cellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            });

            numberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = excelIndex,
                FormatCode = "yyyy/m/d"
            });
            cellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = excelIndex,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            });
            numberingFormats.Count = (uint)numberingFormats.ChildElements.Count;
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;
            cellStyleIndex = cellFormats.Count - 1;

            stylesheet.AppendChild(numberingFormats);
            stylesheet.AppendChild(fonts);
            stylesheet.AppendChild(fills);
            stylesheet.AppendChild(borders);
            stylesheet.AppendChild(cellStyleFormats);
            stylesheet.AppendChild(cellFormats);

            var cellStyles = new CellStyles();
            cellStyles.AppendChild(new CellStyle
            {
                Name = "Normal",
                FormatId = 0,
                BuiltinId = 0
            });
            cellStyles.Count = (uint)cellStyles.ChildElements.Count;
            stylesheet.AppendChild(cellStyles);

            stylesheet.AppendChild(new DifferentialFormats
            {
                Count = 0
            });

            stylesheet.AppendChild(new TableStyles
            {
                Count = 0,
                DefaultTableStyle = "TableStyleMedium9",
                DefaultPivotStyle = "PivotStyleLight16"
            });
        }
    }
}
