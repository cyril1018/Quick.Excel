using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Quick.Excel.Core.Helpers
{
    /// <summary>儲存格位置轉換器</summary>
    internal static class CellReferenceConverter
    {
        const int AlphabetCount = 26;
        /// <summary>將儲存格位置轉換為列索引與欄索引</summary>
        /// <param name="cellReference">儲存格位置</param>
        /// <param name="rowIndex">列索引</param>
        /// <param name="columnIndex">欄索引</param>
        public static (uint columnIndex, uint rowIndex) Convert(string cellReference)
        {
            var pos = SplitCellPosition(cellReference);
            return (AlphabetToNumber(pos.columnName) - 1, uint.Parse(pos.rowNumber.ToString()) - 1);
        }

        /// <summary>字母轉數字</summary>
        /// <param name="columnName">欄</param>
        /// <returns></returns>
        private static uint AlphabetToNumber(string columnName)
        {
            uint _Result = 0;
            foreach (var _Char in columnName)
            {
                _Result *= AlphabetCount;
                _Result += (uint)(_Char - 'A' + 1);
            }
            return _Result;
        }

        /// <summary>拆出列與欄</summary>
        /// <param name="position">儲存格位置</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        private static (string columnName, int rowNumber) SplitCellPosition(string position)
        {
            var match = Regex.Match(position, @"([A-Z]+)(\d+)");

            if (!match.Success)
                throw new ArgumentException("無效的儲存格位置格式。");

            string columnName = match.Groups[1].Value;
            int rowNumber = int.Parse(match.Groups[2].Value);
            return (columnName, rowNumber);
        }

        /// <summary>數字轉字母</summary>
        /// <remarks> 1->A, 2->B</remarks>
        /// <param name="number">數字</param>
        /// <returns></returns>
        public static string NumberToAlphabet(uint number)
        {
            var _Result = string.Empty;
            while (number > 0)
            {
                var _Remainder = (number - 1) % AlphabetCount;
                number = (number - 1) / AlphabetCount;
                _Result = $"{(char)(65 + _Remainder)}" + _Result;
            }
            return _Result;
        }
    }
}
