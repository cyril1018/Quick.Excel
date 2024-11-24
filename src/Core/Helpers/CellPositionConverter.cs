using System.Text.RegularExpressions;

namespace SanChong.Excel.Core.Helpers
{
    /// <summary>Cell position converter</summary>
    internal static class CellReferenceConverter
    {
        const int AlphabetCount = 26;
        /// <summary>Convert cell reference to row index and column index</summary>
        /// <param name="cellReference">Cell reference</param>
        /// <param name="rowIndex">Row index</param>
        /// <param name="columnIndex">Column index</param>
        public static (uint columnIndex, uint rowIndex) Convert(string cellReference)
        {
            var pos = SplitCellPosition(cellReference);
            return (AlphabetToNumber(pos.columnName) - 1, uint.Parse(pos.rowNumber.ToString()) - 1);
        }

        /// <summary>Convert alphabet to number</summary>
        /// <param name="columnName">Column name</param>
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

        /// <summary>Split cell position into column and row</summary>
        /// <param name="position">Cell position</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        private static (string columnName, int rowNumber) SplitCellPosition(string position)
        {
            var match = Regex.Match(position, @"([A-Z]+)(\d+)");

            if (!match.Success)
                throw new ArgumentException("Invalid cell position format.");

            string columnName = match.Groups[1].Value;
            int rowNumber = int.Parse(match.Groups[2].Value);
            return (columnName, rowNumber);
        }

        /// <summary>Convert number to alphabet</summary>
        /// <remarks> 1->A, 2->B</remarks>
        /// <param name="number">Number</param>
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
