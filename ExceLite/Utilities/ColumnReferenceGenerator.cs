namespace ExceLite.Utilities
{
    internal class ColumnReferenceGenerator
    {
        private int _columnIndex = 0;

        /// <summary>
        /// Resets the next column reference to A.
        /// </summary>
        public void Reset()
        {
            _columnIndex = 0;
        }

        /// <summary>
        /// Gets the next column reference in Excel-style (e.g., "A", "B", "AA", "AB", "AAA", "AAB", etc.).
        /// </summary>
        public string Next => GetColumnReference(_columnIndex++);

        /// <summary>
        /// Converts a zero-based column index to an Excel-style column name (e.g., 0 -> "A", 1 -> "B", 26 -> "AA").
        /// </summary>
        /// <param name="columnIndex">The zero-based column index.</param>
        /// <returns>The corresponding Excel-style column name.</returns>
        private static string GetColumnReference(int columnIndex)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string columnName = string.Empty;

            while (columnIndex >= 0)
            {
                columnName = $"{letters[columnIndex % 26]}{columnName}";
                columnIndex = columnIndex / 26 - 1;
            }

            return columnName;
        }
    }
}
