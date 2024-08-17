using System.Collections.Generic;

namespace ExceLite.Utilities
{
    /// <summary>
    /// A custom comparer for Excel column references that orders them according to their position in an Excel worksheet.
    /// This comparer handles the comparison of column references such as "A", "Z", "AA", "ZZ", etc.
    /// </summary>
    internal class ExcelColumnReferenceComparer : IComparer<string>
    {
        /// <summary>
        /// Compares two Excel column references and returns an indication of their relative order.
        /// The comparison is based on the length of the column reference first, and if they are of equal length,
        /// a lexicographical comparison is performed.
        /// </summary>
        /// <param name="x">The first column reference to compare.</param>
        /// <param name="y">The second column reference to compare.</param>
        /// <returns>
        /// An integer that indicates the relative order of the column references:
        /// <list type="bullet">
        /// <item><description>Less than zero: <paramref name="x"/> precedes <paramref name="y"/> in the sort order.</description></item>
        /// <item><description>Zero: <paramref name="x"/> is equal to <paramref name="y"/> in the sort order.</description></item>
        /// <item><description>Greater than zero: <paramref name="x"/> follows <paramref name="y"/> in the sort order.</description></item>
        /// </list>
        /// </returns>
        public int Compare(string x, string y)
        {
            if (x is null)
            {
                return y is null ? 0 : -1;
            }

            if (y is null)
            {
                return 1;
            }

            // First compare by length: longer references are considered greater (e.g., "AA" > "Z")
            int lengthComparison = x.Length.CompareTo(y.Length);
            if (lengthComparison != 0)
            {
                return lengthComparison;
            }

            return string.CompareOrdinal(x, y);
        }
    }
}
