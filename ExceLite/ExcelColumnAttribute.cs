using System;

namespace ExceLite
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        public string ColumnName { get; }
        public string ColumnReference { get; }

        public ExcelColumnAttribute(string columnName = null, string columnReference = null)
        {
            ColumnName = columnName;
            ColumnReference = columnReference;
        }
    }
}
