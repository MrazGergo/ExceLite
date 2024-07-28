using System;

namespace ExceLite
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        public string ColumnName { get; }

        public ExcelColumnAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}
