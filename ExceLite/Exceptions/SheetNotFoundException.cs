using System;

namespace ExceLite.Exceptions
{
    public class SheetNotFoundException : Exception
    {
        public string SheetName { get; }

        public SheetNotFoundException(string sheetName) : base($"Sheet not found with name '{sheetName}'")
        {
            SheetName = sheetName;
        }
    }
}
