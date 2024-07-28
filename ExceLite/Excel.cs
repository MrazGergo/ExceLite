using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExceLite.Exceptions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExceLite
{
    public static class Excel
    {
        /// <summary>
        /// Reads data from an Excel file and maps it to a collection of objects of type T.
        /// </summary>
        /// <typeparam name="T">The type of objects to be returned, which should have public properties corresponding to the Excel columns.</typeparam>
        /// <param name="stream">The stream containing the Excel file data.</param>
        /// <param name="sheetName">The name of the sheet from which to read the data. If null, the first sheet is used.</param>
        /// <returns>An enumerable collection of objects of type T populated with the data from the Excel sheet.</returns>
        /// <exception cref="ArgumentNullException">Thrown if the provided stream is null.</exception>
        /// <exception cref="NoValidPropertyException">Thrown if the type T does not have any valid public properties.</exception>
        /// <exception cref="SheetNotFoundException">Thrown if the specified sheet name does not exist in the Excel file.</exception>
        public static IEnumerable<T> ReadFromExcel<T>(Stream stream, string sheetName = null)
            where T : new()
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            var properties = GetValidProperties<T>();
            if (properties is null || properties.Length == 0)
            {
                throw new NoValidPropertyException(typeof(T));
            }

            var propertyNameDict = properties.ToDictionary(p => p.Name, p => p);

            using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;

                Sheet sheet;
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);

                    if (sheet is null)
                    {
                        throw new SheetNotFoundException(sheetName);
                    }
                }
                else
                {
                    sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
                }

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                var rows = sheetData.Elements<Row>();

                //If it doesn't contain any elements or contains only one element.
                if (!rows.Any() || rows.ElementAtOrDefault(1) == default)
                {
                    yield break;
                }

                var headerRow = rows.First();
                var columnPropertyDict = new Dictionary<string, PropertyInfo>();

                foreach (var cell in headerRow.Elements<Cell>())
                {
                    var cellValue = GetCellValue(spreadsheetDocument, cell);
                    if (propertyNameDict.TryGetValue(cellValue, out var property))
                    {
                        var columnReference = GetColumnReference(cell.CellReference);
                        columnPropertyDict[columnReference] = property;
                    }
                }

                foreach (var row in rows.Skip(1))
                {
                    var obj = new T();
                    foreach (var cell in row.Elements<Cell>())
                    {
                        var columnReference = GetColumnReference(cell.CellReference);
                        if (columnPropertyDict.TryGetValue(columnReference, out var property))
                        {
                            var cellValue = GetCellValue(spreadsheetDocument, cell);
                            object convertedValue;
                            if (property.PropertyType == typeof(DateTime))
                            {
                                convertedValue = DateTime.FromOADate(Convert.ToDouble(cellValue, CultureInfo.InvariantCulture));
                            }
                            else
                            {
                                convertedValue = Convert.ChangeType(cellValue, property.PropertyType, CultureInfo.InvariantCulture);
                            }

                            property.SetValue(obj, convertedValue);
                        }
                    }

                    yield return obj;
                }
            }
        }

        /// <summary>
        /// Extracts the column reference (letters) from the given cell reference (e.g., "A1", "B2").
        /// </summary>
        /// <param name="cellReference">The cell reference from which to extract the column reference.</param>
        /// <returns>The column reference part of the cell reference as a string (e.g., "A", "B").</returns>
        private static string GetColumnReference(string cellReference)
        {
            var match = Regex.Match(cellReference, @"^[A-Z]+");
            return match.Value;
        }

        /// <summary>
        /// Retrieves the cell value from the given cell in an Excel spreadsheet.
        /// </summary>
        /// <param name="document">The spreadsheet document containing the cell.</param>
        /// <param name="cell">The cell from which to retrieve the value.</param>
        /// <returns>The cell value as a string. If the cell contains a shared string, the corresponding value from the shared string table is returned.</returns>
        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            var value = cell.CellValue?.Text;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                if (stringTable != null)
                {
                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }

            return value;
        }

        /// <summary>
        /// Retrieves the public instance properties of the specified type.
        /// </summary>
        /// <typeparam name="T">The type for which to retrieve the properties.</typeparam>
        /// <returns>An array of PropertyInfo objects representing the public instance properties of the specified type.</returns>
        private static PropertyInfo[] GetValidProperties<T>() => typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
    }
}
