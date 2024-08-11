using System.Text.RegularExpressions;

namespace ExceLite.Unit.Tests.ExcelFiles;
internal static class FilePaths
{
    public static string EmptyExcel => GetRelativeTestFilePath("Empty.xlsx");
    public static string OnlyHeaderExcel => GetRelativeTestFilePath("OnlyHeader.xlsx");
    public static string OneLineDataExcel => GetRelativeTestFilePath("OneLineData.xlsx");
    public static string MultiLineDateExcel => GetRelativeTestFilePath("MultiLineDate.xlsx");
    public static string MultiSheetExcel => GetRelativeTestFilePath("MultiSheet.xlsx");
    public static string NoHeaderExcel => GetRelativeTestFilePath("NoHeader.xlsx");

    public static Stream OpenReadStream(string fileName)
    {
        return File.OpenRead(fileName);
    }

    public static Stream OpenReadWriteStream(string fileName)
    {
        var folderPath = Path.GetDirectoryName(fileName)!;
        if (!Directory.Exists(folderPath))
        {
            Directory.CreateDirectory(folderPath);
        }

        return File.Open(fileName, FileMode.OpenOrCreate);
    }

    public static string RandomFilePath
    {
        get
        {
            const string subFolderName = "ExceLite";
            var randomName = Regex.Replace(Path.GetRandomFileName(), ".[a-z]+$", ".xlsx");

            return Path.Combine(Path.GetTempPath(), subFolderName, randomName);
        }
    }

    private static string GetRelativeTestFilePath(string fileName)
    {
        return Path.Combine("ExcelFiles", fileName);
    }
}
