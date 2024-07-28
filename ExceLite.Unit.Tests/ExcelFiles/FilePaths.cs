namespace ExceLite.Unit.Tests.ExcelFiles;
internal static class FilePaths
{
    public static string EmptyExcel => GetRelativeTestFilePath("Empty.xlsx");
    public static string OnlyHeaderExcel => GetRelativeTestFilePath("OnlyHeader.xlsx");
    public static string OneLineDataExcel => GetRelativeTestFilePath("OneLineData.xlsx");
    public static string MultiLineDateExcel => GetRelativeTestFilePath("MultiLineDate.xlsx");

    public static Stream OpenReadStream(string fileName)
    {
        return File.OpenRead(fileName);
    }

    private static string GetRelativeTestFilePath(string fileName)
    {
        return Path.Combine("ExcelFiles", fileName);
    }
}
