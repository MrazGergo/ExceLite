using ExceLite.Exceptions;
using ExceLite.Unit.Tests.ExcelFiles;
using FluentAssertions;

namespace ExceLite.Unit.Tests;

public class ExcelReadTests
{
    [Fact]
    public void ReadFromExcel_ReadsEmptyExcel_ReturnsEmptyIEnumerable()
    {
        //Arrange
        string excelFilePath = FilePaths.EmptyExcel;
        using var stream = FilePaths.OpenReadStream(excelFilePath);

        //Act
        var data = Excel.ReadFromExcel<HeaderTestClass>(stream);

        //Assert
        data.Should().BeEmpty();
    }

    [Fact]
    public void ReadFromExcel_ReadsOnlyHeaderExcel_ReturnsEmptyIEnumerable()
    {
        //Arrange
        string excelFilePath = FilePaths.OnlyHeaderExcel;
        using var stream = FilePaths.OpenReadStream(excelFilePath);

        //Act
        var data = Excel.ReadFromExcel<HeaderTestClass>(stream);

        //Assert
        data.Should().BeEmpty();

    }

    [Fact]
    public void ReadFromExcel_ReadsExcelWithOneLineOfData_ReturnsOneElementWithCorrectValues()
    {
        //Arrange
        string excelFilePath = FilePaths.OneLineDataExcel;
        using var stream = FilePaths.OpenReadStream(excelFilePath);
        var expectedResult = new HeaderTestClass[]
        {
            new()
            {
                StringTest = "Hello world",
                IntTest = 123,
                BoolTest = true,
                FloatTest = 1.1234f,
                DateTimeTest = new DateTime(2024, 2, 16, 14, 55, 45, DateTimeKind.Utc),
                DoubleTest = 3385667.12345678912345,
                CustomStringTest = "Duis aute irure dolor in reprehenderit in voluptate velit."
            }
        };

        //Act
        var data = Excel.ReadFromExcel<HeaderTestClass>(stream).ToArray();

        //Assert
        data.Should().BeEquivalentTo(expectedResult);
    }

    [Fact]
    public void ReadFromExcel_ReadsExcelWithMultipleLineOfData_ReturnsAllElementsWithCorrectValues()
    {
        //Arrange
        string excelFilePath = FilePaths.MultiLineDateExcel;
        using var stream = FilePaths.OpenReadStream(excelFilePath);
        var expectedResult = new HeaderTestClass[]
        {
            new()
            {
                StringTest = "Lorem ipsum",
                IntTest = 743,
                BoolTest = false,
                FloatTest = 0.12345f,
                DateTimeTest = new DateTime(1999, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                DoubleTest = -16.99998,
                CustomStringTest = "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
            },
            new()
            {
                StringTest = "it to make a type specimen book.",
                IntTest = -9876,
                BoolTest = true,
                FloatTest = -5555.123f,
                DateTimeTest = new DateTime(2300, 5, 2, 2, 10, 59, DateTimeKind.Utc),
                DoubleTest = 1000,
                CustomStringTest = "Sed do eiusmod tempor incididunt ut labore et dolore."
            },
            new()
            {
                StringTest = "Ut suscipit ante vitae nisl fringilla, sit amet congue risus congue.",
                IntTest = 0,
                BoolTest = false,
                FloatTest = 987,
                DateTimeTest = new DateTime(2024, 6, 12, 0, 0, 0, DateTimeKind.Utc),
                DoubleTest = 99999.999,
                CustomStringTest = "Ut enim ad minim veniam, quis nostrud exercitation."
            }
        };

        //Act
        var data = Excel.ReadFromExcel<HeaderTestClass>(stream).ToArray();

        //Assert
        data.Should().BeEquivalentTo(expectedResult);
    }

    [Fact]
    public void ReadFromExcel_NoValidPropertyInClass_ThrowsNoValidPropertyException()
    {
        //Arrange
        string excelFilePath = FilePaths.EmptyExcel;
        using var stream = FilePaths.OpenReadStream(excelFilePath);

        //Act
        var act = () => Excel.ReadFromExcel<NoValidPropertyClass>(stream).ToArray();

        //Assert
        act.Should().Throw<NoValidPropertyException>();
    }

    [Fact]
    public void ReadFromExcel_SheetDoesNotExist_ThrowsSheetNotFoundException()
    {
        //Arrange
        string excelFilePath = FilePaths.EmptyExcel;
        using var stream = FilePaths.OpenReadStream(excelFilePath);

        //Act
        var act = () => Excel.ReadFromExcel<HeaderTestClass>(stream, "NotExist").ToArray();

        //Assert
        act.Should().Throw<SheetNotFoundException>();
    }

    [Fact]
    public void ReadFromExcel_HasOneSheet_FindsSheetByName()
    {
        //Arrange
        string excelFilePath = FilePaths.OneLineDataExcel;
        using var stream = FilePaths.OpenReadStream(excelFilePath);

        //Act
        var act = () => Excel.ReadFromExcel<HeaderTestClass>(stream, "Sheet1").ToArray();

        //Assert
        act.Should().NotThrow<SheetNotFoundException>();
    }

    [Fact]
    public void ReadFromExcel_HasMultipleSheets_FindsSheetByName()
    {
        //Arrange
        string excelFilePath = FilePaths.MultiSheetExcel;
        using var stream = FilePaths.OpenReadStream(excelFilePath);

        //Act
        var act = () => Excel.ReadFromExcel<HeaderTestClass>(stream, "Custom Name").ToArray();

        //Assert
        act.Should().NotThrow<SheetNotFoundException>();
    }

    [Fact]
    public void ReadFromExcel_HasNoHeader_ReadsData()
    {
        //Arrange
        string excelFilePath = FilePaths.NoHeaderExcel;
        using var stream = FilePaths.OpenReadStream(excelFilePath);
        var expectedResult = new NoHeaderTestClass[]
        {
            new()
            {
                StringTest = "Hello world",
                DoubleTest = 123.987,
                DateTimeTest = new DateTime(1982, 7, 30, 22, 12, 2, DateTimeKind.Utc)
            }
        };

        //Act
        var data = Excel.ReadFromExcel<NoHeaderTestClass>(stream, hasHeader: false).ToArray();

        //Assert
        data.Should().BeEquivalentTo(expectedResult);
    }

    private class HeaderTestClass
    {
        public int EmptyTest { get; set; }

        public string? StringTest { get; set; }

        public int IntTest { get; set; }

        public bool BoolTest { get; set; }

        public float FloatTest { get; set; }

        public double DoubleTest { get; set; }

        public DateTime DateTimeTest { get; set; }

        [ExcelColumn("Custom string property name")]
        public string? CustomStringTest { get; set; }
    }

    private class NoHeaderTestClass
    {
        [ExcelColumn(columnReference: "C")]
        public DateTime DateTimeTest { get; set; }

        [ExcelColumn(columnReference: "B")]
        public string? StringTest { get; set; }

        [ExcelColumn(columnReference: "A")]
        public double DoubleTest { get; set; }
    }

    private class NoValidPropertyClass
    {
        private int PrivateProperty { get; set; }

        protected int ProtectedProperty { get; set; }

        internal int InternalProperty { get; set; }

        public static string? StaticProperty { get; set; }
    }
}