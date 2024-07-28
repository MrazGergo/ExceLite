using ExceLite.Unit.Tests.ExcelFiles;
using FluentAssertions;

namespace ExceLite.Unit.Tests;

public class ExcelTests
{
    [Fact]
    public void ReadFromExcel_ReadsEmptyExcel_ReturnsEmptyIEnumerable()
    {
        //Arrange
        string excelFilePath = FilePaths.EmptyExcel;
        var stream = FilePaths.OpenReadStream(excelFilePath);

        //Act
        var data = Excel.ReadFromExcel<TestClass>(stream);

        //Assert
        data.Should().BeEmpty();
    }

    [Fact]
    public void ReadFromExcel_ReadsOnlyHeaderExcel_ReturnsEmptyIEnumerable()
    {
        //Arrange
        string excelFilePath = FilePaths.OnlyHeaderExcel;
        var stream = FilePaths.OpenReadStream(excelFilePath);

        //Act
        var data = Excel.ReadFromExcel<TestClass>(stream);

        //Assert
        data.Should().BeEmpty();

    }

    [Fact]
    public void ReadFromExcel_ReadsExcelWithOneLineOfData_ReturnsOneElementWithCorrectValues()
    {
        //Arrange
        string excelFilePath = FilePaths.OneLineDataExcel;
        var stream = FilePaths.OpenReadStream(excelFilePath);
        var expectedResult = new TestClass[]
        {
            new()
            {
                StringTest = "Hello world",
                IntTest = 123,
                BoolTest = true,
                FloatTest = 1.1234f,
                DateTimeTest = new DateTime(2024, 2, 16, 14, 55, 45, DateTimeKind.Utc),
                DoubleTest = 3385667.12345678912345
            }
        };

        //Act
        var data = Excel.ReadFromExcel<TestClass>(stream).ToArray();

        //Assert
        data.Should().BeEquivalentTo(expectedResult);
    }

    [Fact]
    public void ReadFromExcel_ReadsExcelWithMultipleLineOfData_ReturnsAllElementsWithCorrectValues()
    {
        //Arrange
        string excelFilePath = FilePaths.MultiLineDateExcel;
        var stream = FilePaths.OpenReadStream(excelFilePath);
        var expectedResult = new TestClass[]
        {
            new()
            {
                StringTest = "Lorem ipsum",
                IntTest = 743,
                BoolTest = false,
                FloatTest = 0.12345f,
                DateTimeTest = new DateTime(1999, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                DoubleTest = -16.99998
            },
            new()
            {
                StringTest = "it to make a type specimen book.",
                IntTest = -9876,
                BoolTest = true,
                FloatTest = -5555.123f,
                DateTimeTest = new DateTime(2300, 5, 2, 2, 10, 59, DateTimeKind.Utc),
                DoubleTest = 1000
            },
            new()
            {
                StringTest = "Ut suscipit ante vitae nisl fringilla, sit amet congue risus congue.",
                IntTest = 0,
                BoolTest = false,
                FloatTest = 987,
                DateTimeTest = new DateTime(2024, 6, 12, 0, 0, 0, DateTimeKind.Utc),
                DoubleTest = 99999.999
            }
        };

        //Act
        var data = Excel.ReadFromExcel<TestClass>(stream).ToArray();

        //Assert
        data.Should().BeEquivalentTo(expectedResult);
    }

    private class TestClass
    {
        public string? StringTest { get; set; }

        public int IntTest { get; set; }

        public bool BoolTest { get; set; }

        public float FloatTest { get; set; }

        public double DoubleTest { get; set; }

        public DateTime DateTimeTest { get; set; }
    }
}