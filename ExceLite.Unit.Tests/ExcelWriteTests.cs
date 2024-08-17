using ExceLite.Unit.Tests.ExcelFiles;
using FluentAssertions;

namespace ExceLite.Unit.Tests;

public class ExcelWriteTests : IDisposable
{
    private readonly string _excelFilePath;

    public ExcelWriteTests()
    {
        _excelFilePath = FilePaths.RandomFilePath;
    }

    public void Dispose()
    {
        if (File.Exists(_excelFilePath))
        {
            File.Delete(_excelFilePath);
        }
    }

    [Fact]
    public void WriteToExcel_WritesMultipleDataWithoutAttributes_WritesDataInRandomOrder()
    {
        //Arrange
        const bool hasHeader = true;
        var data = new List<TestClass>
        {
            new()
            {
                StringTest = "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
                IntTest = 123,
                BoolTest = true,
                FloatTest = 1.1234f,
                DateTimeTest = new DateTime(2024, 2, 16, 14, 55, 45, DateTimeKind.Utc),
                DoubleTest = 3385667.12345678912345,
            },
            new()
            {
                StringTest = "Duis aute irure dolor in reprehenderit in voluptate velit.",
                IntTest = 0,
                BoolTest = false,
                FloatTest = 9876f,
                DateTimeTest = new DateTime(1900, 10, 1, 1, 2, 3, DateTimeKind.Utc),
                DoubleTest = 12345.6789,
            },
            new()
            {
                StringTest = "Sed do eiusmod tempor incididunt ut labore et dolore.",
                IntTest = -9876,
                BoolTest = true,
                FloatTest = -5555.123f,
                DateTimeTest = new DateTime(2300, 5, 2, 2, 10, 59, DateTimeKind.Utc),
                DoubleTest = 1000,
            }
        };

        //Act
        using (var stream = FilePaths.OpenReadWriteStream(_excelFilePath))
        {
            Excel.WriteToExcel(stream, data, addHeader: hasHeader);
        }

        //Assert
        using var readStream = FilePaths.OpenReadStream(_excelFilePath);
        var readData = Excel.ReadFromExcel<TestClass>(readStream, hasHeader: hasHeader).ToArray();
        readData.Should().BeEquivalentTo(data);
    }

    [Fact]
    public void WriteToExcel_WritesDataWithCustomNames_AddsHeaderWithCustomNames()
    {
        //Arrange
        const bool hasHeader = true;
        var data = new List<CustomHeaderTestClass>
        {
            new()
            {
                StringTest = "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
                IntTest = 123,
                BoolTest = true,
                FloatTest = 1.1234f,
                DateTimeTest = new DateTime(2024, 2, 16, 14, 55, 45, DateTimeKind.Utc),
                DoubleTest = 3385667.12345678912345,
            }
        };

        //Act
        using (var stream = FilePaths.OpenReadWriteStream(_excelFilePath))
        {
            Excel.WriteToExcel(stream, data, addHeader: hasHeader);
        }

        //Assert
        using var readStream = FilePaths.OpenReadStream(_excelFilePath);
        var readData = Excel.ReadFromExcel<CustomHeaderTestClass>(readStream, hasHeader: hasHeader).ToArray();
        readData.Should().BeEquivalentTo(data);
    }

    [Fact]
    public void WriteToExcel_WritesDataWithFullySpecifiedCustomReferences_AddsValuesToCorrectColumns()
    {
        //Arrange
        const bool hasHeader = false;
        var data = new List<CustomReferenceFullTestClass>
        {
            new()
            {
                StringTest = "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
                IntTest = 123,
                BoolTest = true,
                FloatTest = 1.1234f,
                DateTimeTest = new DateTime(2024, 2, 16, 14, 55, 45, DateTimeKind.Utc),
                DoubleTest = 3385667.12345678912345,
            }
        };

        //Act
        using (var stream = FilePaths.OpenReadWriteStream(_excelFilePath))
        {
            Excel.WriteToExcel(stream, data, addHeader: hasHeader);
        }

        //Assert
        using var readStream = FilePaths.OpenReadStream(_excelFilePath);
        var readData = Excel.ReadFromExcel<CustomReferenceFullTestClass>(readStream, hasHeader: hasHeader).ToArray();
        readData.Should().BeEquivalentTo(data);
    }

    [Fact]
    public void WriteToExcel_WritesDataWithPartiallySpecifiedCustomReferences_AddsRandomReferencesToMissingOnes()
    {
        //Arrange
        const bool hasHeader = true;
        var data = new List<CustomReferencePartialTestClass>
        {
            new()
            {
                StringTest = "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
                IntTest = 123,
                BoolTest = true,
                FloatTest = 1.1234f,
                DateTimeTest = new DateTime(2024, 2, 16, 14, 55, 45, DateTimeKind.Utc),
                DoubleTest = 3385667.12345678912345,
            }
        };

        //Act
        using (var stream = FilePaths.OpenReadWriteStream(_excelFilePath))
        {
            Excel.WriteToExcel(stream, data, addHeader: hasHeader);
        }

        //Assert
        using var readStream = FilePaths.OpenReadStream(_excelFilePath);
        var readData = Excel.ReadFromExcel<CustomReferencePartialTestClass>(readStream, hasHeader: hasHeader).ToArray();
        readData.Should().BeEquivalentTo(data);
    }

    [Fact]
    public void WriteToExcel_CustomSheetNameIsProvided_CreatesSheetWithCustomSheetName()
    {
        //Arrange
        const bool hasHeader = true;
        const string sheetName = "Custom Sheet Name";
        var data = new List<TestClass>
        {
            new()
            {
                StringTest = "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
                IntTest = 123,
                BoolTest = true,
                FloatTest = 1.1234f,
                DateTimeTest = new DateTime(2024, 2, 16, 14, 55, 45, DateTimeKind.Utc),
                DoubleTest = 3385667.12345678912345,
            }
        };

        //Act
        using (var stream = FilePaths.OpenReadWriteStream(_excelFilePath))
        {
            Excel.WriteToExcel(stream, data, addHeader: hasHeader, sheetName: sheetName);
        }

        //Assert
        using var readStream = FilePaths.OpenReadStream(_excelFilePath);
        var readData = Excel.ReadFromExcel<TestClass>(readStream, hasHeader: hasHeader, sheetName: sheetName).ToArray();
        readData.Should().BeEquivalentTo(data);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void WriteToExcel_EmptyCollectionProvided_CreatesEmptyExcelFile(bool hasHeader)
    {
        //Arrange
        var data = new List<TestClass>();

        //Act
        using (var stream = FilePaths.OpenReadWriteStream(_excelFilePath))
        {
            Excel.WriteToExcel(stream, data, addHeader: hasHeader);
        }

        //Assert
        File.Exists(_excelFilePath).Should().BeTrue();
    }

    private class CustomReferencePartialTestClass
    {
        [ExcelColumn(columnName: "Empty test", columnReference: "D")]
        public string? EmptyTest { get; set; }

        [ExcelColumn(columnName: "String test", columnReference: "A")]
        public string? StringTest { get; set; }

        public int IntTest { get; set; }

        [ExcelColumn(columnName: "Bool test", columnReference: "C")]
        public bool BoolTest { get; set; }

        public float FloatTest { get; set; }

        public DateTime DateTimeTest { get; set; }

        [ExcelColumn(columnName: "Double test", columnReference: "G")]
        public double DoubleTest { get; set; }
    }

    private class CustomReferenceFullTestClass
    {
        [ExcelColumn(columnName: "Empty test", columnReference: "D")]
        public string? EmptyTest { get; set; }

        [ExcelColumn(columnName: "String test", columnReference: "A")]
        public string? StringTest { get; set; }

        [ExcelColumn(columnName: "Int test", columnReference: "E")]
        public int IntTest { get; set; }

        [ExcelColumn(columnName: "Bool test", columnReference: "C")]
        public bool BoolTest { get; set; }

        [ExcelColumn(columnName: "Float test", columnReference: "B")]
        public float FloatTest { get; set; }

        [ExcelColumn(columnName: "Date test", columnReference: "F")]
        public DateTime DateTimeTest { get; set; }

        [ExcelColumn(columnName: "Double test", columnReference: "G")]
        public double DoubleTest { get; set; }
    }

    private class CustomHeaderTestClass
    {
        [ExcelColumn(columnName: "Empty test")]
        public string? EmptyTest { get; set; }

        [ExcelColumn(columnName: "String test")]
        public string? StringTest { get; set; }

        [ExcelColumn(columnName: "Int test")]
        public int IntTest { get; set; }

        [ExcelColumn(columnName: "Bool test")]
        public bool BoolTest { get; set; }

        [ExcelColumn(columnName: "Float test")]
        public float FloatTest { get; set; }

        [ExcelColumn(columnName: "Date test")]
        public DateTime DateTimeTest { get; set; }

        [ExcelColumn(columnName: "Double test")]
        public double DoubleTest { get; set; }
    }

    private class TestClass
    {
        public string? EmptyTest { get; set; }
        public string? StringTest { get; set; }
        public int IntTest { get; set; }
        public bool BoolTest { get; set; }
        public float FloatTest { get; set; }
        public DateTime DateTimeTest { get; set; }
        public double DoubleTest { get; set; }
    }
}
