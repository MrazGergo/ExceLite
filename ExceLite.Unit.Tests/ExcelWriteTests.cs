using ExceLite.Unit.Tests.ExcelFiles;
using FluentAssertions;

namespace ExceLite.Unit.Tests;

public class ExcelWriteTests : IDisposable
{
    private readonly string excelFilePath;

    public ExcelWriteTests()
    {
        excelFilePath = FilePaths.RandomFilePath;
    }

    public void Dispose()
    {
        if (File.Exists(excelFilePath))
        {
            File.Delete(excelFilePath);
        }
    }

    [Fact]
    public void WritesMultipleDataWithoutAttributes_WritesDataInRandomOrder()
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
        using (var stream = FilePaths.OpenReadWriteStream(excelFilePath))
        {
            Excel.WriteToExcel(stream, data, addHeader: hasHeader);
        }

        //Assert
        using var readStream = FilePaths.OpenReadStream(excelFilePath);
        var readData = Excel.ReadFromExcel<TestClass>(readStream, hasHeader: hasHeader).ToArray();
        data.Should().BeEquivalentTo(readData);
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
