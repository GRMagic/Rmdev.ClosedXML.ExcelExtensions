using ClosedXML.Excel;

namespace Rmdev.ClosedXml.ExcelExtensions.Tests
{
    public class ValuesTests
    {
        [Fact(DisplayName = "The values are in the right cell.")]
        [Trait("Category", "Values")]
        public void Threelines_ExportToExcel_ValuesAtRightPlace()
        {
            // Arrange
            List<SampleModel> data = [new() { ColumnA = "first" }, new(), new() { ColumnD = "last"}];

            // Act
            var excel = data.ExportToExcel();

            // Assert
            var firstValue = excel.Worksheet(1).Cell(2, 1).Value.GetText();
            var lastValue = excel.Worksheet(1).Cell(4, 4).Value.GetText();
            Assert.Equal(data[0].ColumnA, firstValue);
            Assert.Equal(data[2].ColumnD, lastValue);
        }

        [Fact(DisplayName = "Export from anomymous types.")]
        [Trait("Category", "Values")]
        public void AnonymousType_ExportToExcel_NoProblems()
        {
            // Arrange
            var data = Enumerable.Range(1, 3).Select(i => new
            {
                Index = i,
                Description = $"The {i}º item."
            }).ToArray();

            // Act
            var excel = data.ExportToExcel();

            // Assert
            var firstValue = excel.Worksheet(1).Cell(2, 1).Value.GetNumber();
            var lastValue = excel.Worksheet(1).Cell(4, 2).Value.GetText();
            Assert.Equal(1, firstValue);
            Assert.Equal("The 3º item.", lastValue);
        }
    }

    public class WebTests
    {
        [Fact(DisplayName = "Download a file with a name.")]
        [Trait("Category", "Web")]
        public void NameInformed_ToFileStreamResult_NameDefined()
        {
            // Arrange
            List<SampleModel> data = [new()];
            var excel = data.ExportToExcel();
            var fileName = "File.xlsx";

            // Act
            var result = excel.ToFileStreamResult(fileName);

            // Assert
            Assert.Equal(fileName, result.FileDownloadName);

            result.FileStream.Dispose();
        }

        [Fact(DisplayName = "Download a file with a name.")]
        [Trait("Category", "Web")]
        public void Workbook_ToFileStreamResult_HasData()
        {
            // Arrange
            List<SampleModel> data = [new()];
            var sourceExcel = data.ExportToExcel();
            var fileName = "File.xlsx";

            // Act
            var result = sourceExcel.ToFileStreamResult(fileName);

            // Assert
            using var resultExcel = new XLWorkbook(result.FileStream);

            var sourceCellsCount = sourceExcel.Worksheet(1).Cells(true).Count();
            var resultCellsCount = resultExcel.Worksheet(1).Cells(true).Count();
            Assert.Equal(sourceCellsCount, resultCellsCount);
            
            result.FileStream.Dispose();
        }
    }
}