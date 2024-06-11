namespace Rmdev.ClosedXml.ExcelExtensions.Tests
{
    public class HeadersTests
    {
        [Fact(DisplayName = "The headers have been filtered.")]
        [Trait("Category", "Headers")]
        public void FourProperties_ExportToExcel_ThreeColumns()
        {
            // Arrange
            List<SampleModel> data = [new(), new(), new() ];

            // Act
            var excel = data.ExportToExcel(
                [
                    (nameof(SampleModel.ColumnA), "A"),
                    (nameof(SampleModel.ColumnB), "B"),
                    (nameof(SampleModel.ColumnD), "D")
                ]);

            // Assert
            var isBlank = excel.Worksheet(1).Cell(1, 4).Value.IsBlank;
            Assert.True(isBlank);
        }

        [Fact(DisplayName = "The headers have defined name.")]
        [Trait("Category", "Headers")]
        public void DefinedNames_ExportToExcel_UseDefinedNames()
        {
            // Arrange
            List<SampleModel> data = [new(), new(), new()];

            // Act
            var excel = data.ExportToExcel(
                [
                    (nameof(SampleModel.ColumnA), "A"),
                    (nameof(SampleModel.ColumnB), "B"),
                    (nameof(SampleModel.ColumnC), "The Name"),
                    (nameof(SampleModel.ColumnD), "D")
                ]);

            // Assert
            var headerName = excel.Worksheet(1).Cell(1, 3).Value.GetText();
            Assert.Equal("The Name", headerName);
        }

        [Fact(DisplayName = "If the property does not exist, throw an exception.")]
        [Trait("Category", "Headers")]
        public void InvalidName_ExportToExcel_ThrowsException()
        {
            // Arrange
            List<SampleModel> data = [new(), new(), new()];

            // Act
            var act = () => data.ExportToExcel(
                [
                    (nameof(SampleModel.ColumnA), "A"),
                    ("NotExists", "B"),
                    (nameof(SampleModel.ColumnC), "C"),
                    (nameof(SampleModel.ColumnD), "D")
                ]);

            // Assert
            Assert.Throws<ArgumentOutOfRangeException>(act);
        }

        [Fact(DisplayName = "Export from anomymous types.")]
        [Trait("Category", "Headers")]
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
            var secondHeader = excel.Worksheet(1).Cell(1, 2).Value.GetText();
            Assert.Equal("Description", secondHeader);
        }
    }
}