namespace Rmdev.ClosedXml.ExcelExtensions.Tests
{

    public class AttributesTests
    {
        [Fact(DisplayName = "The header name must come from the DisplayAttribute.")]
        [Trait("Category", "Attributes")]
        public void ClassWithAttributes_ExportToExcel_UseDisplayAttribute()
        {
            // Arrange
            List<SampleModel> data = [new(), new(), new()];

            // Act
            var excel = data.ExportToExcel();

            // Assert
            var displayHeader = excel.Worksheet(1).Cell(1, 1).Value.GetText();
            Assert.Equal("DisplayAttribute", displayHeader);
        }

        [Fact(DisplayName = "The header name must come from the DisplayNameAttribute.")]
        [Trait("Category", "Attributes")]
        public void ClassWithAttributes_ExportToExcel_UseDisplayNameAttribute()
        {
            // Arrange
            List<SampleModel> data = [new(), new(), new()];

            // Act
            var excel = data.ExportToExcel();

            // Assert
            var displayHeader = excel.Worksheet(1).Cell(1, 2).Value.GetText();
            Assert.Equal("DisplayNameAttribute", displayHeader);
        }

        [Fact(DisplayName = "The header name must come from the resouces file.")]
        [Trait("Category", "Attributes")]
        public void ClassWithAttributes_ExportToExcel_UseResourcesFile()
        {
            // Arrange
            List<SampleModel> data = [new(), new(), new()];

            // Act
            var excel = data.ExportToExcel();

            // Assert
            var displayHeader = excel.Worksheet(1).Cell(1, 3).Value.GetText();
            Assert.Equal(Translations.NameOfColumnC, displayHeader);
        }

        [Fact(DisplayName = "The header name must be the same as property name.")]
        [Trait("Category", "Attributes")]
        public void ClassWithAttributes_ExportToExcel_UsePropertyName()
        {
            // Arrange
            List<SampleModel> data = [new(), new(), new()];

            // Act
            var excel = data.ExportToExcel();

            // Assert
            var displayHeader = excel.Worksheet(1).Cell(1, 4).Value.GetText();
            Assert.Equal(nameof(SampleModel.ColumnD), displayHeader);
        }
    }
}