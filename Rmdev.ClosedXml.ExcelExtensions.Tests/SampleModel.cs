using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Rmdev.ClosedXml.ExcelExtensions.Tests
{
    public class SampleModel
    {
        [Display(Name = "DisplayAttribute")]
        public string ColumnA { get; set; } = "A";

        [DisplayName("DisplayNameAttribute")]
        public string ColumnB { get; set; } = "B";

        [Display(Name = "NameOfColumnC", ResourceType = typeof(Translations))]
        public string ColumnC { get; set; } = "C";

        // No Attributes
        public string ColumnD { get; set; } = "D";
    }
}