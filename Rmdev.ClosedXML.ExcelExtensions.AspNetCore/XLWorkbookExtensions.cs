using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace Rmdev.ClosedXML.ExcelExtensions
{
    public static class XLWorkbookExtensions
    {
        /// <summary>
        /// Create a new file stream result from XLWorkbook
        /// </summary>
        /// <param name="spreadsheet">Excel document</param>
        /// <param name="fileDownloadName">The file name that will be used in the Content-Disposition header of the response.</param>
        public static FileStreamResult ToFileStreamResult(this XLWorkbook spreadsheet, string fileDownloadName)
        {
            var stream = new MemoryStream();
            spreadsheet.SaveAs(stream);
            stream.Seek(0, SeekOrigin.Begin);
            return new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = fileDownloadName,
            };
        }
    }
}
