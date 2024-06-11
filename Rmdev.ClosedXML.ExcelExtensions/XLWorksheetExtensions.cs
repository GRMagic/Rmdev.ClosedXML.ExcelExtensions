using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace Rmdev.ClosedXML.ExcelExtensions
{
    public static class XLWorksheetExtensions
    {
        /// <summary>
        /// Print a table on worksheet
        /// </summary>
        /// <param name="headers">Property name and display name</param>
        /// <param name="data">Rows</param>
        /// <param name="originRow">Start position for row</param>
        /// <param name="originColumn">Start position for column</param>
        /// <exception cref="ArgumentOutOfRangeException">When there is no property with a specified name.</exception>
        public static void PrintTable<T>(this IXLWorksheet worksheet, IEnumerable<(string PropertyName, string HeaderName)> headers, IEnumerable<T> data, int originRow = 1, int originColumn = 1, IFormatProvider formatProvider = null)
        {
            formatProvider = formatProvider ?? System.Globalization.CultureInfo.CurrentUICulture;

            var column = originColumn;
            var row = originRow;

            foreach (var header in headers)
            {
                var cell = worksheet.Cell(row, column++);
                cell.Value = header.HeaderName;
                cell.Style.Font.Bold = true;
            }

            var properties = headers.Select(h => typeof(T).GetProperty(h.PropertyName)
                                              ?? throw new ArgumentOutOfRangeException(nameof(headers), $"Property {h.PropertyName} not found in type {typeof(T).Name}."))
                                    .ToList();

            foreach (var item in data)
            {
                row++;
                column = originColumn;
                foreach (var property in properties)
                {
                    var value = property.GetValue(item);
                    if (value != null)
                        worksheet.Cell(row, column).Value = XLCellValue.FromObject(value, formatProvider);
                    column++;
                }
            }
        }

        /// <summary>
        /// Creates a new Excel workbook with one worksheet and print an enumerable as a table into it
        /// </summary>
        /// <param name="data">Rows</param>
        /// <param name="headers">Property name and display name </param>
        /// <param name="formatProvider">Used for formatting. When this parameter is null, CurrentUICulture will be used by default.</param>
        public static XLWorkbook ExportToExcel<T>(this IEnumerable<T> data, IEnumerable<(string PropertyName, string HeaderName)> headers, IFormatProvider formatProvider = null)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add();
            worksheet.PrintTable(headers, data, 1, 1, formatProvider);
            worksheet.Columns(1, headers.Count()).AdjustToContents();
            return workbook;
        }

        /// <summary>
        /// Creates a new Excel workbook with one worksheet and print an enumerable as a table into it
        /// </summary>
        /// <param name="data">Rows</param>
        /// <param name="formatProvider">Used for formatting. When this parameter is null, CurrentUICulture will be used by default.</param>
        public static XLWorkbook ExportToExcel<T>(this IEnumerable<T> data, IFormatProvider formatProvider = null)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add();
            var headers = typeof(T).GetProperties().Select(p => (p.Name, p.GetDisplayName()));
            worksheet.PrintTable(headers, data, 1, 1, formatProvider);
            worksheet.Columns(1, headers.Count()).AdjustToContents();
            return workbook;
        }

        private static string GetDisplayName(this MemberInfo memberInfo)
        {
            return memberInfo.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? memberInfo.GetCustomAttribute<DisplayAttribute>()?.GetName() ?? memberInfo.Name;
        }
    }
}
