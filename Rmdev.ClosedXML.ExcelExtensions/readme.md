# RM Dev - ClosedXML Extensions

## Extensions

### PrintTable

Print an enumerable as a table into worksheet.

```charp
void PrintTable<T>(this IXLWorksheet worksheet, IEnumerable<(string PropertyName, string HeaderName)> headers, IEnumerable<T> data, int originRow = 1, int originColumn = 1, IFormatProvider formatProvider = null)
```

### ExportToExcel

Creates a new Excel workbook with one worksheet and print an enumerable as a table into it.

```charp
XLWorkbook ExportToExcel<T>(this IEnumerable<T> data, IFormatProvider formatProvider = null)
```

### ExportToExcel

Creates a new Excel workbook with one worksheet and print an enumerable as a table into it.

The *headers* parameter can be used to filter properties and choose the column name.

```charp
XLWorkbook ExportToExcel<T>(this IEnumerable<T> data, IEnumerable<(string PropertyName, string HeaderName)> headers, IFormatProvider formatProvider = null)
```

## Attributes

You can use [DisplayName("Column Name")] or [Display(Name = "Column Name")] to define columns names. By default, the column names are the same as the property name.


## Examples

Check the tests project for examples!