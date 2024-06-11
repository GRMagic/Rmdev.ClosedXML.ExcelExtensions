# RM Dev - ClosedXML Extensions + Web

## Extensions

**All Extensions** from [Rmdev.ClosedXML.ExcelExtensions](..\Rmdev.ClosedXML.ExcelExtensions\readme.md) plus ToFileStreamResult


### ToFileStreamResult

Create a new file stream result from the XLWorkbook, which can be returned by a controller action.

```charp
FileStreamResult ToFileStreamResult(this XLWorkbook spreadsheet, string fileDownloadName)
```

