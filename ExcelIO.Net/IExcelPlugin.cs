using System;
using System.IO;

namespace ExcelIO.Net
{
    public interface IExcelPlugin : IDisposable
    {
        int LastColumnIndex(object sheet, int rowIndex);
        int LastRowIndex(object sheet);

        object CreateWorkbook(string excelPath);

        object CreateWorkbookByStream(Stream excelStream);

        object CreateWorksheet(object workbook, int sheetIndex, string sheetName, bool appendToLastSheet);

        object GetWorkbook(string excelPath);
        object GetWorkbookByStream(Stream excelStream);
        object GetWorksheet(object workbook, string sheetName);
        object GetWorksheet(object workbook, int sheetIndex);

        string[] GetWorksheetNames(object workbook);

        object GetRow(object worksheet, int rowIndex);
        string GetValue(object row, int columnIndex);
        byte[] GetExcelData(object workbook);
        object SetSheetName(object workbook, int sheetIndex, string sheetName);
        void SetValue(object row, CellProperty cellProperty);
        void Save();
    }

}
