using System;

namespace ExcelIO.NetCore
{
    public interface IExcelPlugin : IDisposable
    {
        int LastColumnIndex(object sheet, int rowIndex);
        int LastRowIndex(object sheet);

        object CreateWorkbook(string excelPath);
        object CreateWorksheet(object workbook, int sheetIndex, string sheetName);

        object GetWorkbook(string excelPath);
        object GetWorksheet(object workbook, string sheetName);
        object GetWorksheet(object workbook, int sheetIndex);

        object GetRow(object worksheet, int rowIndex);
        string GetValue(object row, int columnIndex);
        byte[] GetExcelData(object workbook);
        object SetSheetName(object workbook, int sheetIndex, string sheetName);
        void SetValue(object row, CellProperty cellProperty);
        void Save();
    }

}
