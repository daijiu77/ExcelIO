using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelIO.NetCore
{
    public interface IExcelDataIO : IDisposable
    {
        void ToExcelWithProperty(ExcelSheet excelSheet, string excelPath);
        void ToExcelWithData(ExcelRowChildren excelRowChildren);
        ExcelRowChildren ToExcelWithExcelRowChildren();
        byte[] ToExcelGetBody();

        List<T> FromExcel<T>(ExcelSheet excelSheet, string excelPath);
        DataTable FromExcel(ExcelSheet excelSheet, string excelPath);
        DataTable FromExcel(ExcelSheet excelSheet, string excelPath, bool dataTableHeaderOfExcelHeader);

        void FromExcel(ExcelSheet excelSheet, string excelPath, Action<DataRow> action);
        void FromExcel<T>(ExcelSheet excelSheet, string excelPath, Action<T> action);

        string[] GetRowData(string excelPath, int rowIndex);

    }
}
