using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelIO.Framework
{
    public interface IExcelDataIO : IDisposable
    {
        /// <summary>
        /// Set the relevant properties before exporting the data to Excel
        /// </summary>
        /// <param name="excelSheet">Set the ordinal number of the column header row (starting from 0 by default), the column header text mapping relationship, the column header style, and other related properties</param>
        /// <param name="excelPath">Set the physical path to Excel saved [nullable]</param>
        void ToExcelWithProperty(ExcelSheet excelSheet, string excelPath);

        /// <summary>
        /// Set the relevant properties before exporting the data to Excel
        /// </summary>
        /// <param name="excelSheet">Set the ordinal number of the column header row (starting from 0 by default), the column header text mapping relationship, the column header style, and other related properties</param>
        /// <param name="excelPath">Set the physical path to Excel saved [nullable]</param>
        /// <param name="appendToLastSheet">Whether to add 'WorkSheet' after the last 'WorkSheet', default is No (false)</param>
        void ToExcelWithProperty(ExcelSheet excelSheet, string excelPath, bool appendToLastSheet);

        /// <summary>
        /// Integrate formal data into Excel
        /// </summary>
        /// <param name="excelRowChildren">Insert formal data with column header mappings into Excel</param>
        void ToExcelWithData(ExcelRowChildren excelRowChildren);

        /// <summary>
        /// Gets the assignment object with a column header mapping
        /// </summary>
        /// <returns>An assignment object with a column header mapping</returns>
        ExcelRowChildren ToExcelWithExcelRowChildren();

        /// <summary>
        /// Converts Excel to a byte array
        /// </summary>
        /// <returns>A byte array</returns>
        byte[] ToExcelGetBody();

        /// <summary>
        /// Gets a byte array based on the specified Excel physical path
        /// </summary>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <returns>A byte array</returns>
        byte[] ToExcelGetBody(string excelPath);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <typeparam name="T">The element object type of the List collection</typeparam>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <returns>Returns the List collection of data entities</returns>
        List<T> FromExcel<T>(ExcelSheet excelSheet, string excelPath);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <typeparam name="T">The element object type of the List collection</typeparam>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="execlStream">The data stream of excel</param>
        /// <returns></returns>
        List<T> FromExcelStream<T>(ExcelSheet excelSheet, Stream execlStream);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <typeparam name="T">The element object type of the List collection</typeparam>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelFileData">Provides a byte array of file of excel.</param>
        /// <returns></returns>
        List<T> FromExcelBytes<T>(ExcelSheet excelSheet, byte[] excelFileData);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <returns>Returns a DataTable data collection</returns>
        DataTable FromExcel(ExcelSheet excelSheet, string excelPath);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="execlStream">The data stream of excel</param>
        /// <returns>Returns a DataTable data collection</returns>
        DataTable FromExcelStream(ExcelSheet excelSheet, Stream execlStream);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelFileData">Provides a byte array of file of excel.</param>
        /// <returns>Returns a DataTable data collection</returns>
        DataTable FromExcelBytes(ExcelSheet excelSheet, byte[] excelFileData);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <param name="dataTableHeaderOfExcelHeader">Whether to set Excel's column header file to the column header name of DataTable, the default is No</param>
        /// <returns>Returns a DataTable data collection</returns>
        DataTable FromExcel(ExcelSheet excelSheet, string excelPath, bool dataTableHeaderOfExcelHeader);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="execlStream">The data stream of excel</param>
        /// <param name="dataTableHeaderOfExcelHeader">Whether to set Excel's column header file to the column header name of DataTable, the default is No</param>
        /// <returns>Returns a DataTable data collection</returns>
        DataTable FromExcelStream(ExcelSheet excelSheet, Stream execlStream, bool dataTableHeaderOfExcelHeader);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelFileData">Provides a byte array of file of excel.</param>
        /// <param name="dataTableHeaderOfExcelHeader">Whether to set Excel's column header file to the column header name of DataTable, the default is No</param>
        /// <returns>Returns a DataTable data collection</returns>
        DataTable FromExcelBytes(ExcelSheet excelSheet, byte[] excelFileData, bool dataTableHeaderOfExcelHeader);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <param name="action">Provide an Action parameter that receives the DataRow data row</param>
        void FromExcel(ExcelSheet excelSheet, string excelPath, Action<DataRow> action);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="execlStream">The data stream of excel</param>
        /// <param name="action">Provide an Action parameter that receives the DataRow data row</param>
        void FromExcelStream(ExcelSheet excelSheet, Stream execlStream, Action<DataRow> action);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelFileData">Provides a byte array of file of excel.</param>
        /// <param name="action">Provide an Action parameter that receives the DataRow data row</param>
        void FromExcelBytes(ExcelSheet excelSheet, byte[] excelFileData, Action<DataRow> action);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <typeparam name="T">The type of data entity</typeparam>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <param name="action">Provides an Action parameter for a data behavior data entity type</param>
        void FromExcel<T>(ExcelSheet excelSheet, string excelPath, Action<T> action);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <typeparam name="T">The type of data entity</typeparam>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="execlStream">The data stream of excel</param>
        /// <param name="action">Provides an Action parameter for a data behavior data entity type</param>
        void FromExcelStream<T>(ExcelSheet excelSheet, Stream execlStream, Action<T> action);

        /// <summary>
        /// Gets Excel data based on the properties you set
        /// </summary>
        /// <typeparam name="T">The type of data entity</typeparam>
        /// <param name="excelSheet">Provides information such as WorkSheet and column headers for Excel</param>
        /// <param name="excelFileData">Provides a byte array of file of excel.</param>
        /// <param name="action">Provides an Action parameter for a data behavior data entity type</param>
        void FromExcelBytes<T>(ExcelSheet excelSheet, byte[] excelFileData, Action<T> action);

        /// <summary>
        /// Gets the row data for the specified data row location based on the provided 'Worksheet' position or name information
        /// </summary>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <param name="rowIndex">The sequence number of the data row, which starts at 0 by default</param>
        /// <returns>Returns an array collection of row data</returns>
        string[] GetRowData(string excelPath, int rowIndex);

        /// <summary>
        /// Gets the row data for the specified data row location based on the provided 'Worksheet' position or name information
        /// </summary>
        /// <param name="execlStream">The data stream of excel</param>
        /// <param name="rowIndex">The sequence number of the data row, which starts at 0 by default</param>
        /// <returns>Returns an array collection of row data</returns>
        string[] GetRowDataFromStream(Stream execlStream, int rowIndex);

        /// <summary>
        /// Gets the row data for the specified data row location based on the provided 'Worksheet' position or name information
        /// </summary>
        /// <param name="excelFileData">Provides a byte array of file of excel.</param>
        /// <param name="rowIndex">The sequence number of the data row, which starts at 0 by default</param>
        /// <returns>Returns an array collection of row data</returns>
        string[] GetRowDataFromBytes(byte[] excelFileData, int rowIndex);

        /// <summary>
        /// Gets the row data for the specified data row location based on the provided 'Worksheet' position or name information
        /// </summary>
        /// <param name="excelSheet">Provide the location or name of 'Worksheet'</param>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <param name="rowIndex">The sequence number of the data row, which starts at 0 by default</param>
        /// <returns>Returns an array collection of row data</returns>
        string[] GetRowData(ExcelSheet excelSheet, string excelPath, int rowIndex);

        /// <summary>
        /// Gets the row data for the specified data row location based on the provided 'Worksheet' position or name information
        /// </summary>
        /// <param name="excelSheet">Provide the location or name of 'Worksheet'</param>
        /// <param name="execlStream">The data stream of excel</param>
        /// <param name="rowIndex">The sequence number of the data row, which starts at 0 by default</param>
        /// <returns>Returns an array collection of row data</returns>
        string[] GetRowDataFromStream(ExcelSheet excelSheet, Stream execlStream, int rowIndex);

        /// <summary>
        /// Gets the row data for the specified data row location based on the provided 'Worksheet' position or name information
        /// </summary>
        /// <param name="excelSheet">Provide the location or name of 'Worksheet'</param>
        /// <param name="excelFileData">Provides a byte array of file of excel.</param>
        /// <param name="rowIndex">The sequence number of the data row, which starts at 0 by default</param>
        /// <returns>Returns an array collection of row data</returns>
        string[] GetRowDataFromBytes(ExcelSheet excelSheet, byte[] excelFileData, int rowIndex);

        /// <summary>
        /// Gets the row data of key-value for the specified data row location based on the provided 'Worksheet' position or name information
        /// </summary>
        /// <param name="excelSheet">Provide the location or name of 'Worksheet', and provide the mapping relationship between column headers and field names (if not provided, the column header text will be used as the key)</param>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <param name="rowIndex">The sequence number of the data row, which starts at 0 by default</param>
        /// <returns>Returns row data combined as key-value pairs</returns>
        Dictionary<string, string> GetRowDataKayValue(ExcelSheet excelSheet, string excelPath, int rowIndex);

        /// <summary>
        /// Gets the row data of key-value for the specified data row location based on the provided 'Worksheet' position or name information
        /// </summary>
        /// <param name="excelSheet">Provide the location or name of 'Worksheet', and provide the mapping relationship between column headers and field names (if not provided, the column header text will be used as the key)</param>
        /// <param name="execlStream">The data stream of excel</param>
        /// <param name="rowIndex">The sequence number of the data row, which starts at 0 by default</param>
        /// <returns>Returns row data combined as key-value pairs</returns>
        Dictionary<string, string> GetRowDataKayValueFromStream(ExcelSheet excelSheet, Stream execlStream, int rowIndex);

        /// <summary>
        /// Gets the row data of key-value for the specified data row location based on the provided 'Worksheet' position or name information
        /// </summary>
        /// <param name="excelSheet">Provide the location or name of 'Worksheet', and provide the mapping relationship between column headers and field names (if not provided, the column header text will be used as the key)</param>
        /// <param name="excelFileData">Provides a byte array of file of excel.</param>
        /// <param name="rowIndex">The sequence number of the data row, which starts at 0 by default</param>
        /// <returns>Returns row data combined as key-value pairs</returns>
        Dictionary<string, string> GetRowDataKayValueFromBytes(ExcelSheet excelSheet, byte[] excelFileData, int rowIndex);

        /// <summary>
        ///  Get all the 'Worksheet' names in 'Workbook'
        /// </summary>
        /// <param name="excelPath">The physical path to Excel</param>
        /// <returns>Get all names of worksheet</returns>
        string[] GetWorksheetNames(string excelPath);

        /// <summary>
        /// Get all the 'Worksheet' names in 'Workbook'
        /// </summary>
        /// <param name="execlStream">The data stream of excel</param>
        /// <returns>Get all names of worksheet</returns>
        string[] GetWorksheetNamesFromStream(Stream execlStream);

        /// <summary>
        /// Get all the 'Worksheet' names in 'Workbook'
        /// </summary>
        /// <param name="excelFileData">Provides a byte array of file of excel.</param>
        /// <returns>Get all names of worksheet</returns>
        string[] GetWorksheetNamesFromBytes(byte[] excelFileData);
    }
}
