using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;

namespace ExcelIO.Framework
{
    public class ExcelDataIO : IExcelDataIO
    {
        private IExcelPlugin excelPlugin = null;
        private ExcelRowChildren excelRowChildren = new ExcelRowChildren();

        private object work_book = null;
        private object work_sheet = null;
        private List<CellProperty> headList = null;
        private CellProperty cellP = new CellProperty();
        private int rowIndex = 0;

        public ExcelDataIO(IExcelPlugin excelPlugin)
        {
            this.excelPlugin = excelPlugin;
        }

        public static object ConvertTo(object value, Type type)
        {
            object obj = null;
            object v = value;
            if (type == typeof(Guid?))
            {
                v = v == null ? Guid.Empty.ToString() : v;
                obj = new Guid(v.ToString());
            }
            else if (type == typeof(int?)
                || type == typeof(short?)
                || type == typeof(long?)
                || type == typeof(float?)
                || type == typeof(double?)
                || type == typeof(decimal?))
            {
                v = v == null ? 0 : v;
                value = v;
            }
            else if (type == typeof(bool?))
            {
                v = v == null ? false : v;
                value = v;
            }
            else if (type == typeof(DateTime?))
            {
                v = v == null ? Convert.ToDateTime("1900-01-01 00:00:00") : v;
                value = v;
            }

            if (null == obj)
            {
                string s = type.ToString();
                string typeName = s.Substring(s.LastIndexOf(".") + 1);
                typeName = typeName.Replace("]", "");
                string methodName = "To" + typeName;
                try
                {
                    Type t = Type.GetType("System.Convert");
                    obj = t.InvokeMember(methodName, BindingFlags.InvokeMethod | BindingFlags.Static | BindingFlags.Public, null, null, new object[] { value });
                }
                catch (Exception)
                {

                    //throw;
                }
            }

            return obj;
        }

        private bool isBaseType(Type type)
        {
            bool isObjectOrBaseType = false;
            if (!isObjectOrBaseType) isObjectOrBaseType = type.IsValueType;
            if (!isObjectOrBaseType) isObjectOrBaseType = type == typeof(string)
                    || (false == type.IsClass && false == type.IsInterface
                    && false == type.IsGenericType && false == type.IsAbstract
                    && false == type.IsArray);
            return isObjectOrBaseType;
        }

        private object GetWorkbook(string excelPath, Stream excelStream)
        {
            if (null == excelPlugin) throw new Exception("Plugin cann't is null");
            bool isFilePath = false;
            if (!string.IsNullOrEmpty(excelPath))
            {
                isFilePath = true;
                if (!File.Exists(excelPath)) throw new Exception("The source of excel cann't is empty");
            }

            if ((false == isFilePath) && (null == excelStream)) throw new Exception("The path or stream of excel cann't is empty");

            object workbook = null;
            try
            {
                if (null != excelStream)
                {
                    workbook = excelPlugin.GetWorkbookByStream(excelStream);
                }
                else
                {
                    workbook = excelPlugin.GetWorkbook(excelPath);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
            return workbook;
        }

        private void from_excel<T>(ExcelSheet excelSheet, string excelPath, Stream excelStream, Action<DataRow> action, Action<T> actionT, List<T> list, bool dataTableHeaderOfExcelHeader, ref DataTable dataTable)
        {
            object workbook = GetWorkbook(excelPath, excelStream);
            object worksheet = null;
            if (string.IsNullOrEmpty(excelSheet.SheetName))
            {
                worksheet = excelPlugin.GetWorksheet(workbook, 0);
            }
            else
            {
                worksheet = excelPlugin.GetWorksheet(workbook, excelSheet.SheetName);
            }

            string key = "";
            Dictionary<string, string> mpDic = new Dictionary<string, string>();
            foreach (CellProperty item in excelSheet)
            {
                key = item.headText.Trim();
                if (mpDic.ContainsKey(key)) continue;
                mpDic.Add(key, item.fieldName);
            }

            bool isNullDataTable = null == dataTable;
            if (null != action)
            {
                dataTable = null == dataTable ? new DataTable() : dataTable;
            }

            bool isInitColumn = false;
            if (null != dataTable) isInitColumn = 0 == dataTable.Columns.Count;
            object row = excelPlugin.GetRow(worksheet, excelSheet.HeadRowIndex);
            int len = excelPlugin.LastColumnIndex(worksheet, excelSheet.HeadRowIndex);
            int rowCount = excelPlugin.LastRowIndex(worksheet) + 1;
            len++;
            string[] heads = new string[len];
            string fieldName = "";
            for (int i = 0; i < len; i++)
            {
                key = excelPlugin.GetValue(row, i).Trim();
                heads[i] = key;
                if (!mpDic.ContainsKey(key)) continue;
                fieldName = mpDic[key];
                if (isInitColumn)
                {
                    if (dataTableHeaderOfExcelHeader)
                    {
                        dataTable.Columns.Add(key, typeof(string));
                    }
                    else
                    {
                        dataTable.Columns.Add(fieldName, typeof(string));
                    }
                }
            }

            bool isObjectOrBaseType = typeof(T) == typeof(object);
            if (!isObjectOrBaseType) isObjectOrBaseType = isBaseType(typeof(T));

            bool isList = null != list;

            T t = default(T);
            string cellVal = "";
            object val = null;
            DataRow dataRow = null;
            PropertyInfo pi = null;
            int startY = excelSheet.HeadRowIndex + 1;
            bool isEnabled = false;
            for (int y = startY; y < rowCount; y++)
            {
                row = excelPlugin.GetRow(worksheet, y);
                if (!isObjectOrBaseType)
                {
                    if ((null != actionT) || isList)
                    {
                        t = (T)Activator.CreateInstance(typeof(T));
                        isEnabled = true;
                    }
                }

                if (null != dataTable)
                {
                    if ((null != action) || (false == isNullDataTable))
                    {
                        dataRow = dataTable.NewRow();
                        isEnabled = true;
                    }
                }

                if (false == isEnabled) break;

                for (int x = 0; x < len; x++)
                {
                    key = heads[x];
                    if (!mpDic.ContainsKey(key)) continue;
                    fieldName = mpDic[key];
                    cellVal = excelPlugin.GetValue(row, x);

                    if (null != dataRow)
                    {
                        if (dataTableHeaderOfExcelHeader)
                        {
                            dataRow[key] = cellVal;
                        }
                        else
                        {
                            dataRow[fieldName] = cellVal;
                        }
                    }

                    if (!isObjectOrBaseType)
                    {
                        pi = t.GetType().GetProperty(fieldName);
                        if (null != pi)
                        {
                            if (isBaseType(pi.PropertyType))
                            {
                                val = ConvertTo(cellVal, pi.PropertyType);
                                try
                                {
                                    pi.SetValue(t, val);
                                }
                                catch (Exception) { }
                            }
                        }
                    }
                }

                if (null != action) action(dataRow);
                if (null != actionT && false == isObjectOrBaseType) actionT(t);
                if (false == isNullDataTable) dataTable.Rows.Add(dataRow);
                if (isList && (false == isObjectOrBaseType)) list.Add(t);
            }

            excelPlugin.Dispose();
        }

        void IExcelDataIO.FromExcel(ExcelSheet excelSheet, string excelPath, Action<DataRow> action)
        {
            DataTable dt = null;
            from_excel<object>(excelSheet, excelPath, null, action, null, null, false, ref dt);
        }

        List<T> IExcelDataIO.FromExcel<T>(ExcelSheet excelSheet, string excelPath)
        {
            List<T> list = new List<T>();
            DataTable dt = null;
            from_excel<T>(excelSheet, excelPath, null, null, null, list, false, ref dt);
            return list;
        }

        void IExcelDataIO.FromExcel<T>(ExcelSheet excelSheet, string excelPath, Action<T> action)
        {
            DataTable dt = null;
            from_excel<T>(excelSheet, excelPath, null, null, action, null, false, ref dt);
        }

        DataTable IExcelDataIO.FromExcel(ExcelSheet excelSheet, string excelPath)
        {
            DataTable dt = new DataTable();
            from_excel<object>(excelSheet, excelPath, null, null, null, null, false, ref dt);
            return dt;
        }

        DataTable IExcelDataIO.FromExcel(ExcelSheet excelSheet, string excelPath, bool dataTableHeaderOfExcelHeader)
        {
            DataTable dt = new DataTable();
            from_excel<object>(excelSheet, excelPath, null, null, null, null, dataTableHeaderOfExcelHeader, ref dt);
            return dt;
        }

        private Dictionary<string, string> GetRowData_KayValue(ExcelSheet excelSheet, string excelPath, Stream excelStream, int rowIndex)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            string[] rows = null;
            if (null != excelStream)
            {
                rows = ((IExcelDataIO)this).GetRowDataFromStream(excelSheet, excelStream, rowIndex);
            }
            else
            {
                rows = ((IExcelDataIO)this).GetRowData(excelSheet, excelPath, rowIndex);
            }
            string[][] cols = new string[rows.Length][];
            string[] cols1 = null;
            string txt = "";
            int index = excelSheet.HeadRowIndex, n = 0;
            Dictionary<string, int> keyValues = new Dictionary<string, int>();
            const int maxRowIndex = 10;
            while (index < maxRowIndex)
            {
                if (null != excelStream)
                {
                    cols1 = GetRow_Data(excelSheet, null, excelStream, index);
                }
                else
                {
                    cols1 = ((IExcelDataIO)this).GetRowData(excelSheet, excelPath, index);
                }
                txt = "";
                n = 0;
                keyValues.Clear();
                foreach (var item in cols1)
                {
                    txt = item.Trim();
                    if (string.IsNullOrEmpty(txt)) break;
                    if (n >= cols.Length) break;
                    if (null == cols[n]) cols[n] = new string[2];
                    cols[n][0] = txt;
                    keyValues[txt] = n;
                    n++;
                }
                if (!string.IsNullOrEmpty(txt)) break;
                index++;
            }

            if (keyValues.Count != rows.Length) return dic;

            bool mbool = false;
            if (0 < excelSheet.ExcelColumnsMappings.Count)
            {
                mbool = true;
                foreach (var item in excelSheet.ExcelColumnsMappings)
                {
                    if (!keyValues.ContainsKey(item.headText))
                    {
                        mbool = false;
                        break;
                    }
                    index = keyValues[item.headText];
                    cols[index][1] = item.fieldName;
                }
            }

            if (!mbool)
            {
                int len = cols.Length;
                for (int i = 0; i < len; i++)
                {
                    cols[i][1] = cols[i][0];
                }
            }

            if (cols.Length != rows.Length) return dic;
            index = 0;
            foreach (var item in cols)
            {
                dic[item[1]] = rows[index];
                index++;
            }
            return dic;
        }

        Dictionary<string, string> IExcelDataIO.GetRowDataKayValue(ExcelSheet excelSheet, string excelPath, int rowIndex)
        {
            return GetRowData_KayValue(excelSheet, excelPath, null, rowIndex);
        }

        private string[] GetRow_Data(ExcelSheet excelSheet, string excelPath, Stream excelStream, int rowIndex)
        {
            string[] results = null;
            object workbook = GetWorkbook(excelPath, excelStream);
            object worksheet = null;
            if (!string.IsNullOrEmpty(excelSheet.SheetName))
            {
                worksheet = excelPlugin.GetWorksheet(workbook, excelSheet.SheetName);
            }
            else
            {
                worksheet = excelPlugin.GetWorksheet(workbook, excelSheet.SheetIndex);
            }
            object row = excelPlugin.GetRow(worksheet, rowIndex);
            int len = excelPlugin.LastColumnIndex(worksheet, rowIndex);
            len++;
            results = new string[len];
            for (int i = 0; i < len; i++)
            {
                results[i] = excelPlugin.GetValue(row, i);
            }
            excelPlugin.Dispose();
            return results;
        }

        string[] IExcelDataIO.GetRowData(ExcelSheet excelSheet, string excelPath, int rowIndex)
        {
            return GetRow_Data(excelSheet, excelPath, null, rowIndex);
        }

        string[] IExcelDataIO.GetRowData(string excelPath, int rowIndex)
        {
            ExcelSheet excelSheet = ExcelSheet.Instance;
            excelSheet.SheetName = "";
            excelSheet.SheetIndex = 0;
            return ((IExcelDataIO)this).GetRowData(excelSheet, excelPath, rowIndex);
        }

        void IExcelDataIO.ToExcelWithProperty(ExcelSheet excelSheet, string excelPath)
        {
            ((IExcelDataIO)this).ToExcelWithProperty(excelSheet, excelPath, false);
        }

        void IExcelDataIO.ToExcelWithProperty(ExcelSheet excelSheet, string excelPath, bool appendToLastSheet)
        {
            if (null == excelPlugin) return;

            if (File.Exists(excelPath) && (null == work_book))
            {
                work_book = excelPlugin.GetWorkbook(excelPath);
            }
            else if (null == work_book)
            {
                work_book = excelPlugin.CreateWorkbook(excelPath);
            }
            work_sheet = excelPlugin.CreateWorksheet(work_book, excelSheet.SheetIndex, excelSheet.SheetName, appendToLastSheet);
            if (null == work_sheet)
            {
                throw new Exception("Worksheet is not null!");
            }

            headList = excelSheet.ExcelColumnsMappings;
            headList.Sort();
            rowIndex = excelSheet.HeadRowIndex;
            object row = excelPlugin.GetRow(work_sheet, rowIndex);
            int columnIndex = 0;
            foreach (CellProperty item in headList)
            {
                item.columnIndex = columnIndex;
                item.cellValue = item.headText;
                excelPlugin.SetValue(row, item);
                columnIndex++;
            }
            rowIndex++;
        }

        void IExcelDataIO.ToExcelWithData(ExcelRowChildren excelRowChildren)
        {
            if (null == work_sheet)
            {
                throw new Exception("Worksheet is not null!");
            }

            List<CellProperty> cpList = new List<CellProperty>();
            CellProperty cp = null;
            foreach (CellProperty item in headList)
            {
                if (!excelRowChildren.ContainsKey(item.fieldName)) continue;
                cp = excelRowChildren[item.fieldName];
                cp.columnIndex = item.columnIndex;
                cpList.Add(cp);
            }

            cpList.Sort();

            object row = excelPlugin.GetRow(work_sheet, rowIndex);
            foreach (CellProperty item in cpList)
            {
                excelPlugin.SetValue(row, item);
            }
            rowIndex++;
        }

        ExcelRowChildren IExcelDataIO.ToExcelWithExcelRowChildren()
        {
            if (0 == excelRowChildren.Count)
            {
                int n = 0;
                foreach (CellProperty item in headList)
                {
                    excelRowChildren.Add(new CellProperty()
                    {
                        fieldName = item.fieldName,
                        headText = item.headText,
                        columnIndex = n
                    });
                    n++;
                }
            }

            PropertyInfo[] piArr = cellP.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            object v = null;
            foreach (CellProperty item in excelRowChildren)
            {
                foreach (PropertyInfo p in piArr)
                {
                    if (!p.CanWrite) continue;
                    v = p.GetValue(cellP, null);
                    item.GetType().GetProperty(p.Name).SetValue(item, v);
                }
            }

            return excelRowChildren;
        }

        void IDisposable.Dispose()
        {
            excelRowChildren.Clear();
            work_book = null;
            work_sheet = null;
            excelPlugin.Save();
            excelPlugin.Dispose();
        }

        byte[] IExcelDataIO.ToExcelGetBody()
        {
            if (null == work_book) return null;
            return excelPlugin.GetExcelData(work_book);
        }

        byte[] IExcelDataIO.ToExcelGetBody(string excelPath)
        {
            object workbook = excelPlugin.GetWorkbook(excelPath);
            byte[] data = excelPlugin.GetExcelData(workbook);
            excelPlugin.Dispose();
            return data;
        }

        string[] IExcelDataIO.GetWorksheetNames(string excelPath)
        {
            string[] results = null;
            if (null == excelPlugin) return results;
            if (!File.Exists(excelPath)) return results;

            object workbook = excelPlugin.GetWorkbook(excelPath);
            results = excelPlugin.GetWorksheetNames(workbook);
            excelPlugin.Dispose();
            return results;
        }

        List<T> IExcelDataIO.FromExcelStream<T>(ExcelSheet excelSheet, Stream execlStream)
        {
            List<T> list = new List<T>();
            DataTable dt = null;
            from_excel<T>(excelSheet, null, execlStream, null, null, list, false, ref dt);
            return list;
        }

        DataTable IExcelDataIO.FromExcelStream(ExcelSheet excelSheet, Stream execlStream)
        {
            DataTable dt = new DataTable();
            from_excel<object>(excelSheet, null, execlStream, null, null, null, false, ref dt);
            return dt;
        }

        DataTable IExcelDataIO.FromExcelStream(ExcelSheet excelSheet, Stream execlStream, bool dataTableHeaderOfExcelHeader)
        {
            DataTable dt = new DataTable();
            from_excel<object>(excelSheet, null, execlStream, null, null, null, dataTableHeaderOfExcelHeader, ref dt);
            return dt;
        }

        void IExcelDataIO.FromExcelStream(ExcelSheet excelSheet, Stream execlStream, Action<DataRow> action)
        {
            DataTable dt = null;
            from_excel<object>(excelSheet, null, execlStream, action, null, null, false, ref dt);
        }

        void IExcelDataIO.FromExcelStream<T>(ExcelSheet excelSheet, Stream execlStream, Action<T> action)
        {
            DataTable dt = null;
            from_excel<T>(excelSheet, null, execlStream, null, action, null, false, ref dt);
        }

        string[] IExcelDataIO.GetRowDataFromStream(Stream execlStream, int rowIndex)
        {
            ExcelSheet excelSheet = ExcelSheet.Instance;
            excelSheet.SheetName = "";
            excelSheet.SheetIndex = 0;
            return GetRow_Data(excelSheet, null, execlStream, rowIndex);
        }

        string[] IExcelDataIO.GetRowDataFromStream(ExcelSheet excelSheet, Stream execlStream, int rowIndex)
        {
            return GetRow_Data(excelSheet, null, execlStream, rowIndex);
        }

        Dictionary<string, string> IExcelDataIO.GetRowDataKayValueFromStream(ExcelSheet excelSheet, Stream execlStream, int rowIndex)
        {
            return GetRowData_KayValue(excelSheet, null, execlStream, rowIndex);
        }

        string[] IExcelDataIO.GetWorksheetNamesFromStream(Stream execlStream)
        {
            object workbook = excelPlugin.GetWorkbookByStream(execlStream);
            string[] results = excelPlugin.GetWorksheetNames(workbook);
            excelPlugin.Dispose();
            return results;
        }

        private Stream BytesToStream(byte[] data)
        {
            if (null == data) return null;
            MemoryStream ms = new MemoryStream();
            ms.Write(data, 0, data.Length);
            return ms;
        }

        List<T> IExcelDataIO.FromExcelBytes<T>(ExcelSheet excelSheet, byte[] excelFileData)
        {
            if (null == excelFileData) return null;
            Stream ms = BytesToStream(excelFileData);
            return ((IExcelDataIO)this).FromExcelStream<T>(excelSheet, ms);
        }

        DataTable IExcelDataIO.FromExcelBytes(ExcelSheet excelSheet, byte[] excelFileData)
        {
            if (null == excelFileData) return null;
            Stream ms = BytesToStream(excelFileData);
            return ((IExcelDataIO)this).FromExcelStream(excelSheet, ms);
        }

        DataTable IExcelDataIO.FromExcelBytes(ExcelSheet excelSheet, byte[] excelFileData, bool dataTableHeaderOfExcelHeader)
        {
            if (null == excelFileData) return null;
            Stream ms = BytesToStream(excelFileData);
            return ((IExcelDataIO)this).FromExcelStream(excelSheet, ms, dataTableHeaderOfExcelHeader);
        }

        void IExcelDataIO.FromExcelBytes(ExcelSheet excelSheet, byte[] excelFileData, Action<DataRow> action)
        {
            if (null == excelFileData) return;
            Stream ms = BytesToStream(excelFileData);
            ((IExcelDataIO)this).FromExcelStream(excelSheet, ms, action);
        }

        void IExcelDataIO.FromExcelBytes<T>(ExcelSheet excelSheet, byte[] excelFileData, Action<T> action)
        {
            if (null == excelFileData) return;
            Stream ms = BytesToStream(excelFileData);
            ((IExcelDataIO)this).FromExcelStream<T>(excelSheet, ms, action);
        }

        string[] IExcelDataIO.GetRowDataFromBytes(byte[] excelFileData, int rowIndex)
        {
            if (null == excelFileData) return null;
            Stream ms = BytesToStream(excelFileData);
            return ((IExcelDataIO)this).GetRowDataFromStream(ms, rowIndex);
        }

        string[] IExcelDataIO.GetRowDataFromBytes(ExcelSheet excelSheet, byte[] excelFileData, int rowIndex)
        {
            if (null == excelFileData) return null;
            Stream ms = BytesToStream(excelFileData);
            return ((IExcelDataIO)this).GetRowDataFromStream(excelSheet, ms, rowIndex);
        }

        Dictionary<string, string> IExcelDataIO.GetRowDataKayValueFromBytes(ExcelSheet excelSheet, byte[] excelFileData, int rowIndex)
        {
            if (null == excelFileData) return null;
            Stream ms = BytesToStream(excelFileData);
            return ((IExcelDataIO)this).GetRowDataKayValueFromStream(excelSheet, ms, rowIndex);
        }

        string[] IExcelDataIO.GetWorksheetNamesFromBytes(byte[] excelFileData)
        {
            if (null == excelFileData) return null;
            Stream ms = BytesToStream(excelFileData);
            return ((IExcelDataIO)this).GetWorksheetNamesFromStream(ms);
        }
    }
}
