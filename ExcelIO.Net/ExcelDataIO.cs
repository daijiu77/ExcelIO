using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;

namespace ExcelIO.Net
{
    public class ExcelDataIO : IExcelDataIO
    {
        private IExcelPlugin excelPlugin = null;
        private ExcelRowChildren excelRowChildren = new ExcelRowChildren();

        private object workbook = null;
        private object worksheet = null;
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

        private void from_excel<T>(ExcelSheet excelSheet, string excelPath, Action<DataRow> action, Action<T> actionT, List<T> list, bool dataTableHeaderOfExcelHeader, ref DataTable dataTable)
        {
            if (null == excelPlugin) return;
            if (!File.Exists(excelPath)) return;

            object workbook = excelPlugin.GetWorkbook(excelPath);
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
                    if (false == isList && null != actionT && y == startY)
                    {
                        t = (T)Activator.CreateInstance(typeof(T));
                        isEnabled = true;
                    }
                    else if (isList)
                    {
                        t = (T)Activator.CreateInstance(typeof(T));
                        isEnabled = true;
                    }
                }

                if (null != dataTable)
                {
                    if (isNullDataTable && y == startY)
                    {
                        dataRow = dataTable.NewRow();
                        isEnabled = true;
                    }
                    else if (false == isNullDataTable)
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
            from_excel<object>(excelSheet, excelPath, action, null, null, false, ref dt);
        }

        List<T> IExcelDataIO.FromExcel<T>(ExcelSheet excelSheet, string excelPath)
        {
            List<T> list = new List<T>();
            DataTable dt = null;
            from_excel<T>(excelSheet, excelPath, null, null, list, false, ref dt);
            return list;
        }

        void IExcelDataIO.FromExcel<T>(ExcelSheet excelSheet, string excelPath, Action<T> action)
        {
            DataTable dt = null;
            from_excel<T>(excelSheet, excelPath, null, action, null, false, ref dt);
        }

        DataTable IExcelDataIO.FromExcel(ExcelSheet excelSheet, string excelPath)
        {
            DataTable dt = new DataTable();
            from_excel<object>(excelSheet, excelPath, null, null, null, false, ref dt);
            return dt;
        }

        DataTable IExcelDataIO.FromExcel(ExcelSheet excelSheet, string excelPath, bool dataTableHeaderOfExcelHeader)
        {
            DataTable dt = new DataTable();
            from_excel<object>(excelSheet, excelPath, null, null, null, dataTableHeaderOfExcelHeader, ref dt);
            return dt;
        }

        string[] IExcelDataIO.GetRowData(string excelPath, int rowIndex)
        {
            string[] results = null;
            if (null == excelPlugin) return results;
            if (!File.Exists(excelPath)) return results;

            object workbook = excelPlugin.GetWorkbook(excelPath);
            object worksheet = excelPlugin.GetWorksheet(workbook, 0);
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
            //throw new NotImplementedException();
        }

        void IExcelDataIO.ToExcelWithProperty(ExcelSheet excelSheet, string excelPath)
        {
            if (null == excelPlugin) return;

            workbook = excelPlugin.CreateWorkbook(excelPath);
            worksheet = excelPlugin.CreateWorksheet(workbook, excelSheet.SheetIndex, excelSheet.SheetName);
            if (null == worksheet)
            {
                throw new Exception("Worksheet is not null!");
            }

            headList = excelSheet.ExcelColumnsMappings;
            headList.Sort();
            rowIndex = excelSheet.HeadRowIndex;
            object row = excelPlugin.GetRow(worksheet, rowIndex);
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
            if (null == worksheet)
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

            object row = excelPlugin.GetRow(worksheet, rowIndex);
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
            workbook = null;
            worksheet = null;
            excelPlugin.Save();
            excelPlugin.Dispose();
        }

        byte[] IExcelDataIO.ToExcelGetBody()
        {
            return excelPlugin.GetExcelData(workbook);
        }

    }
}
