using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelIO.Net
{
    public class AsposePlugin : IExcelPlugin
    {
        private Workbook workbook = null;
        private string excelPath = "";

        object IExcelPlugin.CreateWorkbook(string excelPath)
        {
            this.excelPath = null == excelPath ? "" : excelPath;
            workbook = new Workbook();
            return workbook;
        }

        object IExcelPlugin.CreateWorksheet(object workbook, int sheetIndex, string sheetName)
        {
            int n = sheetIndex + 1;
            int x = 0;
            const int max = 100;
            while (((Workbook)workbook).Worksheets.Count < n)
            {
                ((Workbook)workbook).Worksheets.Add();
                x++;
                if (max <= x) break;
            }

            Worksheet worksheet = ((Workbook)workbook).Worksheets[sheetIndex];
            if (!string.IsNullOrEmpty(sheetName))
            {
                worksheet.Name = sheetName;
            }

            return worksheet;
        }

        void IDisposable.Dispose()
        {
            if (null != workbook)
            {
                if (null != workbook as IDisposable)
                {
                    ((IDisposable)workbook).Dispose();
                }
                else
                {
                    MethodInfo mi = workbook.GetType().GetMethod("Dispose");
                    if (null != mi)
                    {
                        try
                        {
                            mi.Invoke(workbook, null);
                        }
                        catch (Exception)
                        {
                            //throw;
                        }
                    }

                    mi = workbook.GetType().GetMethod("Quit");
                    if (null != mi)
                    {
                        try
                        {
                            mi.Invoke(workbook, null);
                        }
                        catch (Exception)
                        {
                            //throw;
                        }
                    }

                    mi = workbook.GetType().GetMethod("Close");
                    if (null != mi)
                    {
                        try
                        {
                            mi.Invoke(workbook, null);
                        }
                        catch (Exception)
                        {
                            //throw;
                        }
                    }
                    //workbook.Dispose();
                }

                workbook = null;
            }

            //RemoveEvaluation();
        }

        private bool IsXlsx(string excelPath, ref bool enable)
        {
            bool isXLSX = false;
            enable = false;
            Regex rg = new Regex(@"\.(?<ExtName>((xls)|(xlsx)))$", RegexOptions.IgnoreCase);
            if (rg.IsMatch(excelPath))
            {
                enable = true;
                string ExtName = rg.Match(excelPath).Groups["ExtName"].Value.ToLower();
                isXLSX = ExtName.Equals("xlsx");
            }
            return isXLSX;
        }

        byte[] IExcelPlugin.GetExcelData(object workbook)
        {
            byte[] datas = null;
            using (MemoryStream ms = ((Workbook)workbook).SaveToStream())
            {
                datas = ms.ToArray();
            }
            return datas;
            //throw new NotImplementedException();
        }

        object IExcelPlugin.GetRow(object worksheet, int rowIndex)
        {
            return new object[] { worksheet, rowIndex };
        }

        string IExcelPlugin.GetValue(object row, int columnIndex)
        {
            object[] arr = row as object[];
            Worksheet worksheet = (Worksheet)arr[0];
            int rowIndex = (int)arr[1];
            object cellVal = worksheet.Cells[rowIndex, columnIndex].Value;
            cellVal = null == cellVal ? "" : cellVal;
            return cellVal.ToString();
        }

        object IExcelPlugin.GetWorkbook(string excelPath)
        {
            this.excelPath = excelPath;
            workbook = new Workbook(excelPath);
            return workbook;
        }

        object IExcelPlugin.GetWorksheet(object workbook, string sheetName)
        {
            return ((IExcelPlugin)this).CreateWorksheet(workbook, 0, sheetName);
        }

        object IExcelPlugin.GetWorksheet(object workbook, int sheetIndex)
        {
            return ((IExcelPlugin)this).CreateWorksheet(workbook, sheetIndex, null);
        }

        int IExcelPlugin.LastColumnIndex(object sheet, int rowIndex)
        {
            return ((Worksheet)sheet).Cells.MaxDataColumn;
        }

        int IExcelPlugin.LastRowIndex(object sheet)
        {
            return ((Worksheet)sheet).Cells.MaxDataRow;
        }

        void IExcelPlugin.Save()
        {
            if (string.IsNullOrEmpty(excelPath)) return;
            bool enable = false;
            bool is_xlsx = IsXlsx(excelPath, ref enable);
            if (enable)
            {
                try
                {
                    if (false == is_xlsx)
                    {
                        workbook.Save(excelPath, SaveFormat.Excel97To2003);
                    }
                    else
                    {
                        workbook.Save(excelPath);
                    }
                }
                catch (Exception)
                {

                    //throw;
                }
            }
            //throw new NotImplementedException();
        }

        object IExcelPlugin.SetSheetName(object workbook, int sheetIndex, string sheetName)
        {
            return ((IExcelPlugin)this).CreateWorksheet(workbook, sheetIndex, sheetName);
        }

        void IExcelPlugin.SetValue(object row, CellProperty cellProperty)
        {
            object[] arr = row as object[];
            Worksheet worksheet = (Worksheet)arr[0];
            int rowIndex = (int)arr[1];
            Cell cell = worksheet.Cells[rowIndex, cellProperty.columnIndex];
            cell.PutValue(cellProperty.cellValue);

            if (0 < cellProperty.width)
            {
                worksheet.Cells.SetColumnWidth(cellProperty.columnIndex, cellProperty.width);
            }

            if (0 < cellProperty.height)
            {
                worksheet.Cells.SetRowHeight(rowIndex, cellProperty.height);
            }

            setCellStyle(cell, cellProperty);
        }

        void setCellStyle(Cell cell, CellProperty cellProperty)
        {
            Style style = cell.GetStyle();
            if (null != cellProperty.backgroundColor)
            {
                //Aspose 背景色是使用 ForegroundColor 来设置
                style.ForegroundColor = (Color)cellProperty.backgroundColor;
                style.Pattern = BackgroundType.Solid;
            }

            if (null != cellProperty.foreColor)
            {
                style.Font.Color = (Color)cellProperty.foreColor;
            }

            if (TextAlign.None != cellProperty.textAlign)
            {
                Dictionary<TextAlign, TextAlignmentType> alignDic = new Dictionary<TextAlign, TextAlignmentType>();
                alignDic.Add(TextAlign.Left, TextAlignmentType.Left);
                alignDic.Add(TextAlign.Center, TextAlignmentType.Center);
                alignDic.Add(TextAlign.Right, TextAlignmentType.Right);
                alignDic.Add(TextAlign.Top, TextAlignmentType.Top);
                alignDic.Add(TextAlign.Middle, TextAlignmentType.Center);
                alignDic.Add(TextAlign.Bottom, TextAlignmentType.Bottom);
                Array arr = Enum.GetValues(typeof(TextAlign));
                TextAlign textAlign = TextAlign.None;
                int num = (int)TextAlign.Top;
                foreach (var item in arr)
                {
                    textAlign = (TextAlign)item;
                    if (textAlign == (cellProperty.textAlign & textAlign))
                    {
                        if ((int)textAlign < num)
                        {
                            if (alignDic.ContainsKey(textAlign)) style.HorizontalAlignment = alignDic[textAlign];
                        }
                        else
                        {
                            if (alignDic.ContainsKey(textAlign)) style.VerticalAlignment = alignDic[textAlign];
                        }
                    }
                }
            }

            if (0 < cellProperty.fontSize)
            {
                style.Font.Size = cellProperty.fontSize;
            }

            if (!string.IsNullOrEmpty(cellProperty.fontFamily))
            {
                style.Font.Name = cellProperty.fontFamily;
            }

            style.Font.IsBold = cellProperty.isBold;
            style.IsTextWrapped = cellProperty.wrapText;

            style.Number = (int)cellProperty.cellDataType;
            cell.SetStyle(style);
        }
    }
}
