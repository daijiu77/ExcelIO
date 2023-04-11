using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace ExcelIO.NetCore
{
    public class CellProperty : IComparable<CellProperty>
    {
        public Color? backgroundColor { get; set; }
        public Color? foreColor { get; set; }
        public int fontSize { get; set; }
        public string fontFamily { get; set; }
        public bool showBorder { get; set; }

        public TextAlign textAlign { get; set; }

        public CellDataType cellDataType { get; set; }

        private int GetIntFromHex(string strHex)
        {
            switch (strHex.ToUpper())
            {
                case ("A"):
                    {
                        return 10;
                    }
                case ("B"):
                    {
                        return 11;
                    }
                case ("C"):
                    {
                        return 12;
                    }
                case ("D"):
                    {
                        return 13;
                    }
                case ("E"):
                    {
                        return 14;
                    }
                case ("F"):
                    {
                        return 15;
                    }
                default:
                    {
                        return int.Parse(strHex);
                    }
            }
        }

        private Color Hex2Color(string hexColor)
        {
            string r, g, b;

            if (string.IsNullOrEmpty(hexColor) == false)
            {
                try
                {
                    hexColor = hexColor.Trim();
                    if (hexColor[0] == '#') hexColor = hexColor.Substring(1, hexColor.Length - 1);
                    while (hexColor.Length < 6)
                    {
                        hexColor += "0";
                    }
                    //MessageBox.Show(hexColor);
                    r = hexColor.Substring(0, 2);
                    g = hexColor.Substring(2, 2);
                    b = hexColor.Substring(4, 2);

                    int r1 = 16 * GetIntFromHex(r.Substring(0, 1)) + GetIntFromHex(r.Substring(1, 1));
                    int g1 = 16 * GetIntFromHex(g.Substring(0, 1)) + GetIntFromHex(g.Substring(1, 1));
                    int b1 = 16 * GetIntFromHex(b.Substring(0, 1)) + GetIntFromHex(b.Substring(1, 1));
                    r = Convert.ToString(r1);
                    g = Convert.ToString(g1);
                    b = Convert.ToString(b1);

                    return Color.FromArgb(Convert.ToInt32(r), Convert.ToInt32(g), Convert.ToInt32(b));
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }

            }

            return Color.Empty;
        }

        public void setBackgroundColor(string HexColor)
        {
            backgroundColor = Hex2Color(HexColor);
        }

        public void setForeColor(string HexColor)
        {
            foreColor = Hex2Color(HexColor);
        }

        /// <summary>
        /// 文本自动转行
        /// </summary>
        public bool wrapText { get; set; }

        /// <summary>
        /// 是否加粗显示文字
        /// </summary>
        public bool isBold { get; set; }

        private Dictionary<string, object> _ofDataRow = new Dictionary<string, object>();
        public Dictionary<string, object> ofDataRow { get { return _ofDataRow; } }

        public int width { get; set; }
        public int height { get; set; }

        public int rowIndex { get; set; }

        private int _columnIndex = -1;
        public int columnIndex
        {
            get { return _columnIndex; }
            set
            {
                if (-1 == _columnIndex) _columnIndex = value;
            }
        }

        private string _fieldName = "";
        public string fieldName
        {
            get { return _fieldName; }
            set
            {
                if (string.IsNullOrEmpty(_fieldName)) _fieldName = value;
            }
        }

        private string _headText = "";
        public string headText
        {
            get { return _headText; }
            set
            {
                if (string.IsNullOrEmpty(_headText)) _headText = value;
            }
        }

        public string cellValue { get; set; }

        int IComparable<CellProperty>.CompareTo(CellProperty other)
        {
            int num = 0;
            if (this.columnIndex < other.columnIndex)
            {
                num = -1;
            }
            else if (this.columnIndex > other.columnIndex)
            {
                num = 1;
            }
            return num;
        }
    }
}
