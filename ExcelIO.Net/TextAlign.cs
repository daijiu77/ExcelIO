using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelIO.Net
{
    public enum TextAlign
    {
        None = 0,
        Left = 1,
        Center = 2 << 0,
        Right = 2 << 1,
        Top = 2 << 2,
        Middle = 2 << 3,
        Bottom = 2 << 4
    }
}
