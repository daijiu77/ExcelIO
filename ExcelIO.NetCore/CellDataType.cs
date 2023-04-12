using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelIO.NetCore
{
    public enum CellDataType
    {
        General = 0,
        Decimal = 4,
        Date = 17,
        Time = 46,
        Text = 49
    }
}
