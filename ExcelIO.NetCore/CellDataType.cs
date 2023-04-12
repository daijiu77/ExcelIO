namespace ExcelIO.NetCore
{
    public enum CellDataType
    {
        /// <summary>
        /// General
        /// </summary>
        General = 0,
        /// <summary>
        /// 0
        /// </summary>
        Number = 1,
        /// <summary>
        /// 0.00
        /// </summary>
        Decimal2 = 2,
        /// <summary>
        /// #,##0
        /// </summary>
        Decimal3 = 3,
        /// <summary>
        /// #,##0.00
        /// </summary>
        Decimal4 = 4,
        /// <summary>
        /// $#,##0_);($#,##0)
        /// </summary>
        Currency1 = 5,
        /// <summary>
        /// $#,##0_);[Red]($#,##0)
        /// </summary>
        Currency2 = 6,
        /// <summary>
        /// $#,##0.00_);($#,##0.00)
        /// </summary>
        Currency3 = 7,
        /// <summary>
        /// $#,##0.00_);[Red]($#,##0.00)
        /// </summary>
        Currency4 = 8,
        /// <summary>
        /// 0%
        /// </summary>
        Percentage1 = 9,
        /// <summary>
        /// 0.00%
        /// </summary>
        Percentage2 = 10,
        /// <summary>
        /// 0.00E+00
        /// </summary>
        Scientific1 = 11,
        /// <summary>
        /// # ?/?
        /// </summary>
        Fraction1 = 12,
        /// <summary>
        /// # ??/??
        /// </summary>
        Fraction2 = 13,
        /// <summary>
        /// m/d/yyyy
        /// </summary>
        Date1 = 14,
        /// <summary>
        /// d-mmm-yy
        /// </summary>
        Date2 = 15,
        /// <summary>
        /// d-mmm
        /// </summary>
        Date3 = 16,
        /// <summary>
        /// mmm-yy
        /// </summary>
        Date4 = 17,
        /// <summary>
        /// h:mm AM/PM
        /// </summary>
        Time1 = 18,
        /// <summary>
        /// h:mm:ss AM/PM
        /// </summary>
        Time2 = 19,
        /// <summary>
        /// h:mm
        /// </summary>
        Time3 = 20,
        /// <summary>
        /// h:mm:ss
        /// </summary>
        Time4 = 21,
        /// <summary>
        /// m/d/yyyy h:mm
        /// </summary>
        Time5 = 22,
        /// <summary>
        /// #,##0_);(#,##0)
        /// </summary>
        Accounting1 = 37,
        /// <summary>
        /// #,##0_);[Red](#,##0)
        /// </summary>
        Accounting2 = 38,
        /// <summary>
        /// #,##0.00_);(#,##0.00)
        /// </summary>
        Accounting3 = 39,
        /// <summary>
        /// #,##0.00_);[Red](#,##0.00)
        /// </summary>
        Accounting4 = 40,
        /// <summary>
        /// _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
        /// </summary>
        Accounting5 = 41,
        /// <summary>
        /// _($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
        /// </summary>
        Currency5 = 42,
        /// <summary>
        /// _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
        /// </summary>
        Accounting6 = 43,
        /// <summary>
        /// _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
        /// </summary>
        Currency6 = 44,
        /// <summary>
        /// mm:ss
        /// </summary>
        Time6 = 45,
        /// <summary>
        /// [h]:mm:ss
        /// </summary>
        Time7 = 46,
        /// <summary>
        /// mm:ss.0
        /// </summary>
        Time8 = 47,
        /// <summary>
        /// ##0.0E+0
        /// </summary>
        Scientific2 = 48,
        /// <summary>
        /// @
        /// </summary>
        Text = 49
    }
}
