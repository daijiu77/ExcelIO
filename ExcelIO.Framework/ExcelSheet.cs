using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ExcelIO.Framework
{
    public class ExcelSheet : IEnumerable<CellProperty>, IEnumerable
    {
        private int _columnIndex = 0;

        private ExcelSheet() { }

        public static ExcelSheet Instance
        {
            get
            {
                return new ExcelSheet();
            }
        }

        public string SheetName { get; set; }

        public int SheetIndex { get; set; }

        public int HeadRowIndex { get; set; }

        private List<CellProperty> _excelColumnsMappings = new List<CellProperty>();
        public List<CellProperty> ExcelColumnsMappings { get { return _excelColumnsMappings; } }

        public void AddMapping(string ExcelHeadText, string DbTableFieldName, int ColumeIndex)
        {
            _excelColumnsMappings.Add(new CellProperty()
            {
                headText = ExcelHeadText,
                fieldName = DbTableFieldName,
                columnIndex = ColumeIndex
            });
        }

        public CellProperty LastCellProperty
        {
            get
            {
                if (0 == _excelColumnsMappings.Count) return null;
                return _excelColumnsMappings[Count - 1];
            }
        }

        public void AddMapping(string ExcelHeadText, string DbTableFieldName)
        {
            AddMapping(ExcelHeadText, DbTableFieldName, -1);
        }

        public void ClearMapping()
        {
            _columnIndex = 0;
            _excelColumnsMappings.Clear();
        }

        public int Count { get { return _excelColumnsMappings.Count; } }

        IEnumerator<CellProperty> IEnumerable<CellProperty>.GetEnumerator()
        {
            return new FieldsMapping(this);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new FieldsMapping(this);
        }

        class FieldsMapping : IEnumerator<CellProperty>, IEnumerator
        {
            private ExcelSheet excelSheet = null;
            private CellProperty excelColumnsMapping = null;
            private int indexNum = 0;

            public FieldsMapping(ExcelSheet excelSheet)
            {
                this.excelSheet = excelSheet;
            }

            object IEnumerator.Current { get { return excelColumnsMapping; } }

            CellProperty IEnumerator<CellProperty>.Current { get { return excelColumnsMapping; } }

            void IDisposable.Dispose()
            {
                excelSheet._excelColumnsMappings.Clear();
            }

            bool IEnumerator.MoveNext()
            {
                if (excelSheet._excelColumnsMappings.Count <= indexNum) return false;
                excelColumnsMapping = excelSheet._excelColumnsMappings[indexNum];
                indexNum++;
                return true;
            }

            void IEnumerator.Reset()
            {
                indexNum = 0;
            }
        }
    }
}
