using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelIO.Net
{
    public class ExcelRowChildren: IEnumerable<CellProperty>, IEnumerable
    {
        private Dictionary<string, CellProperty> dic = new Dictionary<string, CellProperty>();
        private List<string> list = new List<string>();
        public CellProperty this[string fieldName]
        {
            get
            {
                if (!dic.ContainsKey(fieldName)) return null;
                return dic[fieldName];
            }
        }

        public bool ContainsKey(string key)
        {
            return dic.ContainsKey(key);
        }

        public void Add(CellProperty cellProperty)
        {
            if (dic.ContainsKey(cellProperty.fieldName)) return;
            dic.Add(cellProperty.fieldName, cellProperty);
            list.Add(cellProperty.fieldName);
        }

        public void Clear()
        {
            dic.Clear();
        }

        public int Count
        {
            get
            {
                return dic.Count;
            }
        }

        public void RemoveAt(int index)
        {
            if (list.Count <= index) return;
            string fn = list[index];
            dic.Remove(fn);
            list.RemoveAt(index);
        }

        IEnumerator<CellProperty> IEnumerable<CellProperty>.GetEnumerator()
        {
            return new ExcelRowItem(this);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new ExcelRowItem(this);
        }

        class ExcelRowItem : IEnumerator<CellProperty>, IEnumerator
        {
            private ExcelRowChildren excelRow = null;
            private CellProperty cellProperty1 = null;
            private int index = 0;
            public ExcelRowItem(ExcelRowChildren excelRow)
            {
                this.excelRow = excelRow;
            }

            CellProperty IEnumerator<CellProperty>.Current { get { return cellProperty1; } }

            object IEnumerator.Current { get { return cellProperty1; } }

            void IDisposable.Dispose()
            {
                //throw new NotImplementedException();
            }

            bool IEnumerator.MoveNext()
            {
                if (excelRow.dic.Count <= index) return false;
                string fn = excelRow.list[index];
                cellProperty1 = excelRow.dic[fn];
                index++;
                return true;
            }

            void IEnumerator.Reset()
            {
                index = 0;
            }
        }
    }
}
