using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace NETCore.Extensions.Excel.Infrastructure {
    public class IRow {

        public IRow(ISheet _sheet) {
            Cells = new List<ICell>();
            sheet = _sheet;
        }
        public uint RowNum { get;  set; }
        public List<ICell> Cells { get; private set; }
        public ISheet sheet { get; private set; }

        
        public uint LastCellNum { get; private set; }
        public ICell CreateCell(uint columnIndex) {
            var c = Cells.FirstOrDefault(t => t.ColumnIndex == columnIndex);

            if (c == null) {
                c = new ICell(this) { ColumnIndex= columnIndex };
                Cells.Add(c);
                if (columnIndex > LastCellNum)
                    LastCellNum = columnIndex;
            }
            return c;
        }

        public ICell CreateCell(string columnNumber) {
            var index = ICell.ColumnNumberToIndex(columnNumber);
            return CreateCell(index);
        }

       
    }
}
