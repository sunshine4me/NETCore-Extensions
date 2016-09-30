using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace NETCore.Extensions.Excel.Infrastructure {
    public class ICell {
        public ICell(IRow row) {
            this.row = row;
        }
        public string value { get; set; }

        public IRow row { get; private set; }

        public HSSFHyperlink Hyperlink { get; set; }

        private uint _columnIndex;
        public uint ColumnIndex {
            get {
                return _columnIndex;
            }
            set {
                _columnIndex = value;

                string s = string.Empty;
                while (value > 0) {
                    var m = value % 26;
                    if (m == 0) m = 26;
                    s = (char)(m + 64) + s;
                    value = (value - m) / 26;
                }
                ColumnNumber = s;
            }
        }


        public string ColumnNumber { get; private set; }

        public static uint ColumnNumberToIndex(string s) {
            if (string.IsNullOrEmpty(s)) return 0;
            long n = 0;
            for (long i = s.Length - 1, j = 1; i >= 0; i--, j *= 26) {
                char c = Char.ToUpper(s[Convert.ToInt32(i)]);
                if (c < 'A' || c > 'Z') return 0;
                n += (c - 64) * j;
            }
            return Convert.ToUInt32(n);
        }



    }
}
