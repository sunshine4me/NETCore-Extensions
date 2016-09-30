using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

namespace NETCore.Extensions.Excel.Infrastructure
{
    public class SheetWithoutHDR : ISheet
    {
        public SheetWithoutHDR(WorkBook wb, string XmlSource, ExcelStream Excel, SharedStrings stringDictionary)
            : base(wb, Excel, stringDictionary)
        {
            var xd = new XmlDocument();
            xd.LoadXml(XmlSource);
            var rows = xd.GetElementsByTagName("row");
            // 遍历row标签
            foreach (XmlNode x in rows)
            {
                var cols = x.ChildNodes;
                
                var RowNum = Convert.ToUInt32(x.Attributes["r"].Value);
                var row = this.CreateRow(RowNum);

                //更新LastRowNum
                if (LastRowNum < RowNum)
                    LastRowNum = RowNum;

                // 遍历c标签
                foreach (XmlNode y in cols)
                {
                    string value = null;
                    // 如果是字符串类型，则需要从字典中查询
                    if (y.Attributes["t"]?.Value == "s")
                    {
                        var index = Convert.ToUInt32(y.FirstChild.InnerText);
                        value = StringDictionary[index];
                    }
                    else if (y.Attributes["t"]?.Value == "inlineStr")
                    {
                        value = y.FirstChild.FirstChild.InnerText;
                    }
                    // 否则其中的v标签值即为单元格内容
                    else
                    {
                        value = y.InnerText;
                    }
                    var cell = row.CreateCell(y.Attributes["r"].Value);
                    cell.value = value;
                }

                //// 去掉末尾的null
                //while (row.Cells.Last().value == null)
                //    row.Cells.Remove(row.Cells.Last());
            }
            //while (this.Count > 0 && this.Last().Count == 0)
            //    this.RemoveAt(this.Count - 1);

            //TODO hyperlinks 超链接
            //TODO

            GC.Collect();
        }
    }
}