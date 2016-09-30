using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Xml;
using System.Text;
using System.Threading.Tasks;

namespace NETCore.Extensions.Excel.Infrastructure {
    public class ISheet {
        public ISheet(WorkBook wb, ExcelStream Excel, SharedStrings StringDictionary) {
            rows = new List<IRow>();
            this.Id = wb.Id;
            this.Name = wb.Name;
            this.StringDictionary = StringDictionary;
            this.excel = Excel;
            this.ColumnSetting = new Dictionary<uint, uint>();
        }

        private ExcelStream excel;

        private uint Id;

        private string Name;

        protected SharedStrings StringDictionary { get; set; }

        protected Dictionary<uint, uint> ColumnSetting { get; set; }


        private List<IRow> rows;

        public uint LastRowNum { get; protected set; }



        public IRow CreateRow(uint rownum) {

            var r = rows.FirstOrDefault(t => t.RowNum == rownum);

            if (r == null) {
                r = new IRow(this) { RowNum = rownum };
                rows.Add(r);
                if (rownum > LastRowNum)
                    LastRowNum = rownum;
            }

            return r;
        }

        public void SetColumnWidth(uint columnIndex, uint width) {
            ColumnSetting[columnIndex] = width;
        }
        


        public void SaveChanges() {


            // 获取sheetX.xml
            var sheetX = excel.ZipArchive.GetEntry($"xl/worksheets/sheet{Id}.xml");
            using (var stream = sheetX.Open())
            using (var sr = new StreamReader(stream)) {
                // 获取sheetData节点
                var xd = new XmlDocument();
                xd.LoadXml(sr.ReadToEnd());
                var sheetData = xd.GetElementsByTagName("sheetData")
                    .Cast<XmlNode>()
                    .First();

                var worksheet = xd.GetElementsByTagName("worksheet")[0];

                //hyperlinks
                var hyperlinks = xd.GetElementsByTagName("hyperlinks").Cast<XmlNode>().FirstOrDefault();
                if (hyperlinks == null) {
                    hyperlinks = xd.CreateElement("hyperlinks", xd.DocumentElement.NamespaceURI);
                } else {
                    hyperlinks.ParentNode.RemoveChild(hyperlinks);
                }
               

                // 删除全部元素
                sheetData.RemoveAll();
                foreach (var row in rows) {

                    // 添加row节点
                    var element = xd.CreateElement("row", xd.DocumentElement.NamespaceURI);
                    element.SetAttribute("r", row.RowNum.ToString());
                    element.SetAttribute("spans", "1:1");
                    foreach (var cell in row.Cells) {
                        var innerText = cell.value;
                        innerText = StringDictionary._Add(innerText).ToString();

                        

                        var element2 = xd.CreateElement("c", xd.DocumentElement.NamespaceURI);
                        element2.SetAttribute("t", "s");
                        element2.SetAttribute("r", cell.ColumnNumber + row.RowNum.ToString());


                        var element3 = xd.CreateElement("v", xd.DocumentElement.NamespaceURI);
                        element3.InnerText = innerText;
                        element2.AppendChild(element3);
                        element.AppendChild(element2);



                        //添加超链接
                        if (cell.Hyperlink != null) {
                            element2.SetAttribute("s","1");
                            var hl = xd.CreateElement("hyperlink", xd.DocumentElement.NamespaceURI);
                            hl.SetAttribute("ref", cell.ColumnNumber + cell.ColumnIndex);
                            hl.SetAttribute("location", cell.Hyperlink.Address);
                            hl.SetAttribute("display", cell.Hyperlink.Label);
                            hyperlinks.AppendChild(hl);
                        }

                    }
                    sheetData.AppendChild(element);
                }

                if (hyperlinks.FirstChild != null) {
                    var pageMargins = xd.GetElementsByTagName("pageMargins")[0];
                    worksheet.InsertBefore(hyperlinks, pageMargins);
                }

                //列宽设置
                if (ColumnSetting.Count > 0) {
                    var cols = xd.GetElementsByTagName("cols").Cast<XmlNode>().FirstOrDefault();
                    if (cols == null) {
                        cols = xd.CreateElement("cols", xd.DocumentElement.NamespaceURI);
                    }
                    foreach (var cs in ColumnSetting) {
                        var col = xd.CreateElement("col", xd.DocumentElement.NamespaceURI);
                        col.SetAttribute("customWidth", "1");
                        col.SetAttribute("max", cs.Key.ToString());
                        col.SetAttribute("min", cs.Key.ToString());
                        col.SetAttribute("width", cs.Value.ToString());
                        cols.AppendChild(col);
                    }
                    worksheet.InsertBefore(cols,sheetData);
                }

                // 保存sheetX.xml
                stream.Position = 0;
                stream.SetLength(0);
                xd.Save(stream);
            }
            // 回收垃圾
            GC.Collect();

            // 保存sharedStrings.xml
            var sharedStrings = excel.ZipArchive.GetEntry($"xl/sharedStrings.xml");
            using (var stream = sharedStrings.Open())
            using (var sw = new StreamWriter(stream)) {
                var xmlString = new StringBuilder($@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<sst count=""{StringDictionary.LongCount()}"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
                var inner = new StringBuilder();
                foreach (var x in StringDictionary)
                    xmlString.Append($"    <si><t>{x}</t></si>\r\n");
                xmlString.Append("</sst>");
                sw.Write(xmlString);
            }
            // 回收垃圾
            GC.Collect();
        }
    }
}
