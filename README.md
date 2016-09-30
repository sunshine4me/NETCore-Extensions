# NETCore-Extensions
NETCore 初期没有什么好用的库,只能自己造一点轮子


NETCore.Extensions.Excel
一个仿照NPOI创建的Excel 操作库,只支持简单的读写数据,列宽设置和单元格超链接三种功能.

```csharp
using NETCore.Extensions.Excel;
using NETCore.Extensions.Excel.Infrastructure;
public static void Main(string[] args) {
            MemoryStream ms = new MemoryStream(1024 * 1024);

            using (ExcelStream _ExcelStream = new ExcelStream()) {
                _ExcelStream.Create(ms);

                var sheet = _ExcelStream.LoadSheet(1);

                var row = sheet.CreateRow(1);

                var cell = row.CreateCell(row.LastCellNum + 1);
                cell.value = "测试案例1";
                //超链接
                HSSFHyperlink hy = new HSSFHyperlink();
                hy.Address = "'s1'!A1";
                hy.Label = "测试案例1";
                cell.Hyperlink = hy;
                //列宽设置
                sheet.SetColumnWidth(4, 60);
                sheet.SetColumnWidth(2, 30);
                sheet.SaveChanges();

                var sheet2 = _ExcelStream.CreateSheet("s1");
                var r2 = sheet2.CreateRow(1);
                var c2 = r2.CreateCell(1);
                c2.value = "hhhhhhhhhhhhhhhhhhh";

                sheet2.SetColumnWidth(4, 60);
                sheet2.SetColumnWidth(2, 30);
                sheet2.SaveChanges();



            }

            ms.Position = 0;
            using (FileStream fs = new FileStream(@"D:\excel\result.xlsx", FileMode.Create)) {
                ms.CopyTo(fs);
            }
}
```