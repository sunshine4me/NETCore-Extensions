using System.IO;

namespace NETCore.Extensions.Excel
{
    public interface IExcelStream
    {
        /// <summary>
        /// 创建ExcelStream<see cref="ExcelStream"/>
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns></returns>
        ExcelStream Create(string path);




        /// <summary>
        /// 加载Excel,返回ExcelStream
        /// </summary>
        /// <param name="path">excel文件路径</param>
        /// <returns></returns>
        ExcelStream Load(string path);

        /// <summary>
        /// 加载Excel,返回ExcelStream
        /// </summary>
        /// <param name="stream">excel字节流</param>
        /// <returns></returns>
        ExcelStream Load(Stream stream);




        /// <summary>
        /// 删除Sheet
        /// </summary>
        /// <param name="name">Sheet名称</param>
        void RemoveSheet(string name);

        /// <summary>
        /// 删除Sheet
        /// </summary>
        /// <param name="Id">Sheet名称对应的Id</param>
        void RemoveSheet(uint Id);
    }
}
