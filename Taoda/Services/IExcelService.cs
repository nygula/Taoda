using System.Collections.Generic;
using System.Threading.Tasks;
using Taoda.Models;

namespace Taoda.Services;

/// <summary>
/// Excel 文件处理服务接口
/// </summary>
public interface IExcelService
{
    /// <summary>
    /// 异步读取 Excel 文件并返回数据模型
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <returns>Excel 数据模型</returns>
    Task<ExcelDataModel> ReadExcelAsync(string filePath);
    
    /// <summary>
    /// 获取 Excel 文件的列标题
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <returns>列标题列表</returns>
    List<string> GetColumnHeaders(string filePath);
    
    /// <summary>
    /// 获取 Excel 文件的数据行数量
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <returns>数据行数量</returns>
    int GetDataRowCount(string filePath);
}