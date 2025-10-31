using System.Collections.Generic;

namespace Taoda.Models;

/// <summary>
/// 表示从 Excel 文件读取的数据模型
/// </summary>
public class ExcelDataModel
{
    /// <summary>
    /// Excel 文件的列标题（变量名）
    /// </summary>
    public List<string> Headers { get; set; } = new();
    
    /// <summary>
    /// Excel 数据行，每行是一个字典，键为列标题，值为单元格数据
    /// </summary>
    public List<Dictionary<string, object>> Rows { get; set; } = new();
    
    /// <summary>
    /// 数据行总数
    /// </summary>
    public int TotalRows => Rows.Count;
    
    /// <summary>
    /// Excel 文件路径
    /// </summary>
    public string FilePath { get; set; } = string.Empty;
}