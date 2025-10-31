using System.Collections.Generic;
using System.IO;

namespace Taoda.Models;

/// <summary>
/// 表示 Word 模板的数据模型
/// </summary>
public class TemplateModel
{
    /// <summary>
    /// Word 模板文件路径
    /// </summary>
    public string FilePath { get; set; } = string.Empty;
    
    /// <summary>
    /// 模板中检测到的变量列表
    /// </summary>
    public List<string> Variables { get; set; } = new();
    
    /// <summary>
    /// 模板文件名（不含路径）
    /// </summary>
    public string FileName => Path.GetFileName(FilePath);
}