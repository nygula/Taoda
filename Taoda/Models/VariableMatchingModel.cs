using System.Collections.Generic;

namespace Taoda.Models;

/// <summary>
/// 表示 Excel 变量与 Word 模板变量的匹配结果
/// </summary>
public class VariableMatchingModel
{
    /// <summary>
    /// 匹配成功的变量列表
    /// </summary>
    public List<string> MatchedVariables { get; set; } = new();
    
    /// <summary>
    /// 未匹配的 Excel 列标题
    /// </summary>
    public List<string> UnmatchedExcelColumns { get; set; } = new();
    
    /// <summary>
    /// 未匹配的模板变量
    /// </summary>
    public List<string> UnmatchedTemplateVariables { get; set; } = new();
    
    /// <summary>
    /// 是否完全匹配（所有变量都匹配成功）
    /// </summary>
    public bool IsFullyMatched => UnmatchedExcelColumns.Count == 0 && UnmatchedTemplateVariables.Count == 0;
}