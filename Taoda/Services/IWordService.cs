using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Taoda.Models;

namespace Taoda.Services;

/// <summary>
/// Word 文档处理服务接口
/// </summary>
public interface IWordService
{
    /// <summary>
    /// 从 Word 模板中提取变量
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <returns>模板变量列表</returns>
    List<string> ExtractTemplateVariables(string templatePath);
    
    /// <summary>
    /// 异步生成单个 Word 文档
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="data">填充数据</param>
    /// <param name="outputPath">输出文件路径</param>
    /// <returns>异步任务</returns>
    Task GenerateDocumentAsync(string templatePath, Dictionary<string, object> data, string outputPath);
    
    /// <summary>
    /// 异步批量生成 Word 文档
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="dataList">数据列表</param>
    /// <param name="outputDirectory">输出目录</param>
    /// <param name="progress">进度报告</param>
    /// <returns>成功生成的文档数量</returns>
    Task<int> BatchGenerateAsync(string templatePath, List<Dictionary<string, object>> dataList, string outputDirectory, IProgress<int> progress);
    
    /// <summary>
    /// 匹配 Excel 列标题与模板变量
    /// </summary>
    /// <param name="excelHeaders">Excel 列标题列表</param>
    /// <param name="templateVariables">模板变量列表</param>
    /// <returns>变量匹配结果</returns>
    VariableMatchingModel MatchVariables(List<string> excelHeaders, List<string> templateVariables);
    
    /// <summary>
    /// 高级变量匹配，包含模糊匹配和智能匹配
    /// </summary>
    /// <param name="excelHeaders">Excel 列标题列表</param>
    /// <param name="templateVariables">模板变量列表</param>
    /// <param name="enableFuzzyMatching">是否启用模糊匹配</param>
    /// <returns>变量匹配结果</returns>
    VariableMatchingModel MatchVariablesAdvanced(List<string> excelHeaders, List<string> templateVariables, bool enableFuzzyMatching = true);
}