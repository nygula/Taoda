using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO.Compression;
using MiniSoftware;
using Taoda.Models;

namespace Taoda.Services;

/// <summary>
/// Word 文档处理服务实现
/// </summary>
public class WordService : IWordService
{
    /// <summary>
    /// 从 Word 模板中提取变量
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <returns>模板变量列表</returns>
    public List<string> ExtractTemplateVariables(string templatePath)
    {
        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentException("模板文件路径不能为空", nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"模板文件不存在: {templatePath}");

        try
        {
            var variables = new HashSet<string>();

            // Word 文档是一个 ZIP 文件，我们可以解析其中的 XML 内容
            using var archive = ZipFile.OpenRead(templatePath);

            // 查找 document.xml 文件
            var documentEntry = archive.Entries.FirstOrDefault(e => e.FullName == "word/document.xml");
            if (documentEntry != null)
            {
                using var stream = documentEntry.Open();
                using var reader = new StreamReader(stream);
                var content = reader.ReadToEnd();

                // 使用正则表达式提取 {{变量名}} 格式的占位符
                var regex = new Regex(@"\{\{([^}]+)\}\}", RegexOptions.Compiled);
                var matches = regex.Matches(content);

                foreach (Match match in matches)
                {
                    if (match.Groups.Count > 1)
                    {
                        var variableName = match.Groups[1].Value.Trim();
                        if (!string.IsNullOrEmpty(variableName))
                        {
                            variables.Add(variableName);
                        }
                    }
                }
            }

            return variables.OrderBy(v => v).ToList();
        }
        catch (Exception ex) when (!(ex is ArgumentException || ex is FileNotFoundException))
        {
            throw new InvalidOperationException($"提取模板变量时发生错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 异步生成单个 Word 文档
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="data">填充数据</param>
    /// <param name="outputPath">输出文件路径</param>
    /// <returns>异步任务</returns>
    public async Task GenerateDocumentAsync(string templatePath, Dictionary<string, object> data, string outputPath)
    {
        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentException("模板文件路径不能为空", nameof(templatePath));

        if (data == null)
            throw new ArgumentNullException(nameof(data));

        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("输出文件路径不能为空", nameof(outputPath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"模板文件不存在: {templatePath}");

        try
        {
            // 确保输出目录存在
            var outputDirectory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // 使用 MiniWord 生成文档
            await Task.Run(() => { MiniWord.SaveAsByTemplate(outputPath, templatePath, data); });
        }
        catch (Exception ex) when (!(ex is ArgumentException || ex is ArgumentNullException ||
                                     ex is FileNotFoundException))
        {
            throw new InvalidOperationException($"生成文档时发生错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 异步批量生成 Word 文档
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="dataList">数据列表</param>
    /// <param name="outputDirectory">输出目录</param>
    /// <param name="progress">进度报告</param>
    /// <returns>成功生成的文档数量</returns>
    public async Task<int> BatchGenerateAsync(string templatePath, List<Dictionary<string, object>> dataList,
        string outputDirectory, IProgress<int> progress)
    {
        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentException("模板文件路径不能为空", nameof(templatePath));

        if (dataList == null)
            throw new ArgumentNullException(nameof(dataList));

        if (string.IsNullOrEmpty(outputDirectory))
            throw new ArgumentException("输出目录不能为空", nameof(outputDirectory));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"模板文件不存在: {templatePath}");

        // 确保输出目录存在
        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        int successCount = 0;
        int totalCount = dataList.Count;

        try
        {
            for (int i = 0; i < totalCount; i++)
            {
                var data = dataList[i];

                // 生成文件名，优先使用数据中的标识字段，否则使用索引
                string fileName = GenerateFileName(data, i + 1);
                string outputPath = Path.Combine(outputDirectory, fileName);

                try
                {
                    await GenerateDocumentAsync(templatePath, data, outputPath);
                    successCount++;
                }
                catch (Exception ex)
                {
                    // 记录单个文档生成失败，但继续处理其他文档
                    // 这里可以添加日志记录
                    Console.WriteLine($"生成文档 {fileName} 失败: {ex.Message}");
                }

                // 报告进度
                progress?.Report(i + 1);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"批量生成文档时发生错误: {ex.Message}", ex);
        }

        return successCount;
    }

    /// <summary>
    /// 生成输出文件名
    /// </summary>
    /// <param name="data">数据字典</param>
    /// <param name="index">索引</param>
    /// <returns>文件名</returns>
    private static string GenerateFileName(Dictionary<string, object> data, int index)
    {
        // 尝试使用常见的标识字段作为文件名
        var identifierFields = new[] { "姓名", "名称", "编号", "序号", "ID", "id", "Name", "name" };

        foreach (var field in identifierFields)
        {
            if (data.TryGetValue(field, out var value) && value != null)
            {
                var fileName = value.ToString()?.Trim();
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    // 移除文件名中的非法字符
                    fileName = RemoveInvalidFileNameChars(fileName);
                    return $"{fileName}.docx";
                }
            }
        }

        // 如果没有找到合适的标识字段，使用索引
        return $"文档_{index:D4}.docx";
    }

    /// <summary>
    /// 移除文件名中的非法字符
    /// </summary>
    /// <param name="fileName">原始文件名</param>
    /// <returns>清理后的文件名</returns>
    private static string RemoveInvalidFileNameChars(string fileName)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        foreach (var invalidChar in invalidChars)
        {
            fileName = fileName.Replace(invalidChar, '_');
        }

        return fileName;
    }

    /// <summary>
    /// 匹配 Excel 列标题与模板变量
    /// </summary>
    /// <param name="excelHeaders">Excel 列标题列表</param>
    /// <param name="templateVariables">模板变量列表</param>
    /// <returns>变量匹配结果</returns>
    public VariableMatchingModel MatchVariables(List<string> excelHeaders, List<string> templateVariables)
    {
        if (excelHeaders == null)
            throw new ArgumentNullException(nameof(excelHeaders));

        if (templateVariables == null)
            throw new ArgumentNullException(nameof(templateVariables));

        var result = new VariableMatchingModel();

        // 精确匹配
        var exactMatches = excelHeaders.Intersect(templateVariables, StringComparer.OrdinalIgnoreCase).ToList();
        result.MatchedVariables.AddRange(exactMatches);

        // 未匹配的 Excel 列
        result.UnmatchedExcelColumns = excelHeaders
            .Where(header => !exactMatches.Contains(header, StringComparer.OrdinalIgnoreCase))
            .ToList();

        // 未匹配的模板变量
        result.UnmatchedTemplateVariables = templateVariables
            .Where(variable => !exactMatches.Contains(variable, StringComparer.OrdinalIgnoreCase))
            .ToList();

        return result;
    }

    /// <summary>
    /// 高级变量匹配，包含模糊匹配和智能匹配
    /// </summary>
    /// <param name="excelHeaders">Excel 列标题列表</param>
    /// <param name="templateVariables">模板变量列表</param>
    /// <param name="enableFuzzyMatching">是否启用模糊匹配</param>
    /// <returns>变量匹配结果</returns>
    public VariableMatchingModel MatchVariablesAdvanced(List<string> excelHeaders, List<string> templateVariables,
        bool enableFuzzyMatching = true)
    {
        if (excelHeaders == null)
            throw new ArgumentNullException(nameof(excelHeaders));

        if (templateVariables == null)
            throw new ArgumentNullException(nameof(templateVariables));

        var result = new VariableMatchingModel();
        var matchedExcelHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var matchedTemplateVariables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 第一步：精确匹配
        foreach (var header in excelHeaders)
        {
            var exactMatch = templateVariables.FirstOrDefault(v =>
                string.Equals(v, header, StringComparison.OrdinalIgnoreCase));

            if (exactMatch != null)
            {
                result.MatchedVariables.Add(header);
                matchedExcelHeaders.Add(header);
                matchedTemplateVariables.Add(exactMatch);
            }
        }

        // 第二步：模糊匹配（如果启用）
        if (enableFuzzyMatching)
        {
            var unmatchedHeaders = excelHeaders.Where(h => !matchedExcelHeaders.Contains(h)).ToList();
            var unmatchedVariables = templateVariables.Where(v => !matchedTemplateVariables.Contains(v)).ToList();

            foreach (var header in unmatchedHeaders)
            {
                var fuzzyMatch = FindBestFuzzyMatch(header, unmatchedVariables);
                if (fuzzyMatch != null)
                {
                    result.MatchedVariables.Add(header);
                    matchedExcelHeaders.Add(header);
                    matchedTemplateVariables.Add(fuzzyMatch);
                }
            }
        }

        // 设置未匹配的项目
        result.UnmatchedExcelColumns = excelHeaders
            .Where(header => !matchedExcelHeaders.Contains(header))
            .ToList();

        result.UnmatchedTemplateVariables = templateVariables
            .Where(variable => !matchedTemplateVariables.Contains(variable))
            .ToList();

        return result;
    }

    /// <summary>
    /// 查找最佳模糊匹配
    /// </summary>
    /// <param name="target">目标字符串</param>
    /// <param name="candidates">候选字符串列表</param>
    /// <returns>最佳匹配的字符串，如果没有找到合适的匹配则返回 null</returns>
    private static string? FindBestFuzzyMatch(string target, List<string> candidates)
    {
        if (string.IsNullOrEmpty(target) || candidates.Count == 0)
            return null;

        const double minSimilarityThreshold = 0.6; // 最小相似度阈值
        string? bestMatch = null;
        double bestSimilarity = 0;

        foreach (var candidate in candidates)
        {
            if (string.IsNullOrEmpty(candidate))
                continue;

            // 计算相似度
            double similarity = CalculateStringSimilarity(target, candidate);

            if (similarity > bestSimilarity && similarity >= minSimilarityThreshold)
            {
                bestSimilarity = similarity;
                bestMatch = candidate;
            }
        }

        return bestMatch;
    }

    /// <summary>
    /// 计算两个字符串的相似度（使用 Levenshtein 距离）
    /// </summary>
    /// <param name="s1">字符串1</param>
    /// <param name="s2">字符串2</param>
    /// <returns>相似度（0-1之间，1表示完全相同）</returns>
    private static double CalculateStringSimilarity(string s1, string s2)
    {
        if (string.IsNullOrEmpty(s1) && string.IsNullOrEmpty(s2))
            return 1.0;

        if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2))
            return 0.0;

        // 转换为小写进行比较
        s1 = s1.ToLowerInvariant();
        s2 = s2.ToLowerInvariant();

        if (s1 == s2)
            return 1.0;

        // 计算 Levenshtein 距离
        int distance = CalculateLevenshteinDistance(s1, s2);
        int maxLength = Math.Max(s1.Length, s2.Length);

        return 1.0 - (double)distance / maxLength;
    }

    /// <summary>
    /// 计算 Levenshtein 距离
    /// </summary>
    /// <param name="s1">字符串1</param>
    /// <param name="s2">字符串2</param>
    /// <returns>Levenshtein 距离</returns>
    private static int CalculateLevenshteinDistance(string s1, string s2)
    {
        int[,] matrix = new int[s1.Length + 1, s2.Length + 1];

        // 初始化第一行和第一列
        for (int i = 0; i <= s1.Length; i++)
            matrix[i, 0] = i;

        for (int j = 0; j <= s2.Length; j++)
            matrix[0, j] = j;

        // 填充矩阵
        for (int i = 1; i <= s1.Length; i++)
        {
            for (int j = 1; j <= s2.Length; j++)
            {
                int cost = s1[i - 1] == s2[j - 1] ? 0 : 1;

                matrix[i, j] = Math.Min(
                    Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                    matrix[i - 1, j - 1] + cost
                );
            }
        }

        return matrix[s1.Length, s2.Length];
    }
}