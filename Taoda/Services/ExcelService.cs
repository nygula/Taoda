using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using MiniExcelLibs;
using Taoda.Models;

namespace Taoda.Services;

/// <summary>
/// Excel 文件处理服务实现
/// </summary>
public class ExcelService : IExcelService
{
    /// <summary>
    /// 异步读取 Excel 文件并返回数据模型
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <returns>Excel 数据模型</returns>
    /// <exception cref="ArgumentException">文件路径为空或无效</exception>
    /// <exception cref="FileNotFoundException">文件不存在</exception>
    /// <exception cref="InvalidDataException">文件格式无效或数据损坏</exception>
    public async Task<ExcelDataModel> ReadExcelAsync(string filePath)
    {
        // 参数验证
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("文件路径不能为空", nameof(filePath));

        // 文件存在性验证
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Excel 文件不存在: {filePath}");

        // 文件格式验证
        ValidateFileFormat(filePath);

        try
        {
            var result = new ExcelDataModel
            {
                FilePath = filePath
            };

            // 使用 MiniExcel 读取数据
            var rows = await Task.Run(() => MiniExcel.Query(filePath, useHeaderRow: true).ToList());
            
            if (rows.Count == 0)
            {
                throw new InvalidDataException("Excel 文件为空或没有数据行");
            }

            // 获取第一行数据来确定列标题
            var firstRow = rows.First() as IDictionary<string, object>;
            if (firstRow == null)
            {
                throw new InvalidDataException("无法读取 Excel 文件的列标题");
            }

            // 提取列标题
            result.Headers = firstRow.Keys.ToList();

            // 转换所有数据行为字典格式
            foreach (var row in rows)
            {
                if (row is IDictionary<string, object> rowDict)
                {
                    var dataRow = new Dictionary<string, object>();
                    foreach (var header in result.Headers)
                    {
                        // 确保每个列都有值，如果没有则设为空字符串
                        dataRow[header] = rowDict.TryGetValue(header, out var value) ? value : string.Empty;
                    }
                    result.Rows.Add(dataRow);
                }
            }

            return result;
        }
        catch (Exception ex) when (!(ex is ArgumentException || ex is FileNotFoundException || ex is InvalidDataException))
        {
            throw new InvalidDataException($"读取 Excel 文件时发生错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 获取 Excel 文件的列标题
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <returns>列标题列表</returns>
    /// <exception cref="ArgumentException">文件路径为空或无效</exception>
    /// <exception cref="FileNotFoundException">文件不存在</exception>
    /// <exception cref="InvalidDataException">文件格式无效或数据损坏</exception>
    public List<string> GetColumnHeaders(string filePath)
    {
        // 参数验证
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("文件路径不能为空", nameof(filePath));

        // 文件存在性验证
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Excel 文件不存在: {filePath}");

        // 文件格式验证
        ValidateFileFormat(filePath);

        try
        {
            // 只读取第一行来获取列标题
            var firstRow = MiniExcel.Query(filePath, useHeaderRow: true).FirstOrDefault();
            
            if (firstRow == null)
            {
                throw new InvalidDataException("Excel 文件为空或没有数据行");
            }

            if (firstRow is IDictionary<string, object> rowDict)
            {
                return rowDict.Keys.ToList();
            }

            throw new InvalidDataException("无法读取 Excel 文件的列标题");
        }
        catch (Exception ex) when (!(ex is ArgumentException || ex is FileNotFoundException || ex is InvalidDataException))
        {
            throw new InvalidDataException($"读取 Excel 文件列标题时发生错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 获取 Excel 文件的数据行数量
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <returns>数据行数量</returns>
    /// <exception cref="ArgumentException">文件路径为空或无效</exception>
    /// <exception cref="FileNotFoundException">文件不存在</exception>
    /// <exception cref="InvalidDataException">文件格式无效或数据损坏</exception>
    public int GetDataRowCount(string filePath)
    {
        // 参数验证
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("文件路径不能为空", nameof(filePath));

        // 文件存在性验证
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Excel 文件不存在: {filePath}");

        // 文件格式验证
        ValidateFileFormat(filePath);

        try
        {
            // 计算数据行数量
            return MiniExcel.Query(filePath, useHeaderRow: true).Count();
        }
        catch (Exception ex) when (!(ex is ArgumentException || ex is FileNotFoundException || ex is InvalidDataException))
        {
            throw new InvalidDataException($"读取 Excel 文件行数时发生错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 验证文件格式是否为支持的 Excel 格式
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <exception cref="InvalidDataException">文件格式不支持</exception>
    private static void ValidateFileFormat(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension != ".xlsx" && extension != ".xls")
        {
            throw new InvalidDataException($"不支持的文件格式: {extension}。仅支持 .xlsx 和 .xls 格式");
        }
    }
}