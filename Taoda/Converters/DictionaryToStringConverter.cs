using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Avalonia.Data.Converters;

namespace Taoda.Converters;

/// <summary>
/// Dictionary 到字符串转换器，用于在 UI 中显示字典内容
/// </summary>
public class DictionaryToStringConverter : IValueConverter
{
    public static readonly DictionaryToStringConverter Instance = new();

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is Dictionary<string, object> dictionary)
        {
            var sb = new StringBuilder();
            var items = dictionary.Take(10); // 只显示前10个项目，避免过长
            
            foreach (var kvp in items)
            {
                var displayValue = kvp.Value?.ToString() ?? "null";
                // 限制每个值的显示长度
                if (displayValue.Length > 50)
                {
                    displayValue = displayValue.Substring(0, 47) + "...";
                }
                
                sb.AppendLine($"{kvp.Key}: {displayValue}");
            }
            
            if (dictionary.Count > 10)
            {
                sb.AppendLine($"... 还有 {dictionary.Count - 10} 个字段");
            }
            
            return sb.ToString().TrimEnd();
        }
        
        return value?.ToString() ?? "null";
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}