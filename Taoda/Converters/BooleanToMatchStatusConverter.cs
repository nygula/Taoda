using System;
using System.Globalization;
using Avalonia.Data.Converters;

namespace Taoda.Converters;

/// <summary>
/// 布尔值到匹配状态文本转换器
/// </summary>
public class BooleanToMatchStatusConverter : IValueConverter
{
    public static readonly BooleanToMatchStatusConverter Instance = new();

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool boolValue)
        {
            return boolValue ? "✓ 所有变量完全匹配" : "⚠ 存在未匹配的变量";
        }
        return "未知状态";
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}