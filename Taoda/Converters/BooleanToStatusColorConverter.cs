using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace Taoda.Converters;

/// <summary>
/// 布尔值到状态颜色转换器
/// </summary>
public class BooleanToStatusColorConverter : IValueConverter
{
    public static readonly BooleanToStatusColorConverter Instance = new();

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool boolValue)
        {
            return boolValue ? Brushes.Orange : Brushes.Green;
        }
        return Brushes.Gray;
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}