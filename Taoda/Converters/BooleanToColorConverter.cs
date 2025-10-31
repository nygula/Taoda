using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace Taoda.Converters;

/// <summary>
/// 布尔值到颜色转换器
/// </summary>
public class BooleanToColorConverter : IValueConverter
{
    public static readonly BooleanToColorConverter Instance = new();

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool boolValue)
        {
            return boolValue ? Brushes.LightGreen : Brushes.LightCoral;
        }
        return Brushes.Gray;
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}