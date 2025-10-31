using System;
using Avalonia;

namespace Taoda.Desktop;

class Program
{
    // Initialization code. Don't use any Avalonia, third-party APIs or any
    // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
    // yet and stuff might break.
    [STAThread]
    public static void Main(string[] args)
    {
        try
        {
            // 构建并启动 Avalonia 应用程序
            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            // 记录启动异常并提供用户友好的错误信息
            Console.WriteLine($"应用程序启动失败: {ex.Message}");
            Console.WriteLine($"详细错误信息: {ex}");
            
            // 在调试模式下暂停以便查看错误信息
            #if DEBUG
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
            #endif
            
            // 确保应用程序正确退出
            Environment.Exit(1);
        }
    }

    /// <summary>
    /// Avalonia 应用程序配置
    /// 注意：不要移除此方法，它也被可视化设计器使用
    /// </summary>
    /// <returns>配置好的 AppBuilder</returns>
    public static AppBuilder BuildAvaloniaApp()
    {
        return AppBuilder.Configure<App>()
            .UsePlatformDetect()           // 自动检测并使用适当的平台后端
            .WithInterFont()               // 使用 Inter 字体
            .LogToTrace();                 // 启用跟踪日志记录
    }
}
