using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Data.Core.Plugins;
using Avalonia.Markup.Xaml;
using Microsoft.Extensions.DependencyInjection;
using System;

using Taoda.ViewModels;
using Taoda.Views;
using Taoda.Services;

namespace Taoda;

public partial class App : Application
{
    private ServiceProvider? _serviceProvider;

    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
        ConfigureServices();
    }

    /// <summary>
    /// 配置依赖注入服务容器
    /// </summary>
    private void ConfigureServices()
    {
        var services = new ServiceCollection();
        
        // 注册服务接口和实现类
        services.AddSingleton<IExcelService, ExcelService>();
        services.AddSingleton<IWordService, WordService>();
        services.AddSingleton<IFileDialogService, FileDialogService>();
        
        // 注册 ViewModel - 使用 Transient 生命周期以支持多实例
        services.AddTransient<MainViewModel>();
        
        // 构建服务提供者
        _serviceProvider = services.BuildServiceProvider();
    }

    /// <summary>
    /// 获取服务实例
    /// </summary>
    /// <typeparam name="T">服务类型</typeparam>
    /// <returns>服务实例</returns>
    public T? GetService<T>() where T : class
    {
        return _serviceProvider?.GetService<T>();
    }

    /// <summary>
    /// 获取必需的服务实例
    /// </summary>
    /// <typeparam name="T">服务类型</typeparam>
    /// <returns>服务实例</returns>
    public T GetRequiredService<T>() where T : class
    {
        return _serviceProvider?.GetRequiredService<T>() ?? throw new InvalidOperationException($"Service of type {typeof(T).Name} is not registered.");
    }

    public override void OnFrameworkInitializationCompleted()
    {
        // Line below is needed to remove Avalonia data validation.
        // Without this line you will get duplicate validations from both Avalonia and CT
        BindingPlugins.DataValidators.RemoveAt(0);

        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            // 配置应用程序退出时的资源清理
            desktop.Exit += OnApplicationExit;
            
            desktop.MainWindow = new MainWindow
            {
                DataContext = GetRequiredService<MainViewModel>()
            };
        }
        else if (ApplicationLifetime is ISingleViewApplicationLifetime singleViewPlatform)
        {
            singleViewPlatform.MainView = new MainView
            {
                DataContext = GetRequiredService<MainViewModel>()
            };
        }

        base.OnFrameworkInitializationCompleted();
    }

    /// <summary>
    /// 应用程序退出时的清理处理
    /// </summary>
    private void OnApplicationExit(object? sender, ControlledApplicationLifetimeExitEventArgs e)
    {
        // 释放服务提供者资源
        _serviceProvider?.Dispose();
        _serviceProvider = null;
    }
}
